/// <reference types="node" />
// Editing sections and headers/footers: header/footer parts, references and inheritance, page
// setup, section breaks and the even/odd setting. Outputs are written to out/edit-sections/ so
// they can be opened in Word or converted with `node scripts/libreoffice-check.mjs`.
import { mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';
import {
  CT,
  createDocx,
  type Paragraph,
  REL,
  readDocx,
  UsageError,
  validatePackage,
  type WordDocument,
  writeDocx,
} from '../src/index.js';
import {
  copyHeaderFooter,
  createHeaderFooter,
  type HeaderFooter,
  headerFooterParts,
} from '../src/wml/header-footer.js';
import {
  evenAndOddHeaders,
  insertSectionBreak,
  removeSectionBreak,
  SECTPR_ORDER,
  type Section,
  sectionsOf,
  setEvenAndOddHeaders,
} from '../src/wml/section.js';
import { fixture, strictFixtures } from './helpers/fixtures.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W_STRICT = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
const OUT = fileURLToPath(new URL('../out/edit-sections/', import.meta.url));
const F3 = 'docxcompose/docproperties_header_footer_3_sections.docx';
const FMM = 'docx-mailmerge/test_footnote_header_footer.docx';
/** A 1x1 PNG. */
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

beforeAll(() => {
  mkdirSync(OUT, { recursive: true });
});

/** Errors reported by validatePackage, as `code part` strings. */
function errors(doc: WordDocument): string[] {
  return validatePackage(doc)
    .filter((i) => i.severity === 'error')
    .map((i) => `${i.code} ${i.part ?? ''}`);
}

/** Saves (validating), writes out/edit-sections/<name> and reads the result back. */
function roundTrip(doc: WordDocument, name: string): WordDocument {
  expect(errors(doc), name).toEqual([]);
  const bytes = writeDocx(doc, { validate: true });
  writeFileSync(`${OUT}${name}`, bytes);
  const back = readDocx(bytes);
  expect(errors(back), name).toEqual([]);
  return back;
}

/** Sets the text of a header/footer to one paragraph. */
function setText(hf: HeaderFooter, text: string): void {
  hf.clear();
  hf.story.paragraphs()[0]!.appendRun(text);
}

/** Appends a paragraph whose only content is a page break. */
function pageBreak(doc: WordDocument): void {
  const p = doc.body.appendParagraph();
  const xml = doc.mainPart.xml();
  const r = xml.createElement(W, 'r');
  r.append(xml.createElement(W, 'br', { 'w:type': 'page' }));
  p.element.append(r);
}

/** A blank document with `pages` pages ("Page 1", page break, "Page 2", ...). */
function pagedDoc(pages: number): WordDocument {
  const doc = createDocx();
  for (let i = 1; i <= pages; i++) {
    if (i > 1) pageBreak(doc);
    doc.body.appendParagraph(`Page ${i}`);
  }
  return doc;
}

function bodyParagraph(doc: WordDocument, text: string): Paragraph {
  const p = doc.body.paragraphs().find((x) => x.text === text);
  if (p === undefined) throw new Error(`no paragraph ${text}`);
  return p;
}

const texts = (s: Section): string[] => s.paragraphs().map((p) => p.text);

/** Child local names of an element. */
const childNames = (el: { elements(): { local: string }[] } | undefined): string[] =>
  el === undefined ? [] : el.elements().map((e) => e.local);

/** True when the sectPr's known children are in schema order. */
function inSchemaOrder(sectPr: { elements(): { uri: string; local: string }[] }): boolean {
  let last = -1;
  for (const k of sectPr.elements()) {
    const r = SECTPR_ORDER.get(`${k.uri} ${k.local}`);
    if (r === undefined) continue;
    if (r < last) return false;
    last = r;
  }
  return true;
}

// ---- synthetic packages ----------------------------------------------------------------------

const enc = new TextEncoder();
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const CT_BASE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** A minimal package: document.xml with `body`, no styles or settings, optional parts. */
function minimalDocx(body: string, rels = '', parts: Record<string, string> = {}): Uint8Array {
  let overrides = `<Override PartName="/word/document.xml" ContentType="${CT_BASE}.document.main+xml"/>`;
  const files: Record<string, Uint8Array> = {};
  for (const [name, xml] of Object.entries(parts)) {
    const kind = name.includes('header') ? 'header' : 'footer';
    overrides += `<Override PartName="/word/${name}" ContentType="${CT_BASE}.${kind}+xml"/>`;
    files[`word/${name}`] = enc.encode(xml);
  }
  files['[Content_Types].xml'] = enc.encode(
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      `<Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
  );
  files['_rels/.rels'] = enc.encode(
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  files['word/_rels/document.xml.rels'] = enc.encode(
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
  );
  files['word/document.xml'] = enc.encode(
    `${XML_DECL}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`,
  );
  return zipSync(files);
}

// ---- header/footer parts ---------------------------------------------------------------------

describe('createHeaderFooter', () => {
  it('adds a part with content type, relationship and a Word root', () => {
    const doc = createDocx();
    const h = createHeaderFooter(doc, 'header');
    expect(h.kind).toBe('header');
    expect(h.part.name).toBe('word/header1.xml');
    expect(h.part.contentType).toBe(CT.header);
    const rel = doc.mainPart.rels.all().find((r) => r.partName === 'word/header1.xml');
    expect(rel?.type).toBe(REL.header);
    expect(rel?.target).toBe('header1.xml');
    const root = h.story.element;
    expect(root.name).toBe('w:hdr');
    for (const p of ['w', 'r', 'wp', 'mc', 'w14', 'wps', 'v', 'o']) {
      expect(root.lookupNamespace(p), p).toBeDefined();
    }
    expect(root.attr('mc:Ignorable')).toBe('w14 w15 wp14');
    expect(h.story.paragraphs()).toHaveLength(1);
    // createDocx's styles have no Header style.
    expect(h.story.paragraphs()[0]!.style).toBeUndefined();
    expect(h.text()).toBe('');

    const f = createHeaderFooter(doc, 'footer');
    expect(f.part.name).toBe('word/footer1.xml');
    expect(f.part.contentType).toBe(CT.footer);
    expect(f.story.element.name).toBe('w:ftr');
    expect(createHeaderFooter(doc, 'header').part.name).toBe('word/header2.xml');
    // Not referenced by any section yet.
    expect(doc.sections()[0]!.header()).toBeUndefined();
    expect(headerFooterParts(doc)).toHaveLength(3);
    expect(errors(doc)).toEqual([]);
  });

  it('uses the Header/Footer paragraph styles when the document defines them', () => {
    const doc = readDocx(fixture(F3).bytes);
    expect(createHeaderFooter(doc, 'header').story.paragraphs()[0]!.style).toBe('Header');
    const f = createHeaderFooter(doc, 'footer');
    expect(f.story.paragraphs()[0]!.style).toBe('Footer');
    expect(f.part.name).toBe('word/footer5.xml');
  });

  it('follows a Strict document: Strict namespaces and relationship types', () => {
    const [strict] = strictFixtures([F3]);
    const doc = readDocx(strict!.bytes);
    const s = doc.sections()[2]!;
    const h = s.setHeader('default', 'new');
    setText(h, 'Strict header');
    expect(h.story.element.lookupNamespace('w')).toBe(W_STRICT);
    const rel = doc.mainPart.rels.all().find((r) => r.partName === h.part.name);
    expect(rel?.type).toBe('http://purl.oclc.org/ooxml/officeDocument/relationships/header');
    const back = roundTrip(doc, 'strict-new-header.docx');
    expect(back.sections()[2]!.header()?.text()).toBe('Strict header');
    expect(back.sections()[1]!.header()?.text()).toBe('Date: 11.06.2019');
  });

  it('clear() leaves one empty paragraph; the story is editable', () => {
    const doc = readDocx(fixture(F3).bytes);
    const h = doc.sections()[0]!.header()!;
    expect(h.text()).toBe('Text: Foo Bar');
    h.clear();
    expect(h.text()).toBe('');
    expect(h.story.paragraphs().map((p) => p.style)).toEqual(['Header']);
    h.story.appendParagraph('Second', { style: 'Header' });
    expect(h.text()).toBe('\nSecond');
    expect(h.part.dirty).toBe(true);
  });

  it('copyHeaderFooter duplicates content and relationships under a new part', () => {
    const doc = createDocx();
    const h = doc.sections()[0]!.setHeader('default', 'new');
    setText(h, 'Logo:');
    h.appendImage(PNG, { width: 24 });
    const copy = copyHeaderFooter(h);
    expect(copy.part.name).toBe('word/header2.xml');
    // The image went into a second paragraph: the first one already had a run.
    expect(copy.text()).toBe('Logo:\n');
    expect(copy.part.rels.all().map((r) => [r.id, r.partName])).toEqual(
      h.part.rels.all().map((r) => [r.id, r.partName]),
    );
    const ids = (hf: HeaderFooter) =>
      hf.story.element
        .descendants(
          'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
          'docPr',
        )
        .map((d) => d.attr('id'));
    expect(ids(copy)).toHaveLength(1);
    expect(ids(copy)[0]).not.toBe(ids(h)[0]);
    expect(errors(doc)).toEqual([]);
  });
});

// ---- references ----------------------------------------------------------------------------

describe('setHeader / setFooter', () => {
  it("'new' creates and references a header and footer", () => {
    const doc = pagedDoc(2);
    const s = doc.sections()[0]!;
    const h = s.setHeader('default', 'new');
    setText(h, 'My header');
    setText(s.setFooter('default', 'new'), 'My footer');
    expect(s.ownHeader()).toBe(h);
    expect(s.header()).toBe(h);
    const sectPr = s.element!;
    expect(childNames(sectPr).slice(0, 3)).toEqual(['headerReference', 'footerReference', 'pgSz']);
    expect(sectPr.elements(W, 'headerReference')[0]!.attr('w:type')).toBe('default');
    const back = roundTrip(doc, 'new-header-footer.docx');
    expect(back.sections()[0]!.header()?.text()).toBe('My header');
    expect(back.sections()[0]!.footer()?.text()).toBe('My footer');
  });

  it('replaces a reference in place and removes the part nobody references any more', () => {
    const doc = readDocx(fixture(F3).bytes);
    const [s0, s1, s2] = doc.sections();
    const old = s1!.ownHeader()!;
    expect(old.part.name).toBe('word/header4.xml');
    const shared = s0!.ownHeader()!;
    s1!.setHeader('default', shared);
    expect(s1!.ownHeader()).toBe(shared);
    expect(s2!.header()).toBe(shared);
    expect(doc.package.part('word/header4.xml')).toBeUndefined();
    expect(doc.mainPart.rels.all().some((r) => r.partName === 'word/header4.xml')).toBe(false);
    // The shared part keeps its single relationship.
    const rels = doc.mainPart.rels.all().filter((r) => r.partName === shared.part.name);
    expect(rels).toHaveLength(1);
    // Removing one of two references keeps the part.
    s1!.setHeader('default', null);
    expect(doc.package.part(shared.part.name)).toBe(shared.part);
    expect(() => s1!.setHeader('default', old)).toThrow(UsageError);
    roundTrip(doc, 'replace-header.docx');
  });

  it('null removes the own reference so the section inherits again', () => {
    const doc = readDocx(fixture(F3).bytes);
    const [s0, s1, s2] = doc.sections();
    expect(s1!.header()?.text()).toBe('Date: 11.06.2019');
    expect(s1!.setHeader('default', null)).toBeUndefined();
    expect(s1!.ownHeader()).toBeUndefined();
    expect(s1!.header()).toBe(s0!.ownHeader());
    expect(s2!.header()).toBe(s0!.ownHeader());
    s1!.linkToPrevious('footer');
    expect(s1!.ownFooter()).toBeUndefined();
    expect(s1!.footer()).toBe(s0!.ownFooter());
    expect(headerFooterParts(doc)).toHaveLength(6);
    const back = roundTrip(doc, 'link-to-previous.docx');
    const [b0, b1, b2] = back.sections();
    expect(b1!.header()?.text()).toBe('Text: Foo Bar');
    expect(b2!.footer()?.text()).toBe('Number: 123');
    expect(b1!.header()).toBe(b0!.ownHeader());
  });

  it('rejects headers of another document or kind, and unknown types', () => {
    const a = createDocx();
    const b = createDocx();
    const h = createHeaderFooter(b, 'header');
    const s = a.sections()[0]!;
    expect(() => s.setHeader('default', h)).toThrow(UsageError);
    expect(() => s.setFooter('default', createHeaderFooter(a, 'header'))).toThrow(UsageError);
    expect(() => s.setHeader('odd' as 'even', 'new')).toThrow(UsageError);
    expect(s.element!.elements(W, 'headerReference')).toHaveLength(0);
  });
});

describe('unlinkFromPrevious', () => {
  it('gives one section its own copy; editing it leaves the others alone', () => {
    const doc = readDocx(fixture(F3).bytes);
    const [s0, s1, s2] = doc.sections();
    expect(s2!.ownHeader()).toBeUndefined();
    const copy = s2!.unlinkFromPrevious('header', 'default');
    expect(copy).not.toBe(s1!.ownHeader());
    expect(copy.text()).toBe('Date: 11.06.2019');
    expect(s2!.ownHeader()).toBe(copy);
    setText(copy, 'Only section 3');
    expect(s0!.header()?.text()).toBe('Text: Foo Bar');
    expect(s1!.header()?.text()).toBe('Date: 11.06.2019');
    expect(s2!.header()?.text()).toBe('Only section 3');
    // first/even still inherited from section 1.
    expect(s2!.header('first')).toBe(s0!.ownHeader('first'));
    const back = roundTrip(doc, 'unlink-one.docx');
    expect(back.sections().map((s) => s.header()?.text())).toEqual([
      'Text: Foo Bar',
      'Date: 11.06.2019',
      'Only section 3',
    ]);
  });

  it('follows Word: later sections linked to the unlinked one follow its copy', () => {
    const doc = readDocx(fixture(F3).bytes);
    const [s0, s1, s2] = doc.sections();
    const own = s1!.ownHeader();
    // Own and not shared: kept as is.
    expect(s1!.unlinkFromPrevious('header', 'default')).toBe(own);
    const first = s1!.unlinkFromPrevious('header', 'first');
    expect(first).not.toBe(s0!.ownHeader('first'));
    expect(first.text()).toBe(s0!.ownHeader('first')!.text());
    setText(first, 'First page of 2 and 3');
    expect(s0!.header('first')?.text()).not.toBe('First page of 2 and 3');
    expect(s2!.header('first')).toBe(first);
  });

  it('without a type unlinks every type the section shows', () => {
    const doc = readDocx(fixture(F3).bytes);
    const [s0, s1, s2] = doc.sections();
    const before = (['default', 'first', 'even'] as const).map((t) => s2!.footer(t)?.text());
    const result = s2!.unlinkFromPrevious('footer');
    expect(result).toBe(s2!.ownFooter('default'));
    for (const t of ['default', 'first', 'even'] as const) {
      expect(s2!.ownFooter(t), t).toBeDefined();
      expect(s2!.ownFooter(t)).not.toBe(s1!.footer(t));
    }
    expect((['default', 'first', 'even'] as const).map((t) => s2!.footer(t)?.text())).toEqual(
      before,
    );
    expect(s0!.ownFooter('default')?.text()).toBe('Number: 123');
    expect(headerFooterParts(doc)).toHaveLength(11);
    roundTrip(doc, 'unlink-all-footers.docx');
  });

  it('copies a part referenced by another section even when the section references it itself', () => {
    const doc = createDocx();
    doc.body.appendParagraph('one');
    insertSectionBreak(bodyParagraph(doc, 'one'));
    const [s0, s1] = doc.sections();
    const h = s0!.setHeader('default', 'new');
    s1!.setHeader('default', h);
    const copy = s1!.unlinkFromPrevious('header', 'default');
    expect(copy).not.toBe(h);
    expect(s0!.ownHeader()).toBe(h);
    // Nothing to inherit: a new empty one.
    const f = s0!.unlinkFromPrevious('footer', 'default');
    expect(f.text()).toBe('');
    expect(s0!.ownFooter()).toBe(f);
    expect(s0!.unlinkFromPrevious('footer', 'first')).toBeDefined();
    expect(errors(doc)).toEqual([]);
  });

  it('copies header relationships: the copy shows the same image under fresh drawing ids', () => {
    const doc = pagedDoc(1);
    doc.body.appendParagraph('Section 2');
    insertSectionBreak(bodyParagraph(doc, 'Page 1'));
    const [s0, s1] = doc.sections();
    const h = s0!.setHeader('default', 'new');
    h.appendImage(PNG, { width: 36, height: 36 });
    h.story.paragraphs()[0]!.appendRun(' Company');
    expect(s1!.header()).toBe(h);
    const copy = s1!.unlinkFromPrevious('header', 'default');
    setText(copy, 'Section two: ');
    copy.story.paragraphs()[0]!.appendImage(PNG, { width: 18 });
    const relsOf = (hf: HeaderFooter) => hf.part.rels.all().map((r) => r.partName);
    expect(relsOf(copy)).toContain('word/media/image1.png');
    expect(doc.package.parts().filter((p) => p.name.startsWith('word/media/'))).toHaveLength(1);
    const back = roundTrip(doc, 'header-image-unlinked.docx');
    const [b0, b1] = back.sections();
    expect(b0!.header()?.text()).toBe(' Company');
    expect(b1!.header()?.text()).toBe('Section two: ');
    expect(
      b1!
        .header()
        ?.part.rels.all()
        .map((r) => r.partName),
    ).toContain('word/media/image1.png');
  });
});

describe('HeaderFooter.appendImage', () => {
  it('puts the image in the empty paragraph of a new header, then in new paragraphs', () => {
    const doc = createDocx();
    const h = doc.sections()[0]!.setHeader('default', 'new');
    const run = h.appendImage(PNG, { width: 72, description: 'logo' });
    expect(h.story.paragraphs()).toHaveLength(1);
    expect(run.paragraph.element).toBe(h.story.paragraphs()[0]!.element);
    expect(run.element.descendants(W, 'drawing')).toHaveLength(1);
    h.appendImage(PNG);
    expect(h.story.paragraphs()).toHaveLength(2);
    const rel = h.part.rels.all()[0]!;
    expect(rel.partName).toBe('word/media/image1.png');
    roundTrip(doc, 'header-image.docx');
  });
});

// ---- first page and even/odd -------------------------------------------------------------

describe('first page and even/odd headers', () => {
  it('adds a first-page header shown with titlePage', () => {
    const doc = pagedDoc(3);
    const s = doc.sections()[0]!;
    setText(s.setHeader('default', 'new'), 'Default header');
    setText(s.setHeader('first', 'new'), 'First page header');
    setText(s.setFooter('first', 'new'), 'First page footer');
    expect(s.titlePage).toBe(false);
    s.titlePage = true;
    expect(s.titlePage).toBe(true);
    expect(inSchemaOrder(s.element!)).toBe(true);
    const back = roundTrip(doc, 'first-page-header.docx');
    const b = back.sections()[0]!;
    expect(b.titlePage).toBe(true);
    expect(b.header('first')?.text()).toBe('First page header');
    expect(b.header()?.text()).toBe('Default header');
    expect(b.footer('first')?.text()).toBe('First page footer');
    expect(b.footer()).toBeUndefined();
  });

  it('turns even/odd headers on and off in the settings part, in schema order', () => {
    const doc = pagedDoc(4);
    const s = doc.sections()[0]!;
    setText(s.setHeader('default', 'new'), 'Odd page header');
    setText(s.setHeader('even', 'new'), 'Even page header');
    expect(evenAndOddHeaders(doc)).toBe(false);
    setEvenAndOddHeaders(doc, true);
    expect(evenAndOddHeaders(doc)).toBe(true);
    const settings = doc.mainPart.relatedByType('settings')[0]!.xml().root;
    expect(childNames(settings).slice(0, 4)).toEqual([
      'zoom',
      'defaultTabStop',
      'evenAndOddHeaders',
      'characterSpacingControl',
    ]);
    const back = roundTrip(doc, 'even-odd-headers.docx');
    expect(evenAndOddHeaders(back)).toBe(true);
    expect(back.sections()[0]!.header('even')?.text()).toBe('Even page header');
    setEvenAndOddHeaders(back, false);
    expect(evenAndOddHeaders(back)).toBe(false);
    expect(
      back.mainPart.relatedByType('settings')[0]!.xml().root.element(W, 'evenAndOddHeaders'),
    ).toBeUndefined();
  });

  it('reads an existing on/off evenAndOddHeaders and fixes w:val="0" when turning it on', () => {
    expect(evenAndOddHeaders(readDocx(fixture(FMM).bytes))).toBe(true);
    const doc = readDocx(fixture(FMM).bytes);
    const settings = doc.mainPart.relatedByType('settings')[0]!.xml().root;
    settings.element(W, 'evenAndOddHeaders')!.setAttr('w:val', '0');
    expect(evenAndOddHeaders(doc)).toBe(false);
    setEvenAndOddHeaders(doc, true);
    expect(settings.element(W, 'evenAndOddHeaders')!.attr('w:val')).toBeUndefined();
    expect(evenAndOddHeaders(doc)).toBe(true);
  });

  it('creates the settings part when the document has none', () => {
    const doc = readDocx(minimalDocx('<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/>'));
    expect(doc.mainPart.relatedByType('settings')).toEqual([]);
    expect(evenAndOddHeaders(doc)).toBe(false);
    setEvenAndOddHeaders(doc, false); // no-op: nothing created
    expect(doc.package.part('word/settings.xml')).toBeUndefined();
    setEvenAndOddHeaders(doc, true);
    const part = doc.package.part('word/settings.xml')!;
    expect(part.contentType).toBe(CT.settings);
    expect(doc.mainPart.relatedByType('settings')[0]).toBe(part);
    expect(part.xml().root.name).toBe('w:settings');
    const back = roundTrip(doc, 'settings-created.docx');
    expect(evenAndOddHeaders(back)).toBe(true);
  });
});

// ---- page setup ------------------------------------------------------------------------------

describe('page setup', () => {
  it('round-trips every setter, keeping sectPr children in schema order', () => {
    const doc = pagedDoc(2);
    const s = doc.sections()[0]!;
    s.setPageSize(11906, 16838);
    expect(s.pageSize).toEqual({ width: 11906, height: 16838, orientation: 'portrait' });
    s.orientation = 'landscape';
    expect(s.pageSize).toEqual({ width: 16838, height: 11906, orientation: 'landscape' });
    expect(s.orientation).toBe('landscape');
    s.setMargins({ top: 720, left: 1080 });
    s.setColumns(2, 360);
    s.type = 'continuous';
    s.titlePage = true;
    s.pageNumberStart = 5;
    s.pageNumberFormat = 'lowerRoman';
    setText(s.setFooter('default', 'new'), 'Footer');
    expect(inSchemaOrder(s.element!)).toBe(true);
    expect(childNames(s.element!)).toEqual([
      'footerReference',
      'type',
      'pgSz',
      'pgMar',
      'pgNumType',
      'cols',
      'titlePg',
      'docGrid',
    ]);
    const back = roundTrip(doc, 'page-setup.docx');
    const b = back.sections()[0]!;
    expect(b.pageSize).toEqual({ width: 16838, height: 11906, orientation: 'landscape' });
    expect(b.margins).toEqual({
      top: 720,
      right: 1440,
      bottom: 1440,
      left: 1080,
      header: 720,
      footer: 720,
      gutter: 0,
    });
    expect(b.columns).toBe(2);
    expect(b.columnSpace).toBe(360);
    expect(b.type).toBe('continuous');
    expect(b.titlePage).toBe(true);
    expect(b.pageNumberStart).toBe(5);
    expect(b.pageNumberFormat).toBe('lowerRoman');

    b.orientation = 'portrait';
    expect(b.pageSize).toEqual({ width: 11906, height: 16838, orientation: 'portrait' });
    expect(b.element!.element(W, 'pgSz')!.attr('w:orient')).toBeUndefined();
    b.setColumns(1);
    expect(b.columns).toBe(1);
    b.type = 'nextPage';
    b.titlePage = false;
    b.pageNumberStart = undefined;
    b.pageNumberFormat = undefined;
    expect(childNames(b.element!)).toEqual(['footerReference', 'pgSz', 'pgMar', 'cols', 'docGrid']);
    expect(b.type).toBe('nextPage');
    roundTrip(back, 'page-setup-reset.docx');
  });

  it('setPageSize: orientation swaps the dimensions, paper codes go when the size changes', () => {
    const doc = readDocx(
      minimalDocx('<w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/></w:sectPr>'),
    );
    const s = doc.sections()[0]!;
    s.setPageSize(11906, 16838, 'landscape');
    expect(s.pageSize).toEqual({ width: 16838, height: 11906, orientation: 'landscape' });
    expect(s.element!.element(W, 'pgSz')!.attr('w:code')).toBe('9');
    s.setPageSize(12240, 15840);
    expect(s.pageSize?.orientation).toBe('portrait');
    expect(s.element!.element(W, 'pgSz')!.attr('w:code')).toBeUndefined();
    s.setPageSize(15840, 12240);
    expect(s.pageSize?.orientation).toBe('landscape');
    expect(() => s.setPageSize(0, 100)).toThrow(UsageError);
    expect(() => s.setPageSize(Number.NaN, 100)).toThrow(UsageError);
  });

  it('a new pgMar gets every required attribute; unequal columns become equal', () => {
    const doc = readDocx(
      minimalDocx(
        '<w:p/><w:sectPr><w:cols w:equalWidth="0" w:space="100"><w:col w:w="2000"/><w:col w:w="3000"/></w:cols></w:sectPr>',
      ),
    );
    const s = doc.sections()[0]!;
    s.setMargins({ bottom: -720 });
    expect(s.margins).toEqual({
      top: 1440,
      right: 1440,
      bottom: -720,
      left: 1440,
      header: 720,
      footer: 720,
      gutter: 0,
    });
    expect(() => s.setMargins({ left: -1 })).toThrow(UsageError);
    expect(s.columns).toBe(2);
    s.setColumns(3);
    const cols = s.element!.element(W, 'cols')!;
    expect(cols.elements()).toHaveLength(0);
    expect(cols.attr('w:equalWidth')).toBeUndefined();
    expect(cols.attr('w:space')).toBe('100');
    expect(s.columns).toBe(3);
    expect(() => s.setColumns(0)).toThrow(UsageError);
    expect(childNames(s.element!)).toEqual(['pgMar', 'cols']);
    expect(() => (s.type = 'sideways' as 'continuous')).toThrow(UsageError);
    expect(() => (s.pageNumberStart = -1)).toThrow(UsageError);
  });

  it('setters on an implicit final section create the body sectPr as last child', () => {
    const doc = readDocx(minimalDocx('<w:p><w:r><w:t>a</w:t></w:r></w:p>'));
    const s = doc.sections()[0]!;
    expect(s.element).toBeUndefined();
    s.orientation = 'landscape';
    expect(s.element?.parent).toBe(doc.body.element);
    expect(doc.body.element.lastElement()).toBe(s.element);
    expect(s.pageSize).toEqual({ width: 15840, height: 12240, orientation: 'landscape' });
    setText(s.setHeader('default', 'new'), 'H');
    const back = roundTrip(doc, 'implicit-section.docx');
    expect(back.sections()[0]!.header()?.text()).toBe('H');
  });
});

// ---- section breaks ------------------------------------------------------------------------

describe('section breaks', () => {
  function abcd(): WordDocument {
    const doc = createDocx();
    for (const t of ['A', 'B', 'C', 'D']) doc.body.appendParagraph(t);
    return doc;
  }

  it('insertSectionBreak splits a section; headers move so the new section links to it', () => {
    const doc = abcd();
    const s = doc.sections()[0]!;
    setText(s.setHeader('default', 'new'), 'H');
    s.pageNumberStart = 1;
    const ended = insertSectionBreak(bodyParagraph(doc, 'B'), 'nextPage');
    const [s0, s1] = doc.sections();
    expect(doc.sections()).toHaveLength(2);
    expect(texts(s0!)).toEqual(['', 'A', 'B']);
    expect(texts(s1!)).toEqual(['C', 'D']);
    expect(ended.index).toBe(0);
    expect(ended.element).toBe(s0!.element);
    expect(s0!.element!.closest(W, 'pPr')).toBe(s0!.element!.parent);
    expect(s0!.ownHeader()?.text()).toBe('H');
    expect(s1!.ownHeader()).toBeUndefined();
    expect(s1!.header()).toBe(s0!.ownHeader());
    expect(s0!.pageSize).toEqual(s1!.pageSize);
    expect(s0!.pageNumberStart).toBe(1);
    expect(s1!.pageNumberStart).toBeUndefined();
    expect(s1!.type).toBe('nextPage');
    // The Section object of the last section still describes the last section.
    expect(s.index).toBe(1);
    s1!.orientation = 'landscape';
    const back = roundTrip(doc, 'section-break.docx');
    expect(back.sections().map((x) => x.orientation)).toEqual(['portrait', 'landscape']);
    expect(back.sections()[1]!.header()?.text()).toBe('H');
  });

  it('the break type goes to the section after the break', () => {
    const doc = abcd();
    insertSectionBreak(bodyParagraph(doc, 'A'), 'continuous');
    insertSectionBreak(bodyParagraph(doc, 'C'), 'oddPage');
    expect(doc.sections().map((s) => s.type)).toEqual(['nextPage', 'continuous', 'oddPage']);
    expect(doc.sections().map(texts)).toEqual([['', 'A'], ['B', 'C'], ['D']]);
    expect(doc.sections().every((s) => inSchemaOrder(s.element!))).toBe(true);
    roundTrip(doc, 'section-break-types.docx');
  });

  it('adds an empty paragraph when the break follows the last paragraph', () => {
    const doc = abcd();
    insertSectionBreak(bodyParagraph(doc, 'D'));
    const [s0, s1] = doc.sections();
    expect(texts(s0!)).toEqual(['', 'A', 'B', 'C', 'D']);
    expect(texts(s1!)).toEqual(['']);
    roundTrip(doc, 'section-break-at-end.docx');
  });

  it('rejects paragraphs that cannot take a break', () => {
    const doc = readDocx(fixture(F3).bytes);
    const s0 = doc.sections()[0]!;
    const last = s0.paragraphs().at(-1)!;
    expect(() => insertSectionBreak(last)).toThrow(UsageError);
    const hp = s0.header()!.story.paragraphs()[0]!;
    expect(() => insertSectionBreak(hp)).toThrow(UsageError);
    expect(() => insertSectionBreak(s0.paragraphs()[0]!, 'bogus' as 'continuous')).toThrow(
      UsageError,
    );
    const withTable = readDocx(
      minimalDocx(
        '<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/><w:sectPr/>',
      ),
    );
    expect(() => insertSectionBreak(bodyParagraph(withTable, 'cell'))).toThrow(UsageError);
  });

  it('removeSectionBreak merges with the next section, keeping inherited headers', () => {
    const doc = abcd();
    setText(doc.sections()[0]!.setHeader('default', 'new'), 'H');
    insertSectionBreak(bodyParagraph(doc, 'B'));
    const [s0, s1] = doc.sections();
    const pB = bodyParagraph(doc, 'B').element;
    removeSectionBreak(s0!);
    expect(doc.sections()).toHaveLength(1);
    expect(texts(doc.sections()[0]!)).toEqual(['', 'A', 'B', 'C', 'D']);
    expect(pB.element(W, 'pPr')).toBeUndefined();
    expect(s1!.index).toBe(0);
    expect(s1!.ownHeader()?.text()).toBe('H');
    expect(() => s0!.index).toThrow(UsageError);
    expect(() => removeSectionBreak(s1!)).toThrow(UsageError);
    roundTrip(doc, 'section-break-removed.docx');
  });

  it('removeSectionBreak drops headers only the removed section referenced', () => {
    const doc = readDocx(fixture(F3).bytes);
    const [s0, s1] = doc.sections();
    const h4 = s1!.ownHeader()!;
    const first = s0!.ownHeader('first')!;
    const before = headerFooterParts(doc).length;
    removeSectionBreak(s1!);
    const [m0, m1, extra] = doc.sections();
    expect(extra).toBeUndefined();
    // Section 3 inherited header4/footer4 from section 2: they moved to it.
    expect(m1!.ownHeader()).toBe(h4);
    expect(m1!.header('first')).toBe(s0!.ownHeader('first'));
    expect(m0!.header()?.text()).toBe('Text: Foo Bar');
    expect(headerFooterParts(doc)).toHaveLength(before);
    removeSectionBreak(m0!);
    // Section 1's own default header goes (the merged section has its own); its first/even
    // headers, which the next section inherited, move on.
    const [only] = doc.sections();
    expect(only!.header()).toBe(h4);
    expect(only!.ownHeader('first')).toBe(first);
    expect(doc.package.part('word/header2.xml')).toBeUndefined();
    expect(headerFooterParts(doc)).toHaveLength(before - 2);
    roundTrip(doc, 'section-breaks-removed-fixture.docx');
  });

  it('Section views stay valid across edits (scan cached by DOM version)', () => {
    const doc = abcd();
    const last = doc.sections()[0]!;
    expect(sectionsOf(doc)).toHaveLength(1);
    insertSectionBreak(bodyParagraph(doc, 'A'));
    expect(last.index).toBe(1);
    insertSectionBreak(bodyParagraph(doc, 'C'));
    expect(last.index).toBe(2);
    expect(texts(last)).toEqual(['D']);
    const middle = doc.sections()[1]!;
    setText(middle.setHeader('default', 'new'), 'M');
    expect(last.header()?.text()).toBe('M');
    expect(doc.sections()[0]!.header()).toBeUndefined();
    // Raw DOM edits are seen too.
    middle.element!.elements(W, 'headerReference')[0]!.remove();
    expect(last.header()).toBeUndefined();
  });
});

// ---- the fixtures ----------------------------------------------------------------------------

describe('fixtures', () => {
  const cases = [
    { name: F3, bytes: () => fixture(F3).bytes },
    { name: FMM, bytes: () => fixture(FMM).bytes },
    ...strictFixtures([F3, FMM]).map((f) => ({ name: f.name, bytes: () => f.bytes })),
  ];

  for (const c of cases) {
    it(`unlinks, edits and splits sections of ${c.name}`, () => {
      const doc = readDocx(c.bytes());
      const inputErrors = errors(doc);
      const sections = doc.sections();
      const n = sections.length;
      // Split the first section after its first paragraph.
      const ended = insertSectionBreak(sections[0]!.paragraphs()[0]!, 'nextPage');
      expect(doc.sections()).toHaveLength(n + 1);
      expect(ended.index).toBe(0);
      const all = doc.sections();
      const shown = all.map((s) => s.header()?.text());
      for (const s of all) {
        for (const kind of ['header', 'footer'] as const) s.unlinkFromPrevious(kind);
      }
      // Unlinking changes nothing visible.
      expect(all.map((s) => s.header()?.text())).toEqual(shown);
      for (const s of all) {
        const h = s.unlinkFromPrevious('header', 'default');
        h.story.appendParagraph(`Edited S${s.index}`);
      }
      // Every header part is referenced by exactly one section now.
      const owners = new Map<string, number>();
      for (const s of all) {
        for (const t of ['default', 'first', 'even'] as const) {
          const own = s.ownHeader(t);
          if (own) owners.set(own.part.name, (owners.get(own.part.name) ?? 0) + 1);
        }
      }
      expect([...owners.values()].every((v) => v === 1)).toBe(true);
      setEvenAndOddHeaders(doc, true);
      expect(errors(doc)).toEqual(inputErrors);
      const file = `fixture-${c.name.replace(/[/]/g, '-')}`;
      const bytes = writeDocx(doc);
      writeFileSync(`${OUT}${file}`, bytes);
      const back = readDocx(bytes);
      expect(errors(back)).toEqual(inputErrors);
      back.sections().forEach((s, i) => {
        expect(s.header()?.text()).toContain(`Edited S${i}`);
        expect(
          s
            .ownHeader()
            ?.text()
            .startsWith(shown[i] ?? ''),
        ).toBe(true);
      });
      expect(evenAndOddHeaders(back)).toBe(true);
    });
  }

  it('leaves untouched parts byte-identical', () => {
    const src = fixture(F3).bytes;
    const doc = readDocx(src);
    doc.sections()[2]!.unlinkFromPrevious('header', 'default');
    const out = unzipSync(writeDocx(doc));
    const inp = unzipSync(src);
    for (const name of ['word/header1.xml', 'word/header4.xml', 'word/styles.xml']) {
      expect(strFromU8(out[name]!), name).toBe(strFromU8(inp[name]!));
    }
    expect(Object.keys(out)).toContain('word/header5.xml');
  });
});
