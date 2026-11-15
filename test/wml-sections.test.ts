/// <reference types="node" />
import { readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readDocx } from '../src/index.js';
import { HeaderFooter, headerFooterParts } from '../src/wml/header-footer.js';
import { type Section, sectionsOf } from '../src/wml/section.js';

// ---- a tiny docx builder (fflate only, independent of src/zip) ---------------------------

const enc = new TextEncoder();
const W_URI = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NSDECL = `xmlns:w="${W_URI}" xmlns:r="${R_URI}"`;
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_BASE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

interface Rel {
  id: string;
  type: string;
  target: string;
  external?: boolean;
}

/** Builds a .docx whose body is `body`, with header/footer parts `parts` (name → text). */
function buildDocx(
  body: string,
  options: { parts?: Record<string, string>; rels?: Rel[] } = {},
): Uint8Array {
  const parts = options.parts ?? {};
  const files: Record<string, Uint8Array> = {};
  let overrides = `<Override PartName="/word/document.xml" ContentType="${CT_BASE}.document.main+xml"/>`;
  for (const [name, text] of Object.entries(parts)) {
    const kind = name.startsWith('header') ? 'header' : 'footer';
    const root = kind === 'header' ? 'w:hdr' : 'w:ftr';
    overrides += `<Override PartName="/word/${name}" ContentType="${CT_BASE}.${kind}+xml"/>`;
    files[`word/${name}`] = enc.encode(
      `${XML_DECL}<${root} ${NSDECL}><w:p><w:r><w:t>${text}</w:t></w:r></w:p></${root}>`,
    );
  }
  files['[Content_Types].xml'] = enc.encode(
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `${overrides}</Types>`,
  );
  files['_rels/.rels'] = enc.encode(
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_BASE}/officeDocument" Target="word/document.xml"/>` +
      '</Relationships>',
  );
  const rels = (options.rels ?? [])
    .map(
      (r) =>
        `<Relationship Id="${r.id}" Type="${REL_BASE}/${r.type}" Target="${r.target}"` +
        `${r.external === true ? ' TargetMode="External"' : ''}/>`,
    )
    .join('');
  files['word/_rels/document.xml.rels'] = enc.encode(
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
  );
  files['word/document.xml'] = enc.encode(
    `${XML_DECL}<w:document ${NSDECL}><w:body>${body}</w:body></w:document>`,
  );
  return zipSync(files);
}

const p = (text: string, sectPr = ''): string =>
  `<w:p>${sectPr === '' ? '' : `<w:pPr>${sectPr}</w:pPr>`}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const hdr = (type: string, id: string): string =>
  `<w:headerReference w:type="${type}" r:id="${id}"/>`;
const ftr = (type: string, id: string): string =>
  `<w:footerReference w:type="${type}" r:id="${id}"/>`;
const texts = (s: Section): string[] => s.paragraphs().map((x) => x.text);

// ---- the multi-section document ----------------------------------------------------------

const MULTI_BODY =
  p('S0 P1') +
  `<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc>${p('in table')}</w:tc></w:tr></w:tbl>` +
  p(
    'S0 P2',
    '<w:sectPr>' +
      hdr('default', 'rIdH1') +
      hdr('first', 'rIdH2') +
      ftr('default', 'rIdF1') +
      '<w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>' +
      '<w:pgNumType w:start="3"/><w:cols w:num="2" w:space="720"/><w:titlePg/>' +
      '</w:sectPr>',
  ) +
  '<w:sdt><w:sdtPr><w:tag w:val="wrapped"/></w:sdtPr><w:sdtContent>' +
  p('S1 P1') +
  p(
    'S1 P2',
    '<w:sectPr>' +
      hdr('even', 'rIdH3') +
      '<w:type w:val="continuous"/><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>' +
      '<w:titlePg w:val="0"/>' +
      '</w:sectPr>',
  ) +
  '</w:sdtContent></w:sdt>' +
  `<w:customXml w:element="x">${p('S2 P1')}</w:customXml>` +
  p('S2 P2') +
  '<w:sectPr>' +
  ftr('default', 'rIdF2') +
  hdr('first', 'rIdH4') +
  '<w:type w:val="oddPage"/>' +
  '<w:pgMar w:top="1in" w:right="2.54cm" w:bottom="-720" w:left="72pt" w:header="0" w:footer="0" w:gutter="10mm"/>' +
  '<w:cols w:equalWidth="0"><w:col w:w="100"/><w:col w:w="100"/><w:col w:w="100"/></w:cols>' +
  '</w:sectPr>';

const MULTI_PARTS = {
  'header1.xml': 'Header default',
  'header2.xml': 'Header first',
  'header3.xml': 'Header even',
  'header4.xml': 'Header first 2',
  'footer1.xml': 'Footer default',
  'footer2.xml': 'Footer default 2',
};

const MULTI_RELS: Rel[] = [
  { id: 'rIdH1', type: 'header', target: 'header1.xml' },
  { id: 'rIdLink', type: 'hyperlink', target: 'https://example.com/', external: true },
  { id: 'rIdH2', type: 'header', target: 'header2.xml' },
  { id: 'rIdF1', type: 'footer', target: 'footer1.xml' },
  { id: 'rIdH3', type: 'header', target: 'header3.xml' },
  { id: 'rIdF2', type: 'footer', target: 'footer2.xml' },
  { id: 'rIdH4', type: 'header', target: '/word/header4.xml' },
  { id: 'rIdDup', type: 'header', target: 'header1.xml' },
];

function multi() {
  return readDocx(buildDocx(MULTI_BODY, { parts: MULTI_PARTS, rels: MULTI_RELS }));
}

describe('sections', () => {
  it('splits the body at paragraph sectPrs, looking through sdt and customXml', () => {
    const doc = multi();
    const sections = sectionsOf(doc);
    expect(sections).toHaveLength(3);
    expect(sections.map((s) => s.index)).toEqual([0, 1, 2]);
    expect(sections.map(texts)).toEqual([
      ['S0 P1', 'S0 P2'],
      ['S1 P1', 'S1 P2'],
      ['S2 P1', 'S2 P2'],
    ]);
    expect(sections[0]!.element?.closest(W_URI, 'p')).toBe(sections[0]!.paragraphs()[1]!.element);
    expect(sections[2]!.element?.parent).toBe(doc.body.element);
    expect(sections[0]!.paragraphs()[0]!.story).toBe(doc.body);
    expect(doc.sections()).toHaveLength(3);
  });

  it('reads section properties', () => {
    const [s0, s1, s2] = sectionsOf(multi());
    expect(s0!.type).toBe('nextPage');
    expect(s0!.pageSize).toEqual({ width: 12240, height: 15840, orientation: 'portrait' });
    expect(s0!.margins).toEqual({
      top: 1440,
      right: 1800,
      bottom: 1440,
      left: 1800,
      header: 720,
      footer: 720,
      gutter: 0,
    });
    expect(s0!.titlePage).toBe(true);
    expect(s0!.columns).toBe(2);
    expect(s0!.pageNumberStart).toBe(3);

    expect(s1!.type).toBe('continuous');
    expect(s1!.pageSize).toEqual({ width: 15840, height: 12240, orientation: 'landscape' });
    expect(s1!.margins).toBeUndefined();
    expect(s1!.titlePage).toBe(false);
    expect(s1!.columns).toBe(1);
    expect(s1!.pageNumberStart).toBeUndefined();

    expect(s2!.type).toBe('oddPage');
    expect(s2!.pageSize).toBeUndefined();
    // Universal measures (1in, 2.54cm, 72pt, 10mm) are converted to twips.
    expect(s2!.margins).toEqual({
      top: 1440,
      right: 1440,
      bottom: -720,
      left: 1440,
      header: 0,
      footer: 0,
      gutter: 567,
    });
    expect(s2!.columns).toBe(3);
  });

  it('resolves own and effective headers/footers per type, inheriting independently', () => {
    const [s0, s1, s2] = sectionsOf(multi());
    const text = (hf: HeaderFooter | undefined): string | undefined => hf?.text();

    expect(text(s0!.ownHeader())).toBe('Header default');
    expect(text(s0!.ownHeader('first'))).toBe('Header first');
    expect(s0!.ownHeader('even')).toBeUndefined();
    expect(text(s0!.ownFooter())).toBe('Footer default');
    expect(s0!.header('even')).toBeUndefined();
    expect(s0!.footer('first')).toBeUndefined();

    expect(s1!.ownHeader()).toBeUndefined();
    expect(s1!.ownFooter()).toBeUndefined();
    expect(s1!.header()).toBe(s0!.ownHeader());
    expect(s1!.header('first')).toBe(s0!.ownHeader('first'));
    expect(text(s1!.header('even'))).toBe('Header even');
    expect(s1!.footer()).toBe(s0!.ownFooter());

    expect(s2!.header()).toBe(s0!.ownHeader());
    expect(text(s2!.header('first'))).toBe('Header first 2');
    expect(s2!.header('even')).toBe(s1!.ownHeader('even'));
    expect(text(s2!.footer())).toBe('Footer default 2');
    expect(s2!.footer('even')).toBeUndefined();

    const h = s2!.header()!;
    expect(h).toBeInstanceOf(HeaderFooter);
    expect(h.kind).toBe('header');
    expect(h.part.name).toBe('word/header1.xml');
    expect(h.story.kind).toBe('header');
    expect(h.story.part).toBe(h.part);
    expect(h.story.element.local).toBe('hdr');
    expect(h.story).toBe(h.story);
    expect(s2!.footer()!.kind).toBe('footer');
    expect(s2!.footer()!.story.kind).toBe('footer');
  });

  it('lists header/footer parts in rels order, deduped, with stable identity', () => {
    const doc = multi();
    const all = headerFooterParts(doc);
    expect(all.map((hf) => hf.part.name)).toEqual([
      'word/header1.xml',
      'word/header2.xml',
      'word/footer1.xml',
      'word/header3.xml',
      'word/footer2.xml',
      'word/header4.xml',
    ]);
    expect(all.map((hf) => hf.kind)).toEqual([
      'header',
      'header',
      'footer',
      'header',
      'footer',
      'header',
    ]);
    const [s0, s1] = sectionsOf(doc);
    expect(all[0]).toBe(s0!.ownHeader());
    expect(all[0]).toBe(sectionsOf(doc)[2]!.header());
    expect(all[3]).toBe(s1!.ownHeader('even'));
    expect(headerFooterParts(doc)[4]).toBe(all[4]);
    expect(all.map((hf) => hf.text())).toEqual([
      'Header default',
      'Header first',
      'Footer default',
      'Header even',
      'Footer default 2',
      'Header first 2',
    ]);
  });

  it('gives a body without a final sectPr an implicit last section', () => {
    const doc = readDocx(
      buildDocx(p('A', `<w:sectPr>${hdr('default', 'rIdH1')}</w:sectPr>`) + p('B') + p('C'), {
        parts: { 'header1.xml': 'H' },
        rels: [{ id: 'rIdH1', type: 'header', target: 'header1.xml' }],
      }),
    );
    const sections = sectionsOf(doc);
    expect(sections).toHaveLength(2);
    const last = sections[1]!;
    expect(last.element).toBeUndefined();
    expect(texts(last)).toEqual(['B', 'C']);
    expect(last.type).toBe('nextPage');
    expect(last.pageSize).toBeUndefined();
    expect(last.margins).toBeUndefined();
    expect(last.titlePage).toBe(false);
    expect(last.columns).toBe(1);
    expect(last.pageNumberStart).toBeUndefined();
    expect(last.ownHeader()).toBeUndefined();
    expect(last.header()?.text()).toBe('H');
    expect(last.header()).toBe(sections[0]!.ownHeader());
  });

  it('handles an empty body and a body with only a final sectPr', () => {
    const empty = sectionsOf(readDocx(buildDocx('')));
    expect(empty).toHaveLength(1);
    expect(empty[0]!.element).toBeUndefined();
    expect(empty[0]!.paragraphs()).toEqual([]);
    expect(empty[0]!.header()).toBeUndefined();

    const one = sectionsOf(readDocx(buildDocx(`${p('only')}<w:sectPr/>`)));
    expect(one).toHaveLength(1);
    expect(one[0]!.element?.local).toBe('sectPr');
    expect(texts(one[0]!)).toEqual(['only']);
  });

  it('treats a dangling reference as defining the slot (no inheritance past it)', () => {
    const doc = readDocx(
      buildDocx(
        p('A', `<w:sectPr>${hdr('default', 'rIdH1')}</w:sectPr>`) +
          p('B') +
          `<w:sectPr>${hdr('default', 'rIdMissing')}${hdr('first', 'rIdH1')}</w:sectPr>`,
        {
          parts: { 'header1.xml': 'H' },
          rels: [{ id: 'rIdH1', type: 'header', target: 'header1.xml' }],
        },
      ),
    );
    const [s0, s1] = sectionsOf(doc);
    expect(s0!.header()?.text()).toBe('H');
    expect(s1!.ownHeader()).toBeUndefined();
    expect(s1!.header()).toBeUndefined();
    // The same part referenced under another type is the same object.
    expect(s1!.ownHeader('first')).toBe(s0!.ownHeader());
  });

  it('reads on/off values of titlePg', () => {
    const doc = readDocx(
      buildDocx(
        p('a', '<w:sectPr><w:titlePg w:val="true"/></w:sectPr>') +
          p('b', '<w:sectPr><w:titlePg w:val="off"/></w:sectPr>') +
          p('c', '<w:sectPr><w:titlePg w:val="1"/></w:sectPr>') +
          '<w:sectPr><w:titlePg w:val="false"/><w:type w:val="bogus"/></w:sectPr>',
      ),
    );
    const s = sectionsOf(doc);
    expect(s.map((x) => x.titlePage)).toEqual([true, false, true, false]);
    expect(s[3]!.type).toBe('nextPage');
  });

  it('reflects DOM edits on the next call', () => {
    const doc = readDocx(buildDocx(`${p('A')}${p('B')}<w:sectPr/>`));
    expect(sectionsOf(doc)).toHaveLength(1);
    const first = sectionsOf(doc)[0]!.paragraphs()[0]!.element;
    const xml = doc.mainPart.xml();
    const pPr = xml.createElement(W_URI, 'pPr');
    pPr.append(xml.createElement(W_URI, 'sectPr'));
    first.prepend(pPr);
    const sections = sectionsOf(doc);
    expect(sections.map(texts)).toEqual([['A'], ['B']]);
  });
});

// ---- smoke tests over the shared fixtures (when present) ---------------------------------

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url));
let fixtures: string[] = [];
try {
  fixtures = readdirSync(FIXTURES, { recursive: true, encoding: 'utf8' })
    .filter((f) => f.toLowerCase().endsWith('.docx'))
    .sort();
} catch {
  fixtures = [];
}

describe('fixtures smoke', () => {
  it.skipIf(fixtures.length === 0)('reads sections, headers/footers and table grids', () => {
    for (const name of fixtures) {
      const bytes = readFileSync(`${FIXTURES}${name}`);
      // Password-protected .docx files are OLE compound files, not zip packages: skip them.
      if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) continue;
      const doc = readDocx(bytes);
      const sections = sectionsOf(doc);
      expect(sections.length, name).toBeGreaterThan(0);
      const all = new Set(headerFooterParts(doc));
      for (const s of sections) {
        for (const t of ['default', 'first', 'even'] as const) {
          const h = s.header(t);
          const f = s.footer(t);
          if (h !== undefined) {
            expect(all.has(h), `${name}: header ${t} listed`).toBe(true);
            expect(typeof h.text()).toBe('string');
          }
          if (f !== undefined) expect(all.has(f), `${name}: footer ${t} listed`).toBe(true);
        }
        expect(typeof s.type).toBe('string');
        // Every section but the last ends with the paragraph holding its sectPr.
        if (s.index < sections.length - 1) {
          const last = s.paragraphs().at(-1);
          expect(last?.element, `${name}: section ${s.index}`).toBe(s.element?.parent?.parent);
        }
      }
      for (const t of doc.body.tables()) {
        const g = t.grid();
        expect(g.cols, name).toBeGreaterThanOrEqual(0);
        for (let r = 0; r < g.rows; r++) {
          for (let c = 0; c < g.cols; c++) {
            const slot = g.at(r, c);
            if (slot === undefined) continue;
            expect(slot.row).toBeLessThanOrEqual(r);
            expect(slot.col).toBeLessThanOrEqual(c);
            expect(slot.row + slot.rowSpan).toBeGreaterThan(r);
            expect(slot.col + slot.colSpan).toBeGreaterThan(c);
          }
        }
        expect(typeof t.text()).toBe('string');
      }
    }
  });
});

describe('fixtures: known section layouts', () => {
  const load = (name: string) => {
    try {
      return readDocx(readFileSync(`${FIXTURES}${name}`));
    } catch (e) {
      if ((e as { code?: string }).code === 'ENOENT') return undefined;
      throw e;
    }
  };

  it('docxcompose: 3 sections, per-type inheritance of headers/footers', () => {
    const doc = load('docxcompose/docproperties_header_footer_3_sections.docx');
    if (doc === undefined) return;
    const [s0, s1, s2, extra] = sectionsOf(doc);
    expect(extra).toBeUndefined();
    expect(s0!.header()?.text()).toBe('Text: Foo Bar');
    expect(s0!.footer()?.text()).toBe('Number: 123');
    expect(s1!.header()?.text()).toBe('Date: 11.06.2019');
    expect(s1!.footer()?.text()).toContain('Float: 1.1');
    // s1 only overrides `default`: first/even are inherited from s0.
    expect(s1!.header('first')).toBe(s0!.ownHeader('first'));
    expect(s1!.footer('even')).toBe(s0!.ownFooter('even'));
    expect(s2!.type).toBe('continuous');
    expect(s2!.header()).toBe(s1!.ownHeader());
    expect(s2!.footer()).toBe(s1!.ownFooter());
    expect(headerFooterParts(doc)).toHaveLength(8);
  });

  it('docx-mailmerge: titlePg with first/even/default headers and footers', () => {
    const doc = load('docx-mailmerge/test_footnote_header_footer.docx');
    if (doc === undefined) return;
    const [s] = sectionsOf(doc);
    expect(s!.titlePage).toBe(true);
    expect(s!.header()?.text()).toContain('Header on every page');
    expect(s!.header('first')?.text()).toContain('Header on first page');
    expect(s!.header('even')?.text()).toContain('Header even');
    expect(s!.footer('first')?.text()).toContain('Footer on first page');
    expect(s!.footer('even')?.text()).toContain('Footer on even page');
  });

  it('open-xml-sdk: Strict OOXML footer references resolve', () => {
    const doc = load('open-xml-sdk/strict-footer-bookmarkStart.docx');
    if (doc === undefined) return;
    const [s] = sectionsOf(doc);
    expect(s!.footer()?.text()).toBe('Bookmark');
    expect(s!.pageSize).toEqual({ width: 12240, height: 15840, orientation: 'portrait' });
  });
});
