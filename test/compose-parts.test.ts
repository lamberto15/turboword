/// <reference types="node" />
// Compose "parts" steps: notes, comments (five parts), bookmarks, sections + headers/footers,
// settings/fonts/theme.
//
// Unit tests run the steps through a minimal pipeline built here (real WordDocuments, a simple
// PartCopier of our own, the real DocumentIds, importContent and renumberIds), so they exercise
// the steps through their contract only. The last block runs the fixture pairs through the real
// orchestrator (appendDocument / insertDocument / composeDocuments).
// Set COMPOSE_PARTS_OUT=<dir> to also write every merged document there (for LibreOffice).

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { importBookmarks, rewriteInstruction } from '../src/compose/bookmarks.js';
import { importComments } from '../src/compose/comments.js';
import { appendDocument, composeDocuments, insertDocument } from '../src/compose/index.js';
import { importNotes } from '../src/compose/notes.js';
import { adoptNamespaces } from '../src/compose/parts-util.js';
import { renumberIds } from '../src/compose/renumber.js';
import { importSections } from '../src/compose/sections.js';
import { mergeSettings } from '../src/compose/settings.js';
import type { ImportContext, ImportOptions, PartCopier } from '../src/compose/types.js';
import {
  createDocx,
  readDocx,
  validatePackage,
  type WordDocument,
  writeDocx,
} from '../src/index.js';
import type { Part } from '../src/opc/part.js';
import { relativeTarget } from '../src/opc/uri.js';
import { DocumentIds } from '../src/wml/ids.js';
import type { Section } from '../src/wml/section.js';
import type { XElement } from '../src/xml/dom.js';
import { NS } from '../src/xml/ns.js';
import { parseXml } from '../src/xml/parse.js';
import { fixture } from './helpers/fixtures.js';

const OUT = process.env.COMPOSE_PARTS_OUT;

// ---- a minimal pipeline ---------------------------------------------------------------------

const DEFAULTS: Required<Omit<ImportOptions, 'onWarning'>> = {
  styles: 'useDestinationStyles',
  smartStyleBehavior: true,
  lists: 'keepSource',
  restartListsAtImport: false,
  sectionStart: 'nextPage',
  headersFooters: 'keepSource',
  pageNumbering: 'keepSource',
  keepSourceSections: true,
  theme: 'destination',
  bookmarks: 'rename',
  comments: 'keep',
  revisions: 'keep',
  validate: false,
};

/** Copies parts with their relationship graphs (same relationship ids), memoized per part. */
function testCopier(target: WordDocument, added: string[]): PartCopier {
  const memo = new Map<Part, Part>();
  const copyPart = (src: Part): Part => {
    const known = memo.get(src);
    if (known !== undefined) return known;
    const m = /^(.*?)\d*(\.[^./]+)$/.exec(src.name);
    const name = target.package.uniqueName(m ? `${m[1]}{n}${m[2]}` : `${src.name}{n}`);
    const part = target.package.addPart(
      name,
      src.contentType ?? 'application/octet-stream',
      src.bytes().slice(),
    );
    memo.set(src, part);
    added.push(name);
    for (const rel of src.rels.all()) {
      if (rel.external) {
        part.rels.add(rel.type, rel.target, { external: true, id: rel.id });
        continue;
      }
      const t = rel.partName === undefined ? undefined : src.package.part(rel.partName);
      if (t === undefined) continue;
      part.rels.add(rel.type, relativeTarget(part.name, copyPart(t).name), { id: rel.id });
    }
    return part;
  };
  return {
    copyPart,
    copyRelationship(srcPart, relId, tgtPart) {
      const rel = srcPart.rels.get(relId);
      if (rel === undefined) throw new Error(`no relationship ${relId} in ${srcPart.name}`);
      if (rel.external) return tgtPart.rels.add(rel.type, rel.target, { external: true }).id;
      const t = srcPart.package.part(rel.partName!)!;
      return tgtPart.rels.add(rel.type, relativeTarget(tgtPart.name, copyPart(t).name)).id;
    },
  };
}

function makeCtx(target: WordDocument, source: WordDocument, o: ImportOptions = {}): ImportContext {
  const addedParts: string[] = [];
  const warnings: string[] = [];
  const { onWarning: _ignored, ...rest } = o;
  return {
    target,
    source,
    options: { ...DEFAULTS, ...rest },
    ids: DocumentIds.of(target),
    copier: testCopier(target, addedParts),
    warn: (code) => warnings.push(code),
    styleMap: new Map(),
    numIdMap: new Map(),
    bookmarkRenames: new Map(),
    commentIdMap: new Map(),
    footnoteIdMap: new Map(),
    endnoteIdMap: new Map(),
    addedParts,
    importedStyles: [],
    // test-only: the warning codes
    ...({ warnings } as object),
  } as ImportContext;
}

function warningsOf(ctx: ImportContext): string[] {
  return (ctx as unknown as { warnings: string[] }).warnings;
}

function bodyOf(doc: WordDocument): XElement {
  return doc.body.element;
}

interface Composed {
  readonly ctx: ImportContext;
  readonly blocks: XElement[];
  readonly sections: number;
}

/**
 * The compose pipeline minus styles/numbering: clone the source blocks into the target, run the
 * parts steps around a stand-in for the relationship/id steps (which, like remapRelationships,
 * leaves header/footer references to importSections).
 */
function compose(
  target: WordDocument,
  source: WordDocument,
  o: ImportOptions = {},
  v: { insertAfter?: XElement } = {},
): Composed {
  const ctx = makeCtx(target, source, o);
  const body = bodyOf(target);
  const root = target.mainPart.xml().root;
  const blocks = bodyOf(source)
    .elements()
    .filter((e) => !e.is(NS.w, 'sectPr'))
    .map((e) => e.clone());
  if (v.insertAfter !== undefined) v.insertAfter.after(...blocks);
  else {
    const final = body.element(NS.w, 'sectPr') ?? null;
    for (const b of blocks) body.insertBefore(b, final);
  }
  for (const b of blocks) adoptNamespaces(b, root);
  importNotes(ctx, blocks);
  importComments(ctx, blocks);
  importBookmarks(ctx, blocks);
  for (const b of blocks) {
    const stack: XElement[] = [b];
    while (stack.length > 0) {
      const el = stack.pop()!;
      stack.push(...el.elements());
      if (el.is(NS.w, 'headerReference') || el.is(NS.w, 'footerReference')) continue;
      for (const a of el.attributes()) {
        if (a.uri !== NS.r) continue;
        el.setAttr(a.name, ctx.copier.copyRelationship(source.mainPart, a.value, target.mainPart));
      }
    }
  }
  renumberIds(ctx, blocks);
  const sections = importSections(ctx, blocks, {
    mode: v.insertAfter === undefined ? 'append' : 'insert',
    part: target.mainPart,
  });
  mergeSettings(ctx);
  return { ctx, blocks, sections };
}

function open(name: string): WordDocument {
  return readDocx(fixture(name).bytes);
}

/** Saves, re-opens, checks validatePackage has no errors; optionally writes the file. */
function roundTrip(doc: WordDocument, label: string): WordDocument {
  const errors = validatePackage(doc).filter((i) => i.severity === 'error');
  expect(errors, label).toEqual([]);
  const bytes = writeDocx(doc);
  if (OUT !== undefined) {
    mkdirSync(OUT, { recursive: true });
    writeFileSync(join(OUT, `${label}.docx`), bytes);
  }
  const back = readDocx(bytes);
  expect(
    validatePackage(back).filter((i) => i.severity === 'error'),
    label,
  ).toEqual([]);
  return back;
}

const SETTINGS_KIND = 'settings';

function evenAndOdd(doc: WordDocument): boolean {
  const s = doc.mainPart.relatedByType(SETTINGS_KIND)[0];
  const el = s?.xml().root.element(NS.w, 'evenAndOddHeaders');
  if (el === undefined) return false;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined || !['false', '0', 'off'].includes(v);
}

/** What a header/footer shows: text plus number of drawings ('' when there is none). */
function sig(hf: ReturnType<Section['header']>): string {
  if (hf === undefined) return '';
  const root = hf.part.xml().root;
  const pics = root.descendants(NS.w, 'drawing').length + root.descendants(NS.w, 'pict').length;
  const text = hf.text().trim();
  return pics > 0 ? `${text}[${pics} img]` : text;
}

/** What section `i` shows on its first, odd and even pages (headers and footers). */
function shown(doc: WordDocument, i: number): Record<string, string> {
  const s = doc.sections()[i]!;
  const eo = evenAndOdd(doc);
  const out: Record<string, string> = {};
  for (const kind of ['header', 'footer'] as const) {
    const get = (t: 'default' | 'first' | 'even'): string =>
      sig(kind === 'header' ? s.header(t) : s.footer(t));
    out[`${kind}First`] = s.titlePage ? get('first') : get('default');
    out[`${kind}Odd`] = get('default');
    out[`${kind}Even`] = eo ? get('even') : get('default');
  }
  return out;
}

/** Every section of `merged` shows what the corresponding target/source section showed. */
function expectSameAppearance(merged: WordDocument, target: WordDocument, source: WordDocument) {
  const nT = target.sections().length;
  const nS = source.sections().length;
  expect(merged.sections().length).toBe(nT + nS);
  for (let i = 0; i < nT; i++)
    expect(shown(merged, i), `target section ${i}`).toEqual(shown(target, i));
  for (let j = 0; j < nS; j++) {
    expect(shown(merged, nT + j), `source section ${j}`).toEqual(shown(source, j));
  }
}

const THREE = 'docxcompose/docproperties_header_footer_3_sections.docx';
const IMAGE = 'docxcompose/header_with_image.docx';
const MAILMERGE = 'docx-mailmerge/test_footnote_header_footer.docx';
const COMMENTS = 'python-docx/comments-rich-para.docx';
const NOTES = 'open-xml-sdk/TableFntEdnt.docx';
const TOC = 'open-xml-sdk/TOC2.docx';

// ---- sections and headers/footers -----------------------------------------------------------

describe('importSections', () => {
  const pairs: [string, string][] = [
    [IMAGE, THREE],
    [THREE, IMAGE],
    [THREE, THREE],
    [MAILMERGE, IMAGE],
    [IMAGE, MAILMERGE],
    [THREE, MAILMERGE],
    [MAILMERGE, THREE],
    [THREE, NOTES],
    [MAILMERGE, NOTES],
    [MAILMERGE, MAILMERGE],
  ];
  for (const [t, s] of pairs) {
    it(`${t} + ${s}: every section looks as before`, () => {
      const target = open(t);
      const source = open(s);
      const { sections } = compose(target, source);
      expect(sections).toBe(open(s).sections().length);
      expect(source.package.dirty).toBe(false);
      const label = `sections-${t.split('/')[1]}-${s.split('/')[1]}`.replace(/\.docx/g, '');
      const merged = roundTrip(target, label);
      expectSameAppearance(merged, open(t), open(s));
      // The body sectPr is last.
      expect(bodyOf(merged).lastElement()?.is(NS.w, 'sectPr')).toBe(true);
    });
  }

  it('materializes inheritance: a source without headers does not show the target ones', () => {
    const target = open(THREE);
    compose(target, open(NOTES));
    const merged = roundTrip(target, 'sections-materialize');
    const last = shown(merged, merged.sections().length - 1);
    expect(Object.values(last).every((v) => v === '')).toBe(true);
    // ...and the target's own sections keep theirs.
    expect(shown(merged, 0).headerOdd).toBe('Text: Foo Bar');
  });

  it('turns evenAndOddHeaders on for a source that needs it, mirroring the target sections', () => {
    const target = open(THREE);
    expect(evenAndOdd(target)).toBe(false);
    compose(target, open(MAILMERGE));
    const merged = roundTrip(target, 'sections-evenodd');
    expect(evenAndOdd(merged)).toBe(true);
    expect(shown(merged, 0).headerEven).toBe('Text: Foo Bar');
    expect(shown(merged, 3)).toEqual(shown(open(MAILMERGE), 0));
  });

  it('mirrors the imported sections when only the target uses evenAndOddHeaders', () => {
    const target = open(MAILMERGE);
    compose(target, open(THREE));
    const merged = roundTrip(target, 'sections-evenodd-target');
    expect(evenAndOdd(merged)).toBe(true);
    for (let i = 1; i < 4; i++) {
      const s = shown(merged, i);
      expect(s.headerEven).toBe(s.headerOdd);
      expect(s.footerEven).toBe(s.footerOdd);
    }
  });

  it('headersFooters linkToPrevious: imported sections inherit the target ones', () => {
    const target = open(THREE);
    const { sections } = compose(target, open(IMAGE), { headersFooters: 'linkToPrevious' });
    expect(sections).toBe(1);
    const merged = roundTrip(target, 'sections-link');
    const last = merged.sections().at(-1)!;
    expect(last.element?.elements(NS.w, 'headerReference')).toEqual([]);
    expect(shown(merged, 3)).toEqual(shown(open(THREE), 2));
  });

  it('headersFooters none: imported sections show empty headers and footers', () => {
    const target = open(MAILMERGE);
    compose(target, open(THREE), { headersFooters: 'none' });
    const merged = roundTrip(target, 'sections-none');
    for (let i = 1; i < 4; i++) {
      expect(Object.values(shown(merged, i)).every((v) => v === '')).toBe(true);
    }
    expect(shown(merged, 0)).toEqual(shown(open(MAILMERGE), 0));
  });

  it('keepSourceSections false drops the source sectPrs with a warning', () => {
    const target = open(IMAGE);
    const { ctx, sections } = compose(target, open(THREE), { keepSourceSections: false });
    expect(sections).toBe(0);
    expect(warningsOf(ctx)).toContain('SECTIONS_DROPPED');
    const merged = roundTrip(target, 'sections-dropped');
    expect(merged.sections().length).toBe(1);
    expect(shown(merged, 0)).toEqual(shown(open(IMAGE), 0));
  });

  it('pageNumbering restart/continue and sectionStart adjust the first imported section', () => {
    const source = open(IMAGE);
    const sp = bodyOf(source).element(NS.w, 'sectPr')!;
    const pg = sp.ownerDocument()!.createElement(NS.w, 'pgNumType', { 'w:start': '5' });
    sp.insertBefore(pg, sp.element(NS.w, 'cols') ?? null);
    const bytes = writeDocx(source);

    const t1 = open(THREE);
    compose(t1, readDocx(bytes), { pageNumbering: 'restart', sectionStart: 'continuous' });
    const s1 = roundTrip(t1, 'sections-restart').sections().at(-1)!;
    expect(s1.pageNumberStart).toBe(1);
    expect(s1.type).toBe('continuous');

    const t2 = open(THREE);
    compose(t2, readDocx(bytes), { pageNumbering: 'continue' });
    const s2 = roundTrip(t2, 'sections-continue').sections().at(-1)!;
    expect(s2.pageNumberStart).toBeUndefined();
    expect(s2.type).toBe('nextPage');

    const t3 = open(THREE);
    compose(t3, readDocx(bytes));
    expect(roundTrip(t3, 'sections-keep').sections().at(-1)!.pageNumberStart).toBe(5);
  });

  it('moves the target sectPr onto a new paragraph when the target ends with a table', () => {
    const target = open(NOTES);
    const body = bodyOf(target);
    const last = body
      .elements()
      .filter((e) => !e.is(NS.w, 'sectPr'))
      .at(-1)!;
    // Make the target end with a table.
    if (!last.is(NS.w, 'tbl')) {
      const tbl = body.elements(NS.w, 'tbl')[0]!;
      body.insertBefore(tbl, body.element(NS.w, 'sectPr') ?? null);
      for (const p of body.elements(NS.w, 'p'))
        if (body.elements().indexOf(p) > body.elements().indexOf(tbl)) p.remove();
    }
    compose(target, open(IMAGE));
    const merged = roundTrip(target, 'sections-after-table');
    expect(merged.sections().length).toBe(2);
    const holder = merged.sections()[0]!.paragraphs().at(-1)!;
    expect(holder.element.previousElement()?.is(NS.w, 'tbl')).toBe(true);
  });

  it('insert mode splits the host section and pins the remainder headers', () => {
    const target = open(THREE);
    // Insert after the first paragraph of the target's first section.
    const firstP = target.sections()[0]!.paragraphs()[0]!.element;
    const { sections } = compose(target, open(MAILMERGE), {}, { insertAfter: firstP });
    expect(sections).toBe(1);
    const merged = roundTrip(target, 'sections-insert');
    const original = open(THREE);
    expect(merged.sections().length).toBe(original.sections().length + 2);
    // before the insertion: the host section's look; the import: the source's; after: the host.
    const host = shown(original, 0);
    const mirror = (s: Record<string, string>) => ({
      ...s,
      headerEven: s.headerOdd,
      footerEven: s.footerOdd,
    });
    expect(shown(merged, 0)).toEqual(mirror(host));
    expect(shown(merged, 1)).toEqual(shown(open(MAILMERGE), 0));
    expect(shown(merged, 2)).toEqual(mirror(host));
    expect(shown(merged, 3)).toEqual(mirror(shown(original, 1)));
  });
});

// ---- notes ----------------------------------------------------------------------------------

function notesOf(doc: WordDocument, kind: 'footnotes' | 'endnotes'): XElement[] {
  const part = doc.mainPart.relatedByType(kind)[0];
  return part === undefined ? [] : part.xml().root.elements(NS.w, kind.slice(0, -1));
}

describe('importNotes', () => {
  for (const [t, label] of [
    [NOTES, 'self'],
    [THREE, 'existing-parts'],
  ] as const) {
    it(`copies referenced footnotes and endnotes (${label})`, () => {
      const target = open(t);
      const before = {
        f: notesOf(target, 'footnotes').length,
        e: notesOf(target, 'endnotes').length,
      };
      const { ctx, blocks } = compose(target, open(NOTES));
      expect(ctx.footnoteIdMap.size).toBe(3);
      expect(ctx.endnoteIdMap.size).toBe(3);
      const merged = roundTrip(target, `notes-${label}`);
      for (const kind of ['footnotes', 'endnotes'] as const) {
        const notes = notesOf(merged, kind);
        expect(notes.length).toBe((kind === 'footnotes' ? before.f : before.e) + 3);
        const ids = notes.map((n) => n.attrNS(NS.w, 'id'));
        expect(new Set(ids).size).toBe(ids.length);
        const separators = notes.filter((n) => n.attrNS(NS.w, 'type') !== undefined);
        expect(separators.length).toBe(2);
      }
      // References in the imported blocks point at copied notes with the source's text.
      const texts = new Map(merged.footnotes().map((s) => [s.id, s.text()]));
      const refs = blocks.flatMap((b) => b.descendants(NS.w, 'footnoteReference'));
      expect(refs.length).toBe(3);
      expect(refs.map((r) => texts.get(r.attrNS(NS.w, 'id'))?.trim())).toEqual([
        'Footnote 1',
        'Footnote 2',
        'Footnote 3',
      ]);
    });
  }

  it('creates the notes parts (with separators) in a target without them', () => {
    const target = createDocx();
    compose(target, open(NOTES));
    const merged = roundTrip(target, 'notes-created');
    for (const kind of ['footnotes', 'endnotes'] as const) {
      const notes = notesOf(merged, kind);
      expect(notes.map((n) => n.attrNS(NS.w, 'type') ?? 'note')).toEqual([
        'separator',
        'continuationSeparator',
        'note',
        'note',
        'note',
      ]);
    }
    expect(merged.footnotes().length).toBe(3);
    expect(merged.endnotes().length).toBe(3);
  });

  it('copies notes with fields and paragraph ids into a notes part lacking w14', () => {
    const target = open(NOTES);
    compose(target, open(MAILMERGE));
    const merged = roundTrip(target, 'notes-w14');
    const added = merged.footnotes().at(-1)!;
    expect(added.text()).toContain('Merge');
  });

  it('removes references to notes the source does not have', () => {
    const source = open(NOTES);
    const ref = bodyOf(source).descendants(NS.w, 'footnoteReference')[0]!;
    ref.setAttrNS(NS.w, 'id', '77');
    const target = createDocx();
    const { ctx, blocks } = compose(target, readDocx(writeDocx(source)));
    expect(warningsOf(ctx)).toContain('NOTE_MISSING');
    expect(blocks.flatMap((b) => b.descendants(NS.w, 'footnoteReference')).length).toBe(2);
    roundTrip(target, 'notes-missing');
  });
});

// ---- comments -------------------------------------------------------------------------------

function partRootOf(doc: WordDocument, kind: string): XElement | undefined {
  return doc.mainPart.relatedByType(kind)[0]?.xml().root;
}

/**
 * Comment ids unique and referenced markers resolvable, paragraph ids unique, extended-part keys
 * unique. Returns the number of commentsExtended/commentsIds entries matching no comment (stale
 * entries some producers leave behind; never created by an import).
 */
function expectConsistentComments(doc: WordDocument): number {
  let stale = 0;
  const comments = partRootOf(doc, 'comments')!.elements(NS.w, 'comment');
  const ids = comments.map((c) => c.attrNS(NS.w, 'id')!);
  expect(new Set(ids).size).toBe(ids.length);
  const body = bodyOf(doc);
  for (const local of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
    for (const m of body.descendants(NS.w, local)) expect(ids).toContain(m.attrNS(NS.w, 'id'));
  }
  const lastParaIds = new Set(
    comments.map((c) => c.descendants(NS.w, 'p').at(-1)?.attrNS(NS.w14, 'paraId')),
  );
  const allParaIds = [
    ...bodyOf(doc).descendants(NS.w, 'p'),
    ...comments.flatMap((c) => c.descendants(NS.w, 'p')),
  ]
    .map((p) => p.attrNS(NS.w14, 'paraId'))
    .filter((id) => id !== undefined);
  expect(new Set(allParaIds).size).toBe(allParaIds.length);
  const ex = partRootOf(doc, 'commentsExtended');
  if (ex !== undefined) {
    const keys = ex.elements(NS.w15, 'commentEx').map((e) => e.attrNS(NS.w15, 'paraId'));
    expect(new Set(keys).size).toBe(keys.length);
    stale += keys.filter((k) => !lastParaIds.has(k)).length;
  }
  const cid = partRootOf(doc, 'commentsIds');
  if (cid !== undefined) {
    const entries = cid.elements(NS.w16cid, 'commentId');
    const durable = entries.map((e) => e.attrNS(NS.w16cid, 'durableId'));
    expect(new Set(durable).size).toBe(durable.length);
    stale += entries.filter((e) => !lastParaIds.has(e.attrNS(NS.w16cid, 'paraId'))).length;
  }
  return stale;
}

/** Stale extended entries of a fixture as it is (see expectConsistentComments). */
function staleOf(name: string): number {
  return name === '' ? 0 : expectConsistentComments(open(name));
}

describe('importComments', () => {
  it('remaps the comment parts together (merged with itself)', () => {
    const target = open(COMMENTS);
    const { ctx } = compose(target, open(COMMENTS));
    expect([...ctx.commentIdMap.keys()].sort()).toEqual([0, 1, 2, 4]);
    const merged = roundTrip(target, 'comments-self');
    // 4 comments; the extended parts also hold one stale entry, which is not copied.
    expect(merged.comments().length).toBe(4 + 4);
    expect(partRootOf(merged, 'commentsExtended')!.elements().length).toBe(5 + 4);
    expect(partRootOf(merged, 'commentsIds')!.elements().length).toBe(5 + 4);
    expect(expectConsistentComments(merged)).toBe(staleOf(COMMENTS));
    // The copied comment kept its hyperlink and image (relationships of comments.xml).
    const texts = merged.comments().map((c) => c.text());
    expect(texts.filter((t) => t.includes('https://google.com')).length).toBe(2);
  });

  it('creates all comment parts in a target without comments', () => {
    const target = createDocx();
    compose(target, open(COMMENTS));
    const merged = roundTrip(target, 'comments-created');
    expect(merged.comments().length).toBe(4);
    expect(partRootOf(merged, 'commentsExtended')!.elements().length).toBe(4);
    expect(partRootOf(merged, 'commentsIds')!.elements().length).toBe(4);
    expect(expectConsistentComments(merged)).toBe(0);
  });

  it('works with the Strict comment fixture and LibreOffice comments', () => {
    for (const s of [
      'open-xml-sdk/strict-CommentRangeStart.docx',
      'apache-poi/libreoffice-comment.docx',
    ]) {
      const target = open(COMMENTS);
      compose(target, open(s));
      const merged = roundTrip(target, `comments-${s.split('/')[1]!.replace('.docx', '')}`);
      expect(expectConsistentComments(merged)).toBe(staleOf(COMMENTS));
      expect(merged.comments().length).toBe(4 + open(s).comments().length);
    }
  });

  it("comments: 'drop' removes the markers and copies nothing", () => {
    const target = createDocx();
    const { blocks } = compose(target, open(COMMENTS), { comments: 'drop' });
    for (const local of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
      expect(blocks.flatMap((b) => b.descendants(NS.w, local))).toEqual([]);
    }
    const merged = roundTrip(target, 'comments-drop');
    expect(merged.comments()).toEqual([]);
  });

  it('dedupes people.xml authors and remaps commentsExtensible durable ids', () => {
    const make = (): WordDocument => {
      const doc = open(COMMENTS);
      const pkg = doc.package;
      const cid = partRootOf(doc, 'commentsIds')!;
      const durables = cid
        .elements(NS.w16cid, 'commentId')
        .map((e) => e.attrNS(NS.w16cid, 'durableId'));
      const W16CEX = NS.w16cex;
      const ext = `<w16cex:commentsExtensible xmlns:w16cex="${W16CEX}">${durables
        .map(
          (d) =>
            `<w16cex:commentExtensible w16cex:durableId="${d}" w16cex:dateUtc="2025-06-07T11:20:00Z"/>`,
        )
        .join('')}</w16cex:commentsExtensible>`;
      const people = `<w15:people xmlns:w15="${NS.w15}"><w15:person w15:author="Steve Canny"><w15:presenceInfo w15:providerId="None" w15:userId="Steve Canny"/></w15:person></w15:people>`;
      const enc = new TextEncoder();
      pkg.addPart(
        'word/commentsExtensible.xml',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml',
        enc.encode(ext),
      );
      pkg.addPart(
        'word/people.xml',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml',
        enc.encode(people),
      );
      doc.mainPart.rels.add(
        'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
        'commentsExtensible.xml',
      );
      doc.mainPart.rels.add(
        'http://schemas.microsoft.com/office/2011/relationships/people',
        'people.xml',
      );
      return readDocx(writeDocx(doc));
    };
    const target = make();
    compose(target, make());
    const merged = roundTrip(target, 'comments-people');
    expect(partRootOf(merged, 'people')!.elements(NS.w15, 'person').length).toBe(1);
    const ext = partRootOf(merged, 'commentsExtensible')!.elements(NS.w16cex, 'commentExtensible');
    const ids = partRootOf(merged, 'commentsIds')!.elements(NS.w16cid, 'commentId');
    expect(ext.length).toBe(5 + 4);
    const durable = new Set(ids.map((e) => e.attrNS(NS.w16cid, 'durableId')));
    for (const e of ext) expect(durable.has(e.attrNS(NS.w16cex, 'durableId'))).toBe(true);
    expect(new Set(ext.map((e) => e.attrNS(NS.w16cex, 'durableId'))).size).toBe(5 + 4);
  });
});

// ---- bookmarks ------------------------------------------------------------------------------

function bookmarkNames(root: XElement): string[] {
  return root.descendants(NS.w, 'bookmarkStart').map((b) => b.attrNS(NS.w, 'name') ?? '');
}

describe('importBookmarks', () => {
  it('renames conflicting bookmarks and rewrites PAGEREF fields and anchors (TOC merged with itself)', () => {
    const target = open(TOC);
    const { ctx, blocks } = compose(target, open(TOC));
    expect(ctx.bookmarkRenames.size).toBe(6);
    for (const [from, to] of ctx.bookmarkRenames) expect(to).toBe(`${from}_1`);
    const merged = roundTrip(target, 'bookmarks-toc');
    const body = bodyOf(merged);
    const names = bookmarkNames(body);
    expect(new Set(names.map((n) => n.toLowerCase())).size).toBe(names.length);
    const ids = body.descendants(NS.w, 'bookmarkStart').map((b) => b.attrNS(NS.w, 'id'));
    expect(new Set(ids).size).toBe(ids.length);
    // Imported content refers to the renamed bookmarks only.
    const instr = blocks
      .flatMap((b) => b.descendants(NS.w, 'instrText'))
      .map((e) => e.text)
      .join('|');
    const anchors = blocks
      .flatMap((b) => b.descendants(NS.w, 'hyperlink'))
      .map((h) => h.attrNS(NS.w, 'anchor'));
    for (const [, to] of ctx.bookmarkRenames) {
      expect(instr).toContain(`PAGEREF ${to} `);
      expect(anchors).toContain(to);
    }
    // The target's own content still refers to the originals.
    const original = body
      .elements()
      .slice(0, 5)
      .flatMap((b) => b.descendants(NS.w, 'hyperlink'));
    expect(original.every((h) => !(h.attrNS(NS.w, 'anchor') ?? '').endsWith('_1'))).toBe(true);
  });

  it('drops the source _GoBack and keeps the target one', () => {
    const target = open(COMMENTS);
    const { blocks } = compose(target, open(COMMENTS));
    expect(blocks.flatMap((b) => bookmarkNames(b))).not.toContain('_GoBack');
    expect(bookmarkNames(bodyOf(target)).filter((n) => n === '_GoBack').length).toBe(1);
    roundTrip(target, 'bookmarks-goback');
  });

  it("bookmarks: 'drop' removes conflicting bookmarks", () => {
    const target = open(TOC);
    const { ctx, blocks } = compose(target, open(TOC), { bookmarks: 'drop' });
    expect(ctx.bookmarkRenames.size).toBe(0);
    expect(blocks.flatMap((b) => b.descendants(NS.w, 'bookmarkStart'))).toEqual([]);
    expect(blocks.flatMap((b) => b.descendants(NS.w, 'bookmarkEnd'))).toEqual([]);
    roundTrip(target, 'bookmarks-drop');
  });

  it('rewrites instructions split over runs and paragraphs, fldSimple and HYPERLINK \\l', () => {
    const W = `xmlns:w="${NS.w}"`;
    const body = (inner: string) => parseXml(`<w:body ${W}>${inner}</w:body>`).root;
    const target = createDocx();
    const source = createDocx();
    const sb = bodyOf(source);
    sb.clear();
    const xml = body(
      '<w:p><w:bookmarkStart w:id="1" w:name="Intro"/><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE</w:instrText></w:r>' +
        '<w:r><w:instrText>REF In</w:instrText></w:r></w:p><w:p><w:r><w:instrText>tro \\h </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
        '<w:p><w:fldSimple w:instr=" REF Intro \\h "><w:r><w:t>x</w:t></w:r></w:fldSimple>' +
        '<w:hyperlink w:anchor="Intro"><w:r><w:t>go</w:t></w:r></w:hyperlink>' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK \\l "Intro" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    );
    for (const k of xml.elements()) sb.append(k.clone());
    // the target already has an "intro" bookmark (case-insensitive conflict)
    const tp = bodyOf(target).element(NS.w, 'p')!;
    tp.append(parseXml(`<w:bookmarkStart ${W} w:id="0" w:name="intro"/>`).root.clone());
    tp.append(parseXml(`<w:bookmarkEnd ${W} w:id="0"/>`).root.clone());
    const { ctx, blocks } = compose(target, source);
    expect(ctx.bookmarkRenames.get('Intro')).toBe('Intro_1');
    const instr = blocks[1]!.descendants(NS.w, 'instrText').map((e) => e.text);
    expect(
      instr.join('') +
        blocks[2]!
          .descendants(NS.w, 'instrText')
          .map((e) => e.text)
          .join(''),
    ).toBe(' PAGEREF Intro_1 \\h ');
    expect(blocks[3]!.element(NS.w, 'fldSimple')!.attrNS(NS.w, 'instr')).toBe(' REF Intro_1 \\h ');
    expect(blocks[3]!.element(NS.w, 'hyperlink')!.attrNS(NS.w, 'anchor')).toBe('Intro_1');
    expect(blocks[3]!.descendants(NS.w, 'instrText')[0]!.text).toBe(' HYPERLINK \\l "Intro_1" ');
    roundTrip(target, 'bookmarks-fields');
  });

  it('rewriteInstruction', () => {
    const m = new Map([
      ['a', 'a_1'],
      ['long', 'long_2'],
    ]);
    expect(rewriteInstruction(' REF a \\h ', m)).toBe(' REF a_1 \\h ');
    expect(rewriteInstruction('NOTEREF "A" \\f', m)).toBe('NOTEREF "a_1" \\f');
    expect(rewriteInstruction('PAGEREF b', m)).toBe('PAGEREF b');
    expect(rewriteInstruction(' a ', m)).toBe(' a_1 ');
    expect(rewriteInstruction(' PAGE ', new Map([['page', 'x']]))).toBe(' PAGE ');
    expect(rewriteInstruction('HYPERLINK "http://x" \\l long \\o "t"', m)).toBe(
      'HYPERLINK "http://x" \\l long_2 \\o "t"',
    );
    expect(rewriteInstruction('', m)).toBe('');
  });
});

// ---- settings, fonts, theme -----------------------------------------------------------------

describe('mergeSettings', () => {
  it('merges docVars and rsids, keeps the target compat settings', () => {
    const source = createDocx();
    const settings = source.mainPart.relatedByType('settings')[0]!.xml().root;
    const vars = parseXml(
      `<w:docVars xmlns:w="${NS.w}"><w:docVar w:name="a" w:val="1"/><w:docVar w:name="b" w:val="2"/></w:docVars>`,
    ).root.clone();
    settings.insertBefore(vars, settings.element(NS.w, 'decimalSymbol') ?? null);
    const target = createDocx();
    const tset = target.mainPart.relatedByType('settings')[0]!.xml().root;
    tset.insertBefore(
      parseXml(
        `<w:docVars xmlns:w="${NS.w}"><w:docVar w:name="b" w:val="target"/></w:docVars>`,
      ).root.clone(),
      tset.element(NS.w, 'decimalSymbol') ?? null,
    );
    const compat = tset.element(NS.w, 'compat')!.clone();
    mergeSettings(makeCtx(target, source));
    const out = tset.element(NS.w, 'docVars')!.elements(NS.w, 'docVar');
    expect(out.map((v) => `${v.attrNS(NS.w, 'name')}=${v.attrNS(NS.w, 'val')}`)).toEqual([
      'b=target',
      'a=1',
    ]);
    expect(tset.element(NS.w, 'compat')!.elements().length).toBe(compat.elements().length);

    const t2 = open(THREE);
    const before = t2.mainPart
      .relatedByType('settings')[0]!
      .xml()
      .root.element(NS.w, 'rsids')!
      .elements().length;
    mergeSettings(makeCtx(t2, open(MAILMERGE)));
    const after = t2.mainPart
      .relatedByType('settings')[0]!
      .xml()
      .root.element(NS.w, 'rsids')!
      .elements().length;
    expect(after).toBeGreaterThan(before);
    roundTrip(t2, 'settings-rsids');
  });

  it('docVars land at their schema position in a target without them', () => {
    const source = createDocx();
    const settings = source.mainPart.relatedByType('settings')[0]!.xml().root;
    settings.insertBefore(
      parseXml(
        `<w:docVars xmlns:w="${NS.w}"><w:docVar w:name="a" w:val="1"/></w:docVars>`,
      ).root.clone(),
      settings.element(NS.w, 'decimalSymbol') ?? null,
    );
    const target = createDocx();
    mergeSettings(makeCtx(target, source));
    const locals = target.mainPart
      .relatedByType('settings')[0]!
      .xml()
      .root.elements()
      .map((e) => e.local);
    expect(locals.indexOf('docVars')).toBe(locals.indexOf('compat') + 1);
  });

  it('adds missing fonts, copying embedded font parts with their keys', () => {
    const source = open(THREE);
    const fontPart = source.mainPart.relatedByType('fontTable')[0]!;
    const bytes = new Uint8Array(64).map((_, i) => i * 7);
    source.package.addPart(
      'word/fonts/font1.odttf',
      'application/vnd.openxmlformats-officedocument.obfuscatedFont',
      bytes,
    );
    const rid = fontPart.rels.add(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font',
      'fonts/font1.odttf',
    ).id;
    const root = fontPart.xml().root;
    root.append(
      parseXml(
        `<w:font xmlns:w="${NS.w}" xmlns:r="${NS.r}" w:name="Fancy Embedded"><w:embedRegular r:id="${rid}" w:fontKey="{00000000-1111-2222-3333-444444444444}"/></w:font>`,
      ).root.clone(),
    );
    const target = createDocx();
    const srcFonts = root.elements(NS.w, 'font').map((f) => f.attrNS(NS.w, 'name'));
    mergeSettings(makeCtx(target, source));
    const merged = roundTrip(target, 'settings-fonts');
    const table = merged.mainPart.relatedByType('fontTable')[0]!;
    const names = table
      .xml()
      .root.elements(NS.w, 'font')
      .map((f) => f.attrNS(NS.w, 'name'));
    for (const n of srcFonts) expect(names).toContain(n);
    expect(new Set(names).size).toBe(names.length);
    const fancy = table
      .xml()
      .root.elements(NS.w, 'font')
      .find((f) => f.attrNS(NS.w, 'name') === 'Fancy Embedded')!;
    const embed = fancy.element(NS.w, 'embedRegular')!;
    expect(embed.attrNS(NS.w, 'fontKey')).toBe('{00000000-1111-2222-3333-444444444444}');
    expect(table.related(embed.attrNS(NS.r, 'id')!)!.bytes()).toEqual(bytes);
    const settings = merged.mainPart.relatedByType('settings')[0]!.xml().root;
    expect(settings.element(NS.w, 'embedTrueTypeFonts')).toBeDefined();
  });

  it("theme: 'source' replaces the theme, 'destination' keeps it", () => {
    const themeBytes = (d: WordDocument) => d.mainPart.relatedByType('theme')[0]!.bytes();
    const src = open(COMMENTS);
    const t1 = open(THREE);
    const before = themeBytes(t1).slice();
    mergeSettings(makeCtx(t1, src));
    expect(themeBytes(t1)).toEqual(before);
    const t2 = open(THREE);
    mergeSettings(makeCtx(t2, src, { theme: 'source' }));
    const merged = roundTrip(t2, 'settings-theme');
    expect(themeBytes(merged)).toEqual(themeBytes(open(COMMENTS)));
    expect(merged.mainPart.relatedByType('theme').length).toBe(1);
  });
});

// ---- the real orchestrator ------------------------------------------------------------------

describe('the real pipeline', () => {
  const pairs: [string, string][] = [
    [IMAGE, THREE],
    [THREE, MAILMERGE],
    [MAILMERGE, IMAGE],
    [THREE, NOTES],
    [COMMENTS, COMMENTS],
    [NOTES, NOTES],
    [TOC, TOC],
    [MAILMERGE, MAILMERGE],
  ];
  for (const [t, s] of pairs) {
    it(`${t} + ${s}`, () => {
      const target = open(t);
      const result = appendDocument(target, open(s));
      expect(result.sections).toBe(open(s).sections().length);
      const label = `append-${t.split('/')[1]}-${s.split('/')[1]}`.replace(/\.docx/g, '');
      const merged = roundTrip(target, label);
      expectSameAppearance(merged, open(t), open(s));
      if (merged.comments().length > 0) {
        expect(expectConsistentComments(merged)).toBe(t === COMMENTS ? staleOf(COMMENTS) : 0);
      }
      const names = bookmarkNames(bodyOf(merged));
      expect(new Set(names.map((n) => n.toLowerCase())).size).toBe(names.length);
    });
  }

  it('insertDocument after a paragraph splits the host section', () => {
    const target = open(THREE);
    const at = target.sections()[0]!.paragraphs()[0]!;
    const result = insertDocument(target, { paragraph: at, mode: 'after' }, open(MAILMERGE));
    expect(result.sections).toBe(1);
    const merged = roundTrip(target, 'insert-after-paragraph');
    const original = open(THREE);
    expect(merged.sections().length).toBe(original.sections().length + 2);
    const mirror = (s: Record<string, string>) => ({
      ...s,
      headerEven: s.headerOdd,
      footerEven: s.footerOdd,
    });
    expect(shown(merged, 0)).toEqual(mirror(shown(original, 0)));
    expect(shown(merged, 1)).toEqual(shown(open(MAILMERGE), 0));
    expect(shown(merged, 2)).toEqual(mirror(shown(original, 0)));
    for (let i = 1; i < original.sections().length; i++) {
      expect(shown(merged, i + 2)).toEqual(mirror(shown(original, i)));
    }
  });

  it('insertDocument with keepSourceSections false adopts the host section', () => {
    const target = open(THREE);
    const at = target.sections()[1]!.paragraphs()[0]!;
    const warnings: string[] = [];
    const result = insertDocument(target, { paragraph: at, mode: 'before' }, open(MAILMERGE), {
      keepSourceSections: false,
      onWarning: (w) => warnings.push(w.code),
    });
    expect(result.sections).toBe(0);
    expect(warnings).toContain('SECTIONS_DROPPED');
    const merged = roundTrip(target, 'insert-no-sections');
    expect(merged.sections().length).toBe(3);
    for (let i = 0; i < 3; i++) expect(shown(merged, i)).toEqual(shown(open(THREE), i));
  });

  it('composeDocuments keeps every input section looking as it did', () => {
    const inputs = [THREE, MAILMERGE, IMAGE, NOTES, COMMENTS].map(open);
    const out = composeDocuments(inputs, { validate: true });
    const merged = roundTrip(out, 'compose-five');
    let at = 0;
    for (const name of [THREE, MAILMERGE, IMAGE, NOTES, COMMENTS]) {
      const doc = open(name);
      const eo = evenAndOdd(merged);
      for (let i = 0; i < doc.sections().length; i++) {
        const want = shown(doc, i);
        // Documents without even/odd headers show their default header on even pages.
        if (eo && !evenAndOdd(doc))
          Object.assign(want, { headerEven: want.headerOdd, footerEven: want.footerOdd });
        expect(shown(merged, at + i), `${name} section ${i}`).toEqual(want);
      }
      at += doc.sections().length;
    }
    expect(merged.sections().length).toBe(at);
    expect(merged.footnotes().length).toBe(1 + 3);
    expect(merged.comments().length).toBe(4);
  });

  it("theme: 'source' and font merging through appendDocument", () => {
    const target = open(THREE);
    appendDocument(target, open(COMMENTS), { theme: 'source' });
    const merged = roundTrip(target, 'append-theme-source');
    const theme = (d: WordDocument) => d.mainPart.relatedByType('theme')[0]!.bytes();
    expect(theme(merged)).toEqual(theme(open(COMMENTS)));
    const fonts = (d: WordDocument) =>
      d.mainPart
        .relatedByType('fontTable')[0]!
        .xml()
        .root.elements(NS.w, 'font')
        .map((f) => f.attrNS(NS.w, 'name'));
    for (const f of fonts(open(COMMENTS))) expect(fonts(merged)).toContain(f);
    expect(new Set(fonts(merged)).size).toBe(fonts(merged).length);
  });
});
