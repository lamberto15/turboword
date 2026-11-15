// Table editing: newTable, row/column insertion and deletion, merge/unmerge, grid widths, row
// and cell setters. Every operation is followed by an invariant check (cells end with a
// paragraph, rows keep a cell, spans fill the grid, tblGrid has the right column count, vMerge
// chains are valid, range markers stay paired, paragraph ids stay unique), including seeded
// random operation sequences on real fixtures and on a synthetic table with every feature.
// Outputs go to out/edit-tables/ (check them with scripts/libreoffice-check.mjs).

import { mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { UsageError } from '../src/core/errors.js';
import {
  createDocx,
  readDocx,
  type ValidationIssue,
  validatePackage,
  type WordDocument,
  writeDocx,
} from '../src/index.js';
import { DocumentIds } from '../src/wml/ids.js';
import type { Story } from '../src/wml/story.js';
import { newTable, Table, type TableCell } from '../src/wml/table.js';
import { ELEMENT_NODE, type XElement } from '../src/xml/dom.js';
import { NS } from '../src/xml/ns.js';
import { parseXml } from '../src/xml/parse.js';
import { serializeNode } from '../src/xml/serialize.js';
import { fixture, strictFixtures } from './helpers/fixtures.js';

const OUT = fileURLToPath(new URL('../out/edit-tables/', import.meta.url));
mkdirSync(OUT, { recursive: true });

// ---- building documents ------------------------------------------------------------------

const enc = new TextEncoder();
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const DECLS =
  `xmlns:w="${NS.w}" xmlns:w14="${W14}" xmlns:mc="${NS.mc}" xmlns:r="${NS.r}" ` +
  `xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" mc:Ignorable="w14"`;

function docx(body: string, decls = DECLS, w = 'w'): Uint8Array {
  return zipSync({
    '[Content_Types].xml': enc.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': enc.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>',
    ),
    'word/document.xml': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${w}:document ${decls}><${w}:body>${body}<${w}:sectPr/></${w}:body></${w}:document>`,
    ),
  });
}

let nextPid = 0x1000;
const pid = (): string => (nextPid++).toString(16).toUpperCase().padStart(8, '0');
const para = (text: string, rPr = ''): string =>
  `<w:p w14:paraId="${pid()}" w14:textId="77777777">` +
  (text === '' ? '' : `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t>${text}</w:t></w:r>`) +
  '</w:p>';
const tcw = (w: number): string => `<w:tcW w:w="${w}" w:type="dxa"/>`;
const tc = (text: string, tcPr = '', w = 1000): string =>
  `<w:tc><w:tcPr>${tcw(w)}${tcPr}</w:tcPr>${para(text)}</w:tc>`;
const tr = (cells: string, trPr = ''): string =>
  `<w:tr w14:paraId="${pid()}" w14:textId="77777777">${trPr === '' ? '' : `<w:trPr>${trPr}</w:trPr>`}${cells}</w:tr>`;
const grid = (n: number, w = 1000): string =>
  `<w:tblGrid>${`<w:gridCol w:w="${w}"/>`.repeat(n)}</w:tblGrid>`;
const tbl = (inner: string, tblPr = '<w:tblW w:w="0" w:type="auto"/>'): string =>
  `<w:tbl><w:tblPr>${tblPr}</w:tblPr>${inner}</w:tbl>`;
const span = (n: number): string => `<w:gridSpan w:val="${n}"/>`;
const VR = '<w:vMerge w:val="restart"/>';
const VC = '<w:vMerge/>';

function load(name: string): WordDocument {
  return readDocx(fixture(name).bytes);
}

function single(body: string): { doc: WordDocument; t: Table } {
  const doc = readDocx(docx(body));
  return { doc, t: doc.body.tables()[0]! };
}

/** The grid as text: covering cell text per slot ('^' = not the origin, '.' = no cell). */
function show(t: Table): string[] {
  const g = t.grid();
  const out: string[] = [];
  for (let r = 0; r < g.rows; r++) {
    const line: string[] = [];
    for (let c = 0; c < g.cols; c++) {
      const s = g.at(r, c);
      line.push(s === undefined ? '.' : `${s.cell.text()}${s.origin ? '' : '^'}`);
    }
    out.push(line.join('|'));
  }
  return out;
}

const xml = (el: XElement): string => serializeNode(el);

// ---- the invariant checker ---------------------------------------------------------------

const BLOCKS = new Set(['p', 'tbl', 'sdt', 'customXml', 'altChunk']);

function lastBlock(container: XElement): XElement | undefined {
  let c = container;
  for (;;) {
    const kids = c.children;
    let found: XElement | undefined;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === ELEMENT_NODE && k.uri === NS.w && BLOCKS.has(k.local)) {
        found = k;
        break;
      }
    }
    if (found === undefined) return undefined;
    if (found.local === 'sdt') {
      const content = found.element(NS.w, 'sdtContent');
      if (content === undefined) return undefined;
      c = content;
    } else if (found.local === 'customXml') c = found;
    else return found;
  }
}

const intAttr = (el: XElement | undefined, local: string, fallback: number): number => {
  const v = el?.attrNS(NS.w, local);
  return v === undefined ? fallback : Number.parseInt(v, 10);
};

/** Structural invariants of one table (see the file header). */
function tableViolations(t: Table): string[] {
  const out: string[] = [];
  const el = t.element;
  const gridEl = el.element(NS.w, 'tblGrid');
  if (gridEl === undefined) out.push('no tblGrid');
  const width = t.gridColumnCount;
  const rows = t.rows();
  if (rows.length === 0) out.push('table without rows');
  let prev: { col: number; span: number; v: string | undefined }[] = [];
  rows.forEach((row, r) => {
    const cells = row.cells();
    if (cells.length === 0) out.push(`row ${r} has no cell`);
    let col = row.gridBefore;
    const line: { col: number; span: number; v: string | undefined }[] = [];
    for (const cell of cells) {
      const last = lastBlock(cell.element);
      if (last?.is(NS.w, 'p') !== true)
        out.push(`cell at row ${r} col ${col} does not end with w:p`);
      const v = cell.vMerge;
      if (v === 'continue') {
        const up = prev.find((p) => p.col === col);
        if (up === undefined || up.v === undefined || up.span !== cell.gridSpan) {
          out.push(`row ${r} col ${col}: vMerge continue without a matching merge cell above`);
        }
      }
      line.push({ col, span: cell.gridSpan, v });
      col += cell.gridSpan;
    }
    if (col + row.gridAfter !== width) {
      out.push(
        `row ${r}: before ${row.gridBefore} + spans + after ${row.gridAfter} = ${col + row.gridAfter} != grid ${width}`,
      );
    }
    prev = line;
  });
  if (t.grid().cols !== width) out.push(`grid() has ${t.grid().cols} columns, tblGrid ${width}`);
  for (const g of gridEl?.elements(NS.w, 'gridCol') ?? []) {
    if (!(intAttr(g, 'w', -1) >= 0)) out.push('gridCol without a width');
  }
  return out;
}

/** Every table of every story, plus document-wide id and range-marker checks. */
function docViolations(doc: WordDocument): string[] {
  const out: string[] = [];
  const paraIds = new Map<string, number>();
  const parts = new Set<XElement>();
  for (const story of doc.stories()) {
    for (const t of story.element.descendants(NS.w, 'tbl')) {
      for (const v of tableViolations(new Table(t, story))) out.push(`${story.kind}: ${v}`);
    }
    parts.add(story.part.xml().root);
  }
  for (const root of parts) {
    for (const el of root.descendants()) {
      const id = el.attrNS(NS.w14, 'paraId');
      if (id !== undefined) paraIds.set(id, (paraIds.get(id) ?? 0) + 1);
    }
    for (const [start, end] of [
      ['bookmarkStart', 'bookmarkEnd'],
      ['commentRangeStart', 'commentRangeEnd'],
    ] as const) {
      const ids = (local: string): string[] =>
        root
          .descendants(NS.w, local)
          .map((e) => e.attrNS(NS.w, 'id') ?? '')
          .sort();
      const s = ids(start);
      const e = ids(end);
      if (s.join() !== e.join()) out.push(`${start} ids [${s}] vs ${end} ids [${e}]`);
    }
  }
  for (const [id, n] of paraIds) if (n > 1) out.push(`w14:paraId ${id} used ${n} times`);
  const docPr = doc.mainPart
    .xml()
    .root.descendants(NS.wp, 'docPr')
    .map((d) => d.attr('id'));
  if (new Set(docPr).size !== docPr.length) out.push(`duplicate docPr ids ${docPr}`);
  return out;
}

function expectValid(doc: WordDocument): void {
  expect(docViolations(doc)).toEqual([]);
}

const errorCodes = (issues: readonly ValidationIssue[]): string[] =>
  [...new Set(issues.filter((i) => i.severity === 'error').map((i) => i.code))].sort();

/** Saves, re-reads, re-checks and compares validation errors with the input's. */
function saveAndCheck(doc: WordDocument, name: string, input?: Uint8Array): Uint8Array {
  const bytes = writeDocx(doc);
  writeFileSync(`${OUT}${name}.docx`, bytes);
  const again = readDocx(bytes);
  expectValid(again);
  const before = input === undefined ? [] : errorCodes(validatePackage(readDocx(input)));
  const after = errorCodes(validatePackage(again));
  expect(after.filter((c) => !before.includes(c))).toEqual([]);
  return bytes;
}

// ---- newTable ----------------------------------------------------------------------------

describe('newTable', () => {
  it('builds rows × cols with an even grid, borders, tblW and tblLook', () => {
    const doc = createDocx();
    const el = newTable(doc.body, 2, 3, { width: 9242 });
    const t = new Table(el, doc.body);
    expect(t.columnWidths()).toEqual([3080, 3081, 3081]);
    expect(t.rows()).toHaveLength(2);
    expect(
      t
        .rows()[0]!
        .cells()
        .map((c) => c.width),
    ).toEqual([3080, 3081, 3081]);
    const s = xml(el);
    expect(s).toContain('<w:tblW w:w="9242" w:type="dxa"/>');
    expect(s).toContain(
      '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>',
    );
    expect(s).toContain('<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    expect(s).toContain('<w:tblLook w:val="04A0" w:firstRow="1"');
    expect(s).not.toContain('tblStyle');
    expect(tableViolations(t)).toEqual([]);
  });

  it('uses a table style instead of borders, and explicit widths', () => {
    const doc = createDocx();
    const el = newTable(doc.body, 1, 2, { widths: [2000, 3000], style: 'TableGrid' });
    const s = xml(el);
    expect(s).toMatch(/^<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"\/><w:tblW w:w="5000"/);
    expect(s).not.toContain('tblBorders');
    expect(new Table(el, doc.body).columnWidths()).toEqual([2000, 3000]);
    expect(new Table(el, doc.body).style).toBe('TableGrid');
  });

  it('defaults to 9360 twips and names elements with the story prefix', () => {
    const doc = readDocx(docx('<x:p/>', `xmlns:x="${NS.w}"`, 'x'));
    const el = newTable(doc.body, 1, 4);
    expect(el.name).toBe('x:tbl');
    expect(new Table(el, doc.body).columnWidths()).toEqual([2340, 2340, 2340, 2340]);
    expect(el.descendants().every((d) => d.name.startsWith('x:'))).toBe(true);
  });

  it('leaves out the transitional-only tblLook value in a Strict document', () => {
    const [strict] = strictFixtures(['python-docx/tbl-cell-access.docx']);
    const doc = readDocx(strict!.bytes);
    const el = newTable(doc.body, 1, 1);
    expect(xml(el)).toContain('<w:tblLook w:firstRow="1"');
    doc.body.element.firstElement()!.before(el);
    expectValid(doc);
  });

  it('rejects bad arguments', () => {
    const doc = createDocx();
    expect(() => newTable(doc.body, 0, 2)).toThrow(UsageError);
    expect(() => newTable(doc.body, 1, 0)).toThrow(UsageError);
    expect(() => newTable(doc.body, 1, 2, { widths: [1] })).toThrow(UsageError);
    expect(() => newTable(doc.body, 1, 2, { width: -5 })).toThrow(UsageError);
  });

  it('writes a valid document', () => {
    const doc = createDocx();
    const first = doc.body.element.element(NS.w, 'p')!;
    first.before(newTable(doc.body, 3, 4, { style: 'TableNormal' }));
    first.before(doc.body.element.ownerDocument()!.createElement(NS.w, 'p'));
    first.before(newTable(doc.body, 2, 2));
    const [a, b] = doc.body.tables();
    a!.cell(0, 0)!.setText('styled');
    b!.cell(1, 1)!.setText('bordered');
    expectValid(doc);
    expect(validatePackage(doc)).toEqual([]);
    saveAndCheck(doc, 'new-table');
  });
});

// ---- rows --------------------------------------------------------------------------------

describe('cloneRow', () => {
  it('inserts a deep copy after the source with fresh paragraph ids', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[0]!;
    const before = xml(t.rows()[1]!.element);
    const row = t.cloneRow(1);
    expect(show(t)).toEqual(['1|2|3', '4|5|6', '4|5|6', '7|8|9']);
    expect(row.index).toBe(2);
    expect(xml(t.rows()[1]!.element)).toBe(before);
    const ids = (s: string): string[] =>
      [...s.matchAll(/w14:paraId="([0-9A-F]+)"/g)].map((m) => m[1]!);
    const fresh = ids(xml(row.element));
    expect(fresh).toHaveLength(4); // the row and its three paragraphs
    for (const id of fresh) {
      expect(ids(before)).not.toContain(id);
      expect(Number.parseInt(id, 16)).toBeLessThan(0x80000000);
    }
    expectValid(doc);
  });

  it('places the copy after any row, or first with after: -1', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[0]!;
    t.cloneRow(2, { after: -1 });
    expect(show(t)).toEqual(['7|8|9', '1|2|3', '4|5|6', '7|8|9']);
    t.cloneRow(1, { after: t.rows()[3]! });
    expect(show(t)).toEqual(['7|8|9', '1|2|3', '4|5|6', '7|8|9', '1|2|3']);
    expect(() => t.cloneRow(9)).toThrow(UsageError);
    expect(() => t.cloneRow(0, { after: 5 })).toThrow(UsageError);
    const other = doc.body.tables()[1]!.rows()[0]!;
    expect(() => t.cloneRow(0, { after: other })).toThrow(UsageError);
    expectValid(doc);
  });

  it('leaves bookmark, comment and permission markers out of the copy', () => {
    const { doc, t } = single(
      tbl(
        grid(2) +
          tr(
            `<w:tc><w:tcPr>${tcw(1000)}</w:tcPr><w:p><w:bookmarkStart w:id="1" w:name="bm"/>` +
              '<w:commentRangeStart w:id="0"/><w:r><w:t>A</w:t></w:r><w:bookmarkEnd w:id="1"/>' +
              '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>' +
              `<w:permStart w:id="3" w:edGrp="everyone"/><w:permEnd w:id="3"/></w:p></w:tc>${tc('B')}`,
          ),
      ) + para(''),
    );
    const row = t.cloneRow(0);
    const s = xml(row.element);
    for (const m of ['bookmark', 'comment', 'perm']) expect(s).not.toContain(m);
    expect(row.cells()[0]!.text()).toBe('A');
    expect(xml(t.rows()[0]!.element)).toContain('bookmarkStart');
    expectValid(doc);
  });

  it('gives drawings, content controls and revisions fresh ids', () => {
    const drawing =
      '<w:r><w:drawing><wp:inline><wp:extent cx="10" cy="10"/><wp:docPr id="7" name="Pic 7"/>' +
      '<a:graphic><a:graphicData uri="x"/></a:graphic></wp:inline></w:drawing></w:r>';
    const { doc, t } = single(
      tbl(
        grid(2) +
          tr(
            `<w:tc><w:tcPr>${tcw(1000)}</w:tcPr><w:p>${drawing}<w:ins w:id="5" w:author="x"><w:r><w:t>i</w:t></w:r></w:ins></w:p></w:tc>` +
              `<w:sdt><w:sdtPr><w:id w:val="123"/></w:sdtPr><w:sdtContent>${tc('S')}</w:sdtContent></w:sdt>`,
          ),
      ) + para(''),
    );
    const row = t.cloneRow(0);
    const s = xml(row.element);
    expect(s).not.toContain('id="7"');
    expect(s).not.toContain('w:id="5"');
    expect(s).not.toContain('w:val="123"');
    expect(s).toContain('name="Pic 7"');
    expect(validatePackage(doc).filter((i) => i.code === 'DUPLICATE_DOCPR_ID')).toEqual([]);
    expectValid(doc);
  });

  it('duplicates the footnotes and endnotes a copied row references', () => {
    const input = fixture('open-xml-sdk/TableFntEdnt.docx').bytes;
    const doc = readDocx(input);
    const t = doc.body.tables()[0]!;
    const notes = (): number => doc.footnotes().length + doc.endnotes().length;
    const n0 = notes();
    const fn0 = doc.footnotes().map((f) => f.id);
    const row = t.cloneRow(0);
    expect(notes()).toBe(n0 + 2);
    const refs = row.element
      .descendants(NS.w, 'footnoteReference')
      .map((r) => r.attrNS(NS.w, 'id'));
    expect(refs).toHaveLength(1);
    expect(fn0).not.toContain(refs[0]);
    const copy = doc.footnotes().find((f) => f.id === refs[0])!;
    const orig = doc.footnotes().find((f) => f.id === '2')!;
    expect(copy.text()).toBe(orig.text());
    expect(row.element.descendants(NS.w, 'endnoteReference')[0]!.attrNS(NS.w, 'id')).not.toBe('2');
    expectValid(doc);
    saveAndCheck(doc, 'clone-row-notes', input);
  });

  it('keeps paragraph ids unique across the body and duplicated notes', () => {
    // TableFntEdnt with w14:paraId on every paragraph of the body and the notes.
    const parts = unzipSync(fixture('open-xml-sdk/TableFntEdnt.docx').bytes);
    let n = 0x100;
    for (const name of ['word/document.xml', 'word/footnotes.xml', 'word/endnotes.xml']) {
      const text = strFromU8(parts[name]!)
        .replace(/<w:(document|footnotes|endnotes) /, `<w:$1 xmlns:w14="${W14}" `)
        .replace(
          /<w:p( |>)/g,
          (_m, c: string) =>
            `<w:p w14:paraId="${(n++).toString(16).padStart(8, '0')}"${c === '>' ? '>' : ' '}`,
        );
      parts[name] = strToU8(text);
    }
    const input = zipSync(parts);
    const doc = readDocx(input);
    expectValid(doc);
    const t = doc.body.tables()[0]!;
    t.cloneRow(0);
    t.cloneRow(1);
    t.cloneRow(0, { after: 3 });
    expect(doc.footnotes()).toHaveLength(3 + 3);
    expectValid(doc);
    saveAndCheck(doc, 'clone-row-note-ids', input);
  });

  it('works on a table with no document behind it (ids from the table itself)', () => {
    const el = parseXml(`<w:tbl ${DECLS}>${grid(1)}${tr(tc('a'))}</w:tbl>`).root;
    const t = new Table(el, {} as unknown as Story);
    t.cloneRow(0);
    t.cloneRow(1, { after: -1 });
    const ids = el.descendants().flatMap((e) => e.attrNS(NS.w14, 'paraId') ?? []);
    expect(ids).toHaveLength(6);
    expect(new Set(ids).size).toBe(6);
    expect(tableViolations(t)).toEqual([]);
  });

  it('continues a vertical merge when the copy lands inside it, else makes plain cells', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[3]!;
    expect(show(t)).toEqual(['1|2|3', '4|4^|6', '4^|4^|9']);
    t.cloneRow(1); // the restart row, copied inside the merge
    expect(show(t)).toEqual(['1|2|3', '4|4^|6', '4^|4^|6', '4^|4^|9']);
    expect(t.rows()[2]!.cells()[0]!.vMerge).toBe('continue');
    expect(t.rows()[2]!.cells()[0]!.paragraphs()).toHaveLength(1);
    t.cloneRow(3, { after: 3 }); // a continuation row copied below the merge
    expect(show(t)).toEqual(['1|2|3', '4|4^|6', '4^|4^|6', '4^|4^|9', '|^|9']);
    expect(t.rows()[4]!.cells()[0]!.vMerge).toBeUndefined();
    t.cloneRow(1, { after: -1 }); // a restart copied above everything: plain
    expect(show(t)[0]).toBe('4|4^|6');
    expect(t.rows()[0]!.cells()[0]!.vMerge).toBeUndefined();
    expectValid(doc);
  });
});

describe('insertRow', () => {
  it('copies cell layout, widths, borders and formatting from the row above', () => {
    const borders =
      '<w:tcBorders><w:top w:val="double" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders>';
    const { doc, t } = single(
      tbl(
        grid(2, 2000) +
          tr(
            `<w:tc><w:tcPr>${tcw(2000)}${borders}<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr>` +
              `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="00FF00"/></w:rPr><w:t>A</w:t></w:r></w:p></w:tc>` +
              tc('B', '', 2000),
          ) +
          tr(tc('C', '', 2000) + tc('D', '', 2000)),
      ) + para(''),
    );
    const row = t.insertRow(1, ['x', 'y']);
    expect(show(t)).toEqual(['A|B', 'x|y', 'C|D']);
    const s = xml(row.element);
    expect(s).toContain(borders);
    expect(s).toContain('w:fill="EEEEEE"');
    expect(s).toContain('<w:jc w:val="center"/>');
    expect(s).toContain('<w:rPr><w:b/><w:color w:val="00FF00"/></w:rPr><w:t>x</w:t>');
    expect(row.cells().map((c) => c.width)).toEqual([2000, 2000]);
    expect(s).not.toContain('paraId');
    t.insertRow(0, ['top']);
    t.insertRow(4);
    expect(show(t)).toEqual(['top|', 'A|B', 'x|y', 'C|D', '|']);
    expect(() => t.insertRow(9)).toThrow(UsageError);
    expectValid(doc);
  });

  it('continues a vertical merge it lands inside (and ignores text for those cells)', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[3]!;
    t.insertRow(2, ['no', 'z']);
    expect(show(t)).toEqual(['1|2|3', '4|4^|6', '4^|4^|z', '4^|4^|9']);
    t.insertRow(4, ['p', 'q']); // below the merge: plain cells
    expect(show(t)).toEqual(['1|2|3', '4|4^|6', '4^|4^|z', '4^|4^|9', 'p|p^|q']);
    expectValid(doc);
  });

  it('keeps tblHeader only inside the leading header rows', () => {
    const { doc, t } = single(
      tbl(grid(1) + tr(tc('H'), '<w:tblHeader/>') + tr(tc('D'))) + para(''),
    );
    expect(t.insertRow(1).header).toBe(false);
    expect(t.insertRow(0).header).toBe(true);
    expect(t.rows().map((r) => r.header)).toEqual([true, true, false, false]);
    expectValid(doc);
  });

  it('fills a table without rows from its grid', () => {
    const { doc, t } = single(tbl(grid(3, 1500)) + para(''));
    t.insertRow(0, ['a', 'b', 'c']);
    expect(show(t)).toEqual(['a|b|c']);
    expect(t.rows()[0]!.cells()[2]!.width).toBe(1500);
    expectValid(doc);
  });
});

describe('deleteRow', () => {
  it('promotes the next merge cell and moves the merged content into it', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[3]!;
    t.deleteRow(1);
    expect(show(t)).toEqual(['1|2|3', '4|4^|9']);
    expect(t.rows()[1]!.cells()[0]!.vMerge).toBeUndefined();
    const t2 = doc.body.tables()[2]!;
    expect(show(t2)).toEqual(['1|2|3', '4|5|6', '7|5^|9']);
    t2.deleteRow(2);
    expect(show(t2)).toEqual(['1|2|3', '4|5|6']);
    expectValid(doc);
  });

  it('keeps a merge restart when the merge still spans several rows', () => {
    const { doc, t } = single(
      tbl(
        grid(2) + tr(tc('A', VR) + tc('1')) + tr(tc('', VC) + tc('2')) + tr(tc('', VC) + tc('3')),
      ) + para(''),
    );
    t.deleteRow(0);
    expect(show(t)).toEqual(['A|2', 'A^|3']);
    expect(t.rows()[0]!.cells()[0]!.vMerge).toBe('restart');
    expectValid(doc);
  });

  it('keeps ranges paired: markers whose partner stays are moved out of the row', () => {
    const { doc, t } = single(
      `${tbl(
        grid(1) +
          tr(tc('A')) +
          tr(
            `<w:tc><w:tcPr>${tcw(1000)}</w:tcPr><w:p><w:bookmarkStart w:id="4" w:name="b"/><w:bookmarkStart w:id="5" w:name="inner"/><w:r><w:t>B</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p></w:tc>`,
          ),
      )}<w:p><w:r><w:t>after</w:t></w:r><w:bookmarkEnd w:id="4"/></w:p>`,
    );
    t.deleteRow(1);
    const s = xml(t.element);
    expect(s).toContain('<w:bookmarkStart w:id="4" w:name="b"/>');
    expect(s).not.toContain('inner');
    expectValid(doc);
  });

  it('removes a row-level content control left empty, and the table with its last row', () => {
    const { doc, t } = single(
      tbl(
        grid(1) +
          tr(tc('A')) +
          `<w:sdt><w:sdtPr/><w:sdtContent>${tr(tc('B'))}</w:sdtContent></w:sdt>`,
      ) + para('end'),
    );
    t.deleteRow(1);
    expect(xml(t.element)).not.toContain('sdt');
    t.rows()[0]!.remove();
    expect(doc.body.tables()).toHaveLength(0);
    expect(doc.body.text()).toBe('end');
    expectValid(doc);
  });
});

// ---- columns -----------------------------------------------------------------------------

describe('insertColumn', () => {
  it('splits the neighbour column by default, keeping the table width', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[0]!;
    const cells = t.insertColumn(1);
    expect(cells).toHaveLength(3);
    expect(t.columnWidths()).toEqual([1440, 720, 720, 1440]);
    expect(show(t)).toEqual(['1||2|3', '4||5|6', '7||8|9']);
    expect(
      t
        .rows()[0]!
        .cells()
        .map((c) => c.width),
    ).toEqual([1440, 720, 720, 1440]);
    expectValid(doc);
  });

  it('appends a column of a given width and grows a twips table width', () => {
    const { doc, t } = single(
      tbl(
        grid(2) + tr(tc('A') + tc('B')) + tr(tc('C') + tc('D')),
        '<w:tblW w:w="2000" w:type="dxa"/>',
      ) + para(''),
    );
    t.insertColumn(2, { width: 1500 })[1]!.setText('new');
    expect(show(t)).toEqual(['A|B|', 'C|D|new']);
    expect(t.columnWidths()).toEqual([1000, 1000, 1500]);
    expect(xml(t.element)).toContain('<w:tblW w:w="3500" w:type="dxa"/>');
    expect(() => t.insertColumn(5)).toThrow(UsageError);
    expectValid(doc);
  });

  it('widens cells spanning the insertion point and respects gridBefore/gridAfter', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[1]!;
    expect(show(t)).toEqual(['1|2|3', '4|4^|6', '7|8|9']);
    t.insertColumn(1, { width: 500 });
    expect(show(t)).toEqual(['1||2|3', '4|4^|4^|6', '7||8|9']);
    expect(t.rows()[1]!.cells()[0]!.width).toBe(1440 + 500 + 1440);
    const { doc: d2, t: t2 } = single(
      tbl(
        grid(3) +
          tr(tc('A'), '<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>') +
          tr(tc('B') + tc('C') + tc('D')),
      ) + para(''),
    );
    t2.insertColumn(0);
    expect(show(t2)).toEqual(['.|.|A|.', '|B|C|D']);
    t2.insertColumn(4);
    expect(show(t2)).toEqual(['.|.|A|.|.', '|B|C|D|']);
    expect(t2.rows()[0]!.gridAfter).toBe(2);
    expectValid(doc);
    expectValid(d2);
  });
});

describe('deleteColumn', () => {
  it('shrinks spanning cells and removes the others', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[1]!;
    t.deleteColumn(1);
    expect(show(t)).toEqual(['1|3', '4|6', '7|9']);
    expect(t.columnWidths()).toEqual([1440, 1440]);
    expect(t.rows()[1]!.cells()[0]!.gridSpan).toBe(1);
    expect(t.rows()[1]!.cells()[0]!.width).toBe(1440);
    expectValid(doc);
  });

  it('deletes a row whose only cell goes, and the table with its last column', () => {
    const { doc, t } = single(
      tbl(
        grid(3) +
          tr(tc('A'), '<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>') +
          tr(tc('B') + tc('C') + tc('D')),
        '<w:tblW w:w="3000" w:type="dxa"/>',
      ) + para(''),
    );
    t.deleteColumn(1);
    expect(show(t)).toEqual(['B|D']);
    expect(xml(t.element)).toContain('<w:tblW w:w="2000" w:type="dxa"/>');
    t.deleteColumn(0);
    expect(show(t)).toEqual(['D']);
    t.deleteColumn(0);
    expect(doc.body.tables()).toHaveLength(0);
    expectValid(doc);
  });

  it('shrinks gridBefore and gridAfter', () => {
    const { doc, t } = single(
      tbl(
        grid(4) +
          tr(tc('A') + tc('B'), '<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>') +
          tr(tc('C', span(4))),
      ) + para(''),
    );
    t.deleteColumn(0);
    expect(show(t)).toEqual(['A|B|.', 'C|C^|C^']);
    t.deleteColumn(2);
    expect(show(t)).toEqual(['A|B', 'C|C^']);
    expect(t.rows()[0]!.gridBefore + t.rows()[0]!.gridAfter).toBe(0);
    expect(xml(t.rows()[0]!.element)).not.toContain('trPr');
    expectValid(doc);
  });
});

// ---- merge / unmerge ---------------------------------------------------------------------

describe('merge', () => {
  it('merges a rectangle: gridSpan + vMerge, content moved to the origin', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[0]!;
    const cell = t.merge(0, 0, 1, 1);
    expect(show(t)).toEqual(['1\n2\n4\n5|1\n2\n4\n5^|3', '1\n2\n4\n5^|1\n2\n4\n5^|6', '7|8|9']);
    expect(cell.gridSpan).toBe(2);
    expect(cell.vMerge).toBe('restart');
    expect(cell.width).toBe(2880);
    const covered = t.rows()[1]!.cells()[0]!;
    expect(covered.vMerge).toBe('continue');
    expect(covered.gridSpan).toBe(2);
    expect(covered.paragraphs()).toHaveLength(1);
    expect(covered.text()).toBe('');
    expectValid(doc);
  });

  it('merges within a row without vMerge, skipping empty cells', () => {
    const { doc, t } = single(tbl(grid(3) + tr(tc('') + tc('B') + tc('C'))) + para(''));
    const c = t.merge(0, 2, 0, 0);
    expect(show(t)).toEqual(['B\nC|B\nC^|B\nC^']);
    expect(c.vMerge).toBeUndefined();
    expect(c.paragraphs()).toHaveLength(2);
    expect(t.merge(0, 1, 0, 1).element).toBe(c.element); // already merged: no-op
    expectValid(doc);
  });

  it('grows the rectangle over merged cells it touches', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[3]!;
    t.merge(1, 1, 1, 2);
    expect(show(t)).toEqual(['1|2|3', '4\n6\n9|4\n6\n9^|4\n6\n9^', '4\n6\n9^|4\n6\n9^|4\n6\n9^']);
    expectValid(doc);
  });

  it('refuses rectangles with holes or outside the grid, leaving the table valid', () => {
    const { doc, t } = single(
      tbl(
        grid(3) +
          tr(tc('A') + tc('B'), '<w:gridAfter w:val="1"/>') +
          tr(tc('C') + tc('D') + tc('E')),
      ) + para(''),
    );
    expect(() => t.merge(0, 1, 1, 2)).toThrow(UsageError);
    expect(() => t.merge(0, 0, 5, 0)).toThrow(UsageError);
    expect(show(t)).toEqual(['A|B|.', 'C|D|E']);
    expectValid(doc);
  });
});

describe('unmerge', () => {
  it('splits a merged cell into one cell per slot; the top-left keeps the content', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[0]!;
    t.merge(0, 0, 1, 1);
    const cell = t.unmerge(1, 1);
    expect(cell.text()).toBe('1\n2\n4\n5');
    expect(show(t)).toEqual(['1\n2\n4\n5||3', '||6', '7|8|9']);
    expect(
      t
        .rows()[0]!
        .cells()
        .map((c) => c.width),
    ).toEqual([1440, 1440, 1440]);
    expectValid(doc);
  });

  it('splits vertical and horizontal merges from a fixture; plain cells stay', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[3]!;
    t.unmerge(2, 0);
    expect(show(t)).toEqual(['1|2|3', '4||6', '||9']);
    const plain = t.cell(0, 0)!;
    expect(t.unmerge(0, 0).element).toBe(plain.element);
    expect(() => t.unmerge(0, 7)).toThrow(UsageError);
    expectValid(doc);
  });
});

// ---- grid widths -------------------------------------------------------------------------

describe('recomputeGrid and setColumnWidth', () => {
  it('rebuilds tblGrid from cell widths', () => {
    const { doc, t } = single(
      tbl(
        grid(3) +
          tr(tc('A', span(2), 5000) + tc('B', '', 4000)) +
          tr(tc('C', '', 2000) + tc('D', '', 3000) + tc('E', '', 4000)),
      ) + para(''),
    );
    expect(t.recomputeGrid()).toEqual([2000, 3000, 4000]);
    expect(t.columnWidths()).toEqual([2000, 3000, 4000]);
    expectValid(doc);
  });

  it('creates a missing tblGrid and pads short rows', () => {
    const { doc, t } = single(
      `<w:tbl><w:tblPr/>${tr(tc('A', '', 1200) + tc('B', '', 1300))}${tr(tc('C', '', 2500))}</w:tbl>${para('')}`,
    );
    expect(t.recomputeGrid()).toEqual([1200, 1300]);
    expect(t.rows()[1]!.gridAfter).toBe(1);
    expectValid(doc);
  });

  it('creates a missing tblGrid before a column edit', () => {
    const { doc, t } = single(
      `<w:tbl><w:tblPr/>${tr(tc('A', '', 1200) + tc('B', '', 1300))}</w:tbl>${para('')}`,
    );
    t.insertColumn(2, { width: 700 });
    expect(t.columnWidths()).toEqual([1200, 1300, 700]);
    expect(show(t)).toEqual(['A|B|']);
    expectValid(doc);
  });

  it('infers columns only multi-column cells pin down', () => {
    const { t } = single(
      tbl(
        grid(3) +
          tr(tc('A', '', 1000) + tc('B', span(2), 5000)) +
          tr(tc('C', span(2), 3000) + tc('D', '', 3000)),
      ) + para(''),
    );
    expect(t.recomputeGrid()).toEqual([1000, 2000, 3000]);
  });

  it('sets one column width and the widths of the cells over it', () => {
    const doc = load('python-docx/tbl-cell-access.docx');
    const t = doc.body.tables()[1]!;
    t.setColumnWidth(0, 2000);
    expect(t.columnWidths()).toEqual([2000, 1440, 1440]);
    expect(t.rows()[0]!.cells()[0]!.width).toBe(2000);
    expect(t.rows()[1]!.cells()[0]!.width).toBe(3440);
    expect(t.rows()[1]!.cells()[1]!.width).toBe(1440);
    expect(() => t.setColumnWidth(3, 100)).toThrow(UsageError);
    expect(() => t.setColumnWidth(0, 1.5)).toThrow(UsageError);
    expectValid(doc);
  });
});

// ---- remove, legacy hMerge ---------------------------------------------------------------

describe('Table.remove', () => {
  it('removes a nested table and keeps the cell ending with a paragraph', () => {
    const inner = tbl(grid(1) + tr(tc('n')));
    const { doc, t } = single(
      tbl(
        grid(1) + tr(`<w:tc><w:tcPr>${tcw(1000)}</w:tcPr>${para('x')}${inner}${para('')}</w:tc>`),
      ) + para(''),
    );
    const cell = t.cell(0, 0)!;
    // Make the nested table the cell's last block, then remove it.
    cell.element.lastElement()!.remove();
    cell.tables()[0]!.remove();
    expect(cell.tables()).toHaveLength(0);
    expect(lastBlock(cell.element)?.local).toBe('p');
    expectValid(doc);
  });

  it('removes an emptied block-level content control and gives an empty story a paragraph', () => {
    const doc = readDocx(
      docx(`<w:sdt><w:sdtPr/><w:sdtContent>${tbl(grid(1) + tr(tc('a')))}</w:sdtContent></w:sdt>`),
    );
    doc.body.tables()[0]!.remove();
    expect(xml(doc.body.element)).toBe('<w:body><w:p/><w:sectPr/></w:body>');
    expectValid(doc);
  });
});

describe('legacy hMerge', () => {
  it('is turned into gridSpan before column edits, keeping the grid', () => {
    const { doc, t } = single(
      tbl(
        grid(4) +
          tr(tc('A', '<w:hMerge w:val="restart"/>') + tc('', '<w:hMerge/>') + tc('C') + tc('D')) +
          tr(tc('E') + tc('F') + tc('G', '<w:hMerge w:val="restart"/>') + tc('H', '<w:hMerge/>')),
      ) + para(''),
    );
    expect(show(t)).toEqual(['A|A^|C|D', 'E|F|G|G^']);
    t.insertColumn(4, { width: 100 });
    expect(show(t)).toEqual(['A|A^|C|D|', 'E|F|G\nH|G\nH^|']);
    expect(xml(t.element)).not.toContain('hMerge');
    expect(t.rows()[0]!.cells()[0]!.gridSpan).toBe(2);
    expectValid(doc);
  });
});

// ---- row and cell properties -------------------------------------------------------------

describe('TableRow setters', () => {
  it('writes height, hRule, header and cantSplit in Word order after tblPrEx', () => {
    const { doc, t } = single(
      tbl(`${grid(1)}<w:tr><w:tblPrEx><w:tblBorders/></w:tblPrEx>${tc('A')}</w:tr>`) + para(''),
    );
    const row = t.rows()[0]!;
    row.header = true;
    row.height = 500;
    row.cantSplit = true;
    row.heightRule = 'exact';
    expect(xml(row.element)).toContain(
      '<w:tblPrEx><w:tblBorders/></w:tblPrEx><w:trPr><w:cantSplit/><w:trHeight w:val="500" w:hRule="exact"/><w:tblHeader/></w:trPr>',
    );
    expect([row.height, row.heightRule, row.header, row.cantSplit]).toEqual([
      500,
      'exact',
      true,
      true,
    ]);
    row.header = false;
    row.cantSplit = false;
    row.height = undefined;
    expect(xml(row.element)).not.toContain('trPr');
    expect([row.height, row.heightRule, row.header, row.cantSplit]).toEqual([
      undefined,
      undefined,
      false,
      false,
    ]);
    expect(() => {
      row.height = -1;
    }).toThrow(UsageError);
    expect(row.table.element).toBe(t.element);
    expectValid(doc);
  });
});

describe('TableCell setters', () => {
  it('setText keeps the first paragraph and run formatting', () => {
    const { doc, t } = single(
      tbl(
        grid(1) +
          tr(
            `<w:tc><w:tcPr>${tcw(1000)}</w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr>` +
              '<w:r><w:rPr><w:i/></w:rPr><w:t>old</w:t></w:r><w:r><w:t>more</w:t></w:r></w:p>' +
              `${para('second')}</w:tc>`,
          ),
      ) + para(''),
    );
    const cell = t.cell(0, 0)!;
    const p = cell.setText('new\ttext');
    expect(cell.text()).toBe('new\ttext');
    expect(cell.paragraphs()).toHaveLength(1);
    expect(xml(p.element)).toBe(
      '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>new</w:t><w:tab/><w:t>text</w:t></w:r></w:p>',
    );
    cell.setText('');
    expect(cell.text()).toBe('');
    expect(cell.paragraphs()).toHaveLength(1);
    expectValid(doc);
  });

  it('setText keeps range markers whose partner is outside the cell', () => {
    const { doc, t } = single(
      `${tbl(
        grid(1) +
          tr(
            `<w:tc><w:tcPr>${tcw(1000)}</w:tcPr><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>a</w:t></w:r></w:p></w:tc>`,
          ),
      )}<w:p><w:r><w:t>b</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>`,
    );
    t.cell(0, 0)!.setText('z');
    expect(xml(t.element)).toContain('<w:commentRangeStart w:id="1"/>');
    expectValid(doc);
  });

  it('appendParagraph, shading, vAlign and width write ordered tcPr children', () => {
    const { doc, t } = single(tbl(grid(1) + tr(`<w:tc>${para('a')}</w:tc>`)) + para(''));
    const cell: TableCell = t.cell(0, 0)!;
    cell.vAlign = 'center';
    cell.shading = 'FFEE00';
    cell.width = 1234;
    cell.appendParagraph('b', { style: 'Heading1' });
    expect(xml(cell.element)).toMatch(
      /^<w:tc><w:tcPr><w:tcW w:w="1234" w:type="dxa"\/><w:shd w:val="clear" w:color="auto" w:fill="FFEE00"\/><w:vAlign w:val="center"\/><\/w:tcPr>/,
    );
    expect([cell.vAlign, cell.shading, cell.width]).toEqual(['center', 'FFEE00', 1234]);
    expect(cell.paragraphs().map((p) => [p.text, p.style])).toEqual([
      ['a', undefined],
      ['b', 'Heading1'],
    ]);
    cell.vAlign = undefined;
    cell.shading = undefined;
    cell.width = undefined;
    expect(xml(cell.element)).not.toContain('tcPr');
    expect(() => {
      cell.shading = 'red';
    }).toThrow(UsageError);
    expect(cell.row.table.element).toBe(t.element);
    expectValid(doc);
  });
});

// ---- randomized operation sequences ------------------------------------------------------

/** mulberry32: a tiny seeded PRNG (tests only; src never uses randomness). */
function prng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** A synthetic table with spans, vertical merges, gridBefore/After, sdt rows, hMerge, bookmarks and a nested table. */
function richBody(): string {
  const nested = tbl(
    grid(2, 500) + tr(tc('n1', '', 500) + tc('n2', '', 500)) + tr(tc('n3', span(2), 1000)),
  );
  return (
    para('before') +
    tbl(
      grid(5) +
        tr(tc('A', span(2) + VR, 2000) + tc('B') + tc('C', VR) + tc('D')) +
        tr(tc('', span(2) + VC, 2000) + tc('E') + tc('', VC) + tc('F')) +
        tr(tc('G') + tc('H', span(2), 2000), '<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>') +
        `<w:sdt><w:sdtPr><w:tag w:val="row"/></w:sdtPr><w:sdtContent>${tr(
          `<w:tc><w:tcPr>${tcw(1000)}</w:tcPr><w:p><w:bookmarkStart w:id="9" w:name="cross"/><w:r><w:t>I</w:t></w:r></w:p></w:tc>` +
            tc('J') +
            `<w:sdt><w:sdtPr/><w:sdtContent>${tc('K')}</w:sdtContent></w:sdt>` +
            tc('L') +
            tc('M'),
        )}</w:sdtContent></w:sdt>` +
        tr(
          tc('N', '<w:hMerge w:val="restart"/>') +
            tc('O', '<w:hMerge/>') +
            tc('P') +
            tc('Q') +
            tc('R'),
        ) +
        tr(
          `<w:tc><w:tcPr>${tcw(1000)}</w:tcPr>${para('x')}${nested}${para('')}</w:tc>` +
            tc('S', span(3), 3000) +
            tc('T'),
        ),
    ) +
    `<w:p><w:r><w:t>after</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>`
  );
}

type Op = (t: Table, rnd: () => number) => string | undefined;
const pick = (rnd: () => number, n: number): number => Math.floor(rnd() * n);

const OPS: readonly [string, Op][] = [
  [
    'cloneRow',
    (t, rnd) => {
      const n = t.rows().length;
      if (n >= 25) return undefined;
      const i = pick(rnd, n);
      const after = pick(rnd, n + 1) - 1;
      t.cloneRow(i, { after });
      return `cloneRow(${i}, after ${after})`;
    },
  ],
  [
    'insertRow',
    (t, rnd) => {
      const n = t.rows().length;
      if (n >= 25) return undefined;
      const i = pick(rnd, n + 1);
      t.insertRow(i, ['r', 's', 't']);
      return `insertRow(${i})`;
    },
  ],
  [
    'deleteRow',
    (t, rnd) => {
      const n = t.rows().length;
      if (n <= 1) return undefined;
      const i = pick(rnd, n);
      t.deleteRow(i);
      return `deleteRow(${i})`;
    },
  ],
  [
    'insertColumn',
    (t, rnd) => {
      const n = t.gridColumnCount;
      if (n >= 12) return undefined;
      const i = pick(rnd, n + 1);
      const width = rnd() < 0.5 ? undefined : 300 + pick(rnd, 1500);
      t.insertColumn(i, width === undefined ? undefined : { width });
      return `insertColumn(${i}, ${width})`;
    },
  ],
  [
    'deleteColumn',
    (t, rnd) => {
      const n = t.gridColumnCount;
      if (n <= 1) return undefined;
      const i = pick(rnd, n);
      t.deleteColumn(i);
      return `deleteColumn(${i})`;
    },
  ],
  [
    'merge',
    (t, rnd) => {
      const g = t.grid();
      if (g.rows === 0 || g.cols === 0) return undefined;
      const r1 = pick(rnd, g.rows);
      const c1 = pick(rnd, g.cols);
      const r2 = Math.min(g.rows - 1, r1 + pick(rnd, 3));
      const c2 = Math.min(g.cols - 1, c1 + pick(rnd, 3));
      try {
        t.merge(r1, c1, r2, c2);
      } catch (e) {
        if (!(e instanceof UsageError)) throw e;
        return `merge(${r1},${c1},${r2},${c2}) refused`;
      }
      return `merge(${r1},${c1},${r2},${c2})`;
    },
  ],
  [
    'unmerge',
    (t, rnd) => {
      const g = t.grid();
      const r = pick(rnd, g.rows);
      const c = pick(rnd, g.cols);
      if (g.at(r, c) === undefined) return undefined;
      t.unmerge(r, c);
      return `unmerge(${r},${c})`;
    },
  ],
  [
    'setColumnWidth',
    (t, rnd) => {
      const n = t.gridColumnCount;
      const i = pick(rnd, n);
      t.setColumnWidth(i, 200 + pick(rnd, 3000));
      return `setColumnWidth(${i})`;
    },
  ],
  [
    'recomputeGrid',
    (t) => {
      t.recomputeGrid();
      return 'recomputeGrid()';
    },
  ],
  [
    'setText',
    (t, rnd) => {
      const g = t.grid();
      const s = g.at(pick(rnd, g.rows), pick(rnd, g.cols));
      if (s === undefined) return undefined;
      s.cell.setText(`t${pick(rnd, 100)}`);
      return 'setText';
    },
  ],
  [
    'rowProps',
    (t, rnd) => {
      const rows = t.rows();
      const row = rows[pick(rnd, rows.length)]!;
      row.height = rnd() < 0.3 ? undefined : 200 + pick(rnd, 800);
      row.cantSplit = rnd() < 0.5;
      return 'rowProps';
    },
  ],
  [
    'cellProps',
    (t, rnd) => {
      const g = t.grid();
      const s = g.at(pick(rnd, g.rows), pick(rnd, g.cols));
      if (s === undefined) return undefined;
      s.cell.shading = rnd() < 0.5 ? 'DDEEFF' : undefined;
      s.cell.vAlign = rnd() < 0.5 ? 'bottom' : undefined;
      return 'cellProps';
    },
  ],
];

function randomRun(name: string, input: Uint8Array, seed: number, steps: number): void {
  const doc = readDocx(input);
  expectValid(doc);
  const rnd = prng(seed);
  const log: string[] = [];
  let done = 0;
  for (let step = 0; step < steps; step++) {
    const tables = doc.body.element.descendants(NS.w, 'tbl').map((el) => new Table(el, doc.body));
    if (tables.length === 0) break;
    const t = tables[pick(rnd, tables.length)]!;
    const [opName, op] = OPS[pick(rnd, OPS.length)]!;
    let desc: string | undefined;
    try {
      desc = op(t, rnd);
    } catch (e) {
      throw new Error(
        `${name} seed ${seed} step ${step} ${opName}: ${(e as Error).stack}\n${log.slice(-8).join('\n')}`,
      );
    }
    if (desc === undefined) continue;
    done++;
    log.push(desc);
    const v = docViolations(doc);
    if (v.length > 0) {
      throw new Error(
        `${name} seed ${seed} step ${step} ${desc}:\n${v.join('\n')}\nlast ops:\n${log.slice(-8).join('\n')}`,
      );
    }
  }
  expect(done).toBeGreaterThan(steps / 2);
  saveAndCheck(doc, `random-${name.replace(/[/.]/g, '_')}-${seed}`, input);
}

describe('randomized operation sequences keep every invariant', () => {
  it.each([
    ['python-docx/tbl-cell-access.docx', 1],
    ['python-docx/tbl-cell-access.docx', 2],
    ['open-xml-sdk/TableFntEdnt.docx', 3],
    ['apache-poi/bug69628.docx', 4],
  ])('%s (seed %i)', (name, seed) => {
    randomRun(name, fixture(name).bytes, seed, 300);
  });

  it('synthetic Strict tbl-cell-access (seed 8)', () => {
    const [strict] = strictFixtures(['python-docx/tbl-cell-access.docx']);
    randomRun('strict-tbl-cell-access', strict!.bytes, 8, 300);
  });

  it.each([5, 6, 7])('synthetic rich table (seed %i)', (seed) => {
    const input = docx(richBody());
    randomRun('synthetic-rich', input, seed, 300);
  });
});

describe('performance', () => {
  it('row operations on a long table only touch the rows around the change', () => {
    const doc = createDocx();
    doc.body.element.firstElement()!.before(newTable(doc.body, 2, 10));
    const t = doc.body.tables()[0]!;
    for (let c = 0; c < 10; c++) t.cell(1, c)!.setText(`v${c}`);
    const start = performance.now();
    const release = DocumentIds.of(doc).hold();
    for (let i = 0; i < 1000; i++) t.cloneRow(1, { after: 1 + i });
    release();
    for (let i = 0; i < 200; i++) t.insertRow(500, ['a']);
    for (let i = 0; i < 200; i++) t.deleteRow(600);
    expect(t.rows()).toHaveLength(1002);
    expect(performance.now() - start).toBeLessThan(10_000); // ~0.3 s locally
    expectValid(doc);
  });
});

describe('outputs', () => {
  it('writes edited fixtures for the LibreOffice check', () => {
    const input = fixture('python-docx/tbl-cell-access.docx').bytes;
    const doc = readDocx(input);
    const [t0, t1, t2, t3] = doc.body.tables();
    t0!.cloneRow(1);
    t0!.insertColumn(3, { width: 1440 })[0]!.setText('new col');
    t1!.merge(0, 0, 2, 0);
    t2!.deleteRow(0);
    t3!.unmerge(1, 0);
    t3!.insertRow(3, ['x', 'y', 'z']);
    expectValid(doc);
    saveAndCheck(doc, 'edited-tbl-cell-access', input);
  });
});
