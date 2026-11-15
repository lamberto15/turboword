/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { Story } from '../src/wml/story.js';
import { type GridCell, MAX_GRID_COLUMNS, Table } from '../src/wml/table.js';
import { parseXml } from '../src/xml/parse.js';

const W_URI = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W = `xmlns:w="${W_URI}"`;

// The grid logic never touches the story; tests that need a real one go through readDocx.
const story = {} as unknown as Story;

function table(inner: string): Table {
  return new Table(parseXml(`<w:tbl ${W}>${inner}</w:tbl>`).root, story);
}
const tblGrid = (n: number): string =>
  `<w:tblGrid>${'<w:gridCol w:w="1000"/>'.repeat(n)}</w:tblGrid>`;
const tc = (text: string, tcPr = ''): string =>
  `<w:tc>${tcPr === '' ? '' : `<w:tcPr>${tcPr}</w:tcPr>`}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const tr = (cells: string, trPr = ''): string =>
  `<w:tr>${trPr === '' ? '' : `<w:trPr>${trPr}</w:trPr>`}${cells}</w:tr>`;
const span = (n: number): string => `<w:gridSpan w:val="${n}"/>`;

/** The grid as text: each slot shows the covering cell's DOM text ('.' = no cell). */
function gridText(t: Table): string[] {
  const g = t.grid();
  const out: string[] = [];
  for (let r = 0; r < g.rows; r++) {
    const line: string[] = [];
    for (let c = 0; c < g.cols; c++) line.push(g.at(r, c)?.cell.element.text ?? '.');
    out.push(line.join(' '));
  }
  return out;
}

function slot(t: Table, r: number, c: number): Omit<GridCell, 'cell'> & { text: string } {
  const s = t.grid().at(r, c);
  if (s === undefined) throw new Error(`no slot at ${r},${c}`);
  const { cell, ...rest } = s;
  return { text: cell.element.text, ...rest };
}

describe('Table rows and cells', () => {
  it('lists rows and cells, looking through sdt and customXml wrappers', () => {
    const t = table(
      `<w:tblPr/>${tblGrid(3)}` +
        tr(tc('A') + tc('B') + tc('C')) +
        `<w:sdt><w:sdtPr><w:tag w:val="row"/></w:sdtPr><w:sdtContent>${tr(
          tc('D') +
            `<w:sdt><w:sdtPr/><w:sdtContent>${tc('E')}</w:sdtContent></w:sdt>` +
            `<w:customXml w:element="c">${tc('F')}</w:customXml>`,
        )}</w:sdtContent></w:sdt>` +
        `<w:bookmarkStart w:id="0" w:name="x"/>` +
        `<w:customXml w:element="r">${tr(tc('G') + tc('H') + tc('I'))}</w:customXml>` +
        `<w:bookmarkEnd w:id="0"/>`,
    );
    const rows = t.rows();
    expect(rows).toHaveLength(3);
    expect(
      rows.map((r) =>
        r
          .cells()
          .map((c) => c.element.text)
          .join(''),
      ),
    ).toEqual(['ABC', 'DEF', 'GHI']);
    expect(t.gridColumnCount).toBe(3);
    expect(gridText(t)).toEqual(['A B C', 'D E F', 'G H I']);
    expect(rows[0]!.element.local).toBe('tr');
  });

  it('looks through mc:AlternateContent like the text reader (Fallback first)', () => {
    const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    const alt = (choice: string, fallback: string): string =>
      `<mc:AlternateContent><mc:Choice Requires="w14">${choice}</mc:Choice>` +
      `<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
    const t = new Table(
      parseXml(
        `<w:tbl ${W} ${MC}>${tblGrid(2)}${alt(tr(tc('c1') + tc('c2')), tr(tc('f1') + alt(tc('x'), tc('f2'))))}` +
          `<mc:AlternateContent><mc:Choice Requires="w14">${tr(tc('only') + tc('choice'))}</mc:Choice></mc:AlternateContent></w:tbl>`,
      ).root,
      story,
    );
    expect(gridText(t)).toEqual(['f1 f2', 'only choice']);
  });

  it('exposes gridSpan, vMerge, hMerge, gridBefore and gridAfter', () => {
    const t = table(
      tr(
        tc('A', `${span(2)}<w:vMerge w:val="restart"/>`) +
          tc('B', '<w:vMerge/>') +
          tc('C', '<w:hMerge w:val="restart"/>') +
          tc('D', '<w:gridSpan w:val="0"/>'),
        '<w:gridBefore w:val="1"/><w:gridAfter w:val="2"/>',
      ),
    );
    const [row] = t.rows();
    expect(row!.gridBefore).toBe(1);
    expect(row!.gridAfter).toBe(2);
    const [a, b, c, d] = row!.cells();
    expect(a!.gridSpan).toBe(2);
    expect(a!.vMerge).toBe('restart');
    expect(b!.vMerge).toBe('continue');
    expect(b!.gridSpan).toBe(1);
    expect(c!.vMerge).toBeUndefined();
    expect(c!.hMerge).toBe('restart');
    expect(d!.gridSpan).toBe(1); // invalid spans are clamped to 1
    expect(table(tr(tc('X'))).rows()[0]!.gridBefore).toBe(0);
  });
});

describe('Table.grid()', () => {
  it('handles gridSpan', () => {
    const t = table(tblGrid(3) + tr(tc('A', span(2)) + tc('B')) + tr(tc('C') + tc('D') + tc('E')));
    const g = t.grid();
    expect(g.rows).toBe(2);
    expect(g.cols).toBe(3);
    expect(gridText(t)).toEqual(['A A B', 'C D E']);
    expect(slot(t, 0, 0)).toEqual({
      text: 'A',
      row: 0,
      col: 0,
      rowSpan: 1,
      colSpan: 2,
      origin: true,
    });
    expect(slot(t, 0, 1)).toEqual({
      text: 'A',
      row: 0,
      col: 0,
      rowSpan: 1,
      colSpan: 2,
      origin: false,
    });
    expect(slot(t, 0, 2)).toMatchObject({ text: 'B', col: 2, colSpan: 1, origin: true });
    expect(g.at(0, 0)!.cell).toBe(g.at(0, 1)!.cell);
    expect(g.at(2, 0)).toBeUndefined();
    expect(g.at(0, 3)).toBeUndefined();
    expect(g.at(-1, 0)).toBeUndefined();
  });

  it('maps vMerge continuation slots to the origin cell above and computes rowSpan', () => {
    const t = table(
      tblGrid(2) +
        tr(tc('A', '<w:vMerge w:val="restart"/>') + tc('B')) +
        tr(tc('C', '<w:vMerge w:val="continue"/>') + tc('D')) +
        tr(tc('E', '<w:vMerge/>') + tc('F')) +
        tr(tc('G') + tc('H')),
    );
    expect(gridText(t)).toEqual(['A B', 'A D', 'A F', 'G H']);
    expect(slot(t, 0, 0)).toEqual({
      text: 'A',
      row: 0,
      col: 0,
      rowSpan: 3,
      colSpan: 1,
      origin: true,
    });
    expect(slot(t, 2, 0)).toEqual({
      text: 'A',
      row: 0,
      col: 0,
      rowSpan: 3,
      colSpan: 1,
      origin: false,
    });
    expect(slot(t, 3, 0)).toMatchObject({ text: 'G', rowSpan: 1, origin: true });
    expect(t.cell(1, 0)!.element.text).toBe('A');
    expect(t.cell(1, 1)!.element.text).toBe('D');
    expect(t.cell(9, 9)).toBeUndefined();
  });

  it('merges vertically across spanned cells and restarts a new chain', () => {
    const t = table(
      tblGrid(3) +
        tr(tc('A', `${span(2)}<w:vMerge w:val="restart"/>`) + tc('B')) +
        tr(tc('C', `${span(2)}<w:vMerge/>`) + tc('D', '<w:vMerge w:val="restart"/>')) +
        tr(tc('E', `${span(2)}<w:vMerge w:val="restart"/>`) + tc('F', '<w:vMerge/>')),
    );
    expect(gridText(t)).toEqual(['A A B', 'A A D', 'E E D']);
    expect(slot(t, 1, 1)).toEqual({
      text: 'A',
      row: 0,
      col: 0,
      rowSpan: 2,
      colSpan: 2,
      origin: false,
    });
    expect(slot(t, 2, 2)).toEqual({
      text: 'D',
      row: 1,
      col: 2,
      rowSpan: 2,
      colSpan: 1,
      origin: false,
    });
    expect(slot(t, 2, 0)).toMatchObject({
      text: 'E',
      row: 2,
      rowSpan: 1,
      colSpan: 2,
      origin: true,
    });
  });

  it('treats a vMerge continue with nothing above as its own cell', () => {
    const t = table(
      tblGrid(2) +
        tr(tc('A', '<w:vMerge/>'), '<w:gridAfter w:val="1"/>') +
        tr(tc('B') + tc('C', '<w:vMerge/>')),
    );
    expect(gridText(t)).toEqual(['A .', 'B C']);
    expect(slot(t, 0, 0)).toMatchObject({ text: 'A', rowSpan: 1, origin: true });
    expect(slot(t, 1, 1)).toMatchObject({ text: 'C', row: 1, col: 1, origin: true });
  });

  it('honours gridBefore and gridAfter', () => {
    const t = table(
      tblGrid(4) +
        tr(tc('A') + tc('B'), '<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>') +
        tr(tc('C', span(4))),
    );
    expect(t.grid().cols).toBe(4);
    expect(gridText(t)).toEqual(['. A B .', 'C C C C']);
    expect(slot(t, 0, 1)).toMatchObject({ text: 'A', row: 0, col: 1, origin: true });
    expect(t.cell(0, 0)).toBeUndefined();
    expect(t.cell(0, 3)).toBeUndefined();
  });

  it('widens cols beyond tblGrid when rows need more columns', () => {
    const t = table(
      tblGrid(2) + tr(tc('A') + tc('B'), '<w:gridAfter w:val="1"/>') + tr(tc('C', span(4))),
    );
    expect(t.grid().cols).toBe(4);
    expect(gridText(t)).toEqual(['A B . .', 'C C C C']);
    // No tblGrid at all: the widest row decides.
    expect(table(tr(tc('X') + tc('Y'))).grid().cols).toBe(2);
    // An empty table.
    const empty = table(tblGrid(3)).grid();
    expect([empty.rows, empty.cols]).toEqual([0, 3]);
  });

  it('caps hostile spans at MAX_GRID_COLUMNS instead of allocating millions of slots', () => {
    const t = table(
      tblGrid(2) +
        tr(tc('A', span(99_999_999)) + tc('B')) +
        tr(tc('C') + tc('D'), '<w:gridBefore w:val="99999999"/>'),
    );
    const g = t.grid();
    expect(g.cols).toBe(MAX_GRID_COLUMNS);
    expect(g.at(0, 0)?.colSpan).toBe(MAX_GRID_COLUMNS);
    expect(t.rows()[0]!.cells()[1]!.text()).toBe('B'); // still listed, just outside the grid
    expect(g.at(1, 0)).toBeUndefined();
  });

  it('keeps short rows short', () => {
    const t = table(tblGrid(3) + tr(tc('A')) + tr(tc('B') + tc('C') + tc('D')));
    expect(gridText(t)).toEqual(['A . .', 'B C D']);
  });

  it('handles legacy hMerge', () => {
    const t = table(
      tblGrid(4) +
        tr(
          tc('A', '<w:hMerge w:val="restart"/>') +
            tc('B', '<w:hMerge/>') +
            tc('C', '<w:hMerge w:val="continue"/>') +
            tc('D'),
        ) +
        tr(tc('E', '<w:hMerge/>') + tc('F', span(2)) + tc('G')),
    );
    expect(gridText(t)).toEqual(['A A A D', 'E F F G']);
    expect(slot(t, 0, 2)).toEqual({
      text: 'A',
      row: 0,
      col: 0,
      rowSpan: 1,
      colSpan: 3,
      origin: false,
    });
    expect(slot(t, 0, 3)).toMatchObject({ text: 'D', col: 3, origin: true });
    // A leading hMerge continue has nothing to extend.
    expect(slot(t, 1, 0)).toMatchObject({ text: 'E', colSpan: 1, origin: true });
  });

  it('combines hMerge and vMerge', () => {
    const t = table(
      tblGrid(3) +
        tr(
          tc('A', '<w:hMerge w:val="restart"/><w:vMerge w:val="restart"/>') +
            tc('B', '<w:hMerge/><w:vMerge w:val="restart"/>') +
            tc('X'),
        ) +
        tr(
          tc('C', '<w:hMerge w:val="restart"/><w:vMerge/>') +
            tc('D', '<w:hMerge/><w:vMerge/>') +
            tc('Y'),
        ),
    );
    expect(gridText(t)).toEqual(['A A X', 'A A Y']);
    expect(slot(t, 0, 0)).toMatchObject({ rowSpan: 2, colSpan: 2, origin: true });
    expect(slot(t, 1, 1)).toMatchObject({ text: 'A', row: 0, col: 0, origin: false });
  });

  it('includes sdt-wrapped rows and cells in the grid', () => {
    const t = table(
      tblGrid(2) +
        tr(tc('A', '<w:vMerge w:val="restart"/>') + tc('B')) +
        `<w:sdt><w:sdtContent>${tr(
          `<w:sdt><w:sdtContent>${tc('C', '<w:vMerge/>')}</w:sdtContent></w:sdt>${tc('D')}`,
        )}</w:sdtContent></w:sdt>`,
    );
    expect(gridText(t)).toEqual(['A B', 'A D']);
    expect(slot(t, 0, 0).rowSpan).toBe(2);
  });

  it('finds nested tables without flattening them into the outer grid', () => {
    const inner = `<w:tbl>${tblGrid(2)}${tr(tc('n1') + tc('n2'))}${tr(tc('n3', span(2)))}</w:tbl>`;
    const t = table(
      tblGrid(2) +
        tr(`<w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p>${inner}<w:p/></w:tc>${tc('B')}`) +
        tr(`${tc('C')}<w:tc><w:sdt><w:sdtContent>${inner}</w:sdtContent></w:sdt><w:p/></w:tc>`),
    );
    expect(t.rows()).toHaveLength(2);
    expect(t.grid().rows).toBe(2);
    const cell = t.cell(0, 0)!;
    const nested = cell.tables();
    expect(nested).toHaveLength(1);
    expect(nested[0]!.element.parent).toBe(cell.element);
    expect(gridText(nested[0]!)).toEqual(['n1 n2', 'n3 n3']);
    expect(nested[0]!.story).toBe(story);
    // sdt-wrapped nested table
    expect(t.cell(1, 1)!.tables()).toHaveLength(1);
    expect(cell.paragraphs()).toHaveLength(2);
    expect(t.cell(0, 1)!.tables()).toHaveLength(0);
  });

  it('recomputes the grid after DOM changes', () => {
    const t = table(tblGrid(2) + tr(tc('A') + tc('B')));
    expect(t.grid().rows).toBe(1);
    const row = t.rows()[0]!.element;
    t.element.append(row.clone());
    expect(t.grid().rows).toBe(2);
  });

  it('is fast on a large table', () => {
    const rows = Array.from({ length: 300 }, (_, r) =>
      tr(
        Array.from({ length: 20 }, (_, c) =>
          tc(
            `${r}.${c}`,
            r > 0 && c === 0
              ? '<w:vMerge/>'
              : r === 0 && c === 0
                ? '<w:vMerge w:val="restart"/>'
                : '',
          ),
        ).join(''),
      ),
    ).join('');
    const t = table(tblGrid(20) + rows);
    const g = t.grid();
    expect([g.rows, g.cols]).toEqual([300, 20]);
    expect(g.at(0, 0)!.rowSpan).toBe(300);
    expect(g.at(299, 0)!.cell.element.text).toBe('0.0');
    expect(g.at(299, 19)!.cell.element.text).toBe('299.19');
  });
});

// ---- through the public API (needs the docx reader) -------------------------------------

const enc = new TextEncoder();
const DOC_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

function docx(body: string): Uint8Array {
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
        `<Relationship Id="rId1" Type="${DOC_REL}" Target="word/document.xml"/>` +
        '</Relationships>',
    ),
    'word/document.xml': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`,
    ),
  });
}

describe('Table through readDocx', () => {
  it('reads cell and table text, nested tables included', async () => {
    const { readDocx } = await import('../src/index.js');
    const inner = `<w:tbl>${tblGrid(2)}${tr(tc('n1') + tc('n2'))}</w:tbl>`;
    const doc = readDocx(
      docx(
        `<w:p><w:r><w:t>before</w:t></w:r></w:p><w:tbl>${tblGrid(2)}` +
          tr(tc('A', span(2))) +
          tr(
            `<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>` +
              `<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p>${inner}<w:p/></w:tc>`,
          ) +
          '</w:tbl>',
      ),
    );
    const [t] = doc.body.tables();
    expect(t).toBeInstanceOf(Table);
    expect(t!.story).toBe(doc.body);
    const b = t!.cell(1, 0)!;
    expect(b.paragraphs().map((p) => p.text)).toEqual(['B1', 'B2']);
    expect(b.paragraphs()[0]!.story).toBe(doc.body);
    expect(b.text()).toBe('B1\nB2');
    expect(t!.cell(1, 1)!.text()).toBe('C\nn1\tn2\n');
    expect(t!.text()).toBe('A\nB1\nB2\tC\nn1\tn2\n');
    expect(t!.cell(0, 1)!.text()).toBe('A');
  });
});

describe('Table fixtures', () => {
  const file = fileURLToPath(
    new URL('./fixtures/python-docx/tbl-cell-access.docx', import.meta.url),
  );
  let bytes: Uint8Array | undefined;
  try {
    bytes = readFileSync(file);
  } catch {
    bytes = undefined;
  }

  it.skipIf(bytes === undefined)(
    'python-docx tbl-cell-access: spans and vertical merges',
    async () => {
      const { readDocx } = await import('../src/index.js');
      const tables = readDocx(bytes!).body.tables();
      const show = (t: Table): string[] => {
        const g = t.grid();
        const out: string[] = [];
        for (let r = 0; r < g.rows; r++) {
          const line: string[] = [];
          for (let c = 0; c < g.cols; c++) {
            const s = g.at(r, c);
            line.push(s === undefined ? '.' : `${s.cell.text()}${s.origin ? '' : '^'}`);
          }
          out.push(line.join(' '));
        }
        return out;
      };
      expect(tables.map(show)).toEqual([
        ['1 2 3', '4 5 6', '7 8 9'],
        ['1 2 3', '4 4^ 6', '7 8 9'],
        ['1 2 3', '4 5 6', '7 5^ 9'],
        ['1 2 3', '4 4^ 6', '4^ 4^ 9'],
      ]);
      expect(tables[3]!.grid().at(2, 1)).toMatchObject({ row: 1, col: 0, rowSpan: 2, colSpan: 2 });
      expect(tables[3]!.text()).toBe('1\t2\t3\n4\t6\n\t9');
    },
  );
});
