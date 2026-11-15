// acceptAllRevisions() / rejectAllRevisions(): every revision kind on synthetic documents, then
// Apache POI's delins.docx (Word-made: deleted paragraphs, deleted field codes, insertions,
// formatting changes) against a text-level reference of what accepting/rejecting must show.

import { describe, expect, it } from 'vitest';
import { NS, paragraphText, readDocx, type WordDocument, type XElement } from '../src/index.js';
import { acceptAllRevisions, rejectAllRevisions } from '../src/text/index.js';
import { REVISION_ELEMENTS } from '../src/wml/ids.js';
import { blockParagraphs } from '../src/wml/text.js';
import { fixture } from './helpers/fixtures.js';
import {
  docWith,
  docxFrom,
  expectSane,
  originalText,
  p,
  partRoots,
  r,
  revisionsLeft,
  roundTrip,
  xml,
} from './text-helpers.js';

const A = 'w:author="Ann" w:date="2020-01-01T00:00:00Z"';
const ins = (id: number, inner: string): string => `<w:ins w:id="${id}" ${A}>${inner}</w:ins>`;
const del = (id: number, text: string, rPr = ''): string =>
  `<w:del w:id="${id}" ${A}><w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;

/** Accepts or rejects a fresh document built from `body`; returns it and the count. */
function resolved(body: string, accept: boolean): { doc: WordDocument; n: number } {
  const doc = docWith(body);
  const n = accept ? acceptAllRevisions(doc) : rejectAllRevisions(doc);
  expect(revisionsLeft(doc)).toEqual([]);
  expectSane(doc.body.element);
  roundTrip(doc);
  return { doc, n };
}

const paraTexts = (doc: WordDocument): string[] => doc.body.paragraphs().map((x) => x.text);
const bodyXml = (doc: WordDocument): string =>
  doc.body.element
    .elements()
    .filter((e) => !e.is(NS.w, 'sectPr'))
    .map(xml)
    .join('');

describe('run content: w:ins / w:del', () => {
  const body = p(`${r('keep ')}${ins(1, r('added '))}${del(2, 'removed ')}${r('end')}`);

  it('accept keeps insertions and drops deletions', () => {
    const { doc, n } = resolved(body, true);
    expect(n).toBe(2);
    expect(doc.text()).toBe('keep added end');
    expect(bodyXml(doc)).toBe(
      '<w:p><w:r><w:t xml:space="preserve">keep </w:t></w:r><w:r><w:t xml:space="preserve">added </w:t></w:r><w:r><w:t xml:space="preserve">end</w:t></w:r></w:p>',
    );
  });

  it('reject drops insertions and revives deletions (delText → t)', () => {
    const { doc, n } = resolved(body, false);
    expect(n).toBe(2);
    expect(doc.text()).toBe('keep removed end');
    expect(doc.body.element.descendants(NS.w, 't').map((t) => t.text)).toEqual([
      'keep ',
      'removed ',
      'end',
    ]);
  });

  it('nested: text inserted by one author and deleted by another', () => {
    const nested = p(`${r('a')}<w:ins w:id="1" ${A}>${del(2, 'b')}</w:ins>${r('c')}`);
    expect(resolved(nested, true).doc.text()).toBe('ac');
    expect(resolved(nested, false).doc.text()).toBe('ac');
  });

  it('keeps bookmarks and comment ranges found in removed content', () => {
    const body2 = p(
      `${r('a')}<w:ins w:id="1" ${A}><w:bookmarkStart w:id="5" w:name="x"/>${r('b')}</w:ins>` +
        `${r('c')}<w:del w:id="2" ${A}><w:bookmarkEnd w:id="5"/><w:commentRangeStart w:id="0"/>` +
        `<w:r><w:delText>d</w:delText></w:r></w:del><w:commentRangeEnd w:id="0"/>`,
    );
    for (const accept of [true, false]) {
      const { doc } = resolved(body2, accept);
      expect(doc.body.element.descendants(NS.w, 'bookmarkStart')).toHaveLength(1);
      expect(doc.body.element.descendants(NS.w, 'bookmarkEnd')).toHaveLength(1);
      expect(doc.body.element.descendants(NS.w, 'commentRangeStart')).toHaveLength(1);
      expect(doc.text()).toBe(accept ? 'abc' : 'acd');
    }
  });

  it('deleted field codes come back on reject and vanish on accept', () => {
    const body3 = p(
      `${r('See ')}<w:del w:id="1" ${A}><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        '<w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:delText>4</w:delText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:del>' +
        '<w:r><w:delInstrText>stray</w:delInstrText><w:delText>x</w:delText></w:r>',
    );
    const acc = resolved(body3, true).doc;
    expect(acc.text()).toBe('See ');
    expect(acc.body.element.descendants(NS.w, 'fldChar')).toHaveLength(0);
    const rej = resolved(body3, false).doc;
    expect(rej.text()).toBe('See 4x');
    expect(rej.body.element.descendants(NS.w, 'instrText').map((e) => e.text)).toEqual([
      ' PAGE ',
      'stray',
    ]);
    expect(rej.body.element.descendants(NS.w, 'fldChar')).toHaveLength(3);
  });
});

describe('paragraph marks', () => {
  const markDel = (id: number): string => `<w:rPr><w:del w:id="${id}" ${A}/><w:b/></w:rPr>`;
  const markIns = (id: number): string => `<w:rPr><w:ins w:id="${id}" ${A}/></w:rPr>`;

  it('accepting a deleted mark merges with the next paragraph, whose properties survive', () => {
    const body =
      p(r('one '), `<w:jc w:val="center"/>${markDel(1)}`) +
      p(r('two'), '<w:jc w:val="right"/>') +
      p(r('three'));
    const { doc, n } = resolved(body, true);
    expect(n).toBe(1);
    expect(paraTexts(doc)).toEqual(['one two', 'three']);
    expect(
      doc.body
        .paragraphs()[0]!
        .element.element(NS.w, 'pPr')!
        .element(NS.w, 'jc')!
        .attrNS(NS.w, 'val'),
    ).toBe('right');
    const rej = resolved(body, false).doc;
    expect(paraTexts(rej)).toEqual(['one ', 'two', 'three']);
    expect(rej.body.element.descendants(NS.w, 'b')).toHaveLength(1);
  });

  it('a fully deleted paragraph disappears; chains merge into the last paragraph', () => {
    const body =
      p(del(1, 'gone'), markDel(2)) +
      p(r('a'), markDel(3)) +
      p(r('b'), markDel(4)) +
      p(r('c'), '<w:pStyle w:val="Heading1"/>');
    const { doc, n } = resolved(body, true);
    expect(n).toBe(4);
    expect(paraTexts(doc)).toEqual(['abc']);
    expect(doc.body.paragraphs()[0]!.style).toBe('Heading1');
  });

  it('rejecting an inserted mark merges paragraphs', () => {
    const body = p(r('first '), markIns(1)) + p(ins(2, r('new')), markIns(3)) + p(r('last'));
    expect(paraTexts(resolved(body, false).doc)).toEqual(['first last']);
    expect(paraTexts(resolved(body, true).doc)).toEqual(['first ', 'new', 'last']);
  });

  it('a mark that cannot merge (table next, end of cell) is dropped, content kept', () => {
    const cell = (inner: string): string =>
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${inner}</w:tc></w:tr></w:tbl>`;
    const body =
      p(r('before table'), markDel(1)) +
      cell(p(r('in cell'), markDel(2))) +
      p('', markDel(3)) +
      cell(p(r('x'))) +
      p(r('end'));
    const { doc } = resolved(body, true);
    expect(paraTexts(doc)).toEqual(['before table', 'in cell', 'x', 'end']);
  });
});

describe('moves', () => {
  const body =
    p(
      `${r('Start ')}<w:moveFromRangeStart w:id="1" ${A} w:name="move1"/>` +
        `<w:moveFrom w:id="2" ${A}>${r('moved ')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>${r('middle ')}`,
    ) +
    p(
      `${r('end ')}<w:moveToRangeStart w:id="3" ${A} w:name="move1"/>` +
        `<w:moveTo w:id="4" ${A}>${r('moved')}</w:moveTo><w:moveToRangeEnd w:id="3"/>`,
    );

  it('accept lands the move, reject returns it; range markers go both ways', () => {
    const acc = resolved(body, true);
    expect(acc.n).toBe(2);
    expect(paraTexts(acc.doc)).toEqual(['Start middle ', 'end moved']);
    const rej = resolved(body, false);
    expect(paraTexts(rej.doc)).toEqual(['Start moved middle ', 'end ']);
  });
});

describe('property changes', () => {
  it('rPrChange: accept keeps the new formatting, reject restores the old', () => {
    const body = p(
      `<w:r><w:rPr><w:b/><w:sz w:val="28"/><w:rPrChange w:id="1" ${A}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r>`,
      `<w:rPr><w:ins w:id="2" ${A}/><w:u w:val="single"/><w:rPrChange w:id="3" ${A}><w:rPr><w:ins w:id="9" ${A}/><w:strike/></w:rPr></w:rPrChange></w:rPr>`,
    );
    const acc = resolved(body, true);
    expect(acc.n).toBe(3);
    expect(xml(acc.doc.body.element.descendants(NS.w, 'r')[0]!.element(NS.w, 'rPr')!)).toBe(
      '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr>',
    );
    const rej = resolved(body, false).doc;
    expect(xml(rej.body.element.descendants(NS.w, 'r')[0]!.element(NS.w, 'rPr')!)).toBe(
      '<w:rPr><w:i/></w:rPr>',
    );
  });

  it('pPrChange restores old paragraph properties, keeping the mark rPr and sectPr', () => {
    const body =
      p(
        r('x'),
        '<w:jc w:val="center"/><w:rPr><w:b/></w:rPr>' +
          `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr><w:pPrChange w:id="1" ${A}><w:pPr><w:keepNext/><w:spacing w:after="0"/></w:pPr></w:pPrChange>`,
      ) + p(r('y'));
    const rej = resolved(body, false).doc;
    expect(xml(rej.body.paragraphs()[0]!.element.element(NS.w, 'pPr')!)).toBe(
      '<w:pPr><w:keepNext/><w:spacing w:after="0"/><w:rPr><w:b/></w:rPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>',
    );
    const acc = resolved(body, true).doc;
    expect(xml(acc.body.paragraphs()[0]!.element.element(NS.w, 'pPr')!)).toContain(
      '<w:jc w:val="center"/><w:rPr><w:b/></w:rPr><w:sectPr>',
    );
  });

  it('sectPrChange restores old section properties, keeping header/footer references', () => {
    const bytes = docxFrom({ body: p(r('x')), header: p(r('h')) });
    const doc = readDocx(bytes);
    const sectPr = doc.body.element.element(NS.w, 'sectPr')!;
    sectPr.append(
      parseChange(
        `<w:sectPrChange w:id="7" ${A}><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:sectPrChange>`,
      ),
    );
    const rej = doc.fork();
    expect(rejectAllRevisions(rej)).toBe(1);
    expect(xml(rej.body.element.element(NS.w, 'sectPr')!)).toBe(
      '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
    );
    roundTrip(rej);
    const acc = doc.fork();
    expect(acceptAllRevisions(acc)).toBe(1);
    expect(xml(acc.body.element.element(NS.w, 'sectPr')!)).not.toContain('sectPrChange');
  });

  const table = (tblPr: string, grid: string, trPr: string, tcPr: string, ex = ''): string =>
    `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>` +
    `<w:tr>${ex}<w:trPr>${trPr}</w:trPr><w:tc><w:tcPr>${tcPr}</w:tcPr>${p(r('cell'))}</w:tc></w:tr></w:tbl>${p('')}`;

  it('table property changes (tblPr, tblPrEx, tblGrid, trPr, tcPr) restore on reject', () => {
    const body = table(
      `<w:tblW w:w="5000" w:type="pct"/><w:tblPrChange w:id="1" ${A}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange>`,
      `<w:gridCol w:w="4000"/><w:tblGridChange w:id="2"><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid></w:tblGridChange>`,
      `<w:jc w:val="center"/><w:ins w:id="3" ${A}/><w:trPrChange w:id="4" ${A}><w:trPr><w:cantSplit/></w:trPr></w:trPrChange>`,
      `<w:tcW w:w="4000" w:type="dxa"/><w:shd w:val="clear" w:fill="FF0000"/><w:tcPrChange w:id="5" ${A}><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr></w:tcPrChange>`,
      `<w:tblPrEx><w:jc w:val="right"/><w:tblPrExChange w:id="6" ${A}><w:tblPrEx><w:jc w:val="left"/></w:tblPrEx></w:tblPrExChange></w:tblPrEx>`,
    );
    const acc = resolved(body, true);
    expect(acc.n).toBe(6);
    const tbl = acc.doc.body.element.element(NS.w, 'tbl')!;
    expect(xml(tbl.element(NS.w, 'tblPr')!)).toBe(
      '<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>',
    );
    expect(xml(tbl.descendants(NS.w, 'trPr')[0]!)).toBe('<w:trPr><w:jc w:val="center"/></w:trPr>');
    // Rejecting the row insertion removes the only row, hence the table.
    const rej = resolved(body, false);
    expect(rej.n).toBe(6);
    expect(rej.doc.body.element.descendants(NS.w, 'tbl')).toHaveLength(0);
    // Without the row insertion, the old properties come back.
    const keepRow = body.replace(`<w:ins w:id="3" ${A}/>`, '');
    const rej2 = resolved(keepRow, false).doc.body.element.element(NS.w, 'tbl')!;
    expect(xml(rej2.element(NS.w, 'tblPr')!)).toBe(
      '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>',
    );
    expect(xml(rej2.element(NS.w, 'tblGrid')!)).toBe(
      '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>',
    );
    expect(xml(rej2.descendants(NS.w, 'trPr')[0]!)).toBe('<w:trPr><w:cantSplit/></w:trPr>');
    expect(xml(rej2.descendants(NS.w, 'tcPr')[0]!)).toBe(
      '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>',
    );
    expect(xml(rej2.descendants(NS.w, 'tblPrEx')[0]!)).toBe(
      '<w:tblPrEx><w:jc w:val="left"/></w:tblPrEx>',
    );
  });

  it('numberingChange is dropped both ways; numPr/ins keeps or removes the numbering', () => {
    const body = p(
      r('item'),
      `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:numberingChange w:id="1" ${A} w:original="%1."/><w:ins w:id="2" ${A}/></w:numPr>`,
    );
    const acc = resolved(body, true);
    expect(acc.n).toBe(2);
    expect(xml(acc.doc.body.paragraphs()[0]!.element.element(NS.w, 'pPr')!)).toBe(
      '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>',
    );
    const rej = resolved(body, false);
    expect(rej.doc.body.paragraphs()[0]!.element.element(NS.w, 'pPr')).toBeUndefined();
  });
});

describe('rows and cells', () => {
  const tc = (text: string, tcPr = ''): string =>
    `<w:tc><w:tcPr>${tcPr}</w:tcPr>${p(r(text))}</w:tc>`;
  const tr = (cells: string, trPr = ''): string =>
    `<w:tr>${trPr === '' ? '' : `<w:trPr>${trPr}</w:trPr>`}${cells}</w:tr>`;
  const tbl = (rows: string): string =>
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl>${p('')}`;
  const cellTexts = (doc: WordDocument): string[] =>
    doc.body.element.descendants(NS.w, 'tr').map((row) =>
      row
        .elements(NS.w, 'tc')
        .map((c) => paragraphText(c))
        .join('|'),
    );

  it('inserted and deleted rows', () => {
    const body = tbl(
      tr(tc('a') + tc('b')) +
        tr(tc('new') + tc('row'), `<w:ins w:id="1" ${A}/>`) +
        tr(tc('old') + tc('row'), `<w:del w:id="2" ${A}/>`),
    );
    expect(cellTexts(resolved(body, true).doc)).toEqual(['a|b', 'new|row']);
    expect(cellTexts(resolved(body, false).doc)).toEqual(['a|b', 'old|row']);
  });

  it('a table that loses every row goes; a cell keeps a final paragraph', () => {
    const inner = tbl(tr(tc('nested'), `<w:del w:id="1" ${A}/>`)).replace(/<w:p><\/w:p>$/, '');
    const body = tbl(tr(`<w:tc><w:tcPr/>${p(r('x'))}${inner}</w:tc>${tc('y')}`));
    const doc = resolved(body, true).doc;
    expect(doc.body.element.descendants(NS.w, 'tbl')).toHaveLength(1);
    const cell = doc.body.element.descendants(NS.w, 'tc')[0]!;
    expect(cell.lastElement()!.is(NS.w, 'p')).toBe(true);
  });

  it('cellIns, cellDel and cellMerge', () => {
    const body = tbl(
      tr(tc('a') + tc('ins', `<w:cellIns w:id="1" ${A}/>`)) +
        tr(tc('b') + tc('del', `<w:cellDel w:id="2" ${A}/>`)) +
        tr(tc('c', `<w:cellMerge w:id="3" ${A} w:vMerge="rest"/>`) + tc('d')) +
        tr(
          tc('e', `<w:vMerge/><w:cellMerge w:id="4" ${A} w:vMerge="cont" w:vMergeOrig="rest"/>`) +
            tc('f'),
        ),
    );
    const acc = resolved(body, true);
    expect(acc.n).toBe(4);
    expect(cellTexts(acc.doc)).toEqual(['a|ins', 'b', 'c|d', 'e|f']);
    const vm = acc.doc.body.element
      .descendants(NS.w, 'vMerge')
      .map((v) => v.attrNS(NS.w, 'val') ?? '');
    expect(vm).toEqual(['restart', '']);
    const rej = resolved(body, false);
    expect(cellTexts(rej.doc)).toEqual(['a', 'b|del', 'c|d', 'e|f']);
    const vm2 = rej.doc.body.element
      .descendants(NS.w, 'vMerge')
      .map((v) => v.attrNS(NS.w, 'val') ?? '');
    expect(vm2).toEqual(['restart']);
  });
});

describe('targets', () => {
  const bytes = docxFrom({
    body: p(ins(1, r('body'))),
    header: p(ins(2, r('header'))),
    footnotes: `<w:footnote w:id="1">${p(ins(3, r('note')))}</w:footnote>`,
    comments: `<w:comment w:id="0" w:author="a">${p(ins(4, r('comment')))}</w:comment>`,
    styles:
      '<w:style w:type="paragraph" w:styleId="X"><w:name w:val="X"/>' +
      `<w:rPr><w:b/><w:rPrChange w:id="5" ${A}><w:rPr/></w:rPrChange></w:rPr></w:style>`,
  });

  it('a document target resolves every story part, styles included', () => {
    const doc = readDocx(bytes);
    expect(rejectAllRevisions(doc)).toBe(5);
    expect(revisionsLeft(doc)).toEqual([]);
    expect(doc.stories().map((s) => s.text())).toEqual(['', '', '', '']);
    roundTrip(doc, 'revisions-all-stories.docx');
  });

  it('a story target resolves only that story', () => {
    const doc = readDocx(bytes);
    const header = doc.stories(['header'])[0]!;
    expect(acceptAllRevisions(header)).toBe(1);
    expect(revisionsLeft(doc)).toHaveLength(4);
    expect(acceptAllRevisions(doc)).toBe(4);
    expect(doc.stories().map((s) => s.text())).toEqual(['body', 'header', 'note', 'comment']);
  });
});

// ---- delins.docx ----------------------------------------------------------------------------

/** True when the paragraph's mark carries one of `kinds` (pPr/rPr/ins|del|moveFrom|moveTo). */
function markIs(el: XElement, kinds: string[]): boolean {
  const rPr = el.element(NS.w, 'pPr')?.element(NS.w, 'rPr');
  return rPr?.elements(NS.w).some((k) => kinds.includes(k.local)) === true;
}

/**
 * Reference texts of a story after accept/reject: each paragraph's accepted (visible) or
 * original text; a paragraph whose mark goes is joined to the following sibling paragraph.
 */
function expectedTexts(root: XElement, accept: boolean): string[] {
  const out: string[] = [];
  let carry = '';
  for (const el of blockParagraphs(root)) {
    const text = carry + (accept ? paragraphText(el) : originalText(el));
    carry = '';
    const goes = markIs(el, accept ? ['del', 'moveFrom'] : ['ins', 'moveTo']);
    if (goes && el.nextElement()?.is(NS.w, 'p')) carry = text;
    else out.push(text);
  }
  return out;
}

describe('apache-poi/delins.docx', () => {
  const bytes = fixture('apache-poi/delins.docx').bytes;
  const initial = countRevisions(readDocx(bytes));

  it('accept: no revision left, text = the inserted view, paragraphs merged', () => {
    const doc = readDocx(bytes);
    const want = expectedTexts(doc.body.element, true);
    expect(want.length).toBeLessThan(doc.body.paragraphs().length);
    expect(acceptAllRevisions(doc)).toBe(initial);
    expect(revisionsLeft(doc)).toEqual([]);
    expect(paraTexts(doc)).toEqual(want);
    expectSane(doc.body.element);
    const back = roundTrip(doc, 'delins-accepted.docx');
    expect(paraTexts(back)).toEqual(want);
  });

  it('reject: no revision left, text = the original view, deleted field codes revived', () => {
    const doc = readDocx(bytes);
    const want = expectedTexts(doc.body.element, false);
    const notInserted = (e: XElement): boolean =>
      e.closest(NS.w, 'ins') === undefined && e.closest(NS.w, 'moveTo') === undefined;
    const delInstr = doc.body.element.descendants(NS.w, 'delInstrText').filter(notInserted).length;
    const instr = doc.body.element.descendants(NS.w, 'instrText').filter(notInserted).length;
    expect(delInstr).toBeGreaterThan(0);
    expect(rejectAllRevisions(doc)).toBe(initial);
    expect(revisionsLeft(doc)).toEqual([]);
    expect(paraTexts(doc)).toEqual(want);
    expect(doc.body.element.descendants(NS.w, 'instrText')).toHaveLength(instr + delInstr);
    expectSane(doc.body.element);
    const back = roundTrip(doc, 'delins-rejected.docx');
    expect(paraTexts(back)).toEqual(want);
  });

  it('accept then reject is a no-op (nothing left to resolve)', () => {
    const doc = readDocx(bytes);
    acceptAllRevisions(doc);
    const before = bodyXml(doc);
    expect(rejectAllRevisions(doc)).toBe(0);
    expect(bodyXml(doc)).toBe(before);
  });
});

/**
 * Revisions to resolve: revision elements outside the old-state records of property changes
 * (Word stores the old paragraph-mark `w:ins` inside `w:rPrChange`), range markers excluded.
 */
function countRevisions(doc: WordDocument): number {
  let n = 0;
  for (const { root } of partRoots(doc)) {
    const stack: XElement[] = [root];
    while (stack.length > 0) {
      const el = stack.pop()!;
      if (el !== root && el.uri === NS.w && REVISION_ELEMENTS.has(el.local)) {
        if (!el.local.includes('Range')) n++;
        if (el.local.endsWith('Change')) continue;
      }
      stack.push(...el.elements());
    }
  }
  return n;
}

function parseChange(x: string): XElement {
  return docWith(p(`<w:r><w:rPr>${x}</w:rPr></w:r>`))
    .body.element.descendants(NS.w, 'rPr')[0]!
    .firstElement()!;
}
