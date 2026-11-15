// The text engine: flatten(), findText() and every rule of replaceText() (plan: "Text engine
// (find/replace across runs)"). The split-run corpus lives in text-corpus.test.ts, the real
// fixtures in text-fixtures.test.ts, accept/reject in revisions.test.ts.

import { describe, expect, it } from 'vitest';
import {
  NS,
  type Paragraph,
  paragraphText,
  readDocx,
  TurbowordError,
  UsageError,
  type Warning,
} from '../src/index.js';
import {
  acceptAllRevisions,
  findText,
  flatten,
  flattenElement,
  rejectAllRevisions,
  replaceText,
  type TextMatch,
} from '../src/text/index.js';
import {
  docWith,
  docxFrom,
  expectSane,
  originalText,
  p,
  r,
  revisionsLeft,
  roundTrip,
  xml,
} from './text-helpers.js';

const first = (doc: { body: { paragraphs(): Paragraph[] } }): Paragraph =>
  doc.body.paragraphs()[0]!;
const runs = (para: Paragraph): string[] =>
  para.element.descendants(NS.w, 'r').map((x) => paragraphText(x));
const rPrOf = (el: XElementLike | undefined): string => {
  const rPr = el?.element(NS.w, 'rPr');
  return rPr ? xml(rPr) : '';
};
type XElementLike = ReturnType<typeof first>['element'];

/** The run holding visible character `index` of `para`. */
function runAt(para: Paragraph, index: number): XElementLike {
  return flatten(para).at(index)!.node.parent as XElementLike;
}

const FIELD = (instr: string, result: string): string =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  result +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

const TEXTBOX = (text: string, alt = text): string =>
  '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Text Box 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr/><wps:txbx><w:txbxContent>${p(r(text))}</w:txbxContent></wps:txbx>` +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>' +
  `<mc:Fallback><w:pict><v:shape id="tb1" style="width:72pt;height:72pt"><v:textbox><w:txbxContent>${p(r(alt))}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` +
  '</mc:AlternateContent></w:r>';

// ---- flatten ------------------------------------------------------------------------------

describe('flatten()', () => {
  it('maps every visible character to its leaf and offset (UTF-16)', () => {
    const doc = docWith(
      p(
        `${r('ab', '<w:b/>')}<w:r><w:tab/><w:t>c😀d</w:t><w:br/></w:r><w:r><w:sym w:font="Symbol" w:char="03B1"/></w:r>`,
      ),
    );
    const f = flatten(first(doc));
    expect(f.text).toBe('ab\tc😀d\nα');
    expect(f.text.length).toBe(9);
    expect(f.at(1)).toMatchObject({ offset: 1 });
    expect(f.at(1)!.node.local).toBe('t');
    expect(f.at(2)!.node.local).toBe('tab');
    // The emoji is 2 code units: offsets 1 and 2 of "c😀d".
    expect(f.at(4)).toMatchObject({ offset: 1 });
    expect(f.at(5)).toMatchObject({ offset: 2 });
    expect(f.at(6)).toMatchObject({ offset: 3 });
    expect(f.at(7)!.node.local).toBe('br');
    expect(f.at(8)!.node.local).toBe('sym');
    expect(f.at(9)).toBeUndefined();
    expect(f.at(-1)).toBeUndefined();
  });

  it('records zero-width markers and field boundaries with their positions', () => {
    const doc = docWith(
      p(
        `<w:bookmarkStart w:id="1" w:name="a"/>${r('ab')}<w:proofErr w:type="spellStart"/>` +
          `<w:commentRangeStart w:id="2"/>${r('cd')}<w:commentRangeEnd w:id="2"/>` +
          `<w:proofErr w:type="spellEnd"/><w:permStart w:id="3" w:edGrp="everyone"/>` +
          `${FIELD('PAGE', r('7'))}<w:permEnd w:id="3"/>` +
          `<w:fldSimple w:instr=" NUMPAGES ">${r('9')}</w:fldSimple><w:bookmarkEnd w:id="1"/>`,
      ),
    );
    const f = flattenElement(first(doc).element);
    expect(f.text).toBe('abcd79');
    expect(f.markers.map((m) => `${m.node.local}@${m.pos}:${m.kind}`)).toEqual([
      'bookmarkStart@0:start',
      'proofErr@2:proof',
      'commentRangeStart@2:start',
      'commentRangeEnd@4:end',
      'proofErr@4:proof',
      'permStart@4:start',
      'fldChar@4:field',
      'fldChar@4:field',
      'fldChar@5:field',
      'permEnd@5:end',
      'fldSimple@5:field',
      'fldSimple@6:field',
      'bookmarkEnd@6:end',
    ]);
  });

  it('ignores markers inside deleted content and other skipped subtrees', () => {
    const doc = docWith(
      p(
        `${r('a')}<w:del w:id="9" w:author="x"><w:bookmarkStart w:id="1" w:name="b"/>` +
          `<w:r><w:delText>gone</w:delText></w:r></w:del>${r('b')}<w:bookmarkEnd w:id="1"/>`,
      ),
    );
    const f = flattenElement(first(doc).element);
    expect(f.text).toBe('ab');
    expect(f.markers.map((m) => m.node.local)).toEqual(['bookmarkEnd']);
  });
});

// ---- findText -------------------------------------------------------------------------------

describe('findText()', () => {
  const doc = docWith(
    p(
      `${r('Say Won', '<w:b/>')}<w:hyperlink w:anchor="x">${r('der')}</w:hyperlink>` +
        `<w:smartTag w:uri="u" w:element="e">${r('f')}</w:smartTag>` +
        `<w:sdt><w:sdtPr/><w:sdtContent>${r('u')}</w:sdtContent></w:sdt>` +
        `<w:fldSimple w:instr=" X ">${r('l')}</w:fldSimple>${r(' now, wonderful')}`,
    ) +
      p(r('Wonder')) +
      p(r('ful wonderful')),
  );

  it('matches across run, hyperlink, smart tag, content control and field boundaries', () => {
    const ms = findText(doc, 'Wonderful');
    expect(ms.map((m) => [m.start, m.end, m.text])).toEqual([[4, 13, 'Wonderful']]);
    expect(ms[0]!.story).toBe(doc.body);
    expect(ms[0]!.paragraph.element).toBe(first(doc).element);
  });

  it('never matches across paragraphs', () => {
    expect(findText(doc, 'Wonderful', { caseSensitive: false })).toHaveLength(3);
    expect(
      findText(doc, 'Wonderful', { caseSensitive: false }).map((m) => m.paragraph.text),
    ).toEqual(['Say Wonderful now, wonderful', 'Say Wonderful now, wonderful', 'ful wonderful']);
  });

  it('is case-sensitive for strings unless told otherwise; RegExp keeps its flags', () => {
    expect(findText(doc, 'wonderful')).toHaveLength(2);
    expect(findText(doc, /WONDERFUL/)).toHaveLength(0);
    expect(findText(doc, /WONDERFUL/i)).toHaveLength(3);
    expect(findText(doc, /WONDERFUL/, { caseSensitive: false })).toHaveLength(3);
    expect(findText(doc, /wonderful/i, { caseSensitive: true })).toHaveLength(2);
    expect(findText(doc, /wonderful/gy)).toHaveLength(2);
  });

  it('reports capture groups, non-overlapping global matches and skips empty matches', () => {
    const d = docWith(p(r('aaaa x-1 y-22')));
    expect(findText(d, 'aa').map((m) => m.start)).toEqual([0, 2]);
    const ms = findText(d, /(\w)-(\d+)/);
    expect(ms.map((m) => m.groups)).toEqual([
      ['x', '1'],
      ['y', '22'],
    ]);
    expect(findText(d, /z*/)).toEqual([]);
    expect(findText(d, /(?<l>[xy])-/).map((m) => m.text)).toEqual(['x-', 'y-']);
  });

  it('wholeWord rejects matches inside words (Unicode-aware)', () => {
    const d = docWith(p(r('cat concat cat_ café cat. Cat')));
    expect(findText(d, 'cat', { wholeWord: true }).map((m) => m.start)).toEqual([0, 21]);
    expect(findText(d, 'caf', { wholeWord: true })).toEqual([]);
    expect(findText(d, 'café', { wholeWord: true }).map((m) => m.start)).toEqual([16]);
    expect(findText(d, /c\w+/, { wholeWord: true }).map((m) => m.text)).toEqual([
      'cat',
      'concat',
      'cat_',
      'cat',
    ]);
    expect(findText(d, 'cat', { wholeWord: true, caseSensitive: false })).toHaveLength(3);
  });

  it('throws on an empty search string', () => {
    expect(() => findText(doc, '')).toThrow(UsageError);
  });

  it('searches a story or a paragraph target', () => {
    expect(findText(doc.body, /wonder/i)).toHaveLength(4);
    expect(findText(doc.body.paragraphs()[2]!, /wonder/i)).toHaveLength(1);
  });
});

describe('findText() stories', () => {
  const bytes = docxFrom({
    body: p(r('body key')) + p(TEXTBOX('box key')) + p(r('after key')),
    header: p(r('header key')),
    footer: p(r('footer key')),
    footnotes:
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      `<w:footnote w:id="1">${p(r('footnote key'))}</w:footnote>`,
    endnotes: `<w:endnote w:id="1">${p(r('endnote key'))}</w:endnote>`,
    comments: `<w:comment w:id="0" w:author="a">${p(r('comment key'))}</w:comment>`,
  });

  it("searches the body by default and every story with 'all', in document order", () => {
    const doc = readDocx(bytes);
    expect(findText(doc, 'key').map((m) => m.paragraph.text)).toEqual(['body key', 'after key']);
    const all = findText(doc, 'key', { in: 'all' });
    expect(all.map((m) => `${m.story.kind}:${m.paragraph.text}`)).toEqual([
      'body:body key',
      'body:box key',
      'body:after key',
      'header:header key',
      'footer:footer key',
      'footnote:footnote key',
      'endnote:endnote key',
      'comment:comment key',
    ]);
  });

  it('selects story kinds and text boxes', () => {
    const doc = readDocx(bytes);
    const kinds = (o: Parameters<typeof findText>[2]): string[] =>
      findText(doc, 'key', o).map((m) => m.paragraph.text);
    expect(kinds({ in: ['footnote', 'comment'] })).toEqual(['footnote key', 'comment key']);
    expect(kinds({ in: ['body', 'textbox'] })).toEqual(['body key', 'box key', 'after key']);
    expect(kinds({ in: ['textbox'] })).toEqual(['box key']);
    expect(kinds({ in: [] })).toEqual([]);
  });

  it('reports a text box once, from the branch every reader uses (the Fallback)', () => {
    const doc = docWith(p(TEXTBOX('choice text', 'fallback text')));
    expect(findText(doc, 'text', { in: 'all' }).map((m) => m.paragraph.text)).toEqual([
      'fallback text',
    ]);
  });
});

// ---- replaceText: the plan's rules ----------------------------------------------------------

describe('replaceText() rule 1-2: split boundary runs, write into the formatting run', () => {
  it('splits the boundary runs cloning their rPr and writes into the first run', () => {
    const doc = docWith(
      p(r('Say Won', '<w:b/>') + r('der', '<w:i/>') + r('ful now', '<w:color w:val="FF0000"/>')),
    );
    expect(replaceText(doc, 'Wonderful', 'Great')).toBe(1);
    const para = first(doc);
    expect(para.text).toBe('Say Great now');
    expect(runs(para)).toEqual(['Say Great', ' now']);
    expect(rPrOf(runAt(para, 4))).toBe('<w:rPr><w:b/></w:rPr>');
    expect(rPrOf(runAt(para, 10))).toBe('<w:rPr><w:color w:val="FF0000"/></w:rPr>');
    expectSane(para.element);
  });

  it("formatFrom: 'last' takes the last run's formatting", () => {
    const doc = docWith(
      p(r('Say Won', '<w:b/>') + r('der', '<w:i/>') + r('ful now', '<w:color w:val="FF0000"/>')),
    );
    replaceText(doc, 'Wonderful', 'Great', { formatFrom: 'last' });
    const para = first(doc);
    expect(para.text).toBe('Say Great now');
    expect(runs(para)).toEqual(['Say ', 'Great now']);
    expect(rPrOf(runAt(para, 4))).toBe('<w:rPr><w:color w:val="FF0000"/></w:rPr>');
  });

  it('an edit inside one run leaves one run and one w:t', () => {
    const doc = docWith(p(r('Hello World again', '<w:b/>')));
    replaceText(doc, 'World', 'there');
    const para = first(doc);
    expect(para.element.descendants(NS.w, 'r')).toHaveLength(1);
    expect(para.element.descendants(NS.w, 't').map((t) => t.text)).toEqual(['Hello there again']);
  });

  it('adds xml:space="preserve" where a split leaves edge whitespace', () => {
    const doc = docWith(
      p(
        '<w:r><w:t>Hello World</w:t></w:r><w:bookmarkStart w:id="1" w:name="b"/><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="1"/>',
      ),
    );
    replaceText(doc, 'World', 'X');
    const ts = first(doc).element.descendants(NS.w, 't');
    expect(ts.map((t) => [t.text, t.attr('xml:space') ?? ''])).toEqual([
      ['Hello X', 'preserve'],
      ['x', ''],
    ]);
  });

  it('RunInput[] creates formatted runs; TAB/LF become w:tab/w:br', () => {
    const doc = docWith(p(r('Dear NAME,', '<w:rFonts w:ascii="Arial"/><w:sz w:val="20"/>')));
    replaceText(doc, 'NAME', [
      { text: 'Ada', bold: true, italic: true },
      { text: '\tLove\nlace ', underline: true, color: '#00FF00', size: 14, style: 'Strong' },
    ]);
    const para = first(doc);
    expect(para.text).toBe('Dear Ada\tLove\nlace ,');
    const rs = para.element.elements(NS.w, 'r');
    expect(rs).toHaveLength(4);
    expect(rPrOf(rs[1])).toBe(
      '<w:rPr><w:rFonts w:ascii="Arial"/><w:b/><w:i/><w:sz w:val="20"/></w:rPr>',
    );
    expect(rPrOf(rs[2])).toBe(
      '<w:rPr><w:rStyle w:val="Strong"/><w:rFonts w:ascii="Arial"/><w:color w:val="00FF00"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:u w:val="single"/></w:rPr>',
    );
    expect(rs[2]!.elements().map((k) => k.local)).toEqual(['rPr', 'tab', 't', 'br', 't']);
    expect(rs[2]!.elements(NS.w, 't')[1]!.attr('xml:space')).toBe('preserve');
    expectSane(para.element);
  });

  it('a string replacement keeps TAB/LF as w:tab/w:br too', () => {
    const doc = docWith(p(r('a-b')));
    replaceText(doc, '-', '\t|\n');
    expect(first(doc).text).toBe('a\t|\nb');
    expect(first(doc).element.descendants(NS.w, 'tab')).toHaveLength(1);
    expect(first(doc).element.descendants(NS.w, 'br')).toHaveLength(1);
  });

  it('an empty replacement deletes the text and the runs it empties', () => {
    const doc = docWith(p(r('keep ') + r('drop', '<w:b/>') + r('me', '<w:i/>') + r(' keep')));
    replaceText(doc, 'dropme', '');
    expect(first(doc).text).toBe('keep  keep');
    expect(runs(first(doc))).toEqual(['keep ', ' keep']);
  });

  it('keeps non-text content inside the match (note references, page-break hints)', () => {
    const doc = docWith(
      p(
        `${r('ab')}<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>` +
          '<w:r><w:t>c</w:t><w:lastRenderedPageBreak/><w:t>d</w:t></w:r>',
      ),
    );
    replaceText(doc, 'abcd', 'X');
    const para = first(doc);
    expect(para.text).toBe('X');
    expect(para.element.descendants(NS.w, 'footnoteReference')).toHaveLength(1);
    expect(para.element.descendants(NS.w, 'lastRenderedPageBreak')).toHaveLength(1);
    expectSane(para.element);
  });

  it('removes wrappers a replacement leaves empty, keeps content controls', () => {
    const doc = docWith(
      p(
        `${r('one ')}<w:hyperlink w:anchor="a">${r('two')}</w:hyperlink>` +
          `<w:smartTag w:uri="u" w:element="e"><w:smartTagPr/>${r(' three')}</w:smartTag>` +
          `<w:sdt><w:sdtPr/><w:sdtContent>${r(' four')}</w:sdtContent></w:sdt>${r(' five')}`,
      ),
    );
    replaceText(doc, 'one two three four', '1-4');
    const para = first(doc);
    expect(para.text).toBe('1-4 five');
    expect(para.element.descendants(NS.w, 'hyperlink')).toHaveLength(0);
    expect(para.element.descendants(NS.w, 'smartTag')).toHaveLength(0);
    expect(para.element.descendants(NS.w, 'sdt')).toHaveLength(1);
    roundTrip(doc);
  });

  it('replaces several matches in one run and across runs (processed last to first)', () => {
    const doc = docWith(p(r('xx a x', '<w:b/>') + r('x b xx'), ''));
    expect(replaceText(doc, 'x', 'YY')).toBe(6);
    expect(first(doc).text).toBe('YYYY a YYYY b YYYY');
    expect(runs(first(doc))).toEqual(['YYYY a YY', 'YY b YYYY']);
  });
});

describe('replaceText() rule 3: proofErr, bookmarks and comment ranges', () => {
  it('drops proofErr inside the match and moves start/end markers to the match edges', () => {
    const doc = docWith(
      p(
        `<w:bookmarkStart w:id="1" w:name="outer"/>${r('Say Wo')}<w:proofErr w:type="spellStart"/>` +
          `<w:bookmarkStart w:id="2" w:name="inner"/>${r('nder', '<w:i/>')}<w:bookmarkEnd w:id="1"/>` +
          `<w:commentRangeStart w:id="3"/>${r('fu')}<w:bookmarkEnd w:id="2"/><w:proofErr w:type="spellEnd"/>` +
          `${r('l now')}<w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r>`,
      ),
    );
    replaceText(doc, 'Wonderful', 'Great');
    const para = first(doc);
    expect(para.text).toBe('Say Great now');
    expect(para.element.descendants(NS.w, 'proofErr')).toHaveLength(0);
    const f = flattenElement(para.element);
    const at = Object.fromEntries(
      f.markers.map((m) => [`${m.node.local}:${m.node.attrNS(NS.w, 'id')}`, m.pos]),
    );
    expect(at).toEqual({
      'bookmarkStart:1': 0,
      'bookmarkStart:2': 4, // moved to the match start
      'commentRangeStart:3': 4,
      'bookmarkEnd:1': 9, // moved to the match end
      'bookmarkEnd:2': 9,
      'commentRangeEnd:3': 13,
    });
    expectSane(para.element);
    roundTrip(doc);
  });

  it('a bookmark wholly inside the match ends up around the replacement', () => {
    const doc = docWith(
      p(
        `${r('ab')}<w:bookmarkStart w:id="0" w:name="m"/>${r('cd')}<w:bookmarkEnd w:id="0"/>${r('ef')}`,
      ),
    );
    replaceText(doc, 'bcde', 'XY');
    const f = flattenElement(first(doc).element);
    expect(f.text).toBe('aXYf');
    expect(f.markers.map((m) => `${m.node.local}@${m.pos}`)).toEqual([
      'bookmarkStart@1',
      'bookmarkEnd@3',
    ]);
  });

  it('markers at the match edges stay where they are', () => {
    const doc = docWith(
      p(
        `${r('a')}<w:bookmarkStart w:id="0" w:name="m"/>${r('bc')}<w:bookmarkEnd w:id="0"/>${r('d')}`,
      ),
    );
    replaceText(doc, 'bc', 'X');
    expect(xml(first(doc).element)).toContain(
      '<w:bookmarkStart w:id="0" w:name="m"/><w:r><w:t>X</w:t></w:r><w:bookmarkEnd w:id="0"/>',
    );
  });
});

describe('replaceText() rule 4: fields', () => {
  const crossing = (): ReturnType<typeof docWith> =>
    docWith(p(r('Page ') + FIELD('PAGE', r('3')) + r(' of ') + FIELD('NUMPAGES', r('9'))));

  it('skips a match crossing a field boundary with a warning', () => {
    const doc = crossing();
    const before = xml(first(doc).element);
    const warnings: Warning[] = [];
    expect(replaceText(doc, 'Page 3', 'X', { onWarning: (w) => warnings.push(w) })).toBe(0);
    expect(warnings.map((w) => w.code)).toEqual(['TEXT_FIELD_CROSSING']);
    expect(warnings[0]!.part).toBe('word/document.xml');
    expect(xml(first(doc).element)).toBe(before);
  });

  it("throws with onFieldCrossing: 'throw', before changing anything", () => {
    const doc = docWith(p(r('of of')) + p(r('Page ') + FIELD('PAGE', r('3'))));
    const before = doc.text();
    let err: unknown;
    try {
      replaceText(doc, /of|e 3/, 'X', { onFieldCrossing: 'throw' });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(TurbowordError);
    expect((err as TurbowordError).code).toBe('TEXT_FIELD_CROSSING');
    expect(doc.text()).toBe(before);
  });

  it('replaces inside a field result (complex and simple fields)', () => {
    const doc = docWith(
      p(r('Page ') + FIELD('PAGE', r('3') + r('3', '<w:b/>')) + r(' of ')) +
        p(`<w:fldSimple w:instr=" NUMPAGES ">${r('9')}${r('9')}</w:fldSimple>`),
    );
    expect(replaceText(doc, '33', '4')).toBe(1);
    expect(replaceText(doc, '99', '12')).toBe(1);
    expect(doc.text()).toBe('Page 4 of \n12');
    expect(doc.body.element.descendants(NS.w, 'instrText')[0]!.text).toBe(' PAGE ');
    expect(doc.body.element.descendants(NS.w, 'fldChar')).toHaveLength(3);
    roundTrip(doc);
  });

  it('skips matches crossing a fldSimple edge, or containing a whole field', () => {
    const doc = docWith(
      p(`${r('a')}<w:fldSimple w:instr=" X ">${r('b')}</w:fldSimple>${r('c')}`) +
        p(`${r('d')}${FIELD('PAGE', '')}${r('e')}`),
    );
    const warnings: Warning[] = [];
    expect(replaceText(doc, /ab|bc|abc|de/, 'Z', { onWarning: (w) => warnings.push(w) })).toBe(0);
    expect(warnings).toHaveLength(2);
    expect(doc.text()).toBe('abc\nde');
  });
});

describe('replaceText() rule 5: tracked changes', () => {
  it('emits w:del (w:delText) and w:ins with fresh ids, author and date', () => {
    const doc = docWith(
      p(
        `<w:ins w:id="41" w:author="old"><w:r><w:t>Say </w:t></w:r></w:ins>${r('Won', '<w:b/>')}` +
          `${r('derful now')}`,
      ),
    );
    replaceText(doc, 'Wonderful', 'Great', {
      track: { author: 'Ada', date: new Date(Date.UTC(2024, 1, 29, 12, 30, 15, 250)) },
    });
    const para = first(doc);
    expect(para.text).toBe('Say Great now');
    const del = para.element.element(NS.w, 'del')!;
    const ins = para.element.elements(NS.w, 'ins')[1]!;
    expect(del.attrNS(NS.w, 'author')).toBe('Ada');
    expect(del.attrNS(NS.w, 'date')).toBe('2024-02-29T12:30:15Z');
    expect(ins.attrNS(NS.w, 'author')).toBe('Ada');
    const ids = [del, ins].map((e) => Number(e.attrNS(NS.w, 'id')));
    expect(ids[0]).toBeGreaterThan(41);
    expect(ids[1]).toBeGreaterThan(ids[0]!);
    expect(del.descendants(NS.w, 'delText').map((t) => t.text)).toEqual(['Won', 'derful']);
    expect(del.descendants(NS.w, 't')).toHaveLength(0);
    expect(rPrOf(ins.element(NS.w, 'r'))).toBe('<w:rPr><w:b/></w:rPr>');
    expect(originalText(para.element)).toBe('Wonderful now');
  });

  it('writes no w:date unless one is given', () => {
    const doc = docWith(p(r('a b')));
    replaceText(doc, 'b', 'c', { track: { author: 'x', initials: 'X' } });
    for (const el of [
      ...doc.body.element.descendants(NS.w, 'ins'),
      ...doc.body.element.descendants(NS.w, 'del'),
    ]) {
      expect(el.attrNS(NS.w, 'date')).toBeUndefined();
      expect(el.attrNS(NS.w, 'author')).toBe('x');
    }
  });

  it('a tracked deletion has no w:ins; ids stay unique across several edits', () => {
    const doc = docWith(p(r('one two one two')) + p(r('two')));
    replaceText(doc, 'two', '', { track: { author: 'x' } });
    replaceText(doc, 'one', [{ text: '1', bold: true }], { track: { author: 'y' } });
    const ids = doc.body.element
      .descendants(NS.w)
      .filter((e) => e.local === 'ins' || e.local === 'del')
      .map((e) => e.attrNS(NS.w, 'id'));
    expect(ids).toHaveLength(7);
    expect(new Set(ids).size).toBe(7);
    expect(doc.text()).toBe('1  1 \n');
    expect(doc.body.element.descendants(NS.w, 'ins')[0]!.descendants(NS.w, 'b')).toHaveLength(1);
  });

  it('reopens cleanly, accepts to the new text and rejects to the old one', () => {
    const doc = docWith(
      p(
        `<w:bookmarkStart w:id="0" w:name="bm"/>${r('Dear ')}${r('Mr ', '<w:b/>')}` +
          `<w:proofErr w:type="spellStart"/>${r('Smyth')}<w:proofErr w:type="spellEnd"/>` +
          `<w:bookmarkEnd w:id="0"/>${r(', hello')}`,
      ),
    );
    replaceText(doc, /Mr Smyth/, 'Ms Smith', { track: { author: 'Ada' } });
    const back = roundTrip(doc, 'tracked-replace.docx');
    expect(back.text()).toBe('Dear Ms Smith, hello');
    expectSane(back.body.element);
    const accepted = back.fork();
    expect(acceptAllRevisions(accepted)).toBe(2);
    expect(accepted.text()).toBe('Dear Ms Smith, hello');
    expect(revisionsLeft(accepted)).toEqual([]);
    const rejected = back.fork();
    expect(rejectAllRevisions(rejected)).toBe(2);
    expect(rejected.text()).toBe('Dear Mr Smyth, hello');
    expect(revisionsLeft(rejected)).toEqual([]);
    roundTrip(accepted, 'tracked-replace-accepted.docx');
    roundTrip(rejected, 'tracked-replace-rejected.docx');
  });
});

describe('replaceText() rules 6-8: function replacement, limit, count', () => {
  it('calls a function with each match, in document order', () => {
    const doc = docWith(p(r('{a} and {b}')) + p(r('{c}')));
    const seen: unknown[][] = [];
    const n = replaceText(doc, /\{(\w)\}/, (m: TextMatch) => {
      seen.push([m.text, m.start, m.end, m.paragraph.text, m.story.kind]);
      return m.groups[0]!.toUpperCase().repeat(seen.length);
    });
    expect(n).toBe(3);
    expect(doc.text()).toBe('A and BB\nCCC');
    expect(seen).toEqual([
      ['{a}', 0, 3, '{a} and {b}', 'body'],
      ['{b}', 8, 11, '{a} and {b}', 'body'],
      ['{c}', 0, 3, '{c}', 'body'],
    ]);
  });

  it('a function may return RunInput[]', () => {
    const doc = docWith(p(r('x')));
    replaceText(doc, 'x', () => [
      { text: 'b', bold: true },
      { text: 'i', italic: true },
    ]);
    expect(runs(first(doc))).toEqual(['b', 'i']);
  });

  it('stops after limit replacements and returns the count', () => {
    const doc = docWith(p(r('a a a')) + p(r('a a')));
    expect(replaceText(doc, 'a', 'b', { limit: 4 })).toBe(4);
    expect(doc.text()).toBe('b b b\nb a');
    expect(replaceText(doc, 'a', 'b', { limit: 0 })).toBe(0);
    expect(replaceText(doc, 'zzz', 'b')).toBe(0);
  });

  it('a RegExp string replacement expands $ patterns; a string query is literal', () => {
    const doc = docWith(p(r('John Smith')) + p(r('cost: 5')));
    replaceText(doc, /(?<first>\w+) (\w+)/, '$2, $<first> ($&) $$');
    replaceText(doc, '5', '$1 $&');
    expect(doc.text()).toBe('Smith, John (John Smith) $\ncost: $1 $&');
  });

  it('wholeWord and caseSensitive apply to replace as to find', () => {
    const doc = docWith(p(r('Cat cat concat')));
    expect(replaceText(doc, 'cat', 'dog', { wholeWord: true, caseSensitive: false })).toBe(2);
    expect(doc.text()).toBe('dog dog concat');
  });
});

describe('replaceText() stories and text boxes', () => {
  it("replaces in every story with in: 'all'", () => {
    const doc = readDocx(
      docxFrom({
        body: p(r('body key')) + p(TEXTBOX('box key')),
        header: p(r('header key')),
        footer: p(r('footer key')),
        footnotes: `<w:footnote w:id="1">${p(r('footnote key'))}</w:footnote>`,
        endnotes: `<w:endnote w:id="1">${p(r('endnote key'))}</w:endnote>`,
        comments: `<w:comment w:id="0" w:author="a">${p(r('comment key'))}</w:comment>`,
      }),
    );
    expect(replaceText(doc, 'key', 'KEY', { in: 'all' })).toBe(7);
    expect(findText(doc, 'key', { in: 'all' })).toEqual([]);
    expect(findText(doc, 'KEY', { in: 'all' })).toHaveLength(7);
    const back = roundTrip(doc, 'all-stories.docx');
    expect(back.stories().map((s) => s.text())).toEqual([
      'body KEY\n',
      'header KEY',
      'footer KEY',
      'footnote KEY',
      'endnote KEY',
      'comment KEY',
    ]);
  });

  it('edits both branches of a text box (Word shows the Choice) and counts once', () => {
    const doc = docWith(p(r('anchor ') + TEXTBOX('Hello box', 'Hello box')));
    expect(replaceText(doc, 'Hello', 'Bye', { in: 'all' })).toBe(1);
    const boxes = doc.body.element.descendants(NS.w, 'txbxContent');
    expect(boxes.map((b) => paragraphText(b.element(NS.w, 'p')!))).toEqual(['Bye box', 'Bye box']);
    roundTrip(doc, 'textbox.docx');
  });

  it('leaves a branch alone when it does not read the same', () => {
    const doc = docWith(p(TEXTBOX('Hello choice', 'Hello fallback')));
    expect(replaceText(doc, 'Hello', 'Bye', { in: 'all' })).toBe(1);
    const boxes = doc.body.element.descendants(NS.w, 'txbxContent');
    expect(boxes.map((b) => paragraphText(b))).toEqual(['Hello choice', 'Bye fallback']);
  });

  it('text boxes are only searched on request', () => {
    const doc = docWith(p(TEXTBOX('Hello')));
    expect(replaceText(doc, 'Hello', 'Bye')).toBe(0);
    expect(replaceText(doc.body, 'Hello', 'Bye', { in: ['textbox'] })).toBe(1);
  });

  it('a block-level mc:AlternateContent is edited in every branch', () => {
    const doc = docWith(
      `<mc:AlternateContent><mc:Choice Requires="w14">${p(r('alpha'))}</mc:Choice>` +
        `<mc:Fallback>${p(r('alpha'))}</mc:Fallback></mc:AlternateContent>`,
    );
    expect(replaceText(doc, 'alpha', 'beta')).toBe(1);
    expect(doc.body.element.descendants(NS.w, 'p').map((x) => paragraphText(x))).toEqual([
      'beta',
      'beta',
    ]);
  });
});

describe('replaceText() unsupported content', () => {
  it('skips matches touching math or ruby text, with a warning', () => {
    const doc = docWith(
      p(
        `${r('x = ')}<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>` +
          '<w:r><w:ruby><w:rubyPr/><w:rt><w:r><w:t>kan</w:t></w:r></w:rt>' +
          '<w:rubyBase><w:r><w:t>漢</w:t></w:r></w:rubyBase></w:ruby></w:r>',
      ),
    );
    const warnings: Warning[] = [];
    expect(replaceText(doc, /= x|kan漢/, 'y', { onWarning: (w) => warnings.push(w) })).toBe(0);
    expect(warnings.map((w) => w.code)).toEqual(['TEXT_UNSUPPORTED', 'TEXT_UNSUPPORTED']);
    expect(first(doc).text).toBe('x = xkan漢');
    expect(replaceText(doc, 'x = ', 'X=')).toBe(1);
  });
});
