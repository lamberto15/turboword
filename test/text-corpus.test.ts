// The split-run corpus (plan exit criterion for the text engine: 100% must pass).
//
// One paragraph per case: "Say Wonderful now", with "Wonderful" cut into 2-6 runs at every
// possible set of cut points (218 cut sets), each run formatted differently, combined with each
// decoration below (proofing marks, bookmarks, comment ranges, permissions, hyperlinks, smart
// tags, content controls, custom XML, tracked insertions/deletions, complex and simple fields,
// empty runs, non-text run content, ...). Cases also vary whether "Say " / " now" share the
// first/last run of the word (the match starts/ends mid-run) or sit in runs of their own.
//
// Every case is replaced plain, plain with formatFrom 'last', with RunInput[] and tracked, then
// checked: the text, the formatting of the replacement, markers balanced and moved to the match
// edges, no proofErr left inside, field-crossing cases untouched (with a warning), the document
// valid and reopening identical; tracked output accepts to the new text and rejects to the old.

import { describe, expect, it } from 'vitest';
import { NS, type Paragraph, readDocx, type WordDocument, type XElement } from '../src/index.js';
import {
  acceptAllRevisions,
  flattenElement,
  rejectAllRevisions,
  replaceText,
} from '../src/text/index.js';
import {
  docxFrom,
  expectSane,
  originalText,
  r,
  revisionsLeft,
  roundTrip,
  xml,
} from './text-helpers.js';

const WORD = 'Wonderful';
const SOURCE_TEXT = 'Say Wonderful now';
const NEW_TEXT = 'Say Great now';

const FORMATS = [
  '<w:b/>',
  '<w:i/>',
  '',
  '<w:color w:val="C00000"/>',
  '<w:u w:val="single"/>',
  '<w:sz w:val="28"/><w:szCs w:val="28"/>',
  '<w:b/><w:i/>',
  '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>',
];

/** Every way to cut `n` characters into `k` non-empty pieces (cut positions). */
function cutSets(n: number, k: number): number[][] {
  const out: number[][] = [];
  const pick = (from: number, left: number, acc: number[]): void => {
    if (left === 0) {
      out.push(acc);
      return;
    }
    for (let c = from; c <= n - left; c++) pick(c + 1, left - 1, [...acc, c]);
  };
  pick(1, k - 1, []);
  return out;
}

interface Built {
  xml: string;
  crossing: boolean;
  comment: boolean;
}

/** Decorations: (pieces, case id) → paragraph content. `P` are the piece run XML strings. */
type Decorate = (P: string[], id: number, before: string, after: string) => Partial<Built>;

const bm = (id: number, n: number): string =>
  `<w:bookmarkStart w:id="${id * 4 + n}" w:name="b${id}_${n}"/>`;
const bmEnd = (id: number, n: number): string => `<w:bookmarkEnd w:id="${id * 4 + n}"/>`;
const all = (P: string[]): string => P.join('');
const field = (instr: string, result: string): string =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

const DECORATIONS: Readonly<Record<string, Decorate>> = {
  plain: (P, _id, b, a) => ({ xml: b + all(P) + a }),
  proofErr: (P, _id, b, a) => {
    P[1] = `<w:proofErr w:type="spellStart"/>${P[1]}`;
    P[P.length - 1] += '<w:proofErr w:type="spellEnd"/>';
    return {
      xml: `<w:proofErr w:type="gramStart"/>${b}${all(P)}${a}<w:proofErr w:type="gramEnd"/>`,
    };
  },
  bookmarkInside: (P, id, b, a) => {
    P[P.length - 1] = bmEnd(id, 0) + P[P.length - 1];
    P[1] = bm(id, 0) + P[1];
    return { xml: b + all(P) + a };
  },
  bookmarkFromBefore: (P, id, b, a) => {
    P[1] = bmEnd(id, 0) + P[1];
    return { xml: bm(id, 0) + b + all(P) + a };
  },
  bookmarkToAfter: (P, id, b, a) => {
    P[P.length - 1] = bm(id, 0) + P[P.length - 1];
    return { xml: b + all(P) + a + bmEnd(id, 0) };
  },
  bookmarkAround: (P, id, b, a) => ({ xml: b + bm(id, 1) + all(P) + bmEnd(id, 1) + a }),
  comment: (P, id, b, a) => {
    P[1] = `<w:commentRangeStart w:id="${id}"/>${P[1]}`;
    return {
      xml: `${b}${all(P)}${a}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`,
      comment: true,
    };
  },
  permission: (P, id, b, a) => {
    P[P.length - 1] = `<w:permEnd w:id="${id}"/>${P[P.length - 1]}`;
    P[1] = `<w:permStart w:id="${id}" w:edGrp="everyone"/>${P[1]}`;
    return { xml: b + all(P) + a };
  },
  hyperlink: (P, _id, b, a) => {
    P[1] = `<w:hyperlink w:anchor="target" w:history="1">${P[1]}</w:hyperlink>`;
    return { xml: b + all(P) + a };
  },
  hyperlinkToEnd: (P, _id, b, a) => ({
    xml: `${b}${P[0]}<w:hyperlink w:anchor="target">${all(P.slice(1))}${a}</w:hyperlink>`,
  }),
  smartTag: (P, _id, b, a) => {
    const m = P.length >> 1;
    P[m] =
      `<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">` +
      `<w:smartTagPr><w:attr w:name="x" w:val="y"/></w:smartTagPr>${P[m]}</w:smartTag>`;
    return { xml: b + all(P) + a };
  },
  sdtTail: (P, id, b, a) => ({
    xml: `${b}${P[0]}<w:sdt><w:sdtPr><w:id w:val="${1000 + id}"/></w:sdtPr><w:sdtContent>${all(P.slice(1))}</w:sdtContent></w:sdt>${a}`,
  }),
  sdtWhole: (P, id, b, a) => ({
    xml: `<w:sdt><w:sdtPr><w:id w:val="${1000 + id}"/></w:sdtPr><w:sdtContent>${b}${all(P)}${a}</w:sdtContent></w:sdt>`,
  }),
  customXml: (P, _id, b, a) => {
    P[1] = `<w:customXml w:uri="urn:x" w:element="name">${P[1]}</w:customXml>`;
    return { xml: b + all(P) + a };
  },
  dir: (P, _id, b, a) => {
    P[1] = `<w:dir w:val="ltr">${P[1]}</w:dir>`;
    return { xml: b + all(P) + a };
  },
  nested: (P, id, b, a) => {
    P[1] =
      `<w:hyperlink w:anchor="target"><w:smartTag w:uri="u" w:element="e">${bm(id, 2)}${P[1]}` +
      `</w:smartTag></w:hyperlink>`;
    return { xml: b + all(P) + a + bmEnd(id, 2) };
  },
  insertion: (P, id, b, a) => {
    P[1] = `<w:ins w:id="${9000 + id}" w:author="Earlier">${P[1]}</w:ins>`;
    return { xml: b + all(P) + a };
  },
  deletionBetween: (P, id, b, a) => {
    P[1] =
      `<w:del w:id="${9000 + id}" w:author="Earlier"><w:r><w:rPr><w:strike/></w:rPr>` +
      `<w:delText>old</w:delText></w:r></w:del>${P[1]}`;
    return { xml: b + all(P) + a };
  },
  fieldResult: (P, _id, b, a) => ({ xml: field('QUOTE "x"', b + all(P) + a) }),
  fldSimple: (P, _id, b, a) => ({
    xml: `<w:fldSimple w:instr=" QUOTE x ">${b}${all(P)}${a}</w:fldSimple>`,
  }),
  emptyRun: (P, _id, b, a) => {
    P[1] = `<w:r><w:rPr><w:b/></w:rPr></w:r>${P[1]}`;
    return { xml: b + all(P) + a };
  },
  nonTextContent: (P, _id, b, a) => {
    P[1] = P[1]!.replace('<w:t', '<w:lastRenderedPageBreak/><w:t');
    return { xml: b + all(P) + a };
  },
  twoTextLeaves: (P, _id, b, a) => {
    // The last piece's run holds its text in two w:t (split after its first character).
    const last = P.length - 1;
    P[last] = P[last]!.replace(
      /<w:t xml:space="preserve">(.)(.+?)<\/w:t>/,
      '<w:t xml:space="preserve">$1</w:t><w:t xml:space="preserve">$2</w:t>',
    );
    return { xml: b + all(P) + a };
  },
  fieldInside: (P, _id, b, a) => {
    P[1] = field('PAGE', '') + P[1];
    return { xml: b + all(P) + a, crossing: true };
  },
  fieldResultPartly: (P, _id, b, a) => ({
    xml: b + P[0] + field('QUOTE "x"', all(P.slice(1))) + a,
    crossing: true,
  }),
};

interface Case {
  name: string;
  /** Index of the paragraph in the body. */
  index: number;
  crossing: boolean;
  firstFmt: string;
  lastFmt: string;
}

function buildCorpus(): { bytes: Uint8Array; cases: Case[] } {
  const cases: Case[] = [];
  const paras: string[] = [];
  const comments: string[] = [];
  let id = 0;
  for (const [deco, decorate] of Object.entries(DECORATIONS)) {
    for (let k = 2; k <= 6; k++) {
      for (const cuts of cutSets(WORD.length, k)) {
        const bounds = [0, ...cuts, WORD.length];
        const pieces = bounds.slice(0, -1).map((s, i) => WORD.slice(s, bounds[i + 1]));
        const ownPrefix = id % 2 === 0;
        const ownSuffix = (id >> 1) % 2 === 0;
        const fmts = pieces.map((_, i) => FORMATS[(i + id) % FORMATS.length]!);
        const P = pieces.map((t, i) => {
          let text = t;
          if (i === 0 && !ownPrefix) text = `Say ${text}`;
          if (i === pieces.length - 1 && !ownSuffix) text = `${text} now`;
          return r(text, fmts[i]);
        });
        const before = ownPrefix ? r('Say ', '<w:i/>') : '';
        const after = ownSuffix ? r(' now', '<w:u w:val="double"/>') : '';
        const built = decorate(P, id, before, after);
        paras.push(`<w:p>${built.xml}</w:p>`);
        if (built.comment) {
          comments.push(
            `<w:comment w:id="${id}" w:author="c" w:initials="c"><w:p>${r(`note ${id}`)}</w:p></w:comment>`,
          );
        }
        cases.push({
          name: `${deco} k=${k} cuts=${cuts.join(',')}${ownPrefix ? '' : ' mid-start'}${ownSuffix ? '' : ' mid-end'}`,
          index: cases.length,
          crossing: built.crossing === true,
          firstFmt: fmts[0]!,
          lastFmt: fmts[fmts.length - 1]!,
        });
        id++;
      }
    }
  }
  const bytes = docxFrom({
    body: `<w:p><w:bookmarkStart w:id="99999" w:name="target"/><w:bookmarkEnd w:id="99999"/></w:p>${paras.join('')}`,
    comments: comments.join(''),
  });
  return { bytes, cases };
}

const { bytes: CORPUS, cases: CASES } = buildCorpus();

/** The body paragraphs of the corpus (the first one is the hyperlink target). */
const corpusParagraphs = (doc: WordDocument): Paragraph[] => doc.body.paragraphs().slice(1);

const rPrOf = (run: XElement): string => {
  const rPr = run.element(NS.w, 'rPr');
  return rPr ? xml(rPr) : '';
};
const fmtXml = (fmt: string): string => (fmt === '' ? '' : `<w:rPr>${fmt}</w:rPr>`);

/** Checks a replaced (untracked) case; returns failure messages. */
function checkPlain(para: Paragraph, c: Case, source: string, fmt: string): string[] {
  const fail: string[] = [];
  if (c.crossing) {
    if (xml(para.element) !== source) fail.push('crossing case was modified');
    return fail;
  }
  const f = flattenElement(para.element);
  if (f.text !== NEW_TEXT) {
    fail.push(`text ${JSON.stringify(f.text)}`);
    return fail;
  }
  const g = f.at(4)!.node.parent as XElement;
  if (f.at(8)!.node.parent !== g) fail.push('replacement split over several runs');
  if (rPrOf(g) !== fmtXml(fmt)) fail.push(`formatting ${rPrOf(g)} != ${fmtXml(fmt)}`);
  for (const m of f.markers) {
    if (m.pos > 4 && m.pos < 9) fail.push(`${m.node.local} left inside the replacement`);
  }
  try {
    expectSane(para.element);
  } catch (e) {
    fail.push((e as Error).message.split('\n')[0]!);
  }
  return fail;
}

describe('split-run corpus', () => {
  it(`has ${CASES.length} cases reading "${SOURCE_TEXT}"`, () => {
    expect(CASES.length).toBe(218 * Object.keys(DECORATIONS).length);
    const doc = readDocx(CORPUS);
    const paras = corpusParagraphs(doc);
    expect(paras).toHaveLength(CASES.length);
    for (const p of paras) expect(p.text).toBe(SOURCE_TEXT);
    expectSane(doc.body.element);
  });

  it('plain replace: 100% of the corpus', () => {
    const doc = readDocx(CORPUS);
    const sources = corpusParagraphs(doc).map((p) => xml(p.element));
    let warnings = 0;
    const n = replaceText(doc, WORD, 'Great', { onWarning: () => warnings++ });
    const crossing = CASES.filter((c) => c.crossing).length;
    expect(n).toBe(CASES.length - crossing);
    expect(warnings).toBe(crossing);
    const paras = corpusParagraphs(doc);
    const failures: string[] = [];
    for (const c of CASES) {
      for (const f of checkPlain(paras[c.index]!, c, sources[c.index]!, c.firstFmt)) {
        failures.push(`${c.name}: ${f}`);
      }
    }
    expect(failures.slice(0, 20)).toEqual([]);
    const back = roundTrip(doc, 'corpus-plain.docx');
    expect(corpusParagraphs(back).map((p) => p.text)).toEqual(paras.map((p) => p.text));
  });

  it("formatFrom 'last': 100% of the corpus", () => {
    const doc = readDocx(CORPUS);
    const sources = corpusParagraphs(doc).map((p) => xml(p.element));
    replaceText(doc, WORD, 'Great', { formatFrom: 'last' });
    const paras = corpusParagraphs(doc);
    const failures: string[] = [];
    for (const c of CASES) {
      for (const f of checkPlain(paras[c.index]!, c, sources[c.index]!, c.lastFmt)) {
        failures.push(`${c.name}: ${f}`);
      }
    }
    expect(failures.slice(0, 20)).toEqual([]);
    roundTrip(doc, 'corpus-last.docx');
  });

  it('RunInput[] replace: 100% of the corpus', () => {
    const doc = readDocx(CORPUS);
    replaceText(doc, WORD, [
      { text: 'Gr', bold: true },
      { text: 'eat', italic: true },
    ]);
    const failures: string[] = [];
    const paras = corpusParagraphs(doc);
    for (const c of CASES) {
      const para = paras[c.index]!;
      const want = c.crossing ? SOURCE_TEXT : NEW_TEXT;
      if (para.text !== want) failures.push(`${c.name}: text ${para.text}`);
      try {
        expectSane(para.element);
      } catch (e) {
        failures.push(`${c.name}: ${(e as Error).message.split('\n')[0]}`);
      }
    }
    expect(failures.slice(0, 20)).toEqual([]);
    roundTrip(doc, 'corpus-runs.docx');
  });

  it('tracked replace: 100% of the corpus; accepts to the new text, rejects to the old', () => {
    const doc = readDocx(CORPUS);
    const original = corpusParagraphs(doc).map((p) => originalText(p.element));
    replaceText(doc, WORD, 'Great', { track: { author: 'Corpus' } });
    const failures: string[] = [];
    let paras = corpusParagraphs(doc);
    for (const c of CASES) {
      const para = paras[c.index]!;
      const want = c.crossing ? SOURCE_TEXT : NEW_TEXT;
      if (para.text !== want) failures.push(`${c.name}: text ${para.text}`);
      if (originalText(para.element) !== original[c.index]) {
        failures.push(`${c.name}: original view changed to ${originalText(para.element)}`);
      }
      try {
        expectSane(para.element);
      } catch (e) {
        failures.push(`${c.name}: ${(e as Error).message.split('\n')[0]}`);
      }
    }
    expect(failures.slice(0, 20)).toEqual([]);

    const back = roundTrip(doc, 'corpus-tracked.docx');
    const current = corpusParagraphs(back).map((p) => p.text);
    const accepted = back.fork();
    acceptAllRevisions(accepted);
    expect(revisionsLeft(accepted)).toEqual([]);
    paras = corpusParagraphs(accepted);
    expect(paras.map((p) => p.text)).toEqual(current);
    expectSane(accepted.body.element);
    roundTrip(accepted, 'corpus-tracked-accepted.docx');

    const rejected = back.fork();
    rejectAllRevisions(rejected);
    expect(revisionsLeft(rejected)).toEqual([]);
    expect(corpusParagraphs(rejected).map((p) => p.text)).toEqual(original);
    expectSane(rejected.body.element);
    roundTrip(rejected, 'corpus-tracked-rejected.docx');
  });
});
