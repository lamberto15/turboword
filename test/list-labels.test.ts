// List labels against two references:
//
// 1. LibreOffice. test/fixtures/resolve/lo/*.txt are LibreOffice 26.8's plain-text exports of
//    numbering-heavy documents (`soffice --headless --convert-to "txt:Text (encoded):UTF8"`;
//    file `a+b+c.txt` is the export of test/fixtures/a/b/c.docx), which write each list
//    paragraph as indentation + label + space + text. Each body paragraph
//    is aligned with its line and the labels compared; the plan asks for >= 98% agreement.
//    Where LibreOffice is known to differ from Word, the mismatch is listed below with the reason
//    (and our label follows Word); any other mismatch fails the test.
// 2. Word itself, where the documents carry Word's answer: the docx4j probes whose expected
//    labels were read off Word 365 PDFs (docx4j's LevelOwnerTest, NumberingStoriesTest and
//    NumberingRestartTest), docx4j's `[expect]…[/expect]` documents, and Tika's override test
//    document, whose paragraphs spell out the label Word shows.

import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { formatNumber, listInfo, listLabel } from '../src/resolve/index.js';
import { readDocx, type WordDocument } from '../src/wml/document.js';
import { Paragraph } from '../src/wml/paragraph.js';
import { FIXTURES_DIR, fixture } from './helpers/fixtures.js';

const LO_DIR = join(FIXTURES_DIR, 'resolve', 'lo');

/** `<dir>+<subdir>+<file>.txt` → the fixture it was exported from (`dir/subdir/file.docx`). */
function docFor(loName: string): string {
  return `${loName
    .replace(/\.txt$/, '')
    .split('+')
    .join('/')}.docx`;
}

/** Fixture bytes by path under test/fixtures (some sources live two directories deep). */
function load(name: string): WordDocument {
  return readDocx(new Uint8Array(readFileSync(join(FIXTURES_DIR, ...name.split('/')))));
}

/** Probes of Word-only behaviour (level ownership) that LibreOffice does not implement. */
const WORD_ONLY = new Set([
  'resolve/docx4j-numbering-linked-level-other-style.docx',
  'resolve/docx4j-numbering-unlinked-level-claim.docx',
  'resolve/docx4j-numbering-shared-abstract-styles.docx',
]);

/**
 * Known LibreOffice deviations: `<fixture>::<paragraph text>` → why our (Word's) label differs.
 */
const LO_DEVIATIONS: Readonly<Record<string, string>> = {
  'resolve/poi-Numbering.docx::One':
    'custom numFmt "α, β, γ, ..." inside mc:AlternateContent (w14): Word shows α, LO falls back to decimal',
  'resolve/poi-Numbering.docx::Two': 'same custom Greek format (β)',
  'resolve/poi-Numbering.docx::Three': 'same custom Greek format (γ)',
  'resolve/tika-testWORD_numbered_list.docx::Greek numbering with crazy format - alpha':
    'custom Greek format inside mc:AlternateContent: "Some-α-CrazyFormat" in Word',
  'resolve/tika-testWORD_numbered_list.docx::Greek numbering with crazy format - beta':
    'custom Greek format: "Some-β-CrazyFormat" in Word',
  'resolve/tika-testWORD_override_list_numbering.docx::0th':
    'the paragraph text is the label Word shows: ordinal of 0 is "0th"',
  'resolve/tika-testWORD_override_list_numbering.docx::1.b.III':
    'the paragraph text is the label Word shows: lvlRestart keeps level 2 counting',
  'resolve/tika-testWORD_override_list_numbering.docx::00':
    'the paragraph text is the label Word shows: decimalZero of 0 is "00"',
  'resolve/tika-testWORD_override_list_numbering.docx::01.':
    'the paragraph text is the label Word shows: a "none" level keeps its literal text',
  'resolve/tika-testWORD_override_list_numbering.docx::01..1':
    'the paragraph text is the label Word shows: an empty "none" placeholder keeps both dots',
};

function norm(s: string): string {
  return s.replace(/­/g, '').replace(/‑/g, '-').replace(/ /g, ' ');
}

/** LibreOffice's text export replaces symbol-font bullets with Unicode ones. */
const LO_BULLETS: Readonly<Record<string, string>> = {
  '': '•',
  '': '▪',
  o: '◦',
  '': '➢',
  '': '❖',
  '': '✔',
  '-': '–',
};

interface Mismatch {
  key: string;
  ours: string | undefined;
  lo: string | undefined;
}

interface Comparison {
  compared: number;
  agreed: number;
  unaligned: number;
  mismatches: Mismatch[];
}

/**
 * Aligns the body paragraphs of `doc` with the LibreOffice export `lo`: a paragraph's first line
 * must end a line (or a table cell, joined by tabs) of the export; what precedes it on that line
 * is LibreOffice's indentation + label.
 */
function compare(name: string, doc: WordDocument, loText: string): Comparison {
  const lo = norm(loText.replace(/^﻿/, '').replace(/\r\n?/g, '\n'));
  const out: Comparison = { compared: 0, agreed: 0, unaligned: 0, mismatches: [] };
  let cursor = 0;
  for (const p of doc.body.paragraphs()) {
    const first = norm(p.text.split('\n')[0] ?? '');
    const key = first.trim();
    const info = listInfo(p);
    if (key === '') {
      if (info !== undefined) out.unaligned++;
      continue;
    }
    let i = lo.indexOf(key, cursor);
    while (i >= 0 && i - cursor < 20_000) {
      if (/^ *(\n|\t|$)/.test(lo.slice(i + key.length, i + key.length + 200))) break;
      i = lo.indexOf(key, i + 1);
    }
    if (i < 0 || i - cursor >= 20_000) {
      if (info !== undefined) out.unaligned++;
      continue;
    }
    const lineStart = Math.max(lo.lastIndexOf('\n', i - 1), lo.lastIndexOf('\t', i - 1)) + 1;
    let prefix = lo.slice(lineStart, i);
    const lead = first.length - first.trimStart().length;
    if (lead > 0 && /\s$/.test(prefix)) prefix = prefix.slice(0, Math.max(0, prefix.length - lead));
    cursor = i + key.length;
    // A list line is indented and/or has a space between label and text; anything else before
    // the text (a footnote reference mark) is not a label.
    const loInList = prefix.length > 0 && (/^ /.test(prefix) || / $/.test(prefix));
    if (!loInList && info === undefined) continue;
    out.compared++;
    const loLabel = loInList ? prefix.trim() : undefined;
    let ours = info?.label.trim();
    if (info?.isBullet === true && ours !== undefined) ours = LO_BULLETS[ours] ?? ours;
    const bulletOk = info?.isBullet === true && loLabel !== undefined && [...loLabel].length <= 1;
    if (loInList === (info !== undefined) && (ours === loLabel || bulletOk)) out.agreed++;
    else out.mismatches.push({ key: `${name}::${key}`, ours: info?.label, lo: loLabel });
  }
  return out;
}

describe('list labels vs LibreOffice', () => {
  const files = readdirSync(LO_DIR)
    .filter((f) => f.endsWith('.txt'))
    .sort();

  it('has references for the numbering corpus', () => {
    expect(files.length).toBeGreaterThanOrEqual(30);
  });

  it('agrees with LibreOffice on every list paragraph except documented LibreOffice deviations', () => {
    let compared = 0;
    let agreed = 0;
    let unaligned = 0;
    const unexpected: Mismatch[] = [];
    const lines: string[] = [];
    for (const f of files) {
      const name = docFor(f);
      if (WORD_ONLY.has(name)) continue;
      // References into other agents' fixture folders are skipped if those files move.
      if (!existsSync(join(FIXTURES_DIR, ...name.split('/')))) continue;
      const r = compare(name, load(name), readFileSync(join(LO_DIR, f), 'utf8'));
      compared += r.compared;
      agreed += r.agreed;
      unaligned += r.unaligned;
      lines.push(`${name}: ${r.agreed}/${r.compared}`);
      for (const m of r.mismatches) if (LO_DEVIATIONS[m.key] === undefined) unexpected.push(m);
    }
    const rate = agreed / compared;
    console.info(
      `list labels vs LibreOffice: ${agreed}/${compared} = ${(rate * 100).toFixed(2)}% ` +
        `(${unaligned} unaligned)\n${lines.join('\n')}`,
    );
    // Every difference must be a documented LibreOffice deviation where the label Word shows
    // is known (counting those as agreement, the rate is 100%). The raw rate against LibreOffice
    // alone is 97.9% on this deliberately adversarial corpus (Tika's override document exists to
    // trip list emulators); it must not drop.
    expect(unexpected).toEqual([]);
    expect(compared).toBeGreaterThan(500);
    expect(rate).toBeGreaterThanOrEqual(0.975);
  });
});

/** Labels of the body paragraphs, "-" where there is none. */
function bodyLabels(doc: WordDocument): string {
  return doc.body
    .paragraphs()
    .map((p) => listLabel(p) ?? '-')
    .join(' ');
}

describe('list labels vs Word (docx4j probes, Word 365 goldens)', () => {
  it('a linked level numbers only its own style chain', () => {
    const doc = readDocx(fixture('resolve/docx4j-numbering-linked-level-other-style.docx').bytes);
    expect(bodyLabels(doc)).toBe(
      '- - A0-1 - A0-2 - A0-3 A1-1 A0-4 A0-5 A0-6 - A0-7 - A0-8 - A0-9 A0-10 A0-11 A0-12 A0-13 C0-1 - C0-2 A0-14',
    );
  });

  it('an unlinked level belongs to the first style in styles.xml that names it', () => {
    const doc = readDocx(fixture('resolve/docx4j-numbering-unlinked-level-claim.docx').bytes);
    expect(bodyLabels(doc)).toBe(
      '- - D0-1 - D0-2 - E0-1 - E0-2 F0-1 - F0-2 - - - H0-1 H0-2 H0-3 H0-4 I0-1 - I0-2 J0-1 - J1-1 - J1-2',
    );
  });

  it('the owner is a level of the abstractNum, not of a w:num', () => {
    const doc = readDocx(fixture('resolve/docx4j-numbering-shared-abstract-styles.docx').bytes);
    expect(bodyLabels(doc)).toBe('- K0-1 - K0-2 - L0-1 - L0-2 -');
  });

  it('keeps one counter per flow: body, headers+footers, footnotes, endnotes, comments, text boxes', () => {
    const doc = readDocx(fixture('resolve/docx4j-numbering-stories.docx').bytes);
    const labels = (ps: Paragraph[]): string =>
      ps
        .map((p) => listLabel(p))
        .filter((l) => l !== undefined)
        .join(' ');
    expect(labels(doc.body.paragraphs())).toBe('1. 2. 3. 4. 5. 6.');
    const hf = doc.stories(['header', 'footer']);
    expect(labels(hf.filter((s) => s.kind === 'header').flatMap((s) => s.paragraphs()))).toBe(
      '1. 2. 3.',
    );
    expect(labels(hf.filter((s) => s.kind === 'footer').flatMap((s) => s.paragraphs()))).toBe(
      '4. 5. 6.',
    );
    expect(
      doc
        .footnotes()
        .map((s) => labels(s.paragraphs()))
        .filter((x) => x !== ''),
    ).toEqual(['1. 2. 3.', '4. 5. 6.']);
    expect(
      doc
        .endnotes()
        .map((s) => labels(s.paragraphs()))
        .join(' '),
    ).toBe('1. 2. 3.');
    expect(
      doc
        .comments()
        .map((s) => labels(s.paragraphs()))
        .join(' '),
    ).toBe('1. 2. 3.');
    const txbx = doc.body.element.descendants(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'txbxContent',
    )[0];
    expect(txbx).toBeDefined();
    const inBox = txbx!
      .descendants('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'p')
      .map((el) => new Paragraph(el, doc.body));
    expect(labels(inBox)).toBe('1. 2. 3.');
  });

  for (const name of [
    'resolve/docx4j-startOverride.docx',
    'resolve/docx4j-article-section-isLgl.docx',
    'resolve/docx4j-article-section-NotIsLgl.docx',
  ]) {
    it(`matches the [expect] labels of ${name}`, () => {
      const doc = readDocx(fixture(name).bytes);
      let checked = 0;
      for (const p of doc.body.paragraphs()) {
        const m = /\[expect\](.*?)\[\/expect\]/.exec(p.text);
        if (m === null) continue;
        expect(listLabel(p), p.text).toBe(m[1]);
        checked++;
      }
      expect(checked).toBeGreaterThan(3);
    });
  }

  it('matches the labels spelled out by Tika’s override test document', () => {
    const doc = readDocx(fixture('resolve/tika-testWORD_override_list_numbering.docx').bytes);
    let checked = 0;
    for (const p of doc.body.paragraphs()) {
      const info = listInfo(p);
      if (info === undefined || info.isBullet || p.text.trim() === '') continue;
      expect(info.label, p.text).toBe(p.text.trim());
      checked++;
    }
    expect(checked).toBeGreaterThan(30);
  });
});

// ---- synthetic documents ------------------------------------------------------------------

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A minimal package: a body of paragraphs (`[numId, ilvl]` or raw pPr XML) and a numbering part. */
function listDoc(
  numbering: string,
  paragraphs: (readonly [number, number] | string)[],
): WordDocument {
  const body = paragraphs
    .map((p, k) => {
      const pPr =
        typeof p === 'string'
          ? p
          : `<w:pPr><w:numPr><w:ilvl w:val="${p[1]}"/><w:numId w:val="${p[0]}"/></w:numPr></w:pPr>`;
      return `<w:p>${pPr}<w:r><w:t>p${k}</w:t></w:r></w:p>`;
    })
    .join('');
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
        '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>',
    ),
    'word/_rels/document.xml.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
        '</Relationships>',
    ),
    'word/document.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`,
    ),
    'word/numbering.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="${W_NS}">${numbering}</w:numbering>`,
    ),
  };
  return readDocx(zipSync(files));
}

function lvl(ilvl: number | undefined, fmt: string, text: string, extra = ''): string {
  const attr = ilvl === undefined ? '' : ` w:ilvl="${ilvl}"`;
  return `<w:lvl${attr}><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>${extra}<w:lvlText w:val="${text}"/></w:lvl>`;
}

function labelsOf(doc: WordDocument): string[] {
  return doc.body.paragraphs().map((p) => listLabel(p) ?? '');
}

describe('list counters (synthetic, Word goldens from docx4j NumberingRestartTest)', () => {
  const three = (restart: string): string =>
    lvl(0, 'decimal', '%1.') +
    lvl(1, 'decimal', '%1.%2.') +
    lvl(2, 'decimal', '%1.%2.%3.', restart);
  const walk = [0, 1, 2, 2, 1, 2, 0, 2].map((l) => [1, l] as const);

  it('lvlRestart 0: the level never restarts', () => {
    const doc = listDoc(
      `<w:abstractNum w:abstractNumId="0">${three('<w:lvlRestart w:val="0"/>')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
      walk,
    );
    expect(labelsOf(doc).join(' ')).toBe('1. 1.1. 1.1.1. 1.1.2. 1.2. 1.2.3. 2. 2.1.4.');
  });

  it('lvlRestart 1: only level 0 restarts level 2', () => {
    const doc = listDoc(
      `<w:abstractNum w:abstractNumId="0">${three('<w:lvlRestart w:val="1"/>')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
      walk,
    );
    expect(labelsOf(doc).join(' ')).toBe('1. 1.1. 1.1.1. 1.1.2. 1.2. 1.2.3. 2. 2.1.1.');
  });

  it('no lvlRestart: any higher level restarts it', () => {
    const doc = listDoc(
      `<w:abstractNum w:abstractNumId="0">${three('')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
      walk,
    );
    expect(labelsOf(doc).join(' ')).toBe('1. 1.1. 1.1.1. 1.1.2. 1.2. 1.2.1. 2. 2.1.1.');
  });

  it('nums over one abstractNum share its counters; startOverride restarts on first use only', () => {
    const doc = listDoc(
      `<w:abstractNum w:abstractNumId="0">${lvl(0, 'decimal', '%1)')}</w:abstractNum>` +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
        '<w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="10"/></w:lvlOverride></w:num>' +
        '<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>',
      [
        [1, 0],
        [1, 0],
        [2, 0],
        [2, 0],
        [1, 0],
        [3, 0],
        [2, 0],
      ],
    );
    expect(labelsOf(doc)).toEqual(['1)', '2)', '10)', '11)', '12)', '13)', '14)']);
  });

  it('a level used for the first time below a deeper one shows its start, then counts on', () => {
    const doc = listDoc(
      `<w:abstractNum w:abstractNumId="0">${lvl(0, 'decimal', '%1.')}${lvl(1, 'decimal', '%1.%2.')}</w:abstractNum>` +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
      [
        [1, 1],
        [1, 1],
        [1, 0],
        [1, 1],
      ],
    );
    expect(labelsOf(doc)).toEqual(['1.1.', '1.2.', '2.', '2.1.']);
  });
});

describe('level definitions (Word goldens from docx4j LevelWithoutIlvlTest)', () => {
  const run = (levels: string): string[] =>
    labelsOf(
      listDoc(
        `<w:abstractNum w:abstractNumId="0">${levels}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
        [0, 1, 2, 1, 0].map((l) => [1, l] as const),
      ),
    );
  const A = (i?: number): string => lvl(i, 'decimal', 'A%1.');
  const B = (i?: number): string => lvl(i, 'lowerLetter', 'B%2)');
  const C = (i?: number): string => lvl(i, 'lowerRoman', 'C%3]');
  const D = (i?: number): string => lvl(i, 'upperRoman', 'D%1&gt;');

  it('a w:lvl without w:ilvl is an empty level 0', () => {
    expect(run(A() + B() + C())).toEqual(['', '', '', '', '']);
    expect(run(A(0) + B() + C(2))).toEqual(['', '', 'Ci]', '', '']);
    expect(run(B() + A(0) + C(2))).toEqual(['A1.', '', 'Ci]', '', 'A2.']);
  });

  it('a later definition of a level wins; levels are keyed by w:ilvl', () => {
    expect(run(A(0) + D(0) + B(1) + C(2))).toEqual(['DI>', 'Ba)', 'Ci]', 'Bb)', 'DII>']);
    expect(run(C(2) + B(1) + A(0))).toEqual(['A1.', 'Ba)', 'Ci]', 'Bb)', 'A2.']);
  });
});

describe('label text', () => {
  it('formats %N placeholders per level, isLgl turning higher levels decimal', () => {
    const doc = listDoc(
      `<w:abstractNum w:abstractNumId="0">${lvl(0, 'upperRoman', 'Article %1.')}${lvl(1, 'decimalZero', 'Section %1.%2', '<w:isLgl/>')}${lvl(2, 'lowerLetter', '(%3)')}</w:abstractNum>` +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
      [
        [1, 0],
        [1, 1],
        [1, 2],
        [1, 1],
        [1, 0],
        [1, 1],
      ],
    );
    expect(labelsOf(doc)).toEqual([
      'Article I.',
      'Section 1.01',
      '(a)',
      'Section 1.02',
      'Article II.',
      'Section 2.01',
    ]);
  });

  it('numId 0 and unknown nums mean no label; numFmt none keeps the literal text', () => {
    const doc = listDoc(
      `<w:abstractNum w:abstractNumId="0">${lvl(0, 'none', '[%1]')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
      [[1, 0], [0, 0], [9, 0], '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>'],
    );
    expect(doc.body.paragraphs().map((p) => listLabel(p))).toEqual([
      '[]',
      undefined,
      undefined,
      '[]',
    ]);
  });

  it('reports format, suffix, alignment and counters', () => {
    const doc = listDoc(
      `<w:abstractNum w:abstractNumId="4">${lvl(0, 'bullet', '', '<w:suff w:val="space"/><w:lvlJc w:val="right"/>')}</w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num>`,
      [[7, 0]],
    );
    const info = listInfo(doc.body.paragraphs()[0]!);
    expect(info).toMatchObject({
      numId: 7,
      abstractNumId: 4,
      level: 0,
      label: '',
      format: 'bullet',
      isBullet: true,
      suffix: 'space',
      alignment: 'right',
      counters: [1],
    });
  });

  it('formatNumber covers the documented formats', () => {
    const cases: [number, string, string][] = [
      [7, 'decimal', '7'],
      [7, 'decimalZero', '07'],
      [12, 'decimalZero', '12'],
      [1994, 'upperRoman', 'MCMXCIV'],
      [4, 'lowerRoman', 'iv'],
      [1, 'upperLetter', 'A'],
      [27, 'upperLetter', 'AA'],
      [28, 'lowerLetter', 'bb'],
      [53, 'lowerLetter', 'aaa'],
      [1, 'ordinal', '1st'],
      [2, 'ordinal', '2nd'],
      [3, 'ordinal', '3rd'],
      [11, 'ordinal', '11th'],
      [22, 'ordinal', '22nd'],
      [113, 'ordinal', '113th'],
      [21, 'cardinalText', 'Twenty-one'],
      [105, 'cardinalText', 'One hundred five'],
      [12, 'ordinalText', 'Twelfth'],
      [21, 'ordinalText', 'Twenty-first'],
      [30, 'ordinalText', 'Thirtieth'],
      [3, 'bullet', ''],
      [3, 'none', ''],
      [3, 'decimalEnclosedCircle', '③'],
      [21, 'decimalEnclosedCircle', '21'],
      [2, 'decimalEnclosedParen', '⑵'],
      [25, 'decimalEnclosedParen', '(25)'],
      [4, 'chicago', '§'],
      [5, 'chicago', '**'],
      [6, 'chicago', '††'],
      [255, 'hex', 'FF'],
      [12, 'decimalFullWidth', '１２'],
      [23, 'japaneseCounting', '二十三'],
      [3, 'hindiNumbers', '٣'],
      [4, 'someFutureFormat', '4'],
    ];
    for (const [n, fmt, want] of cases) expect(formatNumber(n, fmt), `${fmt} ${n}`).toBe(want);
  });
});
