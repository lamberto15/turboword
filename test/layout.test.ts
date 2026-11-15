// Pagination tests (layout-B): page/column flow, keep rules, sections, headers/footers, notes,
// tables, floats, page fields and the PageInfoProvider.
//
// Most tests drive `layoutIR` with hand-built IR and a deterministic fake paragraph layouter
// (fixed advance = half the font size, line pitch = 1.2 × size, words broken at spaces), so the
// expected positions can be computed by hand and the tests do not depend on fonts or on the
// real line breaker. The end-to-end tests at the bottom run `paginate()` on real documents.

import { describe, expect, it } from 'vitest';
import { readDocx } from '../src/index.js';
import type { FontManager, LayoutEnv, LineBox, ParagraphLayout } from '../src/layout/engine.js';
import { pageCount, pageInfo, paginate } from '../src/layout/index.js';
import { translatePath } from '../src/layout/page/strip.js';
import { layoutIR } from '../src/layout/paginate.js';
import type {
  DisplayItem,
  DocumentLayout,
  Font,
  IRAnchor,
  IRBlock,
  IRCell,
  IRDocument,
  IRInline,
  IRParagraph,
  IRRow,
  IRSection,
  IRTable,
  LineItem,
  Shaper,
  TextItem,
} from '../src/layout/types.js';
import type {
  BorderSpec,
  ResolvedParagraphProps,
  ResolvedRunProps,
  ResolvedTableProps,
} from '../src/resolve/types.js';
import type { WordDocument } from '../src/wml/document.js';
import { XElement } from '../src/xml/dom.js';
import { NS } from '../src/xml/ns.js';
import { fixture, fixtures, strictFixtures } from './helpers/fixtures.js';
import { testFontResolver } from './helpers/test-fonts.js';

// ---- fakes ------------------------------------------------------------------------------------

const fakeFont: Font = {
  family: 'Fake',
  subfamily: 'Regular',
  postscriptName: 'Fake',
  bold: false,
  italic: false,
  metrics: {
    unitsPerEm: 1000,
    ascender: 800,
    descender: -200,
    lineGap: 0,
    winAscent: 800,
    winDescent: 200,
    capHeight: 700,
    xHeight: 500,
    underlinePosition: -100,
    underlineThickness: 50,
    strikeoutPosition: 300,
    strikeoutSize: 50,
    italicAngle: 0,
  },
  outlines: 'standard14',
  glyphId: (cp) => cp,
  advance: () => 500,
  kerning: () => 0,
  bytes: new Uint8Array(),
  numGlyphs: 65536,
};

const fakeFonts: FontManager = {
  want() {},
  async load() {},
  fontFor: () => fakeFont,
  fonts: () => [fakeFont],
  indexOf: () => 0,
};

const fakeShaper: Shaper = {
  shape: (_font, text) =>
    [...text].map((ch, i) => ({ id: ch.codePointAt(0)!, advance: 500, dx: 0, dy: 0, cluster: i })),
};

function makeEnv(warnings: string[] = []): LayoutEnv {
  return {
    doc: {} as WordDocument,
    partOf: () => undefined,
    fonts: fakeFonts,
    shaper: fakeShaper,
    defaultTabStop: 36,
    images: new Map(),
    warn: (code) => warnings.push(code),
  };
}

function textItem(text: string, x: number, size: number): TextItem {
  return {
    kind: 'text',
    x,
    y: 0,
    font: 0,
    size,
    color: '000000',
    glyphs: [...text].map((c) => c.codePointAt(0)!),
    advances: [...text].map(() => size / 2),
    text,
    clusters: [...text].map((_, i) => i),
  };
}

/** Fake layoutParagraph: advance = size/2 per character, pitch = 1.2 × size, breaks at spaces. */
function fakeLayout(
  p: IRParagraph,
  width: number,
  o?: { insets?: (line: number) => { left: number; right: number } },
): ParagraphLayout {
  const size = p.props.markRun.size;
  const pitch =
    p.props.spacing.lineRule === 'exact'
      ? p.props.spacing.line / 20
      : Math.round(size * 1.2 * 100) / 100;
  const indent = p.props.indent.left / 20;
  const baseAvail = width - indent - p.props.indent.right / 20;
  let avail = baseAvail;
  let x0 = 0;
  const lines: LineBox[] = [];
  let cur: LineBox | undefined;
  let x = 0;
  const newLine = (): LineBox => {
    const l: LineBox = {
      height: pitch,
      baseline: size,
      width: 0,
      items: [],
      fields: [],
      noteRefs: [],
      bookmarks: [],
    };
    lines.push(l);
    const ins = o?.insets?.(lines.length - 1) ?? { left: 0, right: 0 };
    x0 = ins.left;
    avail = baseAvail - ins.right;
    x = x0;
    return l;
  };
  const place = (
    text: string,
    field?: Extract<IRInline, { kind: 'field' }>,
    note?: string,
  ): void => {
    const w = (text.length * size) / 2;
    if (cur === undefined) cur = newLine();
    if (x > x0 && x + w > avail + 0.001) cur = newLine();
    if (field !== undefined) cur.fields.push({ item: cur.items.length, field });
    if (note !== undefined) cur.noteRefs.push(note);
    cur.items.push(textItem(text, x, size));
    x += w;
    cur.width = x - x0;
  };
  for (const i of p.inlines) {
    if (i.kind === 'text') {
      for (const word of i.text.split(/(?<= )/)) if (word !== '') place(word);
    } else if (i.kind === 'field') place(i.fallback, i);
    else if (i.kind === 'noteRef') place(i.text, undefined, i.id);
    else if (i.kind === 'break') {
      if (cur === undefined) cur = newLine();
      if (i.type !== 'line') cur.breakAfter = i.type;
      cur = newLine();
    }
  }
  if (lines.length === 0) newLine();
  if (p.bookmarks.length > 0) lines[0]!.bookmarks.push(...p.bookmarks);
  return {
    source: p.source,
    lines,
    spaceBefore: p.props.spacing.before / 20,
    spaceAfter: p.props.spacing.after / 20,
    keepNext: p.props.keepNext,
    keepLines: p.props.keepLines,
    widowControl: p.props.widowControl,
    pageBreakBefore: p.props.pageBreakBefore,
    contextualSpacing: p.props.contextualSpacing,
    styleId: p.props.styleId,
    decorations: [],
    indentLeft: indent,
  };
}

// ---- IR builders ------------------------------------------------------------------------------

function runProps(o: Partial<ResolvedRunProps> = {}): ResolvedRunProps {
  return {
    fonts: { ascii: 'Fake' },
    size: 10,
    sizeCs: 10,
    bold: false,
    italic: false,
    boldCs: false,
    italicCs: false,
    strike: false,
    doubleStrike: false,
    color: 'auto',
    caps: false,
    smallCaps: false,
    hidden: false,
    vertAlign: 'baseline',
    spacing: 0,
    position: 0,
    scale: 100,
    kern: 0,
    rtl: false,
    cs: false,
    ...o,
  };
}

type ParaOpts = Partial<
  Pick<
    ResolvedParagraphProps,
    | 'keepNext'
    | 'keepLines'
    | 'pageBreakBefore'
    | 'widowControl'
    | 'contextualSpacing'
    | 'styleId'
    | 'outlineLevel'
  >
> & {
  before?: number;
  after?: number;
  size?: number;
  exact?: number;
  inlines?: IRInline[];
  bookmarks?: string[];
  anchors?: IRAnchor[];
};

function para(text: string, o: ParaOpts = {}): IRParagraph {
  const rp = runProps(o.size !== undefined ? { size: o.size } : {});
  const props: ResolvedParagraphProps = {
    alignment: 'left',
    indent: { left: 0, right: 0, firstLine: 0, hanging: 0 },
    spacing: {
      before: (o.before ?? 0) * 20,
      after: (o.after ?? 0) * 20,
      beforeAutospacing: false,
      afterAutospacing: false,
      line: o.exact !== undefined ? o.exact * 20 : 240,
      lineRule: o.exact !== undefined ? 'exact' : 'auto',
    },
    contextualSpacing: o.contextualSpacing ?? false,
    keepNext: o.keepNext ?? false,
    keepLines: o.keepLines ?? false,
    pageBreakBefore: o.pageBreakBefore ?? false,
    widowControl: o.widowControl ?? false,
    suppressAutoHyphens: false,
    bidi: false,
    tabs: [],
    borders: {},
    runDefaults: rp,
    markRun: rp,
  };
  if (o.styleId !== undefined) props.styleId = o.styleId;
  if (o.outlineLevel !== undefined) props.outlineLevel = o.outlineLevel;
  const inlines: IRInline[] = o.inlines ?? (text === '' ? [] : [{ kind: 'text', text, props: rp }]);
  return {
    kind: 'paragraph',
    source: new XElement('w:p', NS.w),
    props,
    inlines,
    anchors: o.anchors ?? [],
    bookmarks: o.bookmarks ?? [],
  };
}

/** `n` one-line paragraphs (12pt pitch). */
function lines(n: number, prefix = 'L', o: ParaOpts = {}): IRParagraph[] {
  return Array.from({ length: n }, (_, i) => para(`${prefix}${i}`, o));
}

function section(blocks: IRBlock[], o: Partial<IRSection> = {}): IRSection {
  return {
    index: 0,
    type: 'nextPage',
    pageWidth: 612,
    pageHeight: 792,
    margins: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36, gutter: 0 },
    columns: { count: 1, space: 36, separator: false },
    titlePage: false,
    pageNumberFormat: 'decimal',
    headers: {},
    footers: {},
    blocks,
    vAlign: 'top',
    ...o,
  };
}

function irDoc(sections: IRSection[], o: Partial<IRDocument> = {}): IRDocument {
  return {
    sections,
    evenAndOddHeaders: false,
    defaultTabStop: 36,
    footnotes: new Map(),
    endnotes: new Map(),
    images: new Map(),
    metadata: {},
    ...o,
  };
}

function run(ir: IRDocument, warnings: string[] = [], maxPasses?: number): DocumentLayout {
  const o: Parameters<typeof layoutIR>[2] = {
    layoutParagraph: (p, w, _env, c) => fakeLayout(p, w, c),
  };
  if (maxPasses !== undefined) o.maxPasses = maxPasses;
  return layoutIR(ir, makeEnv(warnings), o);
}

const texts = (items: readonly DisplayItem[]): TextItem[] =>
  items.filter((i): i is TextItem => i.kind === 'text');
const pageText = (l: DocumentLayout, page: number): string[] =>
  texts(l.pages[page]!.items).map((t) => t.text);
const find = (l: DocumentLayout, text: string): { page: number; item: TextItem } | undefined => {
  for (const p of l.pages) {
    const t = texts(p.items).find((i) => i.text === text);
    if (t !== undefined) return { page: p.index, item: t };
  }
  return undefined;
};
const pageOfPara = (l: DocumentLayout, p: IRParagraph): number | undefined =>
  l.paragraphPages.get(p.source);

// body: 612x792, margins 72 → 648pt tall, 468pt wide; 10pt text → 12pt pitch → 54 lines a page

// ---- tests ------------------------------------------------------------------------------------

describe('page flow', () => {
  it('fills pages line by line and records paragraph pages', () => {
    const ps = lines(120);
    const l = run(irDoc([section(ps)]));
    expect(l.pages).toHaveLength(3);
    expect(pageOfPara(l, ps[53]!)).toBe(0);
    expect(pageOfPara(l, ps[54]!)).toBe(1);
    expect(pageOfPara(l, ps[108]!)).toBe(2);
    const first = find(l, 'L0')!;
    expect(first.item.x).toBe(72);
    expect(first.item.y).toBe(72 + 10); // baseline = top + ascent
    expect(find(l, 'L54')!.item.y).toBe(72 + 10);
    expect(l.pages.map((p) => p.number)).toEqual(['1', '2', '3']);
    expect(l.pages[0]!.width).toBe(612);
  });

  it('wraps long paragraphs at the column width', () => {
    // 468pt / 5pt per char = 93 chars per line; "word " = 5 chars
    const p = para('word '.repeat(40).trim());
    const l = run(irDoc([section([p])]));
    const ys = new Set(texts(l.pages[0]!.items).map((t) => t.y));
    expect(ys.size).toBe(3); // 200 chars → 3 lines
  });

  it('drops space before at the top of a page after a soft break but keeps it after a hard one', () => {
    const ps = [...lines(54), para('Spaced', { before: 20 })];
    const l = run(irDoc([section(ps)]));
    expect(find(l, 'Spaced')).toMatchObject({ page: 1, item: { y: 82 } });
    const hard = [para('A'), para('B', { before: 20, pageBreakBefore: true })];
    const l2 = run(irDoc([section(hard)]));
    expect(find(l2, 'B')).toMatchObject({ page: 1, item: { y: 102 } });
  });

  it('separates paragraphs by the larger of space after/before; contextual spacing removes it', () => {
    const ps = [para('A', { after: 6 }), para('B', { before: 4 })];
    expect(find(run(irDoc([section(ps)])), 'B')!.item.y).toBe(72 + 12 + 6 + 10);
    const ps2 = [para('A', { after: 6 }), para('B', { before: 9 })];
    expect(find(run(irDoc([section(ps2)])), 'B')!.item.y).toBe(72 + 12 + 9 + 10);
    const ctx = [
      para('A', { after: 6, styleId: 'List', contextualSpacing: true }),
      para('B', { before: 4, styleId: 'List', contextualSpacing: true }),
      para('C', { before: 4, styleId: 'Other', contextualSpacing: true }),
    ];
    const l = run(irDoc([section(ctx)]));
    expect(find(l, 'B')!.item.y).toBe(72 + 12 + 10);
    expect(find(l, 'C')!.item.y).toBe(72 + 24 + 4 + 10);
  });

  it('keeps a heading with the next paragraph (keepNext)', () => {
    const ps = [...lines(53), para('Heading', { keepNext: true }), para('Body')];
    const l = run(irDoc([section(ps)]));
    expect(find(l, 'Heading')!.page).toBe(1);
    expect(find(l, 'Body')!.page).toBe(1);
    expect(find(l, 'L52')!.page).toBe(0);
  });

  it('breaks a keepNext chain longer than a page where it overflows', () => {
    const ps = lines(70, 'K', { keepNext: true });
    const l = run(irDoc([section(ps)]));
    expect(l.pages).toHaveLength(2);
    expect(find(l, 'K53')!.page).toBe(0);
    expect(find(l, 'K54')!.page).toBe(1);
  });

  it('applies widow and orphan control', () => {
    const long = 'abcd '.repeat(72).trim(); // 18 words a line → 4 lines
    // 53 lines used: only one line of the paragraph would fit → orphan → whole paragraph moves
    const orphan = run(irDoc([section([...lines(53), para(long, { widowControl: true })])]));
    const p1 = orphan.paragraphPages;
    expect(pageText(orphan, 0)).toHaveLength(53);
    expect([...p1.values()].at(-1)).toBe(1);
    // 51 lines used: 3 lines would fit, 1 left over → widow → 2 + 2
    const widow = run(irDoc([section([...lines(51), para(long, { widowControl: true })])]));
    const onFirst = texts(widow.pages[0]!.items).filter((t) => t.text.startsWith('abcd'));
    const ys = new Set(onFirst.map((t) => t.y));
    expect(ys.size).toBe(2);
    // without widow control: 3 + 1
    const none = run(irDoc([section([...lines(51), para(long)])]));
    expect(
      new Set(
        texts(none.pages[0]!.items)
          .filter((t) => t.text.startsWith('abcd'))
          .map((t) => t.y),
      ).size,
    ).toBe(3);
  });

  it('keeps the lines of a keepLines paragraph together', () => {
    const long = 'abcd '.repeat(54).trim(); // 3 lines
    const l = run(irDoc([section([...lines(52), para(long, { keepLines: true })])]));
    expect(texts(l.pages[0]!.items).some((t) => t.text.startsWith('abcd'))).toBe(false);
    expect(l.pages).toHaveLength(2);
  });

  it('honours page breaks (w:br) and pageBreakBefore without creating empty pages', () => {
    const br = para('', {
      inlines: [
        { kind: 'text', text: 'before', props: runProps() },
        { kind: 'break', type: 'page' },
        { kind: 'text', text: 'after', props: runProps() },
      ],
    });
    const l = run(irDoc([section([br, para('X', { pageBreakBefore: true })])]));
    expect(find(l, 'before')!.page).toBe(0);
    expect(find(l, 'after')!.page).toBe(1);
    expect(find(l, 'X')!.page).toBe(2);
    const first = run(irDoc([section([para('Only', { pageBreakBefore: true })])]));
    expect(first.pages).toHaveLength(1);
  });

  it('pushes the body below a tall header and above a tall footer', () => {
    const header = lines(6, 'H'); // 72pt tall, header distance 36 → body top 108
    const footer = lines(5, 'F'); // 60pt, footer distance 36 → body bottom 792-96
    const l = run(
      irDoc([section(lines(60), { headers: { default: header }, footers: { default: footer } })]),
    );
    expect(find(l, 'H0')!.item.y).toBe(36 + 10);
    expect(find(l, 'L0')!.item.y).toBe(108 + 10);
    expect(find(l, 'F4')!.item.y).toBe(792 - 36 - 12 + 10);
    // (696 - 108) / 12 = 49 lines on the first page
    expect(find(l, 'L48')!.page).toBe(0);
    expect(find(l, 'L49')!.page).toBe(1);
  });
});

describe('sections', () => {
  it('starts nextPage sections on a new page and continuous ones on the same page', () => {
    const l = run(
      irDoc([
        section([para('one')]),
        section([para('two')], { type: 'continuous' }),
        section([para('three')], { type: 'nextPage' }),
      ]),
    );
    expect(find(l, 'two')).toMatchObject({ page: 0, item: { y: 72 + 12 + 10 } });
    expect(find(l, 'three')!.page).toBe(1);
    expect(l.pages[1]!.section).toBe(2);
  });

  it('inserts a blank page for oddPage/evenPage starts', () => {
    const odd = run(irDoc([section([para('a')]), section([para('b')], { type: 'oddPage' })]));
    expect(odd.pages).toHaveLength(3);
    expect(find(odd, 'b')!.page).toBe(2);
    expect(pageText(odd, 1)).toEqual([]);
    const even = run(irDoc([section([para('a')]), section([para('b')], { type: 'evenPage' })]));
    expect(even.pages).toHaveLength(2);
  });

  it('restarts and continues page numbers', () => {
    const l = run(
      irDoc([
        section([para('a')], { pageNumberStart: 5 }),
        section([para('b')]),
        section([para('c')], { pageNumberStart: 1 }),
      ]),
    );
    expect(l.pages.map((p) => p.number)).toEqual(['5', '6', '1']);
  });

  it('formats page numbers with pgNumType/@w:fmt and the field switch', () => {
    const field = (format?: string): IRParagraph =>
      para('', {
        inlines: [
          format === undefined
            ? { kind: 'field', type: 'PAGE', fallback: '0', props: runProps() }
            : { kind: 'field', type: 'PAGE', format, fallback: '0', props: runProps() },
        ],
      });
    const l = run(
      irDoc([
        section(lines(60), { pageNumberFormat: 'upperRoman', footers: { default: [field()] } }),
        section([para('x')], { pageNumberStart: 3, footers: { default: [field('ALPHABETIC')] } }),
      ]),
    );
    expect(l.pages.map((p) => p.number)).toEqual(['I', 'II', '3']);
    expect(pageText(l, 0)).toContain('I');
    expect(pageText(l, 1)).toContain('II');
    expect(pageText(l, 2)).toContain('C');
  });

  it('honours the page size and margins of each section', () => {
    const l = run(
      irDoc([
        section([para('a')]),
        section([para('b')], {
          pageWidth: 792,
          pageHeight: 612,
          margins: { top: 36, right: 36, bottom: 36, left: 50, header: 20, footer: 20, gutter: 10 },
        }),
      ]),
    );
    expect(l.pages[1]).toMatchObject({ width: 792, height: 612 });
    expect(find(l, 'b')!.item).toMatchObject({ x: 60, y: 46 });
  });

  it('vertically centers pages with vAlign center', () => {
    const l = run(irDoc([section([para('mid')], { vAlign: 'center' })]));
    // one 12pt line in 648pt → shifted by (648 - 12) / 2
    expect(find(l, 'mid')!.item.y).toBe(72 + 318 + 10);
    const b = run(irDoc([section([para('low')], { vAlign: 'bottom' })]));
    expect(find(b, 'low')!.item.y).toBe(72 + 636 + 10);
  });
});

describe('columns', () => {
  const two = { count: 2, space: 36, separator: true };

  it('flows into the next column, then the next page', () => {
    const l = run(irDoc([section(lines(120), { columns: two })]));
    // column width (468 - 36) / 2 = 216
    expect(find(l, 'L53')!.item.x).toBe(72);
    expect(find(l, 'L54')).toMatchObject({ page: 0, item: { x: 72 + 216 + 36, y: 82 } });
    expect(find(l, 'L108')!.page).toBe(1);
    const sep = l.pages[0]!.items.filter((i): i is LineItem => i.kind === 'line');
    expect(sep).toHaveLength(1);
    expect(sep[0]!.x1).toBe(72 + 216 + 18);
  });

  it('balances columns before a continuous section break', () => {
    const l = run(
      irDoc([
        section(lines(10), { columns: two }),
        section([para('after')], { type: 'continuous' }),
      ]),
    );
    expect(find(l, 'L4')!.item.x).toBe(72);
    expect(find(l, 'L5')).toMatchObject({ item: { x: 324, y: 82 } });
    expect(find(l, 'after')).toMatchObject({ page: 0, item: { x: 72, y: 72 + 60 + 10 } });
  });

  it('supports explicit column widths and column breaks', () => {
    const brk = para('', {
      inlines: [
        { kind: 'text', text: 'c1', props: runProps() },
        { kind: 'break', type: 'column' },
        { kind: 'text', text: 'c2', props: runProps() },
      ],
    });
    const cols = {
      count: 2,
      space: 0,
      separator: false,
      widths: [
        { width: 100, space: 20 },
        { width: 348, space: 0 },
      ],
    };
    const l = run(irDoc([section([brk], { columns: cols })]));
    expect(find(l, 'c1')!.item.x).toBe(72);
    expect(find(l, 'c2')).toMatchObject({ page: 0, item: { x: 192, y: 82 } });
  });
});

describe('headers and footers', () => {
  const pageField = (type: string): IRParagraph =>
    para('', {
      inlines: [
        { kind: 'text', text: 'Page ', props: runProps() },
        { kind: 'field', type, fallback: '0', props: runProps() },
      ],
    });

  it('chooses first/even/default headers and fills PAGE / NUMPAGES per page', () => {
    const sec = section(lines(200), {
      titlePage: true,
      headers: { default: [para('odd-h')], first: [para('first-h')], even: [para('even-h')] },
      footers: {
        default: [pageField('PAGE')],
        even: [pageField('PAGE')],
        first: [pageField('NUMPAGES')],
      },
    });
    const l = run(irDoc([sec], { evenAndOddHeaders: true }));
    expect(l.pages).toHaveLength(4);
    const heads = l.pages.map((p) => texts(p.items).find((t) => t.text.endsWith('-h'))?.text);
    expect(heads).toEqual(['first-h', 'even-h', 'odd-h', 'even-h']);
    const nums = l.pages.map((p) =>
      texts(p.items)
        .filter((t) => /^\d+$/.test(t.text))
        .map((t) => t.text),
    );
    expect(nums).toEqual([['4'], ['2'], ['3'], ['4']]);
  });

  it('leaves the first page header blank when titlePg has no first header', () => {
    const l = run(
      irDoc([section(lines(60), { titlePage: true, headers: { default: [para('dflt')] } })]),
    );
    expect(pageText(l, 0)).not.toContain('dflt');
    expect(pageText(l, 1)).toContain('dflt');
  });

  it('counts SECTIONPAGES per section', () => {
    const l = run(
      irDoc([
        section(lines(60), { footers: { default: [pageField('SECTIONPAGES')] } }),
        section(lines(10), { footers: { default: [pageField('SECTIONPAGES')] } }),
      ]),
    );
    const nums = l.pages.map(
      (p) =>
        texts(p.items)
          .filter((t) => /^\d+$/.test(t.text))
          .map((t) => t.text)[0],
    );
    expect(nums).toEqual(['2', '2', '1']);
  });
});

describe('notes', () => {
  const ref = (id: string): IRInline => ({
    kind: 'noteRef',
    note: 'footnote',
    id,
    props: runProps(),
    text: id,
  });

  it('places footnotes at the bottom of the referencing page with a separator', () => {
    const p = para('', { inlines: [{ kind: 'text', text: 'see ', props: runProps() }, ref('1')] });
    const l = run(
      irDoc([section([p])], { footnotes: new Map([['1', [para('note one', { size: 8 })]]]) }),
    );
    const note = find(l, 'note ')!;
    expect(note.page).toBe(0);
    // note line pitch 9.6, separator 12.2 → note top = 720 - 9.6
    expect(note.item.y).toBeCloseTo(720 - 9.6 + 8, 5);
    const lines = l.pages[0]!.items.filter((i): i is LineItem => i.kind === 'line');
    expect(lines).toHaveLength(1);
    expect(lines[0]!.x2 - lines[0]!.x1).toBe(144);
  });

  it('moves a line to the next page when its footnote does not fit', () => {
    const p = para('', { inlines: [{ kind: 'text', text: 'ref ', props: runProps() }, ref('1')] });
    const body = [...lines(52), p];
    const l = run(irDoc([section(body)], { footnotes: new Map([['1', lines(3, 'N')]]) }));
    // 52 lines + the ref line (12) + separator (12.2) + 3 note lines (36) > 648
    expect(find(l, 'ref ')!.page).toBe(1);
    expect(find(l, 'N0')!.page).toBe(1);
  });

  it('continues footnotes that cannot fit on the next page', () => {
    const p = para('', { inlines: [{ kind: 'text', text: 'ref ', props: runProps() }, ref('1')] });
    const l = run(
      irDoc([section([p, para('next')])], { footnotes: new Map([['1', lines(70, 'N')]]) }),
    );
    expect(find(l, 'N0')!.page).toBe(0);
    expect(find(l, 'N69')!.page).toBe(1);
    const sep = l.pages[1]!.items.filter((i): i is LineItem => i.kind === 'line');
    expect(sep[0]!.x2 - sep[0]!.x1).toBe(468); // continuation separator spans the column
  });

  it('collects endnotes at the end of the document in reference order', () => {
    const e = (id: string): IRInline => ({
      kind: 'noteRef',
      note: 'endnote',
      id,
      props: runProps(),
      text: id,
    });
    const l = run(
      irDoc(
        [section([para('', { inlines: [e('2')] }), para('', { inlines: [e('1')] }), ...lines(3)])],
        {
          endnotes: new Map([
            ['1', [para('end one')]],
            ['2', [para('end two')]],
          ]),
        },
      ),
    );
    const order = pageText(l, 0);
    expect(order.indexOf('end ')).toBeGreaterThan(order.indexOf('L2'));
    const two = find(l, 'two')!;
    const one = find(l, 'one')!;
    expect(two.item.y).toBeLessThan(one.item.y);
  });
});

// ---- tables -------------------------------------------------------------------------------------

const B: BorderSpec = { style: 'single', size: 4, space: 0, color: 'auto' };

function tableProps(o: Partial<ResolvedTableProps> = {}): ResolvedTableProps {
  return {
    alignment: 'left',
    indent: 0,
    layout: 'fixed',
    cellMargins: { top: 0, left: 108, bottom: 0, right: 108 },
    cellSpacing: 0,
    borders: { top: B, left: B, bottom: B, right: B, insideH: B, insideV: B },
    look: {
      firstRow: false,
      lastRow: false,
      firstColumn: false,
      lastColumn: false,
      noHBand: false,
      noVBand: false,
    },
    bidiVisual: false,
    ...o,
  };
}

function cell(blocks: IRBlock[] | string, o: Partial<IRCell> = {}): IRCell {
  return {
    gridSpan: 1,
    blocks: typeof blocks === 'string' ? [para(blocks)] : blocks,
    margins: { top: 0, right: 5.4, bottom: 0, left: 5.4 },
    borders: {},
    vAlign: 'top',
    ...o,
  };
}

function row(cells: IRCell[], o: Partial<IRRow> = {}): IRRow {
  return { cantSplit: false, header: false, gridBefore: 0, gridAfter: 0, cells, ...o };
}

function table(
  rows: IRRow[],
  grid: number[] = [200, 200],
  o: Partial<ResolvedTableProps> = {},
): IRTable {
  return { kind: 'table', source: new XElement('w:tbl', NS.w), props: tableProps(o), grid, rows };
}

describe('tables', () => {
  it('uses the stored grid, cell margins, borders and shading', () => {
    const t = table([
      row([cell('a1', { shading: 'FFFF00' }), cell('b1')]),
      row([cell('a2'), cell('b2')]),
    ]);
    const l = run(irDoc([section([t])]));
    // rows are content + the 0.5pt horizontal border (Word sizes them so)
    expect(find(l, 'a1')!.item).toMatchObject({ x: 72 + 5.4, y: 82.5 });
    expect(find(l, 'b1')!.item.x).toBeCloseTo(72 + 200 + 5.4, 6);
    expect(find(l, 'a2')!.item.y).toBe(95);
    const items = l.pages[0]!.items;
    expect(
      items.some(
        (i) => i.kind === 'rect' && i.fill === 'FFFF00' && i.width === 200 && i.height === 12.5,
      ),
    ).toBe(true);
    const borders = items.filter((i): i is LineItem => i.kind === 'line');
    // 2 rows: top (2) + left/inner/right (3) each; bottom (2) on the last row
    expect(borders).toHaveLength(2 + 3 + 2 + 3 + 2);
    expect(borders[0]!.width).toBe(0.5);
  });

  it('honours row heights, vertical alignment and table alignment', () => {
    const t = table(
      [
        row([cell('x', { vAlign: 'bottom' }), cell('y', { vAlign: 'center' })], {
          height: { value: 48, rule: 'atLeast' },
        }),
      ],
      [100, 100],
      { alignment: 'center' },
    );
    const l = run(irDoc([section([t, para('after')])]));
    const x0 = 72 + (468 - 200) / 2;
    expect(find(l, 'x')!.item).toMatchObject({ x: x0 + 5.4, y: 72 + 35.5 + 10 });
    expect(find(l, 'y')!.item.y).toBe(72 + 18 + 10);
    expect(find(l, 'after')!.item.y).toBe(72 + 48 + 10);
    const exact = table([
      row([cell(lines(3)), cell('z')], { height: { value: 12, rule: 'exact' } }),
    ]);
    expect(find(run(irDoc([section([exact, para('next')])])), 'next')!.item.y).toBe(72 + 12 + 10);
  });

  it('spans merged cells across rows and grid columns', () => {
    const t = table([
      row([cell(lines(4, 'M'), { vMerge: 'restart' }), cell('r1')]),
      row([cell('', { vMerge: 'continue' }), cell('r2')]),
      row([cell('wide', { gridSpan: 2 })]),
    ]);
    const l = run(irDoc([section([t])]));
    // merged content 48pt (+0.5 border) over rows of 12.5 → the second row grows to 36
    expect(find(l, 'r2')!.item.y).toBe(72 + 12.5 + 0.5 + 10);
    expect(find(l, 'wide')!.item.y).toBe(72 + 48.5 + 0.5 + 10);
    expect(find(l, 'M3')!.item.y).toBe(72 + 0.5 + 36 + 10);
  });

  it('repeats header rows on each page and splits rows across pages', () => {
    const rows = [row([cell('Head'), cell('Head2')], { header: true })];
    for (let i = 0; i < 70; i++) rows.push(row([cell(`r${i}`), cell('v')]));
    const l = run(irDoc([section([table(rows)])]));
    expect(l.pages).toHaveLength(2);
    expect(pageText(l, 1)[0]).toBe('Head');
    expect(find(l, 'r53')!.page).toBe(1);
    // a row taller than the space left splits
    const big = [row([cell(lines(10, 'S')), cell('v')])];
    const l2 = run(irDoc([section([...lines(49), table(big)])]));
    expect(find(l2, 'S3')!.page).toBe(0);
    expect(find(l2, 'S4')!.page).toBe(1);
    const cant = [row([cell(lines(10, 'S')), cell('v')], { cantSplit: true })];
    const l3 = run(irDoc([section([...lines(49), table(cant)])]));
    expect(find(l3, 'S0')!.page).toBe(1);
  });

  it('autofits columns when there is no grid', () => {
    const t = table([row([cell('a'), cell('a much longer cell text')])], []);
    const l = run(irDoc([section([t])]));
    const a = find(l, 'a')!.item;
    const b = find(l, 'a ')!;
    // 'a' = 5pt + margins 10.8 → second column starts at 72 + 15.8
    expect(a.x).toBeCloseTo(72 + 5.4, 6);
    expect(b.item.x).toBeCloseTo(72 + 15.8 + 5.4, 6);
  });

  it('caps nesting depth and warns', () => {
    let inner: IRBlock = para('deep');
    for (let i = 0; i < 40; i++) inner = table([row([cell([inner])])], [300]);
    const warnings: string[] = [];
    const l = run(irDoc([section([inner])]), warnings);
    expect(warnings).toContain('LAYOUT_TABLE_DEPTH');
    expect(l.pages).toHaveLength(1);
  });
});

// ---- floats -------------------------------------------------------------------------------------

function anchor(o: Partial<IRAnchor>): IRAnchor {
  return {
    kind: 'image',
    source: new XElement('wp:anchor', NS.wp ?? 'urn:wp'),
    width: 100,
    height: 100,
    image: 'rId1',
    horizontal: { relativeTo: 'margin', offset: 0 },
    vertical: { relativeTo: 'paragraph', offset: 0 },
    wrap: 'square',
    behindText: false,
    distance: { top: 0, right: 0, bottom: 0, left: 0 },
    zIndex: 0,
    ...o,
  };
}

describe('floats', () => {
  it('positions floats relative to page, margin and paragraph', () => {
    const a = anchor({
      horizontal: { relativeTo: 'page', align: 'right' },
      vertical: { relativeTo: 'page', offset: 10 },
    });
    const b = anchor({
      image: 'b',
      horizontal: { relativeTo: 'margin', align: 'center' },
      vertical: { relativeTo: 'paragraph', offset: 5 },
      wrap: 'none',
    });
    const l = run(irDoc([section([para('x'), para('anchored', { anchors: [a, b] })])]));
    const imgs = l.pages[0]!.items.filter((i) => i.kind === 'image');
    expect(imgs).toHaveLength(2);
    expect(imgs.find((i) => i.kind === 'image' && i.image === 'rId1')).toMatchObject({
      x: 512,
      y: 10,
    });
    expect(imgs.find((i) => i.kind === 'image' && i.image === 'b')).toMatchObject({
      x: 72 + 184,
      y: 72 + 12 + 5,
    });
  });

  it('pushes text below a topAndBottom float', () => {
    const a = anchor({ wrap: 'topAndBottom', distance: { top: 0, right: 0, bottom: 6, left: 0 } });
    const l = run(irDoc([section([para('first', { anchors: [a] }), para('second')])]));
    // the anchor paragraph itself is pushed below the float (y 72..172 + 6)
    expect(find(l, 'first')!.item.y).toBe(72 + 106 + 10);
    expect(find(l, 'second')!.item.y).toBe(72 + 118 + 10);
  });

  it('wraps lines beside a square float through per-line insets', () => {
    const a = anchor({ wrap: 'square', distance: { top: 0, right: 10, bottom: 0, left: 0 } });
    const long = 'word '.repeat(150).trim(); // 750 chars, 358pt beside the float → 72 chars a line
    const l = run(irDoc([section([para(long, { anchors: [a] }), ...lines(12)])]));
    const words = texts(l.pages[0]!.items).filter((t) => t.text.startsWith('word'));
    const xs = new Map<number, number>();
    for (const w of words) xs.set(w.y, Math.min(xs.get(w.y) ?? 1e9, w.x));
    const rows = [...xs.entries()].sort((p, q) => p[0] - q[0]);
    // lines 0-8 (y 72..180) start beside the float, the rest at the margin
    expect(rows[0]![1]).toBe(72 + 110);
    expect(rows[8]![1]).toBe(72 + 110);
    expect(rows[9]![1]).toBe(72);
  });

  it('narrows the whole paragraph when the layouter ignores insets', () => {
    const a = anchor({ wrap: 'square', distance: { top: 0, right: 10, bottom: 0, left: 0 } });
    const l = layoutIR(
      irDoc([section([para('beside', { anchors: [a] }), ...lines(12)])]),
      makeEnv(),
      {
        layoutParagraph: (p, w) => fakeLayout(p, w),
      },
    );
    expect(find(l, 'beside')!.item.x).toBe(72 + 110);
    // lines below the float (y ≥ 172) use the full width again
    expect(find(l, 'L9')!.item.x).toBe(72);
  });

  it('draws behind-text floats first and in-front floats last', () => {
    const back = anchor({ image: 'back', wrap: 'none', behindText: true });
    const front = anchor({ image: 'front', wrap: 'none' });
    const l = run(irDoc([section([para('text', { anchors: [front, back] })])]));
    const kinds = l.pages[0]!.items.map((i) => (i.kind === 'image' ? i.image : i.kind));
    expect(kinds[0]).toBe('back');
    expect(kinds.at(-1)).toBe('front');
  });
});

// ---- fields, outline, destinations, PageInfoProvider ---------------------------------------------

describe('page fields and navigation', () => {
  it('resolves PAGEREF and NUMPAGES in the body with the fixed point', () => {
    const ref = para('', {
      inlines: [
        { kind: 'text', text: 'see page ', props: runProps() },
        { kind: 'field', type: 'PAGEREF', arg: 'target', fallback: '99', props: runProps() },
        { kind: 'text', text: ' of ', props: runProps() },
        { kind: 'field', type: 'NUMPAGES', fallback: '1', props: runProps() },
      ],
    });
    const target = para('Target', { bookmarks: ['target'] });
    const l = run(irDoc([section([ref, ...lines(60), target])]));
    expect(find(l, '2')!.page).toBe(0);
    expect(l.destinations.get('target')).toEqual({ page: 1, y: 72 + 7 * 12 });
    expect(pageText(l, 0).filter((t) => /^\d+$/.test(t))).toEqual(['2', '2']);
    // one pass only: the stale fallback is substituted at emission
    const once = run(irDoc([section([ref, ...lines(60), target])]), [], 1);
    expect(pageText(once, 0).filter((t) => /^\d+$/.test(t))).toEqual(['2', '2']);
  });

  it('builds the outline from heading levels', () => {
    const l = run(
      irDoc([
        section([
          para('Intro', { outlineLevel: 0 }),
          ...lines(60),
          para('Details', { outlineLevel: 1 }),
        ]),
      ]),
    );
    expect(l.outline).toEqual([
      { title: 'Intro', level: 0, page: 0, y: 72 },
      { title: 'Details', level: 1, page: 1, y: 72 + 7 * 12 },
    ]);
  });

  it('provides page info for fields (1-based) through paragraphs', () => {
    const ps = [...lines(60)];
    const doc = irDoc([section(ps.slice(0, 30), { pageNumberStart: 3 }), section(ps.slice(30))]);
    const l = run(doc);
    const info = pageInfo(l);
    expect(info.pageCount()).toBe(2);
    expect(info.pageOf(ps[0]!.source)).toBe(1);
    expect(info.pageOf(ps[40]!.source)).toBe(2);
    expect(info.pageNumberText?.(ps[40]!.source)).toBe('4');
    expect(info.sectionPageCount?.(ps[0]!.source)).toBe(1);
    // a node inside a paragraph resolves through its w:p
    const run0 = new XElement('w:r', NS.w);
    ps[40]!.source.append(run0);
    expect(info.pageOf(run0)).toBe(2);
    expect(info.pageOf(new XElement('w:p', NS.w))).toBeUndefined();
  });
});

describe('scale', () => {
  it('lays out a long document in linear time', () => {
    const blocks: IRBlock[] = [];
    for (let i = 0; i < 4000; i++) {
      blocks.push(
        para('word '.repeat(40).trim(), { widowControl: true, after: 6, keepNext: i % 10 === 0 }),
      );
      if (i % 50 === 0)
        blocks.push(table([row([cell('a'), cell('b')]), row([cell('c'), cell('d')])]));
    }
    const t0 = performance.now();
    const l = run(
      irDoc([
        section(blocks, {
          footers: {
            default: [
              para('', {
                inlines: [{ kind: 'field', type: 'PAGE', fallback: '0', props: runProps() }],
              }),
            ],
          },
        }),
      ]),
    );
    const ms = performance.now() - t0;
    expect(l.pages.length).toBeGreaterThan(200);
    expect(ms).toBeLessThan(5000);
  });
});

describe('helpers', () => {
  it('translates SVG path data', () => {
    expect(translatePath('M0 0 L10 5 H20 V30 C1 2 3 4 5 6 Z', 1, 2)).toBe(
      'M 1 2 L 11 7 H 21 V 32 C 2 4 4 6 6 8 Z',
    );
    expect(translatePath('m0 0 l10 5 A5 5 0 0 1 10 10', 1, 2)).toBe(
      'm 1 2 l 10 5 A 5 5 0 0 1 11 12',
    );
  });
});

// ---- end to end ---------------------------------------------------------------------------------

describe('paginate() on real documents', () => {
  const corpus = [...fixtures(), ...strictFixtures()];

  it.each(corpus.map((f) => [f.name, f] as const))('lays out %s consistently', async (_name, f) => {
    const doc = readDocx(f.bytes);
    const fonts = testFontResolver();
    const a = await paginate(doc, { fonts });
    expect(a.pages.length).toBeGreaterThanOrEqual(1);
    for (const p of a.pages) {
      expect(p.width).toBeGreaterThan(0);
      for (const it of p.items) {
        const nums =
          it.kind === 'line'
            ? [it.x1, it.y1, it.x2, it.y2]
            : it.kind === 'path'
              ? []
              : [it.x, it.y];
        for (const n of nums) expect(Number.isFinite(n)).toBe(true);
        if (it.kind === 'text') {
          expect(a.fonts[it.font]).toBeDefined();
          expect(it.glyphs.length).toBe(it.advances.length);
        }
        if (it.kind === 'image') expect(a.images.has(it.image)).toBe(true);
      }
    }
    for (const page of a.paragraphPages.values()) expect(page).toBeLessThan(a.pages.length);
    for (const d of a.destinations.values()) expect(d.page).toBeLessThan(a.pages.length);
    // deterministic: the same document lays out to the same PageModel
    const b = await paginate(readDocx(f.bytes), { fonts });
    expect(JSON.stringify(b.pages)).toBe(JSON.stringify(a.pages));
    expect(await pageCount(readDocx(f.bytes), { fonts })).toBe(a.pages.length);
  });

  it('gives fields 1-based page numbers through pageInfo()', async () => {
    const doc = readDocx(fixture('docx-mailmerge/test_footnote_header_footer.docx').bytes);
    const l = await paginate(doc, { fonts: testFontResolver() });
    const info = pageInfo(l);
    const paras = doc.body.element.descendants(NS.w, 'p');
    expect(info.pageCount()).toBe(3);
    expect(info.pageOf(paras[0]!)).toBe(1);
    expect(info.pageOf(paras[paras.length - 1]!)).toBe(3);
  });
});
