// layoutParagraph: line breaking, tabs, indents, justification, line heights and decorations.
//
// Most tests use a fake face with fixed metrics so every expected position is exact arithmetic:
// 1000 units per em, every glyph 500 units wide except the space (250), win ascent 800 / descent
// 200, no line gap. At 10 pt a character is 5 pt, a space 2.5 pt, a line 10 pt (8 + 2). The
// real UAX #14 breaker and basic shaper are used. The last block uses the real Liberation and
// Carlito test fonts.

import { describe, expect, it } from 'vitest';
import { createDocx, readDocx } from '../src/index.js';
import type { FontManager, LayoutEnv, LineBox, ParagraphLayout } from '../src/layout/engine.js';
import { createFontManager } from '../src/layout/fonts/index.js';
import { buildIR, type IRParagraphExtras, wantFonts } from '../src/layout/ir.js';
import { layoutParagraph } from '../src/layout/paragraph.js';
import { basicShaper } from '../src/layout/text/index.js';
import type {
  DisplayItem,
  Font,
  FontMetrics,
  IRInline,
  IRParagraph,
  TextItem,
} from '../src/layout/types.js';
import type { ResolvedParagraphProps, ResolvedRunProps } from '../src/resolve/types.js';
import { XElement } from '../src/xml/dom.js';
import { NS } from '../src/xml/ns.js';
import { fixtures } from './helpers/fixtures.js';
import { testFontResolver } from './helpers/test-fonts.js';

// ---- fakes ---------------------------------------------------------------------------------------

const METRICS: FontMetrics = {
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
};

function fakeFont(o: { name?: string; bold?: boolean; metrics?: Partial<FontMetrics> } = {}): Font {
  return {
    family: o.name ?? 'Fake',
    subfamily: o.bold ? 'Bold' : 'Regular',
    postscriptName: o.name ?? 'Fake',
    bold: o.bold ?? false,
    italic: false,
    metrics: { ...METRICS, ...o.metrics },
    outlines: 'truetype',
    // Glyph id = code point (so tests can read the glyphs back as text).
    glyphId: (cp) => cp,
    advance: (gid) => (gid === 0x20 ? 250 : 500),
    // "AV" kerns by -100 units.
    kerning: (l, r) => (l === 0x41 && r === 0x56 ? -100 : 0),
    bytes: new Uint8Array(0),
    numGlyphs: 0x110000,
  };
}

const FONT = fakeFont();

function fakeFonts(font: Font = FONT): FontManager {
  const used: Font[] = [];
  return {
    want() {},
    async load() {},
    fontFor() {
      if (!used.includes(font)) used.push(font);
      return font;
    },
    fonts: () => used,
    indexOf: (f) => used.indexOf(f),
  };
}

function env(fonts: FontManager = fakeFonts()): LayoutEnv {
  const doc = createDocx();
  return {
    doc,
    partOf: () => doc.mainPart,
    fonts,
    shaper: basicShaper,
    defaultTabStop: 36,
    images: new Map(),
    warn: () => {},
  };
}

function runProps(o: Partial<ResolvedRunProps> = {}): ResolvedRunProps {
  return {
    fonts: { ascii: 'Fake', hAnsi: 'Fake' },
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

function paraProps(o: Partial<ResolvedParagraphProps> = {}): ResolvedParagraphProps {
  const rp = runProps();
  return {
    alignment: 'left',
    indent: { left: 0, right: 0, firstLine: 0, hanging: 0 },
    spacing: {
      before: 0,
      after: 0,
      beforeAutospacing: false,
      afterAutospacing: false,
      line: 240,
      lineRule: 'auto',
    },
    contextualSpacing: false,
    keepNext: false,
    keepLines: false,
    pageBreakBefore: false,
    widowControl: true,
    suppressAutoHyphens: false,
    bidi: false,
    tabs: [],
    borders: {},
    runDefaults: rp,
    markRun: rp,
    ...o,
  };
}

const SOURCE = new XElement('w:p', NS.w);

function para(
  inlines: (IRInline | string)[],
  o: Partial<ResolvedParagraphProps> = {},
  extra: Partial<IRParagraph> = {},
): IRParagraph {
  return {
    kind: 'paragraph',
    source: SOURCE,
    props: paraProps(o),
    inlines: inlines.map((i) => (typeof i === 'string' ? text(i) : i)),
    anchors: [],
    bookmarks: [],
    ...extra,
  };
}

function text(t: string, o: Partial<ResolvedRunProps> = {}): IRInline {
  return { kind: 'text', text: t, props: runProps(o) };
}

const TAB: IRInline = { kind: 'tab', props: runProps() };

function texts(line: LineBox): TextItem[] {
  return line.items.filter((i): i is TextItem => i.kind === 'text');
}

/** The characters of a line in visual order. */
function lineText(line: LineBox): string {
  return texts(line)
    .sort((a, b) => a.x - b.x)
    .map((t) => t.text)
    .join('');
}

/** x of the glyph showing character `ch` (first occurrence) in a line. */
function xOf(line: LineBox, ch: string): number {
  for (const t of texts(line)) {
    const i = t.text.indexOf(ch);
    if (i < 0) continue;
    let x = t.x;
    for (let g = 0; g < t.glyphs.length && t.clusters[g]! < i; g++) x += t.advances[g]!;
    return x;
  }
  throw new Error(`no ${ch} in line`);
}

function lay(p: IRParagraph, width = 100, e: LayoutEnv = env()): ParagraphLayout {
  return layoutParagraph(p, width, e);
}

// ---- line breaking ---------------------------------------------------------------------------------

describe('layoutParagraph: greedy line breaking', () => {
  it('breaks after the last word that fits; trailing spaces hang', () => {
    // "aaaa bbbb " = 42.5 pt, + "cccc" = 65 > 60.
    const pl = lay(para(['aaaa bbbb cccc dddd']), 60);
    expect(pl.lines.map(lineText)).toEqual(['aaaa bbbb ', 'cccc dddd']);
    expect(xOf(pl.lines[0]!, 'b')).toBeCloseTo(22.5);
    expect(xOf(pl.lines[1]!, 'd')).toBeCloseTo(22.5);
  });

  it('a word that exactly fits stays on the line (spaces after it do not count)', () => {
    const pl = lay(para(['aaaa bbbb   cccc']), 42.5);
    expect(pl.lines.map(lineText)).toEqual(['aaaa bbbb   ', 'cccc']);
  });

  it('breaks inside a word wider than the line, at the overflowing character', () => {
    const pl = lay(para(['aaaaaaa']), 12);
    expect(pl.lines.map(lineText)).toEqual(['aa', 'aa', 'aa', 'a']);
  });

  it('moves a long word to the next line before breaking it', () => {
    const pl = lay(para(['ab cdefghij']), 30);
    expect(pl.lines.map(lineText)).toEqual(['ab ', 'cdefgh', 'ij']);
  });

  it('breaks at hyphens and keeps non-breaking hyphens together', () => {
    expect(lay(para(['aaa-bbb']), 25).lines.map(lineText)).toEqual(['aaa-', 'bbb']);
    const nb = lay(para(['aaa‑bbb ccc']), 40).lines.map(lineText);
    expect(nb[0]).toBe('aaa‑bbb ');
  });

  it('line breaks end lines; a final line break adds an empty line', () => {
    const pl = lay(
      para(['aa', { kind: 'break', type: 'line' }, 'bb', { kind: 'break', type: 'line' }]),
    );
    expect(pl.lines.map(lineText)).toEqual(['aa', 'bb', '']);
    expect(pl.lines.every((l) => l.breakAfter === undefined)).toBe(true);
  });

  it('page and column breaks set breakAfter; a final page break adds no line', () => {
    const pl = lay(
      para(['aa', { kind: 'break', type: 'page' }, 'bb', { kind: 'break', type: 'column' }]),
    );
    expect(pl.lines.map(lineText)).toEqual(['aa', 'bb']);
    expect(pl.lines.map((l) => l.breakAfter)).toEqual(['page', 'column']);
  });

  it('a paragraph holding only a page break is one line', () => {
    const pl = lay(para([{ kind: 'break', type: 'page' }]));
    expect(pl.lines).toHaveLength(1);
    expect(pl.lines[0]!.breakAfter).toBe('page');
  });

  it('shows a soft hyphen only where the line breaks after it', () => {
    const one = lay(para(['aaa­bbb']), 100);
    expect(lineText(one.lines[0]!)).toBe('aaabbb');
    const two = lay(para(['aaa­bbb']), 20);
    expect(two.lines.map(lineText)).toEqual(['aaa-', 'bbb']);
  });

  it('breaks CJK text between ideographs', () => {
    const pl = lay(para(['一二三四五']), 12);
    expect(pl.lines.map(lineText)).toEqual(['一二', '三四', '五']);
  });
});

// ---- indents and alignment ------------------------------------------------------------------------

describe('layoutParagraph: indents and alignment', () => {
  it('applies left/right indents and a first-line indent', () => {
    const p = para(['aaaa bbbb cccc'], {
      indent: { left: 400, right: 200, firstLine: 200, hanging: 0 },
    });
    // Width 60 - 20 - 10 = 30 available; the first line starts 10 pt in.
    const pl = lay(p, 60);
    expect(pl.indentLeft).toBe(20);
    expect(pl.lines.map(lineText)).toEqual(['aaaa ', 'bbbb ', 'cccc']);
    expect(texts(pl.lines[0]!)[0]!.x).toBeCloseTo(10);
    expect(texts(pl.lines[1]!)[0]!.x).toBeCloseTo(0);
    // Widths are measured from the left indent: the first line's indent counts.
    expect(pl.lines.map((l) => l.width)).toEqual([30, 20, 20]);
  });

  it('a hanging indent puts the first line left of the indent', () => {
    const pl = lay(
      para(['aaaa bbbb'], { indent: { left: 720, right: 0, firstLine: 0, hanging: 360 } }),
      50,
    );
    expect(texts(pl.lines[0]!)[0]!.x).toBeCloseTo(-18);
    expect(texts(pl.lines[1]!)[0]!.x).toBeCloseTo(0);
  });

  it('centres and right-aligns lines', () => {
    expect(texts(lay(para(['aaaa'], { alignment: 'center' })).lines[0]!)[0]!.x).toBeCloseTo(40);
    expect(texts(lay(para(['aaaa'], { alignment: 'right' })).lines[0]!)[0]!.x).toBeCloseTo(80);
  });

  it('justifies over the spaces except on the last line', () => {
    const pl = lay(para(['aaaa bbbb cccc dddd'], { alignment: 'both' }), 60);
    // Line 1 "aaaa bbbb": 17.5 pt free goes to its one inner space.
    expect(xOf(pl.lines[0]!, 'b')).toBeCloseTo(40);
    expect(pl.lines[0]!.width).toBeCloseTo(60);
    expect(pl.lines[1]!.width).toBeCloseTo(42.5);
    expect(xOf(pl.lines[1]!, 'd')).toBeCloseTo(22.5);
  });

  it('justifies lines ending in a line break unless doNotExpandShiftReturn', () => {
    const inl: (IRInline | string)[] = ['aa bb', { kind: 'break', type: 'line' }, 'cc'];
    expect(xOf(lay(para(inl, { alignment: 'both' }), 60).lines[0]!, 'b')).toBeCloseTo(50);
    const noExpand = para(inl, { alignment: 'both' }) as IRParagraph & IRParagraphExtras;
    noExpand.compat = { doNotExpandShiftReturn: true };
    expect(xOf(lay(noExpand, 60).lines[0]!, 'b')).toBeCloseTo(12.5);
  });

  it('distribute spreads every line, the last one included', () => {
    const pl = lay(para(['abc'], { alignment: 'distribute' }), 25);
    expect(xOf(pl.lines[0]!, 'b')).toBeCloseTo(10);
    expect(xOf(pl.lines[0]!, 'c')).toBeCloseTo(20);
  });

  it('mirrors RTL paragraphs: start alignment is the right edge', () => {
    const pl = lay(para(['aaaa'], { bidi: true }));
    expect(texts(pl.lines[0]!)[0]!.x).toBeCloseTo(80);
    const right = lay(para(['aaaa'], { bidi: true, alignment: 'right' }));
    expect(texts(right.lines[0]!)[0]!.x).toBeCloseTo(0);
  });

  it('reorders right-to-left text visually', () => {
    // Hebrew alef-bet-gimel in an LTR paragraph: shown right to left.
    const pl = lay(para(['x אבג']));
    const line = pl.lines[0]!;
    expect(xOf(line, 'ג')).toBeLessThan(xOf(line, 'א'));
    expect(xOf(line, 'x')).toBeCloseTo(0);
  });
});

// ---- tabs ----------------------------------------------------------------------------------------

describe('layoutParagraph: tabs', () => {
  it('uses default stops every defaultTabStop', () => {
    const pl = lay(para(['a', TAB, 'b', TAB, 'c']), 200);
    expect(xOf(pl.lines[0]!, 'b')).toBeCloseTo(36);
    expect(xOf(pl.lines[0]!, 'c')).toBeCloseTo(72);
  });

  it('custom left stops clear the default stops before them', () => {
    const pl = lay(para(['a', TAB, 'b', TAB, 'c'], { tabs: [{ pos: 2000, val: 'left' }] }), 300);
    expect(xOf(pl.lines[0]!, 'b')).toBeCloseTo(100);
    // Past the last custom stop the default stops apply again.
    expect(xOf(pl.lines[0]!, 'c')).toBeCloseTo(108);
  });

  it('right, centre and decimal stops', () => {
    const right = lay(para(['a', TAB, 'bbb'], { tabs: [{ pos: 2000, val: 'right' }] }), 200);
    expect(xOf(right.lines[0]!, 'b')).toBeCloseTo(85);
    const center = lay(para(['a', TAB, 'bbbb'], { tabs: [{ pos: 2000, val: 'center' }] }), 200);
    expect(xOf(center.lines[0]!, 'b')).toBeCloseTo(90);
    const dec = lay(para(['a', TAB, '12.34'], { tabs: [{ pos: 2000, val: 'decimal' }] }), 200);
    expect(xOf(dec.lines[0]!, '.')).toBeCloseTo(100);
  });

  it('a right stop past the right indent is honoured', () => {
    const pl = lay(para(['a', TAB, 'bb'], { tabs: [{ pos: 2400, val: 'right' }] }), 100);
    expect(pl.lines).toHaveLength(1);
    expect(xOf(pl.lines[0]!, 'b')).toBeCloseTo(110);
  });

  it('draws leaders and bar tabs', () => {
    const pl = lay(
      para(['a', TAB, 'b'], {
        tabs: [
          { pos: 1000, val: 'bar' },
          { pos: 2000, val: 'right', leader: 'dot' },
        ],
      }),
      200,
    );
    const line = pl.lines[0]!;
    const leader = texts(line).find((t) => /^\.+$/.test(t.text));
    expect(leader).toBeDefined();
    expect(leader!.x).toBeGreaterThanOrEqual(5);
    expect(leader!.x + leader!.advances.reduce((s, v) => s + v, 0)).toBeLessThanOrEqual(95.01);
    const bar = line.items.find((i) => i.kind === 'line' && i.x1 === i.x2);
    expect(bar).toMatchObject({ x1: 50 });
  });

  it('a hanging indent acts as a tab stop for the list label', () => {
    const p = para(
      ['item'],
      { indent: { left: 720, right: 0, firstLine: 0, hanging: 360 } },
      {
        label: { text: '1.', props: runProps(), suffix: 'tab', alignment: 'left' },
      },
    );
    const line = lay(p, 200).lines[0]!;
    expect(xOf(line, '1')).toBeCloseTo(-18);
    expect(xOf(line, 'i')).toBeCloseTo(0);
  });

  it('a label wider than the hanging indent tabs to the next default stop', () => {
    const p = para(
      ['item'],
      { indent: { left: 360, right: 0, firstLine: 0, hanging: 360 } },
      {
        label: { text: '10.10.', props: runProps(), suffix: 'tab', alignment: 'left' },
      },
    );
    expect(xOf(lay(p, 200).lines[0]!, 'i')).toBeCloseTo(36 - 18);
  });

  it('right-aligned labels end at the first-line start; space suffix', () => {
    const p = para(
      ['x'],
      { indent: { left: 720, right: 0, firstLine: 0, hanging: 0 } },
      {
        label: { text: 'iv.', props: runProps(), suffix: 'space', alignment: 'right' },
      },
    );
    const line = lay(p, 200).lines[0]!;
    expect(xOf(line, 'i')).toBeCloseTo(-15);
    expect(xOf(line, 'x')).toBeCloseTo(2.5);
  });

  it('positional tabs align to the margin', () => {
    const ptab = {
      kind: 'tab',
      props: runProps(),
      ptab: { alignment: 'right', relativeTo: 'margin' },
    } as IRInline;
    const line = lay(para(['a', ptab, 'bb']), 100).lines[0]!;
    expect(xOf(line, 'b')).toBeCloseTo(90);
  });
});

// ---- character formatting ------------------------------------------------------------------------

describe('layoutParagraph: character formatting', () => {
  it('superscript and subscript: 2/3 size, shifted baseline', () => {
    const line = lay(
      para(['a', text('b', { vertAlign: 'superscript' }), text('c', { vertAlign: 'subscript' })]),
    ).lines[0]!;
    const [, sup, sub] = texts(line);
    expect(sup!.size).toBeCloseTo(20 / 3);
    expect(sup!.y).toBeCloseTo(-10 / 3);
    expect(sub!.size).toBeCloseTo(20 / 3);
    expect(sub!.y).toBeCloseTo(1.4);
  });

  it('caps and small caps', () => {
    const caps = lay(para([text('abc', { caps: true })])).lines[0]!;
    expect(lineText(caps)).toBe('ABC');
    const sc = texts(lay(para([text('Abc', { smallCaps: true })])).lines[0]!);
    expect(sc.map((t) => [t.text, t.size])).toEqual([
      ['A', 10],
      ['BC', 8],
    ]);
  });

  it('character spacing, scale, position and kerning', () => {
    const sp = lay(para([text('ab', { spacing: 20 })])).lines[0]!;
    expect(xOf(sp, 'b')).toBeCloseTo(6);
    const sc = texts(lay(para([text('ab', { scale: 50 })])).lines[0]!)[0]!;
    expect(sc.scale).toBe(50);
    expect(sc.advances[0]).toBeCloseTo(2.5);
    const pos = texts(lay(para([text('a', { position: 6 })])).lines[0]!)[0]!;
    expect(pos.y).toBeCloseTo(-3);
    expect(xOf(lay(para([text('AV', { kern: 0 })])).lines[0]!, 'V')).toBeCloseTo(5);
    expect(xOf(lay(para([text('AV', { kern: 8 })])).lines[0]!, 'V')).toBeCloseTo(4);
    expect(xOf(lay(para([text('AV', { kern: 12 })])).lines[0]!, 'V')).toBeCloseTo(5);
  });

  it('underline, strike-through and highlight', () => {
    const line = lay(
      para([
        text('ab', { underline: 'single', strike: true, highlight: 'yellow', color: 'FF0000' }),
      ]),
    ).lines[0]!;
    const kinds = line.items.map((i) => i.kind);
    expect(kinds[0]).toBe('rect');
    expect(line.items[0]).toMatchObject({ fill: 'FFFF00', x: 0, width: 10, y: -8, height: 10 });
    const lines = line.items.filter(
      (i): i is Extract<DisplayItem, { kind: 'line' }> => i.kind === 'line',
    );
    expect(lines).toHaveLength(2);
    expect(lines.every((l) => l.x1 === 0 && l.x2 === 10 && l.color === 'FF0000')).toBe(true);
    // Underline below the baseline, strike-through above it.
    expect(lines.some((l) => l.y1 > 0)).toBe(true);
    expect(lines.some((l) => l.y1 < 0)).toBe(true);
    expect(texts(line)[0]!.color).toBe('FF0000');
  });

  it('words underline skips spaces; double underline draws two lines', () => {
    const w = lay(para([text('ab cd', { underline: 'words' })])).lines[0]!;
    expect(w.items.filter((i) => i.kind === 'line')).toHaveLength(2);
    const d = lay(para([text('ab', { underline: 'double' })])).lines[0]!;
    expect(d.items.filter((i) => i.kind === 'line')).toHaveLength(2);
  });

  it('auto colour is white on dark shading', () => {
    expect(texts(lay(para([text('a', { shading: '000080' })])).lines[0]!)[0]!.color).toBe('FFFFFF');
    expect(texts(lay(para([text('a')])).lines[0]!)[0]!.color).toBe('000000');
  });

  it('hyperlinks produce link areas', () => {
    const link: IRInline = {
      kind: 'text',
      text: 'ab',
      props: runProps(),
      link: { url: 'https://x.test/' },
    };
    const line = lay(para(['x ', link])).lines[0]!;
    expect(line.items.find((i) => i.kind === 'link')).toMatchObject({
      x: 7.5,
      width: 10,
      url: 'https://x.test/',
    });
  });
});

// ---- heights -------------------------------------------------------------------------------------

describe('layoutParagraph: line heights and spacing', () => {
  const sp = (line: number, lineRule: 'auto' | 'exact' | 'atLeast') => ({
    spacing: {
      before: 0,
      after: 0,
      beforeAutospacing: false,
      afterAutospacing: false,
      line,
      lineRule,
    },
  });

  it('single, multiple, exact and at-least spacing', () => {
    expect(lay(para(['a'])).lines[0]).toMatchObject({ height: 10, baseline: 8 });
    expect(lay(para(['a'], sp(360, 'auto'))).lines[0]).toMatchObject({ height: 15, baseline: 13 });
    expect(lay(para(['a'], sp(300, 'exact'))).lines[0]).toMatchObject({ height: 15, baseline: 13 });
    expect(lay(para(['a'], sp(100, 'exact'))).lines[0]).toMatchObject({ height: 5, baseline: 3 });
    expect(lay(para(['a'], sp(100, 'atLeast'))).lines[0]!.height).toBe(10);
    expect(lay(para(['a'], sp(400, 'atLeast'))).lines[0]!.height).toBe(20);
  });

  it('uses the largest font on the line and the external leading', () => {
    const line = lay(para(['a', text('B', { size: 20 })])).lines[0]!;
    expect(line).toMatchObject({ height: 20, baseline: 16 });
    // hhea 900/-300/gap 200 vs win 800/200: leading = 200 - (1000 - 1200) = 400 units.
    const leaded = fakeFont({ metrics: { ascender: 900, descender: -300, lineGap: 200 } });
    expect(lay(para(['a']), 100, env(fakeFonts(leaded))).lines[0]).toMatchObject({
      height: 14,
      baseline: 12,
    });
  });

  it('an empty paragraph has the height of its mark', () => {
    const p = para([], { markRun: runProps({ size: 20 }) });
    const pl = lay(p);
    expect(pl.lines).toHaveLength(1);
    expect(pl.lines[0]).toMatchObject({ height: 20, baseline: 16, items: [] });
  });

  it('the paragraph mark counts on the last line only', () => {
    const p = para(['aaaa bbbb'], { markRun: runProps({ size: 20 }) });
    const pl = lay(p, 30);
    expect(pl.lines.map((l) => l.height)).toEqual([10, 20]);
  });

  it('inline images sit on the baseline and raise the line', () => {
    const img: IRInline = {
      kind: 'image',
      image: 'word/media/image1.png',
      width: 30,
      height: 50,
      props: runProps(),
    };
    const line = lay(para(['a', img])).lines[0]!;
    expect(line).toMatchObject({ height: 52, baseline: 50 });
    expect(line.items.find((i) => i.kind === 'image')).toMatchObject({
      x: 5,
      y: -50,
      width: 30,
      height: 50,
    });
  });

  it('multiple spacing scales the text, not inline pictures', () => {
    const img: IRInline = {
      kind: 'image',
      image: 'word/media/image1.png',
      width: 30,
      height: 50,
      props: runProps(),
    };
    // Text 10 pt × 1.5 = 15, plus the 42 pt the picture rises above the text ascent.
    const line = lay(para(['a', img], sp(360, 'auto'))).lines[0]!;
    expect(line).toMatchObject({ height: 57, baseline: 55 });
    // A picture alone on a line: its run's font still sets the text part.
    expect(lay(para([img], sp(276, 'auto'))).lines[0]!.height).toBeCloseTo(11.5 + 42);
  });

  it('multiple spacing does not scale inline charts/drawings either', () => {
    const chart: IRInline = {
      kind: 'drawing',
      source: new XElement('w:drawing', NS.w),
      width: 400,
      height: 315,
      props: runProps(),
    };
    const line = lay(para([chart], sp(259, 'auto')), 500).lines[0]!;
    expect(line.height).toBeCloseTo((10 * 259) / 240 + 307, 6);
    expect(line.baseline).toBeCloseTo(line.height - 2, 6);
  });

  it('exact and at-least spacing with pictures', () => {
    const img: IRInline = { kind: 'image', image: 'i', width: 30, height: 50, props: runProps() };
    expect(lay(para(['a', img], sp(400, 'exact'))).lines[0]).toMatchObject({
      height: 20,
      baseline: 18,
    });
    expect(lay(para(['a', img], sp(400, 'atLeast'))).lines[0]).toMatchObject({
      height: 52,
      baseline: 50,
    });
  });

  it('space before/after in points, autospacing = 14 pt, flags passed through', () => {
    const p = para(['a'], {
      spacing: {
        before: 240,
        after: 120,
        beforeAutospacing: false,
        afterAutospacing: true,
        line: 240,
        lineRule: 'auto',
      },
      keepNext: true,
      contextualSpacing: true,
      styleId: 'Heading1',
    });
    expect(lay(p)).toMatchObject({
      spaceBefore: 12,
      spaceAfter: 14,
      keepNext: true,
      contextualSpacing: true,
      styleId: 'Heading1',
    });
  });

  it('paragraph borders add their space to the first/last line; shading behind', () => {
    const border = { style: 'single', size: 8, space: 4, color: 'auto' };
    const p = para(['aaaa bbbb'], {
      borders: { top: border, bottom: border, left: border },
      shading: 'EEEEEE',
    });
    const pl = lay(p, 30);
    expect(pl.lines[0]).toMatchObject({ height: 15, baseline: 13 });
    expect(pl.lines[1]!.height).toBe(15);
    const rect = pl.decorations.find((d) => d.kind === 'rect');
    expect(rect).toMatchObject({ x: -4, y: 1, width: 34, height: 28, fill: 'EEEEEE' });
    expect(pl.decorations.filter((d) => d.kind === 'line')).toHaveLength(3);
  });
});

// ---- fields, notes, bookmarks ------------------------------------------------------------------

describe('layoutParagraph: fields, note references and bookmarks', () => {
  it('page fields become placeholder text items with a slot', () => {
    const field: IRInline = { kind: 'field', type: 'PAGE', fallback: '12', props: runProps() };
    const line = lay(para(['Page ', field, ' of'])).lines[0]!;
    expect(line.fields).toHaveLength(1);
    const slot = line.fields[0]!;
    const item = line.items[slot.item] as TextItem;
    expect(item.kind).toBe('text');
    expect(item.text).toBe('12');
    expect(item.x).toBeCloseTo(22.5);
    expect(slot.field.type).toBe('PAGE');
  });

  it('reports note references (footnotes and endnotes) and bookmarks on their lines', () => {
    const ref: IRInline = {
      kind: 'noteRef',
      note: 'footnote',
      id: '3',
      text: '1',
      props: runProps({ vertAlign: 'superscript' }),
    };
    const end: IRInline = {
      kind: 'noteRef',
      note: 'endnote',
      id: '7',
      text: 'i',
      props: runProps(),
    };
    const p = para(
      ['aaaa ', 'bbbb', ref, end],
      {},
      { bookmarks: ['start', 'mid'] },
    ) as IRParagraph & IRParagraphExtras;
    p.bookmarkAt = [0, 1];
    const pl = lay(p, 30);
    expect(pl.lines.map((l) => l.noteRefs)).toEqual([[], ['3', '7']]);
    expect(pl.lines.map((l) => l.bookmarks)).toEqual([['start'], ['mid']]);
  });
});

// ---- real fonts --------------------------------------------------------------------------------

describe('layoutParagraph: real test fonts', () => {
  async function realEnv(families: string[]): Promise<LayoutEnv> {
    const doc = createDocx();
    const fonts = createFontManager(doc, { fonts: testFontResolver() });
    for (const f of families) fonts.want(runProps({ fonts: { ascii: f, hAnsi: f } }));
    await fonts.load();
    return {
      doc,
      partOf: () => doc.mainPart,
      fonts,
      shaper: basicShaper,
      defaultTabStop: 36,
      images: new Map(),
      warn: () => {},
    };
  }

  it('Calibri (Carlito) 11 pt single spacing is 13.43 pt; Times New Roman 12 pt is 13.8 pt', async () => {
    const e = await realEnv(['Calibri', 'Times New Roman']);
    const cal = runProps({ fonts: { ascii: 'Calibri', hAnsi: 'Calibri' }, size: 11 });
    const p = para([{ kind: 'text', text: 'Hello', props: cal }], { markRun: cal });
    expect(layoutParagraph(p, 400, e).lines[0]!.height).toBeCloseTo(13.43, 2);
    const tnr = runProps({
      fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman' },
      size: 12,
    });
    const q = para([{ kind: 'text', text: 'Hello', props: tnr }], { markRun: tnr });
    expect(layoutParagraph(q, 400, e).lines[0]!.height).toBeCloseTo(13.8, 1);
  });

  it('breaks Arial (Liberation Sans) text by its real advances', async () => {
    const e = await realEnv(['Arial']);
    const arial = runProps({ fonts: { ascii: 'Arial', hAnsi: 'Arial' }, size: 12 });
    const words = 'The quick brown fox jumps over the lazy dog. '.repeat(4).trim();
    const p = para([{ kind: 'text', text: words, props: arial }], { markRun: arial });
    const pl = layoutParagraph(p, 200, e);
    expect(pl.lines.length).toBeGreaterThan(3);
    for (const l of pl.lines) {
      const t = texts(l);
      const right = Math.max(...t.map((i) => i.x + i.advances.reduce((s, v) => s + v, 0)));
      // Content may only exceed the width by hanging spaces.
      expect(l.width).toBeLessThanOrEqual(200.01);
      expect(right).toBeLessThan(215);
    }
    expect(pl.lines.map(lineText).join('')).toBe(words);
  });
});

// ---- per-line insets (wrapping beside floats) -----------------------------------------------------

describe('layoutParagraph: per-line insets', () => {
  const words = 'aaaa bbbb cccc dddd eeee ffff';

  it('narrows each line by its own insets; items move with a left inset', () => {
    // Line 0: 20 pt free on the left (box 40 pt), line 1: 20 pt on the right, then none.
    const insets = (line: number) =>
      line === 0
        ? { left: 20, right: 0 }
        : line === 1
          ? { left: 0, right: 20 }
          : { left: 0, right: 0 };
    const pl = layoutParagraph(para([words]), 60, env(), { insets });
    expect(pl.lines.map(lineText)).toEqual(['aaaa ', 'bbbb ', 'cccc dddd ', 'eeee ffff']);
    expect(texts(pl.lines[0]!)[0]!.x).toBeCloseTo(20);
    expect(texts(pl.lines[1]!)[0]!.x).toBeCloseTo(0);
    expect(pl.lines[0]!.width).toBeCloseTo(40);
    // Without the option nothing changes.
    expect(lay(para([words]), 60).lines.map(lineText)).toEqual([
      'aaaa bbbb ',
      'cccc dddd ',
      'eeee ffff',
    ]);
    expect(layoutParagraph(para([words]), 60, env(), {}).lines.map(lineText)).toEqual([
      'aaaa bbbb ',
      'cccc dddd ',
      'eeee ffff',
    ]);
  });

  it('insets add to indents and the first-line indent', () => {
    const p = para([words], { indent: { left: 200, right: 0, firstLine: 200, hanging: 0 } });
    const pl = layoutParagraph(p, 80, env(), { insets: () => ({ left: 10, right: 5 }) });
    // Box: 80 - 10 (indent) - 10 - 5 = 55 pt; the first line is 10 pt narrower (45 pt).
    expect(texts(pl.lines[0]!)[0]!.x).toBeCloseTo(20);
    expect(texts(pl.lines[1]!)[0]!.x).toBeCloseTo(10);
    expect(pl.lines.map(lineText)).toEqual(['aaaa bbbb ', 'cccc dddd ', 'eeee ffff']);
    const narrow = layoutParagraph(p, 80, env(), { insets: () => ({ left: 15, right: 5 }) });
    expect(narrow.lines.map(lineText)).toEqual(['aaaa ', 'bbbb cccc ', 'dddd eeee ', 'ffff']);
  });

  it('centres, right-aligns and justifies within the narrowed box', () => {
    const ins = { insets: () => ({ left: 20, right: 10 }) };
    expect(
      texts(layoutParagraph(para(['aaaa'], { alignment: 'center' }), 100, env(), ins).lines[0]!)[0]!
        .x,
    ).toBeCloseTo(20 + 25);
    expect(
      texts(layoutParagraph(para(['aaaa'], { alignment: 'right' }), 100, env(), ins).lines[0]!)[0]!
        .x,
    ).toBeCloseTo(70);
    const just = layoutParagraph(para([words], { alignment: 'both' }), 90, env(), ins);
    // Box 60 pt: "aaaa bbbb" justified to end at 80.
    expect(xOf(just.lines[0]!, 'a')).toBeCloseTo(20);
    expect(xOf(just.lines[0]!, 'b')).toBeCloseTo(60);
    expect(just.lines[0]!.width).toBeCloseTo(80);
  });

  it('tab stops stay measured from the indent origin', () => {
    const pl = layoutParagraph(
      para(['a', TAB, 'b'], { tabs: [{ pos: 1000, val: 'left' }] }),
      200,
      env(),
      {
        insets: () => ({ left: 30, right: 0 }),
      },
    );
    expect(xOf(pl.lines[0]!, 'a')).toBeCloseTo(30);
    expect(xOf(pl.lines[0]!, 'b')).toBeCloseTo(50);
  });

  it('bidi paragraphs: the right inset is the start side', () => {
    const pl = layoutParagraph(para(['aaaa'], { bidi: true }), 100, env(), {
      insets: () => ({ left: 10, right: 30 }),
    });
    // Start-aligned RTL text ends at the right edge minus the right inset.
    expect(texts(pl.lines[0]!)[0]!.x).toBeCloseTo(100 - 30 - 20);
    const end = layoutParagraph(para(['aaaa'], { bidi: true, alignment: 'right' }), 100, env(), {
      insets: () => ({ left: 10, right: 30 }),
    });
    expect(texts(end.lines[0]!)[0]!.x).toBeCloseTo(10);
  });
});

// ---- every fixture paragraph -----------------------------------------------------------------

describe('layoutParagraph: the fixture corpus', () => {
  for (const f of fixtures()) {
    it(`${f.name}: lays out every paragraph within its column`, async () => {
      const doc = readDocx(f.bytes);
      const ir = buildIR(doc);
      const fonts = createFontManager(doc, { fonts: testFontResolver() });
      wantFonts(ir, fonts);
      await fonts.load();
      const e: LayoutEnv = {
        ...env(fonts),
        doc,
        defaultTabStop: ir.defaultTabStop,
        images: ir.images,
      };
      let lines = 0;
      for (const s of ir.sections) {
        const width = s.pageWidth - s.margins.left - s.margins.right;
        const stack: IRParagraph[] = s.blocks.filter(
          (b): b is IRParagraph => b.kind === 'paragraph',
        );
        while (stack.length > 0) {
          const p = stack.pop()!;
          const pl = layoutParagraph(p, width, e);
          expect(pl.lines.length).toBeGreaterThan(0);
          for (const l of pl.lines) {
            expect(l.height).toBeGreaterThan(0);
            expect(l.baseline).toBeLessThanOrEqual(l.height + 1e-9);
            // Content never exceeds the column (justified lines may shrink spaces, so allow none).
            expect(pl.indentLeft + l.width).toBeLessThanOrEqual(
              width + 0.01 + Math.max(0, -p.props.indent.right / 20),
            );
            for (const slot of l.fields) expect(l.items[slot.item]?.kind).toBe('text');
            lines++;
          }
          for (const a of p.anchors)
            for (const b of a.blocks ?? []) if (b.kind === 'paragraph') stack.push(b);
        }
      }
      expect(lines).toBeGreaterThan(0);
    });
  }
});
