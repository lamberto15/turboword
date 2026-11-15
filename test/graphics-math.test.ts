// OMML layout (src/layout/graphics/math.ts) with fake fixed-advance faces: placement of
// fractions, scripts, n-ary operators, radicals, delimiters, matrices and display paragraphs,
// plus robustness (deep nesting) and one run with a real face when the font loader is available.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { createDocx } from '../src/index.js';
import type { FontManager, MathBox } from '../src/layout/engine.js';
import { layoutMath, mathFontRequests, mathRunProps } from '../src/layout/graphics/math.js';
import { runProps } from '../src/layout/graphics/text.js';
import type { DisplayItem, Font, TextItem } from '../src/layout/types.js';
import { parseXml } from '../src/xml/parse.js';
import { fakeEnv, fakeFontManager, fakeShaper } from './graphics-env.js';

const NSDECL =
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const r = (t: string, extra = ''): string => `<m:r>${extra}<m:t>${t}</m:t></m:r>`;
const e = (inner: string): string => `<m:e>${inner}</m:e>`;
const frac = (n: string, d: string, pr = ''): string =>
  `<m:f>${pr}<m:num>${n}</m:num><m:den>${d}</m:den></m:f>`;

const SIZE = 10;

function lay(body: string, para = false, fonts = fakeFontManager()) {
  const xml = para
    ? `<m:oMathPara ${NSDECL}>${body}</m:oMathPara>`
    : `<m:oMath ${NSDECL}>${body}</m:oMath>`;
  const el = parseXml(xml).root;
  const env = fakeEnv(createDocx(), fonts);
  const box = layoutMath(el, runProps({ font: 'Calibri', size: SIZE }), env);
  return { box, fonts, env };
}

const texts = (b: MathBox): TextItem[] => b.items.filter((i): i is TextItem => i.kind === 'text');
const textOf = (b: MathBox, s: string): TextItem => {
  const t = texts(b).find((i) => i.text === s);
  if (t === undefined) throw new Error(`no text item ${s}: ${texts(b).map((i) => i.text)}`);
  return t;
};
const width = (t: TextItem): number => t.advances.reduce((a, b) => a + b, 0);

function numbersOf(it: DisplayItem): number[] {
  const out: number[] = [];
  for (const v of Object.values(it)) {
    if (typeof v === 'number') out.push(v);
    else if (Array.isArray(v)) for (const x of v) if (typeof x === 'number') out.push(x);
  }
  if (it.kind === 'path') for (const m of it.d.match(/-?[\d.]+(?:e-?\d+)?/g) ?? []) out.push(+m);
  return out;
}

function expectFinite(b: MathBox): void {
  for (const n of [b.width, b.ascent, b.descent]) expect(Number.isFinite(n)).toBe(true);
  for (const it of b.items)
    for (const n of numbersOf(it)) expect(Number.isFinite(n), it.kind).toBe(true);
}

describe('math runs', () => {
  it('sets single letters in italic and digits upright, in Cambria Math', () => {
    const { box, fonts } = lay(r('x') + r('2'));
    expect(fonts.requested).toContain('Cambria Math|false|true');
    expect(fonts.requested).toContain('Cambria Math|false|false');
    const x = textOf(box, 'x');
    const two = textOf(box, '2');
    const faces = fonts.fonts();
    expect(faces[x.font]?.italic).toBe(true);
    expect(faces[two.font]?.italic).toBe(false);
    expect(two.x).toBeCloseTo(width(x));
    expect(box.width).toBeCloseTo(width(x) + width(two));
    expect(box.ascent).toBeCloseTo(8);
    expect(box.descent).toBeCloseTo(2);
  });

  it('keeps function names upright and honours m:sty', () => {
    const { box, fonts } = lay(
      `<m:func><m:fName>${r('sin')}</m:fName>${e(r('x'))}</m:func>` +
        r('y', '<m:rPr><m:sty m:val="p"/></m:rPr>') +
        r('b', '<m:rPr><m:sty m:val="bi"/></m:rPr>'),
    );
    const faces = fonts.fonts();
    expect(faces[textOf(box, 'sin').font]?.italic).toBe(false);
    expect(faces[textOf(box, 'x').font]?.italic).toBe(true);
    expect(faces[textOf(box, 'y').font]?.italic).toBe(false);
    const b = faces[textOf(box, 'b').font];
    expect(b?.italic && b.bold).toBe(true);
  });

  it('spaces binary operators and relations but not a leading sign', () => {
    const { box } = lay(r('−a+b=c'));
    const a = textOf(box, 'a');
    const plus = textOf(box, '+');
    expect(a.x).toBeCloseTo(5); // leading minus: no space
    expect(plus.x).toBeGreaterThan(a.x + width(a) + 2);
    expect(textOf(box, '=').x).toBeGreaterThan(textOf(box, 'b').x + 5 + 2.5);
  });

  it('maps m:nor runs to the normal text font and w:rPr size/colour', () => {
    const { box, fonts } = lay(
      r('if', '<m:rPr><m:nor/></m:rPr>') +
        r('x', '<w:rPr><w:sz w:val="40"/><w:color w:val="FF0000"/></w:rPr>'),
    );
    expect(fonts.fonts()[textOf(box, 'if').font]?.family).toBe('Calibri');
    expect(textOf(box, 'x').size).toBe(20);
    expect(textOf(box, 'x').color).toBe('FF0000');
  });

  it('lists the font requests it may make', () => {
    const reqs = mathFontRequests(runProps({ size: 12 }));
    expect(reqs).toHaveLength(6);
    expect(reqs.every((p) => p.fonts.ascii === 'Cambria Math')).toBe(true);
    expect(reqs.map((p) => Math.round(p.size * 10) / 10)).toEqual([12, 12, 8.4, 8.4, 6, 6]);
    expect(mathRunProps(runProps({ size: 9 }), true).italic).toBe(true);
  });
});

describe('fractions and scripts', () => {
  it('stacks numerator above and denominator below a bar', () => {
    const { box } = lay(frac(r('a'), r('bc')));
    const num = textOf(box, 'a');
    const den = textOf(box, 'bc');
    expect(num.y).toBeLessThan(0);
    expect(den.y).toBeGreaterThan(0);
    // Inline fractions use script size (70%).
    expect(num.size).toBeCloseTo(7);
    const bar = box.items.find((i) => i.kind === 'rect');
    expect(bar).toBeDefined();
    if (bar?.kind === 'rect') {
      expect(bar.y + bar.height / 2).toBeCloseTo(-2.5); // math axis
      expect(bar.y).toBeGreaterThan(num.y);
      expect(bar.y + bar.height).toBeLessThan(den.y);
    }
    // width = max(num, den) + 2 × 0.1 em padding; both centred.
    expect(box.width).toBeCloseTo(width(den) + 2);
    expect(num.x + width(num) / 2).toBeCloseTo(box.width / 2);
    expect(den.x + width(den) / 2).toBeCloseTo(box.width / 2);
    expect(box.ascent).toBeGreaterThan(-num.y);
    expect(box.descent).toBeGreaterThan(den.y);
  });

  it('draws no bar for noBar and inline text for lin', () => {
    const noBar = lay(frac(r('a'), r('b'), '<m:fPr><m:type m:val="noBar"/></m:fPr>')).box;
    expect(noBar.items.some((i) => i.kind === 'rect')).toBe(false);
    const lin = lay(frac(r('a'), r('b'), '<m:fPr><m:type m:val="lin"/></m:fPr>')).box;
    expect(texts(lin).map((t) => t.text)).toEqual(['a', '/', 'b']);
    expect(texts(lin).every((t) => t.y === 0)).toBe(true);
  });

  it('keeps display fractions at full size and shrinks nested ones', () => {
    const { box } = lay(`<m:oMath>${frac(frac(r('a'), r('b')), r('c'))}</m:oMath>`, true);
    expect(textOf(box, 'c').size).toBe(10);
    expect(textOf(box, 'a').size).toBeCloseTo(7);
  });

  it('raises superscripts and lowers subscripts at script size', () => {
    const sup = lay(`<m:sSup>${e(r('x'))}<m:sup>${r('2')}</m:sup></m:sSup>`).box;
    const two = textOf(sup, '2');
    expect(two.y).toBeLessThan(-3);
    expect(two.size).toBeCloseTo(7);
    expect(two.x).toBeGreaterThanOrEqual(width(textOf(sup, 'x')));
    const sub = lay(`<m:sSub>${e(r('x'))}<m:sub>${r('i')}</m:sub></m:sSub>`).box;
    expect(textOf(sub, 'i').y).toBeGreaterThan(1);
    expect(textOf(sub, 'i').size).toBeCloseTo(7);
  });

  it('separates sub- and superscript of sSubSup and nests scripts down to 50%', () => {
    const { box } = lay(
      `<m:sSubSup>${e(r('x'))}<m:sub>${r('i')}</m:sub><m:sup>` +
        `<m:sSup>${e(r('n'))}<m:sup>${r('k')}</m:sup></m:sSup></m:sup></m:sSubSup>`,
    );
    const i = textOf(box, 'i');
    const n = textOf(box, 'n');
    const k = textOf(box, 'k');
    expect(n.y).toBeLessThan(0);
    expect(i.y).toBeGreaterThan(0);
    expect(i.x).toBeCloseTo(n.x);
    expect(k.size).toBeCloseTo(5);
    expect(k.y).toBeLessThan(n.y);
    // Visible gap between the superscript's bottom and the subscript's top.
    expect(i.y - 0.5 * 7 - n.y).toBeGreaterThan(0);
  });

  it('puts pre-scripts before the base', () => {
    const { box } = lay(
      `<m:sPre><m:sub>${r('1')}</m:sub><m:sup>${r('2')}</m:sup>${e(r('X'))}</m:sPre>`,
    );
    expect(textOf(box, 'X').x).toBeGreaterThan(textOf(box, '1').x);
    expect(textOf(box, '2').y).toBeLessThan(0);
    expect(textOf(box, '1').y).toBeGreaterThan(0);
  });
});

describe('n-ary, radicals and delimiters', () => {
  it('places sum limits above and below in display style', () => {
    const { box } = lay(
      `<m:oMath><m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${r('i=1')}</m:sub>` +
        `<m:sup>${r('n')}</m:sup>${e(r('i'))}</m:nary></m:oMath>`,
      true,
    );
    const op = textOf(box, '∑');
    expect(op.size).toBeCloseTo(16);
    const n = textOf(box, 'n');
    const lo = texts(box).filter((t) => t.size < 8 && t.text !== 'n');
    expect(n.y).toBeLessThan(op.y - 8);
    for (const t of lo) expect(t.y).toBeGreaterThan(op.y);
    // limits centred on the operator column
    const colCenter = op.x + width(op) / 2;
    expect(n.x + width(n) / 2).toBeCloseTo(colCenter);
    // the body comes after the operator column
    const body = texts(box).find((t) => t.text === 'i' && t.size === 10);
    expect(body?.x).toBeGreaterThan(op.x + width(op));
    expectFinite(box);
  });

  it('places integral limits as sub/superscripts by default', () => {
    const { box } = lay(
      `<m:nary><m:sub>${r('0')}</m:sub><m:sup>${r('1')}</m:sup>${e(r('f'))}</m:nary>`,
    );
    const op = textOf(box, '∫');
    expect(op.size).toBeCloseTo(12);
    const top = textOf(box, '1');
    const bot = textOf(box, '0');
    expect(top.x).toBeGreaterThan(op.x + width(op) * 0.9);
    expect(bot.x).toBeGreaterThan(op.x);
    expect(top.y).toBeLessThan(bot.y);
    expect(textOf(box, 'f').x).toBeGreaterThan(top.x + width(top));
  });

  it('hides limits with subHide/supHide', () => {
    const { box } = lay(
      `<m:nary><m:naryPr><m:chr m:val="∏"/><m:subHide m:val="1"/><m:supHide/></m:naryPr>` +
        `<m:sub>${r('a')}</m:sub><m:sup>${r('b')}</m:sup>${e(r('c'))}</m:nary>`,
    );
    expect(texts(box).map((t) => t.text)).toEqual(['∏', 'c']);
  });

  it('draws a radical over the radicand, with an optional degree', () => {
    const plain = lay(
      `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${e(r('x'))}</m:rad>`,
    ).box;
    const sign = plain.items.find((i) => i.kind === 'path');
    expect(sign?.kind === 'path' && sign.stroke !== undefined).toBe(true);
    const x = textOf(plain, 'x');
    expect(x.x).toBeGreaterThan(4);
    expect(plain.width).toBeGreaterThan(x.x + width(x));
    expect(plain.ascent).toBeGreaterThan(7);
    expect(texts(plain)).toHaveLength(1);

    const cube = lay(`<m:rad><m:deg>${r('3')}</m:deg>${e(r('x'))}</m:rad>`).box;
    const three = textOf(cube, '3');
    expect(three.size).toBeCloseTo(5);
    expect(three.y).toBeLessThan(0);
    expect(three.x).toBeLessThan(textOf(cube, 'x').x);
  });

  it('stretches delimiters around tall content and keeps glyphs for short content', () => {
    const short = lay(`<m:d>${e(r('x'))}</m:d>`).box;
    expect(texts(short).map((t) => t.text)).toEqual(['(', 'x', ')']);
    const tall = lay(
      `<m:oMath><m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val=""/></m:dPr>` +
        `${e(frac(r('a'), r('b')))}</m:d></m:oMath>`,
      true,
    ).box;
    expect(texts(tall).map((t) => t.text)).toEqual(['a', 'b']);
    const paths = tall.items.filter((i) => i.kind === 'path');
    expect(paths).toHaveLength(1);
    // the bracket spans the fraction
    const ys = (paths[0]?.kind === 'path' ? paths[0].d : '')
      .split(/[MLCZ ]+/)
      .filter((s) => s !== '')
      .map(Number)
      .filter((_, i) => i % 2 === 1);
    expect(Math.min(...ys)).toBeLessThan(textOf(tall, 'a').y - 5);
    expect(Math.max(...ys)).toBeGreaterThan(textOf(tall, 'b').y);
    expect(textOf(tall, 'a').x).toBeGreaterThan(0);
  });

  it('separates several arguments with the separator', () => {
    const { box } = lay(
      `<m:d><m:dPr><m:begChr m:val="{"/><m:sepChr m:val=","/><m:endChr m:val="}"/></m:dPr>` +
        `${e(r('a'))}${e(r('b'))}</m:d>`,
    );
    expect(texts(box).map((t) => t.text)).toEqual(['{', 'a', ',', 'b', '}']);
  });
});

describe('matrices, arrays and other structures', () => {
  it('lays out a 2×2 matrix in centred columns around the axis', () => {
    const { box } = lay(
      `<m:m><m:mr>${e(r('1'))}${e(r('22'))}</m:mr><m:mr>${e(r('333'))}${e(r('4'))}</m:mr></m:m>`,
    );
    const [a, b, c, d] = ['1', '22', '333', '4'].map((s) => textOf(box, s)) as TextItem[];
    if (!a || !b || !c || !d) throw new Error('missing cells');
    expect(a.y).toBeCloseTo(b.y);
    expect(c.y).toBeCloseTo(d.y);
    expect(c.y).toBeGreaterThan(a.y);
    expect(a.x + width(a) / 2).toBeCloseTo(c.x + width(c) / 2);
    expect(b.x).toBeGreaterThan(c.x + width(c));
    // centred on the math axis (y = -2.5): the box extends equally above and below it
    expect((box.ascent - box.descent) / 2).toBeCloseTo(2.5);
    expect(c.y - a.y).toBeCloseTo(2 + 2 + 8); // descent + row gap + ascent
    expect(box.width).toBeCloseTo(width(c) + 10 + width(b));
  });

  it('aligns equation array rows at m:aln and stacks them', () => {
    const aln = '<m:rPr><m:aln/></m:rPr>';
    const { box } = lay(
      `<m:eqArr>${e(r('x') + r('=1', aln))}${e(r('yyy') + r('=2', aln))}</m:eqArr>`,
    );
    const eqs = texts(box).filter((t) => t.text === '=');
    expect(eqs).toHaveLength(2);
    expect(eqs[0]?.x).toBeCloseTo(eqs[1]?.x ?? 0);
    expect(eqs[1]?.y).toBeGreaterThan(eqs[0]?.y ?? 0);
  });

  it('draws accents, bars, boxes and group characters', () => {
    const { box } = lay(
      `<m:acc>${e(r('a'))}</m:acc>` +
        `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${e(r('b'))}</m:bar>` +
        `<m:borderBox>${e(r('c'))}</m:borderBox>` +
        `<m:groupChr>${e(r('d+e'))}</m:groupChr>` +
        `<m:limLow>${e(r('lim', '<m:rPr><m:sty m:val="p"/></m:rPr>'))}<m:lim>${r('n→∞')}</m:lim></m:limLow>`,
    );
    const a = textOf(box, 'a');
    const hat = box.items.find((i) => i.kind === 'path');
    expect(hat).toBeDefined();
    expect(box.items.filter((i) => i.kind === 'line')).toHaveLength(4);
    expect(box.items.some((i) => i.kind === 'rect' && i.y < -4)).toBe(true);
    const lim = textOf(box, 'lim');
    const n = texts(box).find((t) => t.text.startsWith('n'));
    expect(n?.y).toBeGreaterThan(lim.y);
    expect(a.x).toBe(0);
    expectFinite(box);
  });

  it('centres each line of an m:oMathPara', () => {
    const { box } = lay(`<m:oMath>${r('x=1')}</m:oMath><m:oMath>${r('y')}</m:oMath>`, true);
    const y = textOf(box, 'y');
    const x = textOf(box, 'x');
    expect(x.x).toBe(0);
    expect(y.x).toBeCloseTo((box.width - width(y)) / 2);
    expect(y.y).toBeGreaterThan(x.y);
    expect(box.descent).toBeGreaterThan(y.y);
  });

  it('honours left justification of an m:oMathPara', () => {
    const { box } = lay(
      `<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>` +
        `<m:oMath>${r('x=1')}</m:oMath><m:oMath>${r('y')}</m:oMath>`,
      true,
    );
    expect(textOf(box, 'y').x).toBe(0);
  });

  it('survives 200 nested fractions with finite output', () => {
    let s = r('1');
    for (let i = 0; i < 200; i++) s = frac(s, r(String(i % 10)));
    const { box } = lay(s);
    expectFinite(box);
    expect(box.width).toBeGreaterThan(0);
    expect(texts(box).length).toBeGreaterThan(40);
  });

  it('lays out tracked insertions and skips deletions', () => {
    const { box } = lay(`<w:ins>${r('a')}</w:ins><w:del>${r('b')}</w:del>${r('c')}`);
    expect(texts(box).map((t) => t.text)).toEqual(['a', 'c']);
  });
});

// ---- a real face ------------------------------------------------------------------------------

const FONT_PATH = fileURLToPath(
  new URL('./fixtures/fonts/Caladea-Regular.ttf.gz', import.meta.url),
);

async function realFont(): Promise<Font | undefined> {
  try {
    const { loadFont } = await import('../src/layout/fonts/index.js');
    return loadFont(gunzipSync(new Uint8Array(readFileSync(FONT_PATH))));
  } catch {
    return undefined;
  }
}

const caladea = await realFont();

describe('with a real face', () => {
  it.skipIf(caladea === undefined)('lays out a fraction with real metrics', () => {
    const font = caladea as Font;
    const fonts: FontManager = {
      want() {},
      async load() {},
      fontFor: () => font,
      fonts: () => [font],
      indexOf: () => 0,
    };
    const { box } = lay(
      `<m:sSup>${e(frac(r('a+b'), r('c')))}<m:sup>${r('2')}</m:sup></m:sSup>`,
      false,
      fonts as ReturnType<typeof fakeFontManager>,
    );
    expectFinite(box);
    expect(textOf(box, 'c').glyphs[0]).toBe(font.glyphId(99));
    expect(textOf(box, '2').y).toBeLessThan(textOf(box, 'a').y + 1);
    expect(box.width).toBeGreaterThan(5);
    void fakeShaper;
  });
});
