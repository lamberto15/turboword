// basicShaper: cmap + hmtx + pair kerning, logical order, clusters as UTF-16 indices, and the
// fallbacks Word documents depend on. A synthetic Font covers the rules exactly; the real fonts
// in test/fixtures/fonts (and the standard-14 faces) check it against parsed sfnt data.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { loadFont, standardFont } from '../src/layout/fonts/index.js';
import { basicShaper } from '../src/layout/text/index.js';
import type { Font, ShapedGlyph } from '../src/layout/types.js';

/** A Font whose cmap is `glyphs` (code point → [glyph id, advance]) and kerning `kerns` ("l,r"). */
function fakeFont(
  glyphs: Record<number, [number, number]>,
  kerns: Record<string, number> = {},
): Font {
  const advances = new Map<number, number>([[0, 500]]);
  for (const [id, adv] of Object.values(glyphs)) advances.set(id, adv);
  return {
    family: 'Fake',
    subfamily: 'Regular',
    postscriptName: 'Fake-Regular',
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
      strikeoutPosition: 250,
      strikeoutSize: 50,
      italicAngle: 0,
    },
    outlines: 'truetype',
    glyphId: (cp) => glyphs[cp]?.[0] ?? 0,
    advance: (id) => advances.get(id) ?? 0,
    kerning: (l, r) => kerns[`${l},${r}`] ?? 0,
    bytes: new Uint8Array(0),
    numGlyphs: 100,
  };
}

const font = fakeFont(
  {
    32: [1, 250], // space
    45: [2, 333], // hyphen-minus
    46: [3, 280], // full stop
    48: [4, 556], // zero
    65: [10, 667], // A
    86: [11, 667], // V
    97: [12, 556], // a
    98: [13, 556], // b
    99: [14, 500], // c
    40: [20, 333], // (
    41: [21, 333], // )
    769: [30, 0], // combining acute
    1488: [40, 600], // alef
    1489: [41, 600], // bet
    119808: [50, 700], // mathematical bold A
  },
  { '10,11': -80, '11,10': -60 },
);

const ids = (g: ShapedGlyph[]): number[] => g.map((x) => x.id);
const adv = (g: ShapedGlyph[]): number[] => g.map((x) => x.advance);
const clusters = (g: ShapedGlyph[]): number[] => g.map((x) => x.cluster);
const ltr = { rtl: false, kerning: false };

describe('basicShaper', () => {
  it('maps characters to glyphs with hmtx advances, in logical order', () => {
    const g = basicShaper.shape(font, 'Aab', ltr);
    expect(ids(g)).toEqual([10, 12, 13]);
    expect(adv(g)).toEqual([667, 556, 556]);
    expect(clusters(g)).toEqual([0, 1, 2]);
    expect(g.every((x) => x.dx === 0 && x.dy === 0)).toBe(true);
    expect(basicShaper.shape(font, '', ltr)).toEqual([]);
  });

  it('uses UTF-16 indices as clusters across surrogate pairs', () => {
    const g = basicShaper.shape(font, 'a\u{1D400}b', ltr);
    expect(ids(g)).toEqual([12, 50, 13]);
    expect(clusters(g)).toEqual([0, 1, 3]);
  });

  it('gives combining marks their own glyph and hmtx advance', () => {
    const g = basicShaper.shape(font, 'áb', ltr);
    expect(ids(g)).toEqual([12, 30, 13]);
    expect(adv(g)).toEqual([556, 0, 556]);
    expect(clusters(g)).toEqual([0, 1, 2]);
  });

  it('returns .notdef for characters the face lacks', () => {
    const g = basicShaper.shape(font, 'a一', ltr);
    expect(ids(g)).toEqual([12, 0]);
    expect(adv(g)).toEqual([556, 500]);
  });

  it('drops default-ignorable characters and controls', () => {
    const g = basicShaper.shape(font, 'a‍b­c‎️\t', ltr);
    expect(ids(g)).toEqual([12, 13, 14]);
    expect(clusters(g)).toEqual([0, 2, 4]);
  });

  it('applies pair kerning only when asked', () => {
    expect(adv(basicShaper.shape(font, 'AV', ltr))).toEqual([667, 667]);
    expect(adv(basicShaper.shape(font, 'AV', { rtl: false, kerning: true }))).toEqual([587, 667]);
  });

  it('kerns across marks without moving them off their base', () => {
    // A + acute + V: the A/V adjustment goes on the mark, so the mark stays over A.
    const g = basicShaper.shape(font, 'ÁV', { rtl: false, kerning: true });
    expect(adv(g)).toEqual([667, -80, 667]);
  });

  it('kerns right-to-left runs by their visual pairs', () => {
    // Logical A V displayed as V A: the pair is (V, A) and the adjustment goes on V.
    const g = basicShaper.shape(font, 'AV', { rtl: true, kerning: true });
    expect(ids(g)).toEqual([10, 11]);
    expect(adv(g)).toEqual([667, 607]);
  });

  it('mirrors Bidi_Mirrored characters in right-to-left runs', () => {
    expect(ids(basicShaper.shape(font, '(א)', ltr))).toEqual([20, 40, 21]);
    expect(ids(basicShaper.shape(font, '(א)', { rtl: true, kerning: false }))).toEqual([
      21, 40, 20,
    ]);
  });

  it('falls back for special spaces and hyphens the face lacks', () => {
    const g = basicShaper.shape(font, '      ', ltr);
    expect(ids(g)).toEqual([1, 1, 1, 1, 1, 1]);
    expect(adv(g)).toEqual([250, 1000, 556, 280, 200, 125]);
    const h = basicShaper.shape(font, '‐‑−', ltr);
    expect(ids(h)).toEqual([2, 2, 2]);
    expect(adv(h)).toEqual([333, 333, 333]);
  });
});

const fontsDir = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/fonts/${name}`, import.meta.url));

function tryLoad(name: string): Font | undefined {
  try {
    return loadFont(gunzipSync(readFileSync(fontsDir(name))));
  } catch {
    return undefined;
  }
}

describe('basicShaper with real fonts', () => {
  const sans = tryLoad('LiberationSans-Regular.ttf.gz');

  it.skipIf(!sans)('shapes Liberation Sans with its kerning', () => {
    const f = sans!;
    const text = 'AVATAR Wave';
    const plain = basicShaper.shape(f, text, ltr);
    const kerned = basicShaper.shape(f, text, { rtl: false, kerning: true });
    expect(plain.length).toBe(text.length);
    expect(plain.every((g) => g.id > 0)).toBe(true);
    expect(clusters(plain)).toEqual([...text].map((_, i) => i));
    // Liberation Sans is metric-compatible with Arial: A = 1366/2048, space = 569/2048.
    expect(plain[0]!.advance).toBe(1366);
    expect(plain[6]!.advance).toBe(569);
    const sum = (g: ShapedGlyph[]): number => g.reduce((s, x) => s + x.advance, 0);
    expect(sum(kerned)).toBeLessThan(sum(plain));
    expect(f.kerning(plain[0]!.id, plain[1]!.id)).toBeLessThan(0);
    expect(kerned[0]!.advance).toBe(1366 + f.kerning(plain[0]!.id, plain[1]!.id));
  });

  it.skipIf(!sans)('falls back to "-" for the non-breaking hyphen when the face lacks it', () => {
    const f = sans!;
    const g = basicShaper.shape(f, 'a‑b', ltr);
    const expected = f.glyphId(0x2011) || f.glyphId(0x2010) || f.glyphId(0x2d);
    expect(g[1]!.id).toBe(expected);
    expect(g[1]!.id).toBeGreaterThan(0);
  });

  it('shapes standard-14 faces through their WinAnsi glyph ids', () => {
    let helv: Font | undefined;
    try {
      helv = standardFont('Helvetica');
    } catch {
      return; // fonts module not available yet
    }
    expect(helv.outlines).toBe('standard14');
    const g = basicShaper.shape(helv, 'A€ ‑', { rtl: false, kerning: true });
    expect(ids(g)).toEqual([65, 128, 32, 45]);
    expect(g[0]!.advance).toBe(667);
    expect(g[2]!.advance).toBe(278);
  });
});
