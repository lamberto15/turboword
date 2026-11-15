// The 14 standard PDF fonts as metrics-only faces: the last-resort fallback (core ships no font
// files, plan decision D7).
//
// Agreed convention with the PDF renderer: a standard-14 face has no glyph ids. glyphId(cp)
// returns the character's byte code in the font's PDF encoding (WinAnsiEncoding for Helvetica,
// Times and Courier; the built-in encoding for Symbol and ZapfDingbats), 0 when the character
// cannot be encoded, and advance(code) takes that code. unitsPerEm is 1000, bytes are empty.
//
// Widths come from the Adobe Core14 AFMs via pdf.js (see standard14-data.ts for attribution).
// Ascender/descender/cap/x heights and bboxes are the AFM header values; winAscent/winDescent
// and lineGap are those of the fonts Word documents actually name (Arial, Times New Roman,
// Courier New, scaled to 1000 units), so line heights match Word even on the fallback.

import { UsageError } from '../../core/errors.js';
import type { Font, FontMetrics } from '../types.js';
import type { FontDetails } from './sfnt.js';
import { DINGBATS_UNICODE, SYMBOL_UNICODE, WIDTHS } from './standard14-data.js';

/** PostScript names of the 14 standard fonts. */
export const STANDARD14_NAMES: readonly string[] = Object.keys(WIDTHS);

interface Face {
  family: string;
  subfamily: string;
  bold: boolean;
  italic: boolean;
  italicAngle: number;
  bbox: [number, number, number, number];
  ascender: number;
  descender: number;
  capHeight: number;
  xHeight: number;
}

// AFM headers: FontBBox, ItalicAngle, Ascender, Descender, CapHeight, XHeight.
const FACES: Readonly<Record<string, Face>> = {
  Courier: face('Courier', '', 0, [-23, -250, 715, 805], 629, -157, 562, 426),
  'Courier-Bold': face('Courier', 'B', 0, [-113, -250, 749, 801], 629, -157, 562, 439),
  'Courier-Oblique': face('Courier', 'O', -12, [-27, -250, 849, 805], 629, -157, 562, 426),
  'Courier-BoldOblique': face('Courier', 'BO', -12, [-57, -250, 869, 801], 629, -157, 562, 439),
  Helvetica: face('Helvetica', '', 0, [-166, -225, 1000, 931], 718, -207, 718, 523),
  'Helvetica-Bold': face('Helvetica', 'B', 0, [-170, -228, 1003, 962], 718, -207, 718, 532),
  'Helvetica-Oblique': face('Helvetica', 'O', -12, [-170, -225, 1116, 931], 718, -207, 718, 523),
  'Helvetica-BoldOblique': face(
    'Helvetica',
    'BO',
    -12,
    [-174, -228, 1114, 962],
    718,
    -207,
    718,
    532,
  ),
  'Times-Roman': face('Times', '', 0, [-168, -218, 1000, 898], 683, -217, 662, 450),
  'Times-Bold': face('Times', 'B', 0, [-168, -218, 1000, 935], 683, -217, 676, 461),
  'Times-Italic': face('Times', 'I', -15.5, [-169, -217, 1010, 883], 683, -217, 653, 441),
  'Times-BoldItalic': face('Times', 'BI', -15, [-200, -218, 996, 921], 683, -217, 669, 462),
  Symbol: face('Symbol', '', 0, [-180, -293, 1090, 1010], 1010, -293, 700, 500),
  ZapfDingbats: face('ZapfDingbats', '', 0, [-1, -143, 981, 820], 820, -143, 700, 500),
};

/** Arial / Times New Roman / Courier New hhea+win metrics per 1000 units (2048-unit fonts). */
const WORD_VERTICAL: Readonly<Record<string, [number, number, number]>> = {
  Helvetica: [905, 212, 33],
  Times: [891, 216, 42],
  Courier: [833, 300, 0],
};

function face(
  family: string,
  style: '' | 'B' | 'O' | 'I' | 'BO' | 'BI',
  italicAngle: number,
  bbox: [number, number, number, number],
  ascender: number,
  descender: number,
  capHeight: number,
  xHeight: number,
): Face {
  const bold = style.startsWith('B');
  const slant = style.endsWith('O') ? 'Oblique' : style.endsWith('I') ? 'Italic' : '';
  const subfamily = [bold ? 'Bold' : '', slant].filter(Boolean).join(' ') || 'Regular';
  return {
    family,
    subfamily,
    bold,
    italic: slant !== '',
    italicAngle,
    bbox,
    ascender,
    descender,
    capHeight,
    xHeight,
  };
}

// ---- encodings ---------------------------------------------------------------------------------

/** cp1252 / WinAnsiEncoding codes 0x80-0x9F (0 = unused). */
const WIN_ANSI_HIGH = [
  0x20ac, 0, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0,
  0x17d, 0, 0, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a,
  0x153, 0, 0x17e, 0x178,
];

let winAnsi: Map<number, number> | undefined;
let symbol: Map<number, number> | undefined;
let dingbats: Map<number, number> | undefined;

/** WinAnsiEncoding byte code of a code point (0 = not encodable). */
export function winAnsiCode(cp: number): number {
  if ((cp >= 0x20 && cp <= 0x7e) || (cp >= 0xa0 && cp <= 0xff)) return cp;
  if (winAnsi === undefined) {
    winAnsi = new Map([
      [0x2010, 0x2d],
      [0x2011, 0x2d],
      [0x2212, 0x2d],
    ]);
    WIN_ANSI_HIGH.forEach((u, i) => {
      if (u !== 0) winAnsi!.set(u, 0x80 + i);
    });
  }
  return winAnsi.get(cp) ?? 0;
}

function pairsMap(pairs: readonly number[]): Map<number, number> {
  const m = new Map<number, number>();
  for (let i = 0; i + 1 < pairs.length; i += 2) m.set(pairs[i]!, pairs[i + 1]!);
  return m;
}

/** Code in a symbolic font's built-in encoding; Word's U+F020-U+F0FF symbol range included. */
function symbolicCode(map: Map<number, number>, widths: Uint16Array, cp: number): number {
  const direct = map.get(cp);
  if (direct !== undefined) return direct;
  if (cp >= 0xf020 && cp <= 0xf0ff) return widths[cp - 0xf000]! > 0 ? cp - 0xf000 : 0;
  return 0;
}

// ---- faces -------------------------------------------------------------------------------------

const cache = new Map<string, StandardFont>();

/** A standard-14 face (metrics only; see the file header for the code convention). */
export class StandardFont implements Font, FontDetails {
  readonly family: string;
  readonly subfamily: string;
  readonly postscriptName: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly metrics: FontMetrics;
  readonly outlines: 'standard14' = 'standard14';
  readonly bytes: Uint8Array = new Uint8Array(0);
  readonly numGlyphs: number = 256;
  readonly bbox: readonly [number, number, number, number];
  readonly weight: number;
  readonly fixedPitch: boolean;
  readonly symbolic: boolean;
  readonly fsType: number = 0;
  readonly typographicFamily: string;
  readonly fullName: string;
  readonly faceIndex: number = 0;
  readonly #widths: Uint16Array;
  readonly #encode: (cp: number) => number;

  /** @internal */
  constructor(name: string, f: Face) {
    this.postscriptName = name;
    this.family = f.family;
    this.typographicFamily = f.family;
    this.subfamily = f.subfamily;
    this.fullName = f.subfamily === 'Regular' ? f.family : `${f.family} ${f.subfamily}`;
    this.bold = f.bold;
    this.italic = f.italic;
    this.bbox = f.bbox;
    this.weight = f.bold ? 700 : 400;
    this.fixedPitch = f.family === 'Courier';
    this.symbolic = f.family === 'Symbol' || f.family === 'ZapfDingbats';
    const data = WIDTHS[name]!;
    const widths = new Uint16Array(256);
    for (let code = 32; code < 256; code++) {
      widths[code] = Number.parseInt(data.slice((code - 32) * 2, (code - 32) * 2 + 2), 36);
    }
    this.#widths = widths;
    if (f.family === 'Symbol') {
      symbol ??= pairsMap(SYMBOL_UNICODE);
      const m = symbol;
      this.#encode = (cp) => symbolicCode(m, widths, cp);
    } else if (f.family === 'ZapfDingbats') {
      dingbats ??= pairsMap(DINGBATS_UNICODE);
      const m = dingbats;
      this.#encode = (cp) => symbolicCode(m, widths, cp);
    } else {
      this.#encode = winAnsiCode;
    }
    const vertical = WORD_VERTICAL[f.family];
    this.metrics = {
      unitsPerEm: 1000,
      ascender: f.ascender,
      descender: f.descender,
      lineGap: vertical?.[2] ?? 0,
      winAscent: vertical?.[0] ?? f.bbox[3],
      winDescent: vertical?.[1] ?? -f.bbox[1],
      capHeight: f.capHeight,
      xHeight: f.xHeight,
      underlinePosition: -100,
      underlineThickness: 50,
      strikeoutPosition: Math.round(f.xHeight / 2),
      strikeoutSize: 50,
      italicAngle: f.italicAngle,
    };
  }

  glyphId(codePoint: number): number {
    return this.#encode(codePoint);
  }

  advance(code: number): number {
    return code >= 0 && code < 256 ? this.#widths[code]! : 0;
  }

  kerning(_left: number, _right: number): number {
    return 0;
  }
}

/**
 * A PDF base-14 face (metrics only) — the last-resort fallback. `name` is a PostScript name
 * (`Helvetica-Bold`, `Times-Roman`, `Symbol`, ..., case-insensitive); `Times`, `Arial`,
 * `Times New Roman` and `Courier New` (optionally with `,Bold`/`-Italic`... suffixes) are accepted
 * as aliases. Throws UsageError for anything else.
 */
export function standardFont(name: string): StandardFont {
  const ps = standardName(name);
  if (ps === undefined) throw new UsageError(`Not a standard PDF font: ${name}`);
  let f = cache.get(ps);
  if (f === undefined) {
    f = new StandardFont(ps, FACES[ps]!);
    cache.set(ps, f);
  }
  return f;
}

const LOWER: ReadonlyMap<string, string> = /* @__PURE__ */ new Map(
  Object.keys(FACES).map((k) => [k.toLowerCase(), k]),
);

const ALIAS_FAMILY: Readonly<Record<string, string>> = {
  helvetica: 'Helvetica',
  arial: 'Helvetica',
  times: 'Times',
  'times new roman': 'Times',
  timesnewroman: 'Times',
  courier: 'Courier',
  'courier new': 'Courier',
  couriernew: 'Courier',
  symbol: 'Symbol',
  zapfdingbats: 'ZapfDingbats',
};

function standardName(name: string): string | undefined {
  const lower = name.trim().toLowerCase();
  const exact = LOWER.get(lower);
  if (exact !== undefined) return exact;
  const m = /^(.*?)(?:[-,]\s*(bold|italic|oblique|bolditalic|boldoblique|roman|regular))?$/.exec(
    lower,
  );
  const family = m === null ? undefined : ALIAS_FAMILY[m[1]!.trim()];
  if (family === undefined) return undefined;
  const style = m?.[2] ?? '';
  return standard14For(family, style.includes('bold'), /italic|oblique/.test(style));
}

/** PostScript name of the standard-14 face of `family` (Helvetica/Times/Courier/Symbol/ZapfDingbats). */
export function standard14For(family: string, bold: boolean, italic: boolean): string {
  switch (family) {
    case 'Times':
      return bold
        ? italic
          ? 'Times-BoldItalic'
          : 'Times-Bold'
        : italic
          ? 'Times-Italic'
          : 'Times-Roman';
    case 'Courier':
    case 'Helvetica': {
      const style = (bold ? 'Bold' : '') + (italic ? 'Oblique' : '');
      return style === '' ? family : `${family}-${style}`;
    }
    case 'ZapfDingbats':
    case 'Symbol':
      return family;
    default:
      return standard14For('Helvetica', bold, italic);
  }
}
