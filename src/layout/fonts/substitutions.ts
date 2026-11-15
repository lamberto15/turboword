// Font substitution: metric-compatible replacements, generic classes, standard-14 last resort.
//
// When a family named by the document is not available (not embedded, no resolver has it), the
// font manager asks the resolvers for a metric-compatible family instead (same advance widths,
// so line breaks and page counts match Word): Carlito for Calibri, Liberation/Croscore for the
// Microsoft core fonts, and so on. Reverse pairs are listed too, for documents written with the
// open families on machines that only have the Microsoft ones. Failing that, the generic class of
// the family (from the font table's w:family/w:pitch, else from its name) gives a list of common
// families, and finally a standard-14 PDF face.

import type { FontTableEntry } from './embedded.js';

/** Generic family classes. */
export type GenericFamily = 'serif' | 'sans-serif' | 'monospace' | 'symbol';

const SERIF_GENERIC = [
  'Liberation Serif',
  'Tinos',
  'Times New Roman',
  'DejaVu Serif',
  'Noto Serif',
];
const SANS_GENERIC = ['Liberation Sans', 'Arimo', 'Arial', 'DejaVu Sans', 'Noto Sans'];
const MONO_GENERIC = [
  'Liberation Mono',
  'Cousine',
  'Courier New',
  'DejaVu Sans Mono',
  'Noto Sans Mono',
];

/**
 * Metric-compatible substitutions (family → replacement families, best first), plus the generic
 * fallbacks under the keys `serif`, `sans-serif` and `monospace`. Keys match case-insensitively.
 */
export const FONT_SUBSTITUTIONS: Readonly<Record<string, readonly string[]>> = {
  Calibri: ['Carlito'],
  'Calibri Light': ['Carlito'],
  Cambria: ['Caladea'],
  'Cambria Math': ['Caladea'],
  Arial: ['Liberation Sans', 'Arimo'],
  Helvetica: ['Liberation Sans', 'Arimo', 'Arial'],
  'Times New Roman': ['Liberation Serif', 'Tinos'],
  Times: ['Liberation Serif', 'Tinos', 'Times New Roman'],
  'Courier New': ['Liberation Mono', 'Cousine'],
  Courier: ['Liberation Mono', 'Cousine', 'Courier New'],
  Georgia: ['Gelasio'],
  'Arial Narrow': ['Liberation Sans Narrow'],
  'Segoe UI': ['Selawik'],
  Symbol: ['OpenSymbol'],
  // Reverse pairs.
  Carlito: ['Calibri'],
  Caladea: ['Cambria'],
  'Liberation Sans': ['Arimo', 'Arial'],
  Arimo: ['Liberation Sans', 'Arial'],
  'Liberation Serif': ['Tinos', 'Times New Roman'],
  Tinos: ['Liberation Serif', 'Times New Roman'],
  'Liberation Mono': ['Cousine', 'Courier New'],
  Cousine: ['Liberation Mono', 'Courier New'],
  'Liberation Sans Narrow': ['Arial Narrow'],
  Gelasio: ['Georgia'],
  Selawik: ['Segoe UI'],
  OpenSymbol: ['Symbol'],
  // Generic fallbacks.
  serif: SERIF_GENERIC,
  'sans-serif': SANS_GENERIC,
  monospace: MONO_GENERIC,
};

/** Vertical metrics of a family, in its own font units (see SUBSTITUTE_METRICS). */
export interface VerticalMetrics {
  unitsPerEm: number;
  /** The ascender/descender/lineGap Word uses (hhea, or OS/2 typo with USE_TYPO_METRICS). */
  ascender: number;
  descender: number;
  lineGap: number;
  winAscent: number;
  winDescent: number;
  underlinePosition?: number;
  underlineThickness?: number;
}

const vm = (
  ascender: number,
  descender: number,
  lineGap: number,
  winAscent: number,
  winDescent: number,
  underlinePosition: number,
  underlineThickness?: number,
): VerticalMetrics =>
  underlineThickness === undefined
    ? { unitsPerEm: 2048, ascender, descender, lineGap, winAscent, winDescent, underlinePosition }
    : {
        unitsPerEm: 2048,
        ascender,
        descender,
        lineGap,
        winAscent,
        winDescent,
        underlinePosition,
        underlineThickness,
      };

/**
 * Vertical metrics of the Microsoft families, measured from their regular faces (Office 365 for
 * Mac / macOS copies; Times New Roman corrected to the Windows metrics). A face standing in for one of these families (a substitute, a generic
 * fallback or a standard-14 face) reports these metrics, scaled to its own unitsPerEm, so line
 * heights match Word even when the substitute's own vertical metrics differ (Caladea's win
 * metrics are 1.30 em against Cambria's 1.17 em; Gelasio's are 1.66 em against Georgia's 1.14).
 * Advances always stay the substitute's. Faces whose line metrics already match (Carlito,
 * Arial → Liberation Sans, Courier New → Liberation Mono) are used as they are. Underline
 * thickness is only listed where it is the same in the regular and bold faces.
 */
export const SUBSTITUTE_METRICS: Readonly<Record<string, VerticalMetrics>> = {
  Calibri: vm(1950, -550, 0, 1950, 550, -232, 134),
  'Calibri Light': vm(1950, -550, 0, 1950, 550, -261, 104),
  Cambria: vm(1946, -455, 0, 1946, 455, -125),
  'Cambria Math': vm(1593, -455, 353, 6383, 5045, -125), // USE_TYPO_METRICS; huge win metrics
  Arial: vm(1854, -434, 67, 1854, 434, -217),
  'Arial Narrow': vm(1916, -434, 0, 1888, 431, -217),
  // The Office-for-Mac copy has lineGap 0; Windows Word lays TNR 12 pt out at 13.8 pt, i.e.
  // (1825 + 443 + 87) / 2048 em, the metrics Liberation Serif copies.
  'Times New Roman': vm(1825, -443, 87, 1825, 443, -223),
  'Courier New': vm(1705, -615, 0, 1705, 615, -477),
  Georgia: vm(1878, -449, 0, 1878, 449, -181),
  Symbol: vm(2059, -450, 0, 2059, 450, -223),
};

const LOWER_METRICS: ReadonlyMap<string, VerticalMetrics> = /* @__PURE__ */ new Map(
  Object.entries(SUBSTITUTE_METRICS).map(([k, v]) => [k.toLowerCase(), v]),
);

/** The original vertical metrics of `family` (case-insensitive) when a stand-in must report them. */
export function substituteMetricsOf(family: string): VerticalMetrics | undefined {
  return LOWER_METRICS.get(family.trim().toLowerCase());
}

/** Last-resort standard-14 family per generic class (Helvetica/Times/Courier/Symbol). */
export const STANDARD14_FALLBACKS: Readonly<Record<GenericFamily, string>> = {
  serif: 'Times',
  'sans-serif': 'Helvetica',
  monospace: 'Courier',
  symbol: 'Symbol',
};

const LOWER_SUBSTITUTIONS: ReadonlyMap<string, readonly string[]> = /* @__PURE__ */ new Map(
  Object.entries(FONT_SUBSTITUTIONS).map(([k, v]) => [k.toLowerCase(), v]),
);

/** Built-in substitutes of `family` (case-insensitive), [] when none. */
export function substitutesOf(family: string): readonly string[] {
  return LOWER_SUBSTITUTIONS.get(family.trim().toLowerCase()) ?? [];
}

const SYMBOL_FAMILIES = new Set([
  'symbol',
  'wingdings',
  'wingdings 2',
  'wingdings 3',
  'webdings',
  'zapfdingbats',
  'zapf dingbats',
  'itc zapf dingbats',
  'opensymbol',
  'marlett',
  'ms outlook',
  'bookshelf symbol 7',
  'ms reference specialty',
  'mt extra',
]);

const MONO_NAMES =
  /mono|courier|consol|typewriter|lucida console|lucida sans typewriter|fixedsys|terminal|\bcode\b|inconsolata|menlo|monaco|ocr[ -]?a|ocr[ -]?b/;
const SERIF_NAMES =
  /serif|times|roman|cambria|caladea|georgia|gelasio|garamond|palatino|book antiqua|bookman|century|baskerville|bodoni|didot|caslon|constantia|minion|cochin|charter|perpetua|goudy|rockwell|sitka|tinos|liberation serif|dejavu serif|noto serif|droid serif|mincho|song|simsun|batang|ming|pmingliu|mingliu|nsimsun|fangsong|kaiti|cambo|lora|merriweather|playfair|crimson|libre baskerville|pt serif|source serif|bell mt|elephant|footlight|high tower|modern no|niagara|poor richard|californian|centaur|bernard|calisto|imprint|plantagenet/;

/** True for families whose characters are pictographs at Latin code points (Symbol, Wingdings). */
export function isSymbolFamily(family: string): boolean {
  return SYMBOL_FAMILIES.has(family.trim().toLowerCase());
}

/**
 * Generic class of a family: from the font table's w:family / w:pitch / w:charset when known,
 * else from the name.
 */
export function genericFamily(family: string, entry?: FontTableEntry): GenericFamily {
  const lower = family.trim().toLowerCase();
  if (lower === 'serif' || lower === 'sans-serif' || lower === 'monospace') return lower;
  if (isSymbolFamily(lower) || entry?.charset?.toLowerCase() === '02') return 'symbol';
  if (entry !== undefined) {
    if (entry.pitch === 'fixed' || entry.family === 'modern') return 'monospace';
    if (entry.family === 'roman') return 'serif';
    if (entry.family === 'swiss') return 'sans-serif';
  }
  if (lower.includes('sans')) return MONO_NAMES.test(lower) ? 'monospace' : 'sans-serif';
  if (MONO_NAMES.test(lower)) return 'monospace';
  if (SERIF_NAMES.test(lower)) return 'serif';
  return 'sans-serif';
}
