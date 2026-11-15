// The sfnt `cmap` table: character code → glyph id.
//
// A font usually carries several subtables. They are ranked (full-Unicode format 12/13 first,
// then the BMP format 4, then legacy symbol and Mac Roman tables) and kept as a chain, so a code
// point missing from the best subtable can still be found in a lesser one. Formats 0, 4, 6, 12 and
// 13 are decoded; 14 (variation sequences) and the rare 2/8/10 are skipped. Segment arrays are
// copied into typed arrays once, so lookups are a binary search with no further bounds checks
// except the glyphIdArray reads of format 4, which can point anywhere and are checked per read.
//
// Symbol fonts (Windows platform, encoding 0) map their characters at U+F020-U+F0FF; like Word,
// a lookup of U+0020-U+00FF in such a font also tries U+F000 + code.

import { fontError, type Reader } from './reader.js';

/** code point → glyph id (0 = missing). */
export type CmapLookup = (codePoint: number) => number;

/** Unicode code points of Mac OS Roman codes 0x80-0xFF. */
export const MAC_ROMAN_HIGH: readonly number[] = [
  0xc4, 0xc5, 0xc7, 0xc9, 0xd1, 0xd6, 0xdc, 0xe1, 0xe0, 0xe2, 0xe4, 0xe3, 0xe5, 0xe7, 0xe9, 0xe8,
  0xea, 0xeb, 0xed, 0xec, 0xee, 0xef, 0xf1, 0xf3, 0xf2, 0xf4, 0xf6, 0xf5, 0xfa, 0xf9, 0xfb, 0xfc,
  0x2020, 0xb0, 0xa2, 0xa3, 0xa7, 0x2022, 0xb6, 0xdf, 0xae, 0xa9, 0x2122, 0xb4, 0xa8, 0x2260, 0xc6,
  0xd8, 0x221e, 0xb1, 0x2264, 0x2265, 0xa5, 0xb5, 0x2202, 0x2211, 0x220f, 0x3c0, 0x222b, 0xaa, 0xba,
  0x3a9, 0xe6, 0xf8, 0xbf, 0xa1, 0xac, 0x221a, 0x192, 0x2248, 0x2206, 0xab, 0xbb, 0x2026, 0xa0,
  0xc0, 0xc3, 0xd5, 0x152, 0x153, 0x2013, 0x2014, 0x201c, 0x201d, 0x2018, 0x2019, 0xf7, 0x25ca,
  0xff, 0x178, 0x2044, 0x20ac, 0x2039, 0x203a, 0xfb01, 0xfb02, 0x2021, 0xb7, 0x201a, 0x201e, 0x2030,
  0xc2, 0xca, 0xc1, 0xcb, 0xc8, 0xcd, 0xce, 0xcf, 0xcc, 0xd3, 0xd4, 0xf8ff, 0xd2, 0xda, 0xdb, 0xd9,
  0x131, 0x2c6, 0x2dc, 0xaf, 0x2d8, 0x2d9, 0x2da, 0xb8, 0x2dd, 0x2db, 0x2c7,
];

let macRomanReverse: Map<number, number> | undefined;

/** Mac OS Roman code of a code point, or -1. */
export function macRomanCode(cp: number): number {
  if (cp < 0x80) return cp;
  if (macRomanReverse === undefined) {
    macRomanReverse = new Map();
    for (let i = 0; i < MAC_ROMAN_HIGH.length; i++)
      macRomanReverse.set(MAC_ROMAN_HIGH[i]!, 0x80 + i);
  }
  return macRomanReverse.get(cp) ?? -1;
}

interface Candidate {
  rank: number;
  lookup: CmapLookup;
}

/** The parsed cmap: a lookup chain and whether the best subtable is a symbol-font table. */
export interface Cmap {
  lookup: CmapLookup;
  symbol: boolean;
}

/**
 * Parses the cmap table. Throws FONT_FORMAT when no usable subtable exists; a malformed subtable
 * is skipped as long as another one works.
 */
export function parseCmap(r: Reader, numGlyphs: number): Cmap {
  const count = r.u16(2);
  r.array(4, count, 8);
  const candidates: Candidate[] = [];
  let symbol = false;
  let bestRank = -1;
  const seen = new Set<number>();
  for (let i = 0; i < count; i++) {
    const platform = r.u16(4 + i * 8);
    const encoding = r.u16(6 + i * 8);
    const offset = r.u32(8 + i * 8);
    if (offset >= r.length || seen.has(offset * 8 + platform)) continue;
    seen.add(offset * 8 + platform);
    let format: number;
    try {
      format = r.u16(offset);
    } catch {
      continue;
    }
    const rank = rankOf(platform, encoding, format);
    if (rank < 0) continue;
    let lookup: CmapLookup | undefined;
    try {
      lookup = subtable(r.sub(offset), format, numGlyphs);
    } catch {
      continue; // a broken subtable is skipped; another one may work
    }
    if (lookup === undefined) continue;
    if (platform === 1) lookup = macRoman(lookup);
    const isSymbol = platform === 3 && encoding === 0;
    if (isSymbol) lookup = symbolLookup(lookup);
    candidates.push({ rank, lookup });
    if (rank > bestRank) {
      bestRank = rank;
      symbol = isSymbol;
    }
  }
  if (candidates.length === 0) fontError('cmap: no supported Unicode subtable');
  candidates.sort((a, b) => b.rank - a.rank);
  const chain = candidates.map((c) => c.lookup);
  if (chain.length === 1) return { lookup: chain[0]!, symbol };
  return {
    symbol,
    lookup: (cp) => {
      for (const f of chain) {
        const g = f(cp);
        if (g !== 0) return g;
      }
      return 0;
    },
  };
}

function rankOf(platform: number, encoding: number, format: number): number {
  const full = format === 12 || format === 13;
  const bmp = format === 4;
  const small = format === 0 || format === 6;
  if (platform === 3) {
    if (encoding === 10 && full) return 100;
    if (encoding === 1 && bmp) return 90;
    if (encoding === 1 && (full || small)) return 85;
    if (encoding === 0 && (bmp || full || small)) return 70; // symbol
    return -1;
  }
  if (platform === 0) {
    if (encoding === 14 || format === 14) return -1; // variation sequences
    if (full) return 95;
    if (bmp) return 80;
    if (small) return 60;
    return -1;
  }
  if (platform === 1 && encoding === 0 && (small || bmp)) return 50;
  return -1;
}

function subtable(r: Reader, format: number, numGlyphs: number): CmapLookup | undefined {
  const valid = (g: number): number => (g < numGlyphs ? g : 0);
  switch (format) {
    case 0: {
      r.array(6, 256, 1);
      const ids = r.bytes.slice(r.start + 6, r.start + 262);
      return (cp) => (cp >= 0 && cp < 256 ? valid(ids[cp]!) : 0);
    }
    case 6: {
      const first = r.u16(6);
      const n = r.u16(8);
      r.array(10, n, 2);
      const ids = new Uint16Array(n);
      for (let i = 0; i < n; i++) ids[i] = r.u16(10 + i * 2);
      return (cp) => {
        const i = cp - first;
        return i >= 0 && i < n ? valid(ids[i]!) : 0;
      };
    }
    case 4:
      return format4(r, valid);
    case 12:
    case 13: {
      const n = r.u32(12);
      r.array(16, n, 12);
      const groups = new Uint32Array(n * 3);
      for (let i = 0; i < n * 3; i++) groups[i] = r.u32(16 + i * 4);
      const constant = format === 13;
      return (cp) => {
        let lo = 0;
        let hi = n - 1;
        while (lo <= hi) {
          const mid = (lo + hi) >>> 1;
          const start = groups[mid * 3]!;
          const end = groups[mid * 3 + 1]!;
          if (cp < start) hi = mid - 1;
          else if (cp > end) lo = mid + 1;
          else return valid(constant ? groups[mid * 3 + 2]! : groups[mid * 3 + 2]! + (cp - start));
        }
        return 0;
      };
    }
    default:
      return undefined;
  }
}

function format4(r: Reader, valid: (g: number) => number): CmapLookup {
  // The subtable's own length field is unreliable in real fonts (it overflows 16 bits), so the
  // arrays are checked against the cmap table end instead.
  const segCount = r.u16(6) >>> 1;
  if (segCount === 0) fontError('cmap format 4: no segments');
  const ends = 14;
  const starts = ends + segCount * 2 + 2;
  const deltas = starts + segCount * 2;
  const rangeOffsets = deltas + segCount * 2;
  r.array(ends, segCount * 4 + 1, 2);
  const end = new Uint16Array(segCount);
  const start = new Uint16Array(segCount);
  const delta = new Uint16Array(segCount);
  const range = new Uint16Array(segCount);
  for (let i = 0; i < segCount; i++) {
    end[i] = r.u16(ends + i * 2);
    start[i] = r.u16(starts + i * 2);
    delta[i] = r.u16(deltas + i * 2);
    range[i] = r.u16(rangeOffsets + i * 2);
  }
  return (cp) => {
    if (cp < 0 || cp > 0xffff) return 0;
    let lo = 0;
    let hi = segCount - 1;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (end[mid]! < cp) lo = mid + 1;
      else hi = mid;
    }
    if (end[lo]! < cp || start[lo]! > cp) return 0;
    const ro = range[lo]!;
    if (ro === 0) return valid((cp + delta[lo]!) & 0xffff);
    const at = rangeOffsets + lo * 2 + ro + (cp - start[lo]!) * 2;
    if (!r.has(at, 2)) return 0;
    const g = r.u16(at);
    return g === 0 ? 0 : valid((g + delta[lo]!) & 0xffff);
  };
}

function macRoman(inner: CmapLookup): CmapLookup {
  return (cp) => {
    const code = macRomanCode(cp);
    return code < 0 ? 0 : inner(code);
  };
}

function symbolLookup(inner: CmapLookup): CmapLookup {
  return (cp) => {
    if (cp >= 0x20 && cp <= 0xff) return inner(0xf000 + cp) || inner(cp);
    return inner(cp);
  };
}
