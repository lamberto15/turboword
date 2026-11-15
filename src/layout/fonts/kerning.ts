// Pair kerning: GPOS PairPos (lookup type 2, also behind type 9 extension wrappers) of the `kern`
// feature, else the legacy `kern` table (format 0 subtables, Microsoft and Apple headers).
//
// Only what a cmap + pair-kerning shaper needs: the horizontal advance adjustment of the first
// glyph of a pair (ValueRecord XAdvance of value1). Lookups are those of the `kern` feature of the
// DFLT script's default language system (else latn, else the first script), applied in LookupList
// order and summed; within a lookup the first subtable that applies wins, as in OpenType.
// Everything is parsed lazily on the first kerning() call; a malformed GPOS/kern table disables
// kerning instead of failing the font (kerning is an enhancement, not required to lay out text).

import type { Reader } from './reader.js';

/** Kerning adjustment of a glyph pair, in font units. */
export type Kerner = (left: number, right: number) => number;

const NO_KERNING: Kerner = () => 0;
const CACHE_LIMIT = 1 << 16;

/** Builds the kerner from the GPOS and kern table windows (either may be undefined). */
export function buildKerner(gpos: Reader | undefined, kern: Reader | undefined): Kerner {
  let inner: Kerner | undefined;
  if (gpos !== undefined) {
    try {
      inner = gposKerner(gpos);
    } catch {
      inner = undefined; // malformed GPOS: fall back to the kern table
    }
  }
  if (inner === undefined && kern !== undefined) {
    try {
      inner = kernTableKerner(kern);
    } catch {
      inner = undefined;
    }
  }
  if (inner === undefined) return NO_KERNING;
  const run = inner;
  const cache = new Map<number, number>();
  return (left, right) => {
    if (left < 0 || right < 0 || left > 0xffff || right > 0xffff) return 0;
    const key = left * 0x10000 + right;
    const hit = cache.get(key);
    if (hit !== undefined) return hit;
    let v: number;
    try {
      v = run(left, right);
    } catch {
      v = 0;
    }
    if (cache.size < CACHE_LIMIT) cache.set(key, v);
    return v;
  };
}

// ---- legacy kern table ------------------------------------------------------------------------

function kernTableKerner(r: Reader): Kerner | undefined {
  const pairs = new Map<number, number>();
  const version = r.u16(0);
  let nTables: number;
  let at: number;
  const apple = version === 1;
  if (apple) {
    nTables = r.u32(4);
    at = 8;
  } else if (version === 0) {
    nTables = r.u16(2);
    at = 4;
  } else {
    return undefined;
  }
  for (let t = 0; t < nTables && r.has(at, apple ? 8 : 6); t++) {
    let length: number;
    let format: number;
    let usable: boolean;
    let override = false;
    let header: number;
    if (apple) {
      length = r.u32(at);
      const coverage = r.u16(at + 4);
      format = coverage & 0xff;
      usable = (coverage & 0xe000) === 0; // not vertical, cross-stream or variation
      header = 8;
    } else {
      length = r.u16(at + 2);
      const coverage = r.u16(at + 4);
      format = coverage >>> 8;
      usable = (coverage & 0x7) === 0x1; // horizontal, not minimum, not cross-stream
      override = (coverage & 0x8) !== 0;
      header = 6;
    }
    let size = length;
    if (format === 0 && r.has(at + header, 8)) {
      const n = r.u16(at + header);
      const data = at + header + 8;
      // A 16-bit length overflows for big subtables: trust nPairs when the length is truncated.
      const real = header + 8 + n * 6;
      if (!apple && (real & 0xffff) === length) size = real;
      if (usable && r.has(data, n * 6)) {
        for (let i = 0; i < n; i++) {
          const key = r.u16(data + i * 6) * 0x10000 + r.u16(data + i * 6 + 2);
          const v = r.i16(data + i * 6 + 4);
          pairs.set(key, override ? v : (pairs.get(key) ?? 0) + v);
        }
      }
    }
    if (size <= 0) break;
    at += size;
  }
  if (pairs.size === 0) return undefined;
  return (left, right) => pairs.get(left * 0x10000 + right) ?? 0;
}

// ---- GPOS --------------------------------------------------------------------------------------

type CoverageFn = (glyph: number) => number;
type ClassFn = (glyph: number) => number;

interface PairSubtable {
  coverage: CoverageFn;
  apply(left: number, coverageIndex: number, right: number): number | undefined;
}

function gposKerner(r: Reader): Kerner | undefined {
  if (r.u16(0) !== 1) return undefined;
  const scriptList = r.sub(r.u16(4));
  const featureList = r.sub(r.u16(6));
  const lookupList = r.sub(r.u16(8));

  const langSys = defaultLangSys(scriptList);
  if (langSys === undefined) return undefined;
  const featureCount = featureList.u16(0);
  featureList.array(2, featureCount, 6);
  const lookupIndices = new Set<number>();
  const required = langSys.u16(2);
  const n = langSys.u16(4);
  langSys.array(6, n, 2);
  const featureIndices: number[] = [];
  if (required !== 0xffff) featureIndices.push(required);
  for (let i = 0; i < n; i++) featureIndices.push(langSys.u16(6 + i * 2));
  for (const fi of featureIndices) {
    if (fi >= featureCount) continue;
    if (featureList.tag(2 + fi * 6) !== 'kern') continue;
    const feature = featureList.sub(featureList.u16(2 + fi * 6 + 4));
    const count = feature.u16(2);
    feature.array(4, count, 2);
    for (let i = 0; i < count; i++) lookupIndices.add(feature.u16(4 + i * 2));
  }
  if (lookupIndices.size === 0) return undefined;

  const lookupCount = lookupList.u16(0);
  lookupList.array(2, lookupCount, 2);
  const lookups: PairSubtable[][] = [];
  for (const li of [...lookupIndices].sort((a, b) => a - b)) {
    if (li >= lookupCount) continue;
    const subs = pairLookup(lookupList.sub(lookupList.u16(2 + li * 2)));
    if (subs.length > 0) lookups.push(subs);
  }
  if (lookups.length === 0) return undefined;
  return (left, right) => {
    let total = 0;
    for (const subs of lookups) {
      for (const st of subs) {
        const ci = st.coverage(left);
        if (ci < 0) continue;
        const v = st.apply(left, ci, right);
        if (v !== undefined) {
          total += v;
          break;
        }
      }
    }
    return total;
  };
}

function defaultLangSys(scriptList: Reader): Reader | undefined {
  const count = scriptList.u16(0);
  scriptList.array(2, count, 6);
  let pick = -1;
  for (const want of ['DFLT', 'latn']) {
    for (let i = 0; i < count && pick < 0; i++) {
      if (scriptList.tag(2 + i * 6) === want) pick = i;
    }
    if (pick >= 0) break;
  }
  if (pick < 0 && count > 0) pick = 0;
  if (pick < 0) return undefined;
  const script = scriptList.sub(scriptList.u16(2 + pick * 6 + 4));
  const def = script.u16(0);
  if (def !== 0) return script.sub(def);
  // No default LangSys: use the first language system of the script.
  const langCount = script.u16(2);
  script.array(4, langCount, 6);
  return langCount > 0 ? script.sub(script.u16(4 + 4)) : undefined;
}

function pairLookup(lookup: Reader): PairSubtable[] {
  const type = lookup.u16(0);
  const count = lookup.u16(4);
  lookup.array(6, count, 2);
  const out: PairSubtable[] = [];
  for (let i = 0; i < count; i++) {
    let st = lookup.sub(lookup.u16(6 + i * 2));
    let stType = type;
    if (type === 9) {
      if (st.u16(0) !== 1) continue;
      stType = st.u16(2);
      st = st.sub(st.u32(4));
    }
    if (stType !== 2) continue;
    const parsed = pairPos(st);
    if (parsed !== undefined) out.push(parsed);
  }
  return out;
}

const popcount = (v: number): number => {
  let c = 0;
  for (let x = v & 0xff; x !== 0; x &= x - 1) c++;
  return c;
};

function pairPos(st: Reader): PairSubtable | undefined {
  const format = st.u16(0);
  const coverage = coverageOf(st.sub(st.u16(2)));
  const vf1 = st.u16(4);
  const vf2 = st.u16(6);
  const size1 = popcount(vf1) * 2;
  const size2 = popcount(vf2) * 2;
  const hasXAdv = (vf1 & 0x4) !== 0;
  const xAdvAt = popcount(vf1 & 0x3) * 2;
  if (format === 1) {
    const setCount = st.u16(8);
    st.array(10, setCount, 2);
    const record = 2 + size1 + size2;
    return {
      coverage,
      apply(_left, ci, right) {
        if (ci >= setCount) return undefined;
        const set = st.sub(st.u16(10 + ci * 2));
        const n = set.u16(0);
        set.array(2, n, record);
        let lo = 0;
        let hi = n - 1;
        while (lo <= hi) {
          const mid = (lo + hi) >>> 1;
          const g = set.u16(2 + mid * record);
          if (g < right) lo = mid + 1;
          else if (g > right) hi = mid - 1;
          else return hasXAdv ? set.i16(2 + mid * record + 2 + xAdvAt) : 0;
        }
        return undefined;
      },
    };
  }
  if (format === 2) {
    const class1 = classDefOf(st.sub(st.u16(8)));
    const class2 = classDefOf(st.sub(st.u16(10)));
    const c1Count = st.u16(12);
    const c2Count = st.u16(14);
    const record = size1 + size2;
    st.array(16, c1Count * c2Count, record);
    return {
      coverage,
      apply(left, _ci, right) {
        const k1 = class1(left);
        const k2 = class2(right);
        if (k1 >= c1Count || k2 >= c2Count) return undefined;
        if (!hasXAdv) return 0;
        return st.i16(16 + (k1 * c2Count + k2) * record + xAdvAt);
      },
    };
  }
  return undefined;
}

function coverageOf(r: Reader): CoverageFn {
  const format = r.u16(0);
  const n = r.u16(2);
  if (format === 1) {
    r.array(4, n, 2);
    return (g) => {
      let lo = 0;
      let hi = n - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >>> 1;
        const v = r.u16(4 + mid * 2);
        if (v < g) lo = mid + 1;
        else if (v > g) hi = mid - 1;
        else return mid;
      }
      return -1;
    };
  }
  if (format === 2) {
    r.array(4, n, 6);
    return (g) => {
      let lo = 0;
      let hi = n - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >>> 1;
        const at = 4 + mid * 6;
        if (r.u16(at + 2) < g) lo = mid + 1;
        else if (r.u16(at) > g) hi = mid - 1;
        else return r.u16(at + 4) + g - r.u16(at);
      }
      return -1;
    };
  }
  return () => -1;
}

function classDefOf(r: Reader): ClassFn {
  const format = r.u16(0);
  if (format === 1) {
    const first = r.u16(2);
    const n = r.u16(4);
    r.array(6, n, 2);
    return (g) => (g >= first && g < first + n ? r.u16(6 + (g - first) * 2) : 0);
  }
  if (format === 2) {
    const n = r.u16(2);
    r.array(4, n, 6);
    return (g) => {
      let lo = 0;
      let hi = n - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >>> 1;
        const at = 4 + mid * 6;
        if (r.u16(at + 2) < g) lo = mid + 1;
        else if (r.u16(at) > g) hi = mid - 1;
        else return r.u16(at + 4);
      }
      return 0;
    };
  }
  return () => 0;
}
