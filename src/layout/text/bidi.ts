// UAX #9 (Unicode Bidirectional Algorithm, Unicode 16.0): P1–P3, X1–X10, W1–W7, N0–N2, I1–I2, L1
// and L2, following the structure of the Unicode reference implementation (BidiReference.java):
// explicit levels first, then each isolating run sequence is resolved on its own with characters
// removed by X9 skipped. Conformance: BidiTest.txt and BidiCharacterTest.txt pass in full
// (test/unicode-bidi.test.ts).
//
// Levels are computed per code point and spread to both halves of surrogate pairs. Characters
// removed by X9 (embedding controls, BN) get the level of the preceding character so they never
// split a reversal in L2.

import { BIDI_BRACKETS, BIDI_CLASSES, BIDI_MIRRORS, BIDI_PALETTE, BIDI_RUNS } from './data/bidi.js';
import { decodeRuns, decodeVarints, lazy, lookup, type RangeTable } from './table.js';

// Bidi_Class values, numbered as in BIDI_CLASSES (scripts/gen-unicode.mjs).
const L = 0;
const R = 1;
const AL = 2;
const EN = 3;
const ES = 4;
const ET = 5;
const AN = 6;
const CS = 7;
const NSM = 8;
const BN = 9;
const B = 10;
const S = 11;
const WS = 12;
const ON = 13;
const LRE = 14;
const LRO = 15;
const RLE = 16;
const RLO = 17;
const PDF = 18;
const LRI = 19;
const RLI = 20;
const FSI = 21;
const PDI = 22;

const MAX_DEPTH = 125;
const BRACKET_STACK = 63;

const classes: () => RangeTable = lazy(() => decodeRuns(BIDI_RUNS, decodeVarints(BIDI_PALETTE)));

/** Opening bracket → closing bracket, and closing → opening (as negative numbers). */
const brackets: () => Map<number, number> = lazy(() => {
  const nums = decodeVarints(BIDI_BRACKETS);
  const map = new Map<number, number>();
  let open = 0;
  for (let i = 0; i < nums.length; i += 2) {
    open += nums[i]!;
    const close = open + unzigzag(nums[i + 1]!);
    map.set(open, close);
    map.set(close, -open);
  }
  return map;
});

const mirrors: () => Map<number, number> = lazy(() => {
  const nums = decodeVarints(BIDI_MIRRORS);
  const map = new Map<number, number>();
  let cp = 0;
  for (let i = 0; i < nums.length; i += 2) {
    cp += nums[i]!;
    map.set(cp, cp + unzigzag(nums[i + 1]!));
  }
  return map;
});

function unzigzag(n: number): number {
  return n % 2 === 0 ? n / 2 : -(n + 1) / 2;
}

/** Bidi_Class names, indexed by class number. */
const BIDI_NAMES: readonly string[] = BIDI_CLASSES.split(' ');

/** The Bidi_Class of a code point ('L', 'R', 'AL', 'EN', 'NSM', 'ON', ...). */
export function bidiClass(codePoint: number): string {
  return BIDI_NAMES[lookup(classes(), codePoint)]!;
}

/** Bidi_Mirroring_Glyph: the mirrored counterpart of a code point (e.g. '(' ↔ ')'), if any. */
export function mirrorOf(codePoint: number): number | undefined {
  return mirrors().get(codePoint);
}

/**
 * Bidi_Paired_Bracket: for an opening bracket its closing bracket (> 0), for a closing bracket
 * minus its opening bracket (< 0), otherwise 0.
 */
export function pairedBracket(codePoint: number): number {
  return brackets().get(codePoint) ?? 0;
}

/** Bidi_Class number of a code point (internal; used by the shaper to spot nonspacing marks). */
export function bidiClassId(codePoint: number): number {
  return lookup(classes(), codePoint);
}

/** Bidi_Class number of a nonspacing mark (see bidiClassId). */
export const BIDI_NSM: number = NSM;

const isIsolateInit = (t: number): boolean => t === LRI || t === RLI || t === FSI;
/** Removed by X9. */
const isRemoved = (t: number): boolean => t === BN || (t >= LRE && t <= PDF);
/** Neutral or isolate (NI in N1/N2). */
const isNI = (t: number): boolean => t === B || t === S || t === WS || t === ON || t >= LRI;
/** Whitespace for L1 (WS, isolate controls, and characters removed by X9). */
const isL1Space = (t: number): boolean => t === WS || t >= LRI || isRemoved(t);

/** Code points of `text` with their UTF-16 offsets. */
function codePoints(text: string): { cps: Int32Array; offsets: Int32Array; count: number } {
  const cps = new Int32Array(text.length);
  const offsets = new Int32Array(text.length + 1);
  let count = 0;
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    cps[count] = cp;
    offsets[count++] = i;
    i += cp > 0xffff ? 2 : 1;
  }
  offsets[count] = text.length;
  return { cps, offsets, count };
}

/**
 * P2/P3: 1 if the first strong character (skipping isolates) in types[from, to) is R or AL, 0 if
 * it is L or there is none. `matchingPdi` jumps over isolates.
 */
function firstStrong(types: Uint8Array, from: number, to: number, matchingPdi: Int32Array): 0 | 1 {
  for (let i = from; i < to; i++) {
    const t = types[i]!;
    if (t === L) return 0;
    if (t === R || t === AL) return 1;
    if (isIsolateInit(t)) {
      const m = matchingPdi[i]!;
      if (m < 0) return 0;
      i = m;
    }
  }
  return 0;
}

/** BD9: the matching PDI of each isolate initiator (-1 if none), within paragraph [from, to). */
function matchIsolates(types: Uint8Array, from: number, to: number, matchingPdi: Int32Array): void {
  const stack: number[] = [];
  for (let i = from; i < to; i++) {
    const t = types[i]!;
    matchingPdi[i] = -1;
    if (isIsolateInit(t)) stack.push(i);
    else if (t === PDI && stack.length > 0) matchingPdi[stack.pop()!] = i;
  }
}

/**
 * The base level of the first paragraph of `text` by rules P2–P3: 1 when its first strong
 * character (outside isolates) is R or AL, else 0.
 */
export function paragraphLevel(text: string): 0 | 1 {
  const { cps, count } = codePoints(text);
  const t = classes();
  const types = new Uint8Array(count);
  let end = count;
  for (let i = 0; i < count; i++) {
    types[i] = lookup(t, cps[i]!);
    if (types[i] === B && end === count) end = i;
  }
  const matchingPdi = new Int32Array(count);
  matchIsolates(types, 0, end, matchingPdi);
  return firstStrong(types, 0, end, matchingPdi);
}

/**
 * Embedding level per UTF-16 index (UAX #9), base direction from `rtl` ('auto' = rules P2–P3 per
 * paragraph). Paragraph separators split paragraphs (P1). L1 is applied as if the text were one
 * line; for a paragraph broken into several lines, take each line's levels from `lineLevels`.
 */
export function bidiLevels(text: string, rtl: boolean | 'auto'): Uint8Array {
  const out = new Uint8Array(text.length);
  const { cps, offsets, count } = codePoints(text);
  if (count === 0) return out;
  const t = classes();
  const initial = new Uint8Array(count);
  for (let i = 0; i < count; i++) initial[i] = lookup(t, cps[i]!);
  const types = initial.slice();
  const levels = new Uint8Array(count);
  const matchingPdi = new Int32Array(count);

  for (let start = 0; start < count; ) {
    let end = start;
    while (end < count && initial[end] !== B) end++;
    if (end < count) end++; // the separator belongs to its paragraph
    matchIsolates(initial, start, end, matchingPdi);
    const para = rtl === 'auto' ? firstStrong(initial, start, end, matchingPdi) : rtl ? 1 : 0;
    resolveParagraph(cps, initial, types, levels, matchingPdi, start, end, para);
    start = end;
  }
  for (let i = 0; i < count; i++) out.fill(levels[i]!, offsets[i]!, offsets[i + 1]!);
  return out;
}

function resolveParagraph(
  cps: Int32Array,
  initial: Uint8Array,
  types: Uint8Array,
  levels: Uint8Array,
  matchingPdi: Int32Array,
  start: number,
  end: number,
  para: number,
): void {
  explicitLevels(initial, types, levels, matchingPdi, start, end, para);
  // sos/eos compare embedding levels, so keep them before I1/I2 raise resolved ones.
  const embedding = levels.slice(start, end);

  // X10: level runs (removed characters skipped) chained into isolating run sequences.
  const runs: number[][] = [];
  const runOf = new Int32Array(end - start).fill(-1);
  let run: number[] | null = null;
  let runLevel = -1;
  for (let i = start; i < end; i++) {
    if (isRemoved(initial[i]!)) continue;
    if (run === null || levels[i] !== runLevel) {
      run = [];
      runs.push(run);
      runLevel = levels[i]!;
    }
    run.push(i);
    runOf[i - start] = runs.length - 1;
  }
  // A run starting with a matched PDI continues the sequence of its isolate initiator.
  const closesIsolate = new Uint8Array(end - start);
  for (let i = start; i < end; i++) {
    const m = isIsolateInit(initial[i]!) ? matchingPdi[i]! : -1;
    if (m >= 0) closesIsolate[m - start] = 1;
  }
  for (const first of runs) {
    if (closesIsolate[first[0]! - start]) continue;
    const seq: number[] = [];
    let r: number[] | undefined = first;
    while (r) {
      for (const i of r) seq.push(i);
      const last: number = r[r.length - 1]!;
      const pdi: number = isIsolateInit(initial[last]!) ? matchingPdi[last]! : -1;
      r = pdi >= 0 ? runs[runOf[pdi - start]!] : undefined;
    }
    resolveSequence(seq, cps, initial, types, levels, embedding, start, end, para);
  }

  // Removed characters take the level of the preceding character (or the paragraph level).
  for (let i = start; i < end; i++) {
    if (isRemoved(initial[i]!)) levels[i] = i > start ? levels[i - 1]! : para;
  }

  // L1: separators, and whitespace before them or at the end of the line, go to the paragraph level.
  let trailing = true;
  for (let i = end - 1; i >= start; i--) {
    const c = initial[i]!;
    if (c === B || c === S) {
      levels[i] = para;
      trailing = true;
    } else if (trailing && isL1Space(c)) levels[i] = para;
    else trailing = false;
  }
}

/** X1–X8: explicit embedding levels and directional overrides. */
function explicitLevels(
  initial: Uint8Array,
  types: Uint8Array,
  levels: Uint8Array,
  matchingPdi: Int32Array,
  start: number,
  end: number,
  para: number,
): void {
  // Directional status stack: level, override (-1 neutral, else L or R), isolate flag.
  const stLevel = new Uint8Array(MAX_DEPTH + 2);
  const stOverride = new Int8Array(MAX_DEPTH + 2);
  const stIsolate = new Uint8Array(MAX_DEPTH + 2);
  let top = 0;
  stLevel[0] = para;
  stOverride[0] = -1;
  stIsolate[0] = 0;
  let overflowIsolates = 0;
  let overflowEmbeddings = 0;
  let validIsolates = 0;

  for (let i = start; i < end; i++) {
    const t = initial[i]!;
    const cur = stLevel[top]!;
    switch (t) {
      case RLE:
      case LRE:
      case RLO:
      case LRO: {
        const next = t === RLE || t === RLO ? (cur + 1) | 1 : (cur + 2) & ~1;
        if (next <= MAX_DEPTH && overflowIsolates === 0 && overflowEmbeddings === 0) {
          top++;
          stLevel[top] = next;
          stOverride[top] = t === RLO ? R : t === LRO ? L : -1;
          stIsolate[top] = 0;
        } else if (overflowIsolates === 0) overflowEmbeddings++;
        levels[i] = cur;
        break;
      }
      case RLI:
      case LRI:
      case FSI: {
        levels[i] = cur;
        if (stOverride[top]! >= 0) types[i] = stOverride[top]!;
        let rtl = t === RLI;
        if (t === FSI) {
          const m = matchingPdi[i]!;
          rtl = firstStrong(initial, i + 1, m < 0 ? end : m, matchingPdi) === 1;
        }
        const next = rtl ? (cur + 1) | 1 : (cur + 2) & ~1;
        if (next <= MAX_DEPTH && overflowIsolates === 0 && overflowEmbeddings === 0) {
          validIsolates++;
          top++;
          stLevel[top] = next;
          stOverride[top] = -1;
          stIsolate[top] = 1;
        } else overflowIsolates++;
        break;
      }
      case PDI:
        if (overflowIsolates > 0) overflowIsolates--;
        else if (validIsolates > 0) {
          overflowEmbeddings = 0;
          while (!stIsolate[top]) top--;
          top--;
          validIsolates--;
        }
        levels[i] = stLevel[top]!;
        if (stOverride[top]! >= 0) types[i] = stOverride[top]!;
        break;
      case PDF:
        if (overflowIsolates > 0) {
          // ignored
        } else if (overflowEmbeddings > 0) overflowEmbeddings--;
        else if (!stIsolate[top] && top > 0) top--;
        levels[i] = cur;
        break;
      case B:
        levels[i] = para;
        break;
      case BN:
        levels[i] = cur;
        break;
      default:
        levels[i] = cur;
        if (stOverride[top]! >= 0) types[i] = stOverride[top]!;
    }
  }
}

/** W1–W7, N0–N2 and I1–I2 on one isolating run sequence (indices of its characters, in order). */
function resolveSequence(
  seq: number[],
  cps: Int32Array,
  initial: Uint8Array,
  types: Uint8Array,
  levels: Uint8Array,
  embedding: Uint8Array,
  start: number,
  end: number,
  para: number,
): void {
  const n = seq.length;
  const first = seq[0]!;
  const last = seq[n - 1]!;
  const level = embedding[first - start]!;

  // sos/eos from the neighbouring non-removed characters.
  let before = first - 1;
  while (before >= start && isRemoved(initial[before]!)) before--;
  let after = last + 1;
  while (after < end && isRemoved(initial[after]!)) after++;
  const prevLevel = before >= start ? embedding[before - start]! : para;
  const nextLevel =
    after < end && !isIsolateInit(initial[last]!) ? embedding[after - start]! : para;
  const sos = Math.max(prevLevel, level) & 1 ? R : L;
  const eos = Math.max(nextLevel, level) & 1 ? R : L;

  const ty = new Uint8Array(n);
  for (let k = 0; k < n; k++) ty[k] = types[seq[k]!]!;
  const beforeW = ty.slice(); // types before W1 (N0 needs the original NSMs)

  // W1: NSM takes the type of the previous character (ON after an isolate initiator or PDI).
  let prev = sos;
  for (let k = 0; k < n; k++) {
    const c = ty[k]!;
    if (c === NSM) ty[k] = prev;
    else prev = c >= LRI ? ON : c;
  }
  // W2, W3: EN after AL becomes AN; AL becomes R.
  let strong = sos;
  for (let k = 0; k < n; k++) {
    const c = ty[k]!;
    if (c === L || c === R || c === AL) strong = c;
    if (c === EN && strong === AL) ty[k] = AN;
  }
  for (let k = 0; k < n; k++) if (ty[k] === AL) ty[k] = R;
  // W4: a single separator between two numbers of the same kind.
  for (let k = 1; k < n - 1; k++) {
    const c = ty[k]!;
    const p = ty[k - 1]!;
    const q = ty[k + 1]!;
    if (c === ES && p === EN && q === EN) ty[k] = EN;
    else if (c === CS && p === q && (p === EN || p === AN)) ty[k] = p;
  }
  // W5: terminators next to European numbers.
  for (let k = 0; k < n; k++) {
    if (ty[k] !== ET) continue;
    let e = k;
    while (e < n && ty[e] === ET) e++;
    if ((k > 0 && ty[k - 1] === EN) || (e < n && ty[e] === EN)) ty.fill(EN, k, e);
    k = e - 1;
  }
  // W6: remaining separators and terminators become ON.
  for (let k = 0; k < n; k++) {
    const c = ty[k]!;
    if (c === ES || c === ET || c === CS) ty[k] = ON;
  }
  // W7: EN after L (or an L sos) becomes L.
  strong = sos;
  for (let k = 0; k < n; k++) {
    const c = ty[k]!;
    if (c === L || c === R) strong = c;
    else if (c === EN && strong === L) ty[k] = L;
  }

  // N0: paired brackets.
  const dirE = level & 1 ? R : L; // embedding direction
  const pairs = bracketPairs(seq, cps, ty);
  const strongOf = (c: number): number => (c === L ? L : c === R || c === EN || c === AN ? R : -1);
  for (const [open, close] of pairs) {
    let foundE = false;
    let foundO = false;
    for (let k = open + 1; k < close; k++) {
      const s = strongOf(ty[k]!);
      if (s === dirE) {
        foundE = true;
        break;
      }
      if (s >= 0) foundO = true;
    }
    let dir = -1;
    if (foundE) dir = dirE;
    else if (foundO) {
      // The opposite direction only if the context before the brackets has it too (N0 c).
      dir = sos;
      for (let k = open - 1; k >= 0; k--) {
        const s = strongOf(ty[k]!);
        if (s >= 0) {
          dir = s;
          break;
        }
      }
    }
    if (dir < 0) continue;
    ty[open] = dir;
    ty[close] = dir;
    for (let k = open + 1; k < n && beforeW[k] === NSM; k++) ty[k] = dir;
    for (let k = close + 1; k < n && beforeW[k] === NSM; k++) ty[k] = dir;
  }

  // N1, N2: neutrals between same-direction strong types take it; else the e direction.
  for (let k = 0; k < n; k++) {
    if (!isNI(ty[k]!)) continue;
    let e = k;
    while (e < n && isNI(ty[e]!)) e++;
    const lead = k === 0 ? sos : strongOf(ty[k - 1]!);
    const trail = e === n ? eos : strongOf(ty[e]!);
    ty.fill(lead === trail ? lead : dirE, k, e);
    k = e - 1;
  }

  // I1, I2
  for (let k = 0; k < n; k++) {
    const c = ty[k]!;
    const i = seq[k]!;
    if (level & 1) {
      if (c === L || c === EN || c === AN) levels[i] = level + 1;
    } else if (c === R) levels[i] = level + 1;
    else if (c === AN || c === EN) levels[i] = level + 2;
  }
}

/** BD16: bracket pairs (positions within the sequence), sorted by opening position. */
function bracketPairs(seq: number[], cps: Int32Array, ty: Uint8Array): [number, number][] {
  const map = brackets();
  const pairs: [number, number][] = [];
  const stack: [number, number][] = []; // [closing code point, position]
  for (let k = 0; k < seq.length; k++) {
    if (ty[k] !== ON) continue;
    const cp = canonicalBracket(cps[seq[k]!]!);
    const v = map.get(cp);
    if (v === undefined) continue;
    if (v > 0) {
      if (stack.length === BRACKET_STACK) break;
      stack.push([canonicalBracket(v), k]);
    } else {
      for (let j = stack.length - 1; j >= 0; j--) {
        if (stack[j]![0] === cp) {
          pairs.push([stack[j]![1], k]);
          stack.length = j;
          break;
        }
      }
    }
  }
  return pairs.sort((x, y) => x[0] - y[0]);
}

/** U+2329/U+232A are canonically equivalent to U+3008/U+3009 (BD16 compares canonically). */
function canonicalBracket(cp: number): number {
  return cp === 0x2329 ? 0x3008 : cp === 0x232a ? 0x3009 : cp;
}

/**
 * Levels of the line text[start, end) of a paragraph whose levels came from `bidiLevels`: rule L1
 * again for the line end, so whitespace (and isolate controls) before a line break go back to
 * the paragraph level. Returns a new array of end - start levels.
 */
export function lineLevels(
  text: string,
  levels: Uint8Array,
  start: number,
  end: number,
  paragraphLevel: number,
): Uint8Array {
  const out = levels.slice(start, end);
  const t = classes();
  for (let i = end; i > start; ) {
    // Step back one code point (both halves of a surrogate pair).
    let k = i - 1;
    const lo = text.charCodeAt(k);
    if (lo >= 0xdc00 && lo < 0xe000 && k > start) {
      const hi = text.charCodeAt(k - 1);
      if (hi >= 0xd800 && hi < 0xdc00) k--;
    }
    const c = lookup(t, text.codePointAt(k)!);
    if (!isL1Space(c) && c !== B && c !== S) break;
    out.fill(paragraphLevel, k - start, i - start);
    i = k;
  }
  return out;
}

/** Visual order of `levels` (indices), per UAX #9 rule L2. */
export function visualOrder(levels: Uint8Array): number[] {
  const n = levels.length;
  const order: number[] = new Array(n);
  let max = 0;
  let minOdd = 255;
  for (let i = 0; i < n; i++) {
    order[i] = i;
    const l = levels[i]!;
    if (l > max) max = l;
    if (l & 1 && l < minOdd) minOdd = l;
  }
  for (let level = max; level >= minOdd && level > 0; level--) {
    for (let i = 0; i < n; ) {
      if (levels[order[i]!]! < level) {
        i++;
        continue;
      }
      let e = i;
      while (e < n && levels[order[e]!]! >= level) e++;
      for (let a = i, b = e - 1; a < b; a++, b--) {
        const tmp = order[a]!;
        order[a] = order[b]!;
        order[b] = tmp;
      }
      i = e;
    }
  }
  return order;
}

/** A maximal run of one level, as returned by visualRuns. */
export interface BidiRun {
  /** UTF-16 range [start, end) in logical order. */
  start: number;
  end: number;
  level: number;
}

/**
 * The level runs of `levels`, in visual order (L2 applied to whole runs): what a line layout
 * needs to place shaped runs left to right; glyphs inside a run with an odd level go right to
 * left.
 */
export function visualRuns(levels: Uint8Array): BidiRun[] {
  const runs: BidiRun[] = [];
  for (let i = 0; i < levels.length; ) {
    let e = i + 1;
    while (e < levels.length && levels[e] === levels[i]) e++;
    runs.push({ start: i, end: e, level: levels[i]! });
    i = e;
  }
  const order = visualOrder(Uint8Array.from(runs, (r) => r.level));
  return order.map((k) => runs[k]!);
}
