// UAX #14 (Unicode Line Breaking Algorithm, Unicode 16.0) with the default rules LB2–LB31.
//
// Why a rule-by-rule evaluator instead of the old pair table: since Unicode 15.1 several rules
// look further than one character (LB15a/b quotation marks, LB19a East Asian context, LB20a,
// LB21a, LB25 numbers, LB28a Brahmic syllables), which a pair table cannot express. We first fold
// combining sequences into "units" (LB9/LB10: X CM* behaves as X), then evaluate the rules in
// order for each unit boundary with O(1) look-behind state, so long runs of spaces or digits stay
// linear. LineBreakTest.txt passes in full (test/unicode-linebreak.test.ts).

import { LB_CLASSES, LB_PALETTE, LB_RUNS } from './data/linebreak.js';
import { decodeRuns, decodeVarints, lazy, lookup, type RangeTable } from './table.js';

// Resolved Line_Break classes, numbered as in LB_CLASSES (scripts/gen-unicode.mjs).
const BK = 0;
const CR = 1;
const LF = 2;
const NL = 3;
const SP = 4;
const ZW = 5;
const ZWJ = 6;
const CM = 7;
const WJ = 8;
const GL = 9;
const BA = 10;
const HY = 11;
const CL = 12;
const CP = 13;
const EX = 14;
const SY = 15;
const OP = 16;
const QU = 17;
const IS = 18;
const NS = 19;
const B2 = 20;
const CB = 21;
const BB = 22;
const HL = 23;
const AL = 24;
const NU = 25;
const PR = 26;
const PO = 27;
const ID = 28;
const EB = 29;
const EM = 30;
const IN = 31;
const JL = 32;
const JV = 33;
const JT = 34;
const H2 = 35;
const H3 = 36;
const RI = 37;
const AK = 38;
const AP = 39;
const AS = 40;
const VF = 41;
const VI = 42;

// Flag bits stored above the class (value & 63).
const F_EA = 0x40; // East_Asian_Width F, W or H ($EastAsian)
const F_PI = 0x80; // QU with General_Category Pi
const F_PF = 0x100; // QU with General_Category Pf
const F_EPCN = 0x200; // Extended_Pictographic & General_Category Cn

const DOTTED_CIRCLE = 0x25cc;
const HYPHEN = 0x2010;

const table: () => RangeTable = lazy(() => decodeRuns(LB_RUNS, decodeVarints(LB_PALETTE)));

/** Line_Break class names (resolved per LB1), indexed by class number. */
export const LINE_BREAK_CLASSES: readonly string[] = LB_CLASSES.split(' ');

/**
 * The resolved Line_Break class of a code point (LB1 applied: AI/SG/XX → AL, SA → CM or AL by
 * General_Category, CJ → NS), e.g. 'AL', 'ID', 'SP', 'CM'.
 */
export function lineBreakClass(codePoint: number): string {
  return LINE_BREAK_CLASSES[lookup(table(), codePoint) & 63]!;
}

/** East_Asian_Width is Fullwidth, Wide or Halfwidth (UAX #14's $EastAsian). */
export function isEastAsianWidth(codePoint: number): boolean {
  return (lookup(table(), codePoint) & F_EA) !== 0;
}

/** Hard breaks (BK, CR, LF, NL) never absorb combining marks (LB9). */
const isHard = (c: number): boolean => c <= NL;
const isNoBase = (c: number): boolean => c <= ZW; // BK CR LF NL SP ZW
const isAlHl = (c: number): boolean => c === AL || c === HL;
const isAkDot = (c: number, cp: number): boolean => c === AK || c === AS || cp === DOTTED_CIRCLE;

/**
 * Per UTF-16 index: 0 = no break before this index, 1 = break allowed before it, 2 = mandatory
 * break before it (after BK, CR, LF, NL). Index 0 is always 0 (LB2) and the low half of a
 * surrogate pair is always 0; the break at the end of the text (LB3) has no index.
 */
export function lineBreakOpportunities(text: string): Uint8Array {
  const n = text.length;
  const out = new Uint8Array(n);
  if (n === 0) return out;
  const t = table();

  // Units after LB9/LB10: a base plus the CM/ZWJ characters it absorbs.
  const start = new Int32Array(n);
  const cls = new Uint8Array(n);
  const flags = new Uint16Array(n);
  const base = new Int32Array(n);
  const endsZwj = new Uint8Array(n); // the unit's last character is ZWJ (LB8a)
  let m = 0;
  for (let i = 0; i < n; ) {
    const cp = text.codePointAt(i)!;
    const v = lookup(t, cp);
    const c = v & 63;
    if (c === CM || c === ZWJ) {
      if (m > 0 && !isNoBase(cls[m - 1]!)) {
        endsZwj[m - 1] = c === ZWJ ? 1 : 0; // LB9: absorbed, no break before it
      } else {
        // LB10: a lone CM/ZWJ behaves as U+0041 (AL, not East Asian, not pictographic).
        start[m] = i;
        cls[m] = AL;
        flags[m] = 0;
        base[m] = 0x41;
        endsZwj[m] = c === ZWJ ? 1 : 0;
        m++;
      }
    } else {
      start[m] = i;
      cls[m] = c;
      flags[m] = v & ~63;
      base[m] = cp;
      endsZwj[m] = 0;
      m++;
    }
    i += cp > 0xffff ? 2 : 1;
  }

  const ea = (k: number): boolean => (flags[k]! & F_EA) !== 0;
  const clsAt = (k: number): number => (k < m ? cls[k]! : -1);

  // State of the boundary being decided: units a = b - 1 and b, their classes A and B.
  let a = 0;
  let b = 0;
  let A = 0;
  let B = 0;
  let nonSp = -1; // last unit before b that is not SP
  let ri = 0; // RI units ending at a
  let numPrev = false; // unit a-1 ends "NU (SY | IS)*"
  let num = false; // unit a ends "NU (SY | IS)*"

  /** The rules after LB3, in order; the first that applies decides. */
  const rule = (): number => {
    // LB4, LB5: hard breaks.
    if (A === BK) return 2;
    if (A === CR) return B === LF ? 0 : 2;
    if (A === LF || A === NL) return 2;
    // LB6, LB7
    if (isHard(B) || B === SP || B === ZW) return 0;
    // LB8: ZW SP* ÷
    const S = nonSp >= 0 ? cls[nonSp]! : -1;
    if (S === ZW) return 1;
    // LB8a: ZWJ ×
    if (endsZwj[a]) return 0;
    // LB11, LB12, LB12a
    if (A === WJ || B === WJ || A === GL) return 0;
    if (B === GL && A !== SP && A !== BA && A !== HY) return 0;
    // LB13
    if (B === CL || B === CP || B === EX || B === SY) return 0;
    // LB14: OP SP* ×
    if (S === OP) return 0;
    // LB15a: (sot | BK | CR | LF | NL | OP | QU | GL | SP | ZW) [Pi & QU] SP* ×
    if (S === QU && flags[nonSp]! & F_PI) {
      const p = nonSp > 0 ? cls[nonSp - 1]! : -1;
      if (p < 0 || p <= ZW || p === OP || p === QU || p === GL) return 0;
    }
    // LB15b: × [Pf & QU] (SP | GL | WJ | CL | QU | CP | EX | IS | SY | BK | CR | LF | NL | ZW | eot)
    if (B === QU && flags[b]! & F_PF) {
      const q = clsAt(b + 1);
      if (q < 0 || q <= ZW || q === GL || q === WJ || q === CL || q === QU || q === CP) return 0;
      if (q === EX || q === IS || q === SY) return 0;
    }
    // LB15c, LB15d
    if (B === IS) return A === SP && clsAt(b + 1) === NU ? 1 : 0;
    // LB16, LB17
    if (B === NS && (S === CL || S === CP)) return 0;
    if (B === B2 && S === B2) return 0;
    // LB18
    if (A === SP) return 1;
    // LB19
    if (B === QU && !(flags[b]! & F_PI)) return 0;
    if (A === QU && !(flags[a]! & F_PF)) return 0;
    // LB19a
    if (B === QU && (!ea(a) || b + 1 >= m || !ea(b + 1))) return 0;
    if (A === QU && (!ea(b) || a === 0 || !ea(a - 1))) return 0;
    // LB20
    if (A === CB || B === CB) return 1;
    // LB20a: (sot | BK | CR | LF | NL | SP | ZW | CB | GL) (HY | U+2010) × AL
    if ((A === HY || base[a] === HYPHEN) && B === AL) {
      const p = a > 0 ? cls[a - 1]! : -1;
      if (p < 0 || p <= ZW || p === CB || p === GL) return 0;
    }
    // LB21
    if (B === BA || B === HY || B === NS || A === BB) return 0;
    // LB21a: HL (HY | [BA - $EastAsian]) × [^HL]
    if (a > 0 && cls[a - 1] === HL && (A === HY || (A === BA && !ea(a))) && B !== HL) return 0;
    // LB21b, LB22
    if ((A === SY && B === HL) || B === IN) return 0;
    // LB23, LB23a, LB24
    if ((isAlHl(A) && B === NU) || (A === NU && isAlHl(B))) return 0;
    if (A === PR && (B === ID || B === EB || B === EM)) return 0;
    if ((A === ID || A === EB || A === EM) && B === PO) return 0;
    if ((A === PR || A === PO) && isAlHl(B)) return 0;
    if (isAlHl(A) && (B === PR || B === PO)) return 0;
    // LB25
    if (B === PO || B === PR) {
      if (num || ((A === CL || A === CP) && numPrev)) return 0;
    }
    if (A === PO || A === PR) {
      if (B === NU) return 0;
      if (B === OP) {
        const q = clsAt(b + 1);
        if (q === NU || (q === IS && clsAt(b + 2) === NU)) return 0;
      }
    }
    if (B === NU && (A === HY || A === IS || num)) return 0;
    // LB26, LB27
    if (A === JL && (B === JL || B === JV || B === H2 || B === H3)) return 0;
    if ((A === JV || A === H2) && (B === JV || B === JT)) return 0;
    if ((A === JT || A === H3) && B === JT) return 0;
    if (A >= JL && A <= H3 && B === PO) return 0;
    if (A === PR && B >= JL && B <= H3) return 0;
    // LB28
    if (isAlHl(A) && isAlHl(B)) return 0;
    // LB28a
    const akA = isAkDot(A, base[a]!);
    const akB = isAkDot(B, base[b]!);
    if (A === AP && akB) return 0;
    if (akA && (B === VF || B === VI)) return 0;
    if (A === VI && a > 0 && isAkDot(cls[a - 1]!, base[a - 1]!)) {
      if (B === AK || base[b] === DOTTED_CIRCLE) return 0;
    }
    if (akA && akB && clsAt(b + 1) === VF) return 0;
    // LB29, LB30
    if (A === IS && isAlHl(B)) return 0;
    if ((isAlHl(A) || A === NU) && B === OP && !ea(b)) return 0;
    if (A === CP && !ea(a) && (isAlHl(B) || B === NU)) return 0;
    // LB30a: break between RIs only after an even number of them
    if (A === RI && B === RI && ri % 2 === 1) return 0;
    // LB30b
    if (B === EM && (A === EB || flags[a]! & F_EPCN)) return 0;
    // LB31
    return 1;
  };

  for (b = 1; b < m; b++) {
    a = b - 1;
    A = cls[a]!;
    B = cls[b]!;
    if (A !== SP) nonSp = a;
    ri = A === RI ? ri + 1 : 0;
    numPrev = num;
    num = A === NU || ((A === SY || A === IS) && numPrev);
    out[start[b]!] = rule();
  }
  return out;
}
