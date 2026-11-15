// Script property (Scripts.txt, Unicode 16.0) and script itemization.
//
// scriptOf is the raw property: Common (Zyyy), Inherited (Zinh) and Unknown (Zzzz) come back as
// such. scriptRuns does the usual resolution for font selection and shaping (as ICU's
// UScriptRun and HarfBuzz do): Inherited and Common characters join the run around them, leading
// ones join the first real script, and a closing bracket takes the script of its opening bracket
// so "(text)" never splits its parentheses across fonts.

import { pairedBracket } from './bidi.js';
import { IGNORABLE_PALETTE, IGNORABLE_RUNS, SCRIPT_CODES, SCRIPT_RUNS } from './data/scripts.js';
import { decodeRuns, decodeVarints, lazy, lookup, type RangeTable } from './table.js';

const scripts: () => RangeTable = lazy(() => decodeRuns(SCRIPT_RUNS, null));
const ignorables: () => RangeTable = lazy(() =>
  decodeRuns(IGNORABLE_RUNS, decodeVarints(IGNORABLE_PALETTE)),
);

// SCRIPT_CODES starts with Unknown, Common, Inherited (scripts/gen-unicode.mjs keeps them first).
const ZZZZ = 0;
const ZYYY = 1;
const ZINH = 2;

const code = (i: number): string => SCRIPT_CODES.slice(4 * i, 4 * i + 4);

/**
 * ISO 15924 code of a code point's Script property: Latn, Arab, Hebr, Hani, Hira, Kana, Hang,
 * Thai, Deva, ..., Zyyy (Common), Zinh (Inherited) or Zzzz (Unknown).
 */
export function scriptOf(codePoint: number): string {
  return code(lookup(scripts(), codePoint));
}

/** Default_Ignorable_Code_Point (ZWJ, ZWNJ, soft hyphen, bidi controls, variation selectors, ...). */
export function isDefaultIgnorable(codePoint: number): boolean {
  return lookup(ignorables(), codePoint) === 1;
}

/** A run of one script, as returned by scriptRuns. */
export interface ScriptRun {
  /** UTF-16 range [start, end). */
  start: number;
  end: number;
  /** ISO 15924 code; Zyyy only when the whole text is Common/Inherited/Unknown. */
  script: string;
}

/** Opening quotation marks that pair like brackets (they are not Bidi_Paired_Bracket). */
const QUOTES = /* @__PURE__ */ new Map([
  [0x2018, 0x2019],
  [0x201c, 0x201d],
  [0xab, 0xbb],
  [0x2039, 0x203a],
]);

/**
 * Splits `text` into runs of one script, Common and Inherited characters merged into their
 * neighbours (see the file header). Runs cover the text without gaps.
 */
export function scriptRuns(text: string): ScriptRun[] {
  const t = scripts();
  const runs: ScriptRun[] = [];
  let runStart = 0;
  let current = -1; // script index of the open run; -1 while only Common has been seen
  // Open brackets awaiting their closing pair: [closing code point, script index at the opening].
  const stack: [number, number][] = [];
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    const len = cp > 0xffff ? 2 : 1;
    let s = lookup(t, cp);
    if (s === ZINH || s === ZZZZ) s = ZYYY;
    if (s === ZYYY) {
      const close = QUOTES.get(cp) ?? pairedBracket(cp);
      if (close > 0) {
        if (stack.length < 64) stack.push([close, current]);
      } else {
        for (let j = stack.length - 1; j >= 0; j--) {
          if (stack[j]![0] !== cp) continue;
          s = stack[j]![1];
          stack.length = j;
          break;
        }
      }
    }
    if (s !== ZYYY && s >= 0 && s !== current) {
      if (current >= 0) {
        runs.push({ start: runStart, end: i, script: code(current) });
        runStart = i;
      }
      // Brackets opened while the script was still unknown belong to this one.
      for (const entry of stack) if (entry[1] < 0) entry[1] = s;
      current = s;
    }
    i += len;
  }
  if (text.length > 0) {
    runs.push({ start: runStart, end: text.length, script: code(current < 0 ? ZYYY : current) });
  }
  return runs;
}
