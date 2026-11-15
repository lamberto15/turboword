// Decoder and lookup for the generated Unicode range tables (src/layout/text/data/*.ts, written by
// scripts/gen-unicode.mjs). A table is the run-length list of a property's values over
// U+0000..U+10FFFF, packed as base64url 5-bit varints. It is decoded once, on first use, into a
// sorted Uint32Array of run starts plus the run values, then binary-searched; code points below
// DIRECT get a flat array so Latin/Greek/Cyrillic/Hebrew/Arabic text never searches.

const DIRECT = 0x800;

/** Base64url digit values; 32..63 are continuation digits (another 5-bit group follows). */
const DIGIT = /* @__PURE__ */ (() => {
  const t = new Int8Array(128).fill(-1);
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
  for (let i = 0; i < 64; i++) t[alphabet.charCodeAt(i)] = i;
  return t;
})();

/** Decodes a varint string into its integers. */
export function decodeVarints(data: string): number[] {
  const out: number[] = [];
  let n = 0;
  for (let i = 0; i < data.length; i++) {
    const d = DIGIT[data.charCodeAt(i)] ?? -1;
    if (d < 0) throw new Error('bad Unicode table data');
    if (d >= 32) n = n * 32 + (d - 32);
    else {
      out.push(n * 32 + d);
      n = 0;
    }
  }
  return out;
}

/** A decoded property table. */
export interface RangeTable {
  /** Run start code points, ascending; starts[0] = 0. */
  readonly starts: Uint32Array;
  /** Value of each run. */
  readonly values: Uint16Array;
  /** Values of U+0000..DIRECT-1. */
  readonly direct: Uint16Array;
}

/** Decodes runs (`palette index, length - 1` varint pairs) with `palette` mapping index → value. */
export function decodeRuns(runs: string, palette: readonly number[] | null): RangeTable {
  const nums = decodeVarints(runs);
  const n = nums.length >> 1;
  const starts = new Uint32Array(n);
  const values = new Uint16Array(n);
  const direct = new Uint16Array(DIRECT);
  let cp = 0;
  for (let i = 0; i < n; i++) {
    const idx = nums[2 * i]!;
    const v = palette ? palette[idx]! : idx;
    const len = nums[2 * i + 1]! + 1;
    starts[i] = cp;
    values[i] = v;
    if (cp < DIRECT) direct.fill(v, cp, Math.min(cp + len, DIRECT));
    cp += len;
  }
  if (cp !== 0x110000) throw new Error('bad Unicode table data');
  return { starts, values, direct };
}

/** The value of code point `cp` in `t` (out-of-range values are clamped to U+0000..U+10FFFF). */
export function lookup(t: RangeTable, cp: number): number {
  if (cp < DIRECT) return t.direct[Math.max(cp, 0)]!;
  const { starts, values } = t;
  let lo = 0;
  let hi = starts.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (starts[mid]! <= cp) lo = mid;
    else hi = mid - 1;
  }
  return values[lo]!;
}

/** Wraps a table builder so it runs once, on first use (keeps module load free and tree-shakable). */
export function lazy<T>(build: () => T): () => T {
  let value: T | undefined;
  return () => {
    if (value === undefined) value = build();
    return value;
  };
}
