// The generated Unicode 16.0 tables (src/layout/text/data, scripts/gen-unicode.mjs): class
// numbering the runtime relies on, decoding, the Script property and script itemization.

import { describe, expect, it } from 'vitest';
import { BIDI_CLASSES } from '../src/layout/text/data/bidi.js';
import { LB_CLASSES } from '../src/layout/text/data/linebreak.js';
import { SCRIPT_CODES } from '../src/layout/text/data/scripts.js';
import { isDefaultIgnorable, scriptOf, scriptRuns } from '../src/layout/text/index.js';
import { decodeRuns, decodeVarints, lookup } from '../src/layout/text/table.js';

describe('generated tables', () => {
  it('keep the class numbering hard-coded in linebreak.ts and bidi.ts', () => {
    expect(LB_CLASSES).toBe(
      'BK CR LF NL SP ZW ZWJ CM WJ GL BA HY CL CP EX SY OP QU IS NS B2 CB BB HL AL NU PR PO ID EB ' +
        'EM IN JL JV JT H2 H3 RI AK AP AS VF VI',
    );
    expect(BIDI_CLASSES).toBe(
      'L R AL EN ES ET AN CS NSM BN B S WS ON LRE LRO RLE RLO PDF LRI RLI FSI PDI',
    );
    expect(SCRIPT_CODES.startsWith('ZzzzZyyyZinh')).toBe(true);
    expect(SCRIPT_CODES.length % 4).toBe(0);
    expect(SCRIPT_CODES.length / 4).toBeGreaterThan(160);
  });

  it('decode varints and runs', () => {
    // 'B' = 1; 'h' = 33 (continuation 1) then 'A' = 0 → 32.
    expect(decodeVarints('BhA')).toEqual([1, 32]);
    // One run of value 7 over all code points: palette index 0, length - 1 = 0x10FFFF.
    const n = 0x10ffff;
    const digits = [n >> 20, (n >> 15) & 31, (n >> 10) & 31, (n >> 5) & 31, n & 31];
    const abc = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
    const enc = `A${digits.map((d, i) => abc[i < 4 ? d + 32 : d]).join('')}`;
    const t = decodeRuns(enc, [7]);
    expect(lookup(t, 0)).toBe(7);
    expect(lookup(t, 0x10ffff)).toBe(7);
    expect(() => decodeRuns('AA', [7])).toThrow();
    expect(() => decodeVarints('!')).toThrow();
  });
});

describe('scriptOf', () => {
  it('returns ISO 15924 codes', () => {
    expect(scriptOf(0x41)).toBe('Latn');
    expect(scriptOf(0x3b1)).toBe('Grek');
    expect(scriptOf(0x430)).toBe('Cyrl');
    expect(scriptOf(0x5d0)).toBe('Hebr');
    expect(scriptOf(0x627)).toBe('Arab');
    expect(scriptOf(0x915)).toBe('Deva');
    expect(scriptOf(0xe01)).toBe('Thai');
    expect(scriptOf(0x4e00)).toBe('Hani');
    expect(scriptOf(0x3042)).toBe('Hira');
    expect(scriptOf(0x30a2)).toBe('Kana');
    expect(scriptOf(0xac00)).toBe('Hang');
    expect(scriptOf(0x20000)).toBe('Hani');
    expect(scriptOf(0x1f600)).toBe('Zyyy');
  });

  it('leaves Common, Inherited and Unknown unresolved', () => {
    expect(scriptOf(0x20)).toBe('Zyyy');
    expect(scriptOf(0x31)).toBe('Zyyy');
    expect(scriptOf(0x301)).toBe('Zinh');
    expect(scriptOf(0x378)).toBe('Zzzz');
    expect(scriptOf(0x10ffff)).toBe('Zzzz');
  });
});

describe('isDefaultIgnorable', () => {
  it('matches Default_Ignorable_Code_Point', () => {
    for (const cp of [0xad, 0x34f, 0x200b, 0x200d, 0x202e, 0x2066, 0xfe0f, 0xfeff, 0xe0001]) {
      expect(isDefaultIgnorable(cp)).toBe(true);
    }
    for (const cp of [0x20, 0x41, 0xa0, 0x2010, 0x300]) expect(isDefaultIgnorable(cp)).toBe(false);
  });
});

describe('scriptRuns', () => {
  const runs = (s: string): [string, string][] =>
    scriptRuns(s).map((r) => [s.slice(r.start, r.end), r.script]);

  it('merges Common and Inherited characters into their neighbours', () => {
    expect(runs('')).toEqual([]);
    expect(runs('Hello, world!')).toEqual([['Hello, world!', 'Latn']]);
    expect(runs('123 ')).toEqual([['123 ', 'Zyyy']]);
    expect(runs('  abc')).toEqual([['  abc', 'Latn']]);
    expect(runs('café שלום 42')).toEqual([
      ['café ', 'Latn'],
      ['שלום 42', 'Hebr'],
    ]);
  });

  it('splits at script changes and keeps surrogate pairs whole', () => {
    expect(runs('abc一\u{20000}def')).toEqual([
      ['abc', 'Latn'],
      ['一\u{20000}', 'Hani'],
      ['def', 'Latn'],
    ]);
  });

  it('gives a closing bracket the script of its opening bracket', () => {
    expect(runs('abc (אב) def')).toEqual([
      ['abc (', 'Latn'],
      ['אב', 'Hebr'],
      [') def', 'Latn'],
    ]);
    expect(runs('“привет” x')).toEqual([
      ['“привет” ', 'Cyrl'],
      ['x', 'Latn'],
    ]);
  });
});
