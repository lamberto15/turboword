// UAX #14 conformance (LineBreakTest.txt, Unicode 16.0) plus behaviour checks that matter for
// paragraph layout. The test file is stored gzipped under test/fixtures/unicode (Unicode License
// v3, see LICENSE-UNICODE.txt there).

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import {
  isEastAsianWidth,
  lineBreakClass,
  lineBreakOpportunities,
} from '../src/layout/text/index.js';

const fixture = (name: string): string =>
  gunzipSync(
    readFileSync(fileURLToPath(new URL(`./fixtures/unicode/${name}.gz`, import.meta.url))),
  ).toString('utf8');

/** "× 0041 ÷ 0020 ×" → the string and, per code point after the first, whether a break precedes it. */
function parseCase(line: string): { text: string; breaks: boolean[]; offsets: number[] } {
  const tokens = line.trim().split(/\s+/);
  let text = '';
  const breaks: boolean[] = [];
  const offsets: number[] = [];
  for (let i = 1; i < tokens.length; i += 2) {
    offsets.push(text.length);
    text += String.fromCodePoint(Number.parseInt(tokens[i]!, 16));
    if (i > 1) breaks.push(tokens[i - 1] === '÷');
  }
  return { text, breaks, offsets };
}

/** Renders opportunities in the test file's notation, for readable failures. */
function render(text: string, ops: Uint8Array): string {
  let s = '×';
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    if (i > 0) s += ops[i] ? ' ÷' : ' ×';
    s += ` ${cp.toString(16).toUpperCase().padStart(4, '0')}`;
    i += cp > 0xffff ? 2 : 1;
  }
  return `${s} ÷`;
}

describe('LineBreakTest.txt (Unicode 16.0)', () => {
  const lines = fixture('LineBreakTest.txt')
    .split('\n')
    .map((l) => l.replace(/#.*/, '').trim())
    .filter(Boolean);

  it('passes every case', () => {
    expect(lines.length).toBeGreaterThan(15000);
    const failures: string[] = [];
    for (const line of lines) {
      const { text, breaks, offsets } = parseCase(line);
      const ops = lineBreakOpportunities(text);
      let ok = ops[0] === 0;
      for (let k = 0; k < breaks.length; k++) {
        if ((ops[offsets[k + 1]!]! !== 0) !== breaks[k]) ok = false;
      }
      if (!ok) failures.push(`expected ${line}\n     got ${render(text, ops)}`);
    }
    expect(failures.slice(0, 20).join('\n')).toBe('');
    expect(failures.length).toBe(0);
  });
});

describe('lineBreakOpportunities', () => {
  const breaksOf = (s: string): number[] => [...lineBreakOpportunities(s)];

  it('returns one entry per UTF-16 unit and never breaks at index 0', () => {
    expect(breaksOf('')).toEqual([]);
    expect(breaksOf('a')).toEqual([0]);
    expect(breaksOf('ab cd')).toEqual([0, 0, 0, 1, 0]);
  });

  it('marks mandatory breaks with 2 (BK, CR, LF, CR LF, NL)', () => {
    expect(breaksOf('a\nb')).toEqual([0, 0, 2]);
    expect(breaksOf('a\r\nb')).toEqual([0, 0, 0, 2]);
    expect(breaksOf('a\rb')).toEqual([0, 0, 2]);
    expect(breaksOf('a b')).toEqual([0, 0, 2]);
    expect(breaksOf('a\u0085b')).toEqual([0, 0, 2]);
  });

  it('keeps the low half of a surrogate pair at 0', () => {
    // U+1F600 is ID: breaks are allowed around it, never inside it.
    const ops = breaksOf('a\u{1F600}\u{1F600}b');
    expect(ops).toEqual([0, 1, 0, 1, 0, 1]);
  });

  it('breaks after hyphens and spaces but not before punctuation', () => {
    const s = 'well-known (x), y!';
    const ops = lineBreakOpportunities(s);
    expect(ops[s.indexOf('known')]).toBe(1);
    expect(ops[s.indexOf('(')]).toBe(1);
    expect(ops[s.indexOf(')')]).toBe(0);
    expect(ops[s.indexOf(',')]).toBe(0);
    expect(ops[s.indexOf('!')]).toBe(0);
  });

  it('honours no-break space, word joiner and zero-width space', () => {
    expect(breaksOf('a b')).toEqual([0, 0, 0]);
    expect(breaksOf('a ⁠b')).toEqual([0, 0, 0, 0]);
    expect(breaksOf('ab​cd')).toEqual([0, 0, 0, 1, 0]);
  });

  it('allows breaks between ideographs but not before small kana or closing marks', () => {
    const s = '日本語のテキスト、です。';
    const ops = lineBreakOpportunities(s);
    expect(ops[1]).toBe(1);
    expect(ops[s.indexOf('、')]).toBe(0);
    expect(ops[s.indexOf('。')]).toBe(0);
    expect(breaksOf('キャ')).toEqual([0, 0]); // ャ is CJ → NS
  });

  it('keeps numbers with their prefixes, suffixes and separators together', () => {
    for (const s of ['$(12.35)', '2,1234', '(12)¢', '12.54¢', '₹1,00,000.00', '-1/12', '50%']) {
      expect(breaksOf(s).every((v) => v === 0)).toBe(true);
    }
  });

  it('does not split combining sequences, emoji ZWJ sequences or flags', () => {
    expect(breaksOf('éx')).toEqual([0, 0, 0]);
    const family = '\u{1F468}‍\u{1F469}‍\u{1F467}';
    expect(breaksOf(family).every((v) => v === 0)).toBe(true);
    const flags = '\u{1F1EB}\u{1F1F7}\u{1F1E9}\u{1F1EA}'; // FR DE
    expect(breaksOf(flags)).toEqual([0, 0, 0, 0, 1, 0, 0, 0]);
  });

  it('is linear on long runs of spaces and digits', () => {
    const s = `(${' '.repeat(200_000)}x${'1'.repeat(200_000)}`;
    const t0 = performance.now();
    const ops = lineBreakOpportunities(s);
    expect(performance.now() - t0).toBeLessThan(2000);
    expect(ops[200_001]).toBe(0); // LB14: OP SP* ×
  });
});

describe('Line_Break data', () => {
  it('resolves LB1 classes', () => {
    expect(lineBreakClass(0x41)).toBe('AL');
    expect(lineBreakClass(0x20)).toBe('SP');
    expect(lineBreakClass(0x4e00)).toBe('ID');
    expect(lineBreakClass(0x0e01)).toBe('AL'); // SA (Thai letter) → AL
    expect(lineBreakClass(0x0e31)).toBe('CM'); // SA with gc=Mn → CM
    expect(lineBreakClass(0x30e3)).toBe('NS'); // CJ → NS
    expect(lineBreakClass(0x00a7)).toBe('AL'); // AI → AL
    expect(lineBreakClass(0xd800)).toBe('AL'); // SG → AL
    expect(lineBreakClass(0xe000)).toBe('AL'); // XX (private use) → AL
    expect(lineBreakClass(0x2fffd)).toBe('ID'); // unassigned plane 2 defaults to ID
    expect(lineBreakClass(0x10ffff)).toBe('AL');
  });

  it('knows East Asian width', () => {
    expect(isEastAsianWidth(0x41)).toBe(false);
    expect(isEastAsianWidth(0x4e00)).toBe(true);
    expect(isEastAsianWidth(0xff01)).toBe(true); // fullwidth
    expect(isEastAsianWidth(0xff61)).toBe(true); // halfwidth
    expect(isEastAsianWidth(0x20000)).toBe(true);
  });
});
