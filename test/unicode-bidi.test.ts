// UAX #9 conformance (BidiCharacterTest.txt and BidiTest.txt, Unicode 16.0) plus the helpers the
// line layout uses (visualOrder, visualRuns, lineLevels). Test files are stored gzipped under
// test/fixtures/unicode (Unicode License v3, see LICENSE-UNICODE.txt there).

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import {
  bidiClass,
  bidiLevels,
  lineLevels,
  mirrorOf,
  paragraphLevel,
  visualOrder,
  visualRuns,
} from '../src/layout/text/index.js';

const fixture = (name: string): string =>
  gunzipSync(
    readFileSync(fileURLToPath(new URL(`./fixtures/unicode/${name}.gz`, import.meta.url))),
  ).toString('utf8');

/** Per-code-point levels from per-UTF-16 levels. */
function cpLevels(text: string, levels: Uint8Array): number[] {
  const out: number[] = [];
  for (let i = 0; i < text.length; ) {
    out.push(levels[i]!);
    i += text.codePointAt(i)! > 0xffff ? 2 : 1;
  }
  return out;
}

/** Checks levels (ignoring 'x') and the visual order of the non-'x' characters. */
function check(
  text: string,
  rtl: boolean | 'auto',
  expectLevels: string[],
  expectOrder: string,
): string | null {
  const levels = cpLevels(text, bidiLevels(text, rtl));
  for (let k = 0; k < expectLevels.length; k++) {
    const e = expectLevels[k]!;
    if (e !== 'x' && Number(e) !== levels[k]) return `levels ${levels.join(' ')}`;
  }
  const order = visualOrder(Uint8Array.from(levels)).filter((k) => expectLevels[k] !== 'x');
  if (order.join(' ') !== expectOrder) return `order ${order.join(' ')}`;
  return null;
}

describe('BidiCharacterTest.txt (Unicode 16.0)', () => {
  it('passes every case (levels and order)', () => {
    const lines = fixture('BidiCharacterTest.txt')
      .split('\n')
      .filter((l) => l && !l.startsWith('#'));
    expect(lines.length).toBeGreaterThan(90000);
    const failures: string[] = [];
    for (const line of lines) {
      const [cps, dir, para, lv, order] = line.split(';');
      const text = String.fromCodePoint(...cps!.split(' ').map((h) => Number.parseInt(h, 16)));
      const rtl = dir === '0' ? false : dir === '1' ? true : 'auto';
      if (rtl === 'auto' && String(paragraphLevel(text)) !== para) {
        failures.push(`${line}\n  paragraph level ${paragraphLevel(text)}`);
        continue;
      }
      const err = check(text, rtl, lv!.trim().split(' '), order!.trim());
      if (err) failures.push(`${line}\n  got ${err}`);
    }
    expect(failures.slice(0, 10).join('\n')).toBe('');
    expect(failures.length).toBe(0);
  });
});

/** A representative character per Bidi_Class (none of them brackets). */
const SAMPLE: Record<string, number> = {
  L: 0x41,
  R: 0x5d0,
  AL: 0x627,
  EN: 0x30,
  ES: 0x2b,
  ET: 0x24,
  AN: 0x660,
  CS: 0x2c,
  NSM: 0x300,
  BN: 0xad,
  B: 0x2029,
  S: 0x09,
  WS: 0x20,
  ON: 0x21,
  LRE: 0x202a,
  RLE: 0x202b,
  PDF: 0x202c,
  LRO: 0x202d,
  RLO: 0x202e,
  LRI: 0x2066,
  RLI: 0x2067,
  FSI: 0x2068,
  PDI: 0x2069,
};

describe('BidiTest.txt (Unicode 16.0)', () => {
  it('passes every case (levels and order)', () => {
    let levels: string[] = [];
    let order = '';
    let cases = 0;
    const failures: string[] = [];
    for (const raw of fixture('BidiTest.txt').split('\n')) {
      const line = raw.replace(/#.*/, '').trim();
      if (!line) continue;
      if (line.startsWith('@Levels:')) {
        levels = line.slice(8).trim().split(/\s+/);
        continue;
      }
      if (line.startsWith('@Reorder:')) {
        order = line.slice(9).trim().split(/\s+/).join(' ');
        continue;
      }
      if (line.startsWith('@')) continue;
      const [classes, bits] = line.split(';');
      const text = String.fromCodePoint(
        ...classes!
          .trim()
          .split(/\s+/)
          .map((c) => SAMPLE[c]!),
      );
      const set = Number(bits);
      for (const [bit, rtl] of [
        [1, 'auto'],
        [2, false],
        [4, true],
      ] as const) {
        if (!(set & bit)) continue;
        cases++;
        const err = check(text, rtl, levels, order);
        if (err) failures.push(`${classes} (${rtl}) @Levels ${levels.join(' ')}: got ${err}`);
      }
    }
    expect(cases).toBeGreaterThan(700000);
    expect(failures.slice(0, 10).join('\n')).toBe('');
    expect(failures.length).toBe(0);
  });
});

describe('bidi helpers', () => {
  it('resolves mixed Hebrew/Latin with numbers', () => {
    const text = 'abc אב 123 ג';
    const levels = [...bidiLevels(text, false)];
    expect(levels).toEqual([0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 1, 1]);
    expect(visualOrder(Uint8Array.from(levels))).toEqual([0, 1, 2, 3, 11, 10, 7, 8, 9, 6, 5, 4]);
  });

  it('keeps surrogate pairs at one level', () => {
    const text = 'א\u{10900}x'; // Phoenician is R
    expect([...bidiLevels(text, false)]).toEqual([1, 1, 1, 0]);
  });

  it('splits paragraphs at paragraph separators and detects each base direction', () => {
    const text = 'א a b ב';
    expect([...bidiLevels(text, 'auto')]).toEqual([1, 1, 2, 1, 0, 0, 1]);
    expect(paragraphLevel('א a')).toBe(1);
    expect(paragraphLevel('a א')).toBe(0);
    expect(paragraphLevel('⁧א⁩ a')).toBe(0); // isolates are skipped (P2)
    expect(paragraphLevel('123')).toBe(0);
  });

  it('applies L1 per line with lineLevels', () => {
    const text = 'אב גד ';
    const levels = bidiLevels(text, false);
    // As one line, the trailing space is at the paragraph level; the inner one is not.
    expect([...levels]).toEqual([1, 1, 1, 1, 1, 0]);
    // Broken after the first space, that space ends the line and goes back to level 0.
    expect([...lineLevels(text, levels, 0, 3, 0)]).toEqual([1, 1, 0]);
    expect([...lineLevels(text, levels, 3, 6, 0)]).toEqual([1, 1, 0]);
    // Trailing characters removed by X9 (here a supplementary BN) reset with both halves.
    const tagged = '\u05d0\u05d1\u{E0001}\u05d2';
    const taggedLevels = bidiLevels(tagged, false);
    expect([...taggedLevels]).toEqual([1, 1, 1, 1, 1]);
    expect([...lineLevels(tagged, taggedLevels, 0, 4, 0)]).toEqual([1, 1, 0, 0]);
  });

  it('orders level runs visually', () => {
    const levels = Uint8Array.from([0, 0, 1, 1, 2, 2, 1, 0]);
    expect(visualRuns(levels)).toEqual([
      { start: 0, end: 2, level: 0 },
      { start: 6, end: 7, level: 1 },
      { start: 4, end: 6, level: 2 },
      { start: 2, end: 4, level: 1 },
      { start: 7, end: 8, level: 0 },
    ]);
    expect(visualRuns(new Uint8Array(0))).toEqual([]);
  });

  it('exposes Bidi_Class and mirroring data', () => {
    expect(bidiClass(0x41)).toBe('L');
    expect(bidiClass(0x5d0)).toBe('R');
    expect(bidiClass(0x627)).toBe('AL');
    expect(bidiClass(0x300)).toBe('NSM');
    expect(bidiClass(0x5ff)).toBe('R'); // unassigned in the Hebrew block defaults to R
    expect(bidiClass(0x20cf)).toBe('ET'); // unassigned currency symbol defaults to ET
    expect(mirrorOf(0x28)).toBe(0x29);
    expect(mirrorOf(0x29)).toBe(0x28);
    expect(mirrorOf(0x2264)).toBe(0x2265);
    expect(mirrorOf(0x41)).toBeUndefined();
  });
});
