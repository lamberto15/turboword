// SVG path data → PDF path operators (y flipped against the page height).

import { describe, expect, it } from 'vitest';
import { svgPathToPdf } from '../src/render/pdf/path.js';

const H = 100;
const ops = (d: string): string[] => svgPathToPdf(d, H).split('\n').filter(Boolean);
/** Numbers of one operator line, y converted back to SVG space. */
const pts = (line: string): number[] =>
  line
    .split(' ')
    .slice(0, -1)
    .map(Number)
    .map((v, i) => (i % 2 ? H - v : v));

describe('svgPathToPdf', () => {
  it('absolute move/line/close with y flipped', () => {
    expect(ops('M10 20L30 40Z')).toEqual(['10 80 m', '30 60 l', 'h']);
  });

  it('relative commands, H/V and implicit repeats', () => {
    expect(ops('m10 10 l 5 5 2 0 h 10 v -5 z')).toEqual([
      '10 90 m',
      '15 85 l',
      '17 85 l',
      '27 85 l',
      '27 90 l',
      'h',
    ]);
    expect(ops('M 0 0 10 10 20 0')).toEqual(['0 100 m', '10 90 l', '20 100 l']);
    expect(ops('m 1 1 2 2')).toEqual(['1 99 m', '3 97 l']);
  });

  it('compact number syntax', () => {
    expect(ops('M1.5.5L-1-2l1e1,0')).toEqual(['1.5 99.5 m', '-1 102 l', '9 102 l']);
  });

  it('current point returns to the subpath start after Z', () => {
    expect(ops('M 10 10 L 20 10 Z l 5 0')).toEqual(['10 90 m', '20 90 l', 'h', '15 90 l']);
  });

  it('cubic and smooth cubic (reflected control point)', () => {
    const o = ops('M0 0 C 0 10 10 10 10 0 S 20 -10 20 0');
    expect(pts(o[1]!)).toEqual([0, 10, 10, 10, 10, 0]);
    expect(pts(o[2]!)).toEqual([10, -10, 20, -10, 20, 0]);
    // S without a preceding curve uses the current point.
    expect(pts(ops('M0 0 S 5 5 10 0')[1]!)).toEqual([0, 0, 5, 5, 10, 0]);
  });

  it('quadratic and smooth quadratic become cubics', () => {
    expect(pts(ops('M0 0 Q 3 3 6 0')[1]!)).toEqual([2, 2, 4, 2, 6, 0]);
    const o = ops('M0 0 Q 5 10 10 0 T 20 0');
    expect(pts(o[2]!).map((v) => Math.round(v * 1000) / 1000)).toEqual([
      13.333, -6.667, 16.667, -6.667, 20, 0,
    ]);
    expect(pts(ops('m0 0 q 3 3 6 0 t 6 0')[2]!)).toEqual([8, -2, 10, -2, 12, 0]);
  });

  it('arcs become ≤ 90° cubic segments ending exactly at the target', () => {
    // Semicircle over the top, r = 50: two quarter segments through (50, 0).
    const o = ops('M 0 50 A 50 50 0 0 1 100 50');
    expect(o).toHaveLength(3);
    const a = pts(o[1]!);
    const b = pts(o[2]!);
    const k = 0.5523 * 50;
    expect(a.map((v) => Math.round(v))).toEqual([
      0,
      Math.round(50 - k),
      Math.round(50 - k),
      0,
      50,
      0,
    ]);
    expect(b.slice(4)).toEqual([100, 50]);
    // Large-arc flag picks the long way round (3 segments for 270°).
    expect(ops('M 50 0 A 50 50 0 1 0 100 50')).toHaveLength(4);
    // Relative arc, radii too small get scaled up; zero radius is a straight line.
    expect(pts(ops('M 0 0 a 1 1 0 0 0 10 0').at(-1)!).slice(-2)).toEqual([10, 0]);
    expect(ops('M 0 0 A 0 5 0 0 0 10 0')).toEqual(['0 100 m', '10 100 l']);
    // Flags may be written without separators.
    expect(ops('M0 0a5 5 0 1010 0').at(-1)!.endsWith('10 100 c')).toBe(true);
  });

  it('rotated elliptical arc reaches its end point', () => {
    const o = ops('M 10 10 A 30 10 45 0 1 50 40');
    expect(pts(o.at(-1)!).slice(-2)).toEqual([50, 40]);
  });

  it('stops at malformed data, keeping what came before', () => {
    expect(ops('M 0 0 L 10')).toEqual(['0 100 m']);
    expect(ops('10 10')).toEqual([]);
    expect(ops('M 0 0 Z 5 5')).toEqual(['0 100 m', 'h']);
  });
});
