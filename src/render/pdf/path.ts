// SVG path data → PDF path construction operators. Supports every command (M L H V C S Q T A Z,
// absolute and relative, implicit repeats), turns quadratic curves into cubics and elliptical arcs
// into at most 90° cubic segments (SVG 1.1 F.6.5 endpoint → centre conversion), and flips y from
// the PageModel's top-left origin to PDF's bottom-left. Parsing stops at the first malformed
// command, keeping what came before (as SVG renderers do).

import { num } from './writer.js';

const NUM = /[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;

/** PDF path operators (`m l c h`) for SVG path data `d`, on a page `pageHeight` points tall. */
export function svgPathToPdf(d: string, pageHeight: number): string {
  const out: string[] = [];
  const P = (x: number, y: number): string => `${num(x)} ${num(pageHeight - y)}`;
  let i = 0;
  const skip = (): void => {
    while (i < d.length && /[\s,]/.test(d[i]!)) i++;
  };
  const read = (): number | undefined => {
    skip();
    NUM.lastIndex = i;
    const m = NUM.exec(d);
    if (!m) return undefined;
    i = NUM.lastIndex;
    return Number.parseFloat(m[0]);
  };
  const flag = (): number | undefined => {
    skip();
    const c = d[i];
    if (c !== '0' && c !== '1') return undefined;
    i++;
    return c === '1' ? 1 : 0;
  };
  const nums = (n: number): number[] | undefined => {
    const r: number[] = [];
    for (let k = 0; k < n; k++) {
      const v = read();
      if (v === undefined) return undefined;
      r.push(v);
    }
    return r;
  };

  let cx = 0;
  let cy = 0;
  let sx = 0;
  let sy = 0;
  // Reflection points for S/T (the previous control point when the previous command matched).
  let qx = 0;
  let qy = 0;
  let c2x = 0;
  let c2y = 0;
  let prev = '';
  let cmd = '';
  const cubic = (x1: number, y1: number, x2: number, y2: number, x: number, y: number): void => {
    out.push(`${P(x1, y1)} ${P(x2, y2)} ${P(x, y)} c`);
    c2x = x2;
    c2y = y2;
    cx = x;
    cy = y;
  };
  const quad = (x1: number, y1: number, x: number, y: number): void => {
    cubic(
      cx + (2 / 3) * (x1 - cx),
      cy + (2 / 3) * (y1 - cy),
      x + (2 / 3) * (x1 - x),
      y + (2 / 3) * (y1 - y),
      x,
      y,
    );
    qx = x1;
    qy = y1;
  };

  for (;;) {
    skip();
    if (i >= d.length) break;
    const ch = d[i]!;
    if (/[MmLlHhVvCcSsQqTtAaZz]/.test(ch)) {
      cmd = ch;
      i++;
    } else if (!cmd || cmd === 'Z' || cmd === 'z') break;
    const rel = cmd === cmd.toLowerCase();
    const ox = rel ? cx : 0;
    const oy = rel ? cy : 0;
    const up = cmd.toUpperCase();
    if (up === 'Z') {
      out.push('h');
      cx = sx;
      cy = sy;
      prev = 'Z';
      continue;
    }
    if (up === 'M') {
      const a = nums(2);
      if (!a) break;
      cx = sx = ox + a[0]!;
      cy = sy = oy + a[1]!;
      out.push(`${P(cx, cy)} m`);
      cmd = rel ? 'l' : 'L'; // further pairs are implicit line-tos
    } else if (up === 'L' || up === 'H' || up === 'V') {
      const v = read();
      if (v === undefined) break;
      if (up === 'L') {
        const y = read();
        if (y === undefined) break;
        cx = ox + v;
        cy = oy + y;
      } else if (up === 'H') cx = ox + v;
      else cy = oy + v;
      out.push(`${P(cx, cy)} l`);
    } else if (up === 'C') {
      const a = nums(6);
      if (!a) break;
      cubic(ox + a[0]!, oy + a[1]!, ox + a[2]!, oy + a[3]!, ox + a[4]!, oy + a[5]!);
    } else if (up === 'S') {
      const a = nums(4);
      if (!a) break;
      const r = prev === 'C' || prev === 'S';
      cubic(
        r ? 2 * cx - c2x : cx,
        r ? 2 * cy - c2y : cy,
        ox + a[0]!,
        oy + a[1]!,
        ox + a[2]!,
        oy + a[3]!,
      );
    } else if (up === 'Q') {
      const a = nums(4);
      if (!a) break;
      quad(ox + a[0]!, oy + a[1]!, ox + a[2]!, oy + a[3]!);
    } else if (up === 'T') {
      const a = nums(2);
      if (!a) break;
      const r = prev === 'Q' || prev === 'T';
      quad(r ? 2 * cx - qx : cx, r ? 2 * cy - qy : cy, ox + a[0]!, oy + a[1]!);
    } else {
      const a = nums(3);
      const fa = flag();
      const fs = fa === undefined ? undefined : flag();
      const e = fs === undefined ? undefined : nums(2);
      if (!a || fa === undefined || fs === undefined || !e) break;
      arc(cx, cy, a[0]!, a[1]!, a[2]!, fa, fs, ox + e[0]!, oy + e[1]!, cubic, (x, y) => {
        out.push(`${P(x, y)} l`);
        cx = x;
        cy = y;
      });
      cx = ox + e[0]!;
      cy = oy + e[1]!;
    }
    prev = up;
  }
  return out.join('\n');
}

type Cubic = (x1: number, y1: number, x2: number, y2: number, x: number, y: number) => void;

/** Emits an SVG elliptical arc from (x1, y1) to (x2, y2) as cubic Béziers. */
function arc(
  x1: number,
  y1: number,
  rxIn: number,
  ryIn: number,
  angle: number,
  large: number,
  sweep: number,
  x2: number,
  y2: number,
  cubic: Cubic,
  line: (x: number, y: number) => void,
): void {
  if (x1 === x2 && y1 === y2) return;
  let rx = Math.abs(rxIn);
  let ry = Math.abs(ryIn);
  if (rx === 0 || ry === 0) {
    line(x2, y2);
    return;
  }
  const phi = (angle * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const dx = (x1 - x2) / 2;
  const dy = (y1 - y2) / 2;
  const xp = cos * dx + sin * dy;
  const yp = -sin * dx + cos * dy;
  const lambda = (xp * xp) / (rx * rx) + (yp * yp) / (ry * ry);
  if (lambda > 1) {
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  const rx2 = rx * rx;
  const ry2 = ry * ry;
  const den = rx2 * yp * yp + ry2 * xp * xp;
  const k = (large === sweep ? -1 : 1) * Math.sqrt(Math.max(0, (rx2 * ry2 - den) / (den || 1)));
  const cxp = (k * rx * yp) / ry;
  const cyp = (-k * ry * xp) / rx;
  const ccx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const ccy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const ang = (ux: number, uy: number, vx: number, vy: number): number =>
    Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
  const t1 = ang(1, 0, (xp - cxp) / rx, (yp - cyp) / ry);
  let dt = ang((xp - cxp) / rx, (yp - cyp) / ry, (-xp - cxp) / rx, (-yp - cyp) / ry);
  if (!sweep && dt > 0) dt -= 2 * Math.PI;
  else if (sweep && dt < 0) dt += 2 * Math.PI;
  const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
  const step = dt / n;
  const kappa = (4 / 3) * Math.tan(step / 4);
  const pt = (t: number): [number, number] => [
    ccx + rx * cos * Math.cos(t) - ry * sin * Math.sin(t),
    ccy + rx * sin * Math.cos(t) + ry * cos * Math.sin(t),
  ];
  const dpt = (t: number): [number, number] => [
    -rx * cos * Math.sin(t) - ry * sin * Math.cos(t),
    -rx * sin * Math.sin(t) + ry * cos * Math.cos(t),
  ];
  let a = t1;
  for (let s = 0; s < n; s++) {
    const b = a + step;
    const [ax, ay] = pt(a);
    const [bx, by] = s === n - 1 ? [x2, y2] : pt(b);
    const [dax, day] = dpt(a);
    const [dbx, dby] = dpt(b);
    cubic(ax + kappa * dax, ay + kappa * day, bx - kappa * dbx, by - kappa * dby, bx, by);
    a = b;
  }
}
