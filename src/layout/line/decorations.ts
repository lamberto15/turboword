// Text decorations (underline, strike-through, highlight/shading) and paragraph borders/shading as
// display items.
//
// Underlines and strike-through lines use the run font's own metrics (post.underlinePosition /
// underlineThickness, OS/2 yStrikeoutPosition / yStrikeoutSize) at the run's size. The ~18
// `w:u` styles map onto solid, dashed and dotted lines, double lines and wavy paths. Highlight
// and character shading are rectangles over the font's ascent+descent behind the text.
//
// Paragraph borders sit outside the text box: the left/right ones `space` points beyond the
// indents, the top/bottom ones above/below the lines (layoutParagraph adds their space and width
// to the first/last line heights). Paragraph shading fills the box between the borders.

import type { BorderSpec, Borders } from '../../resolve/types.js';
import type { DisplayItem, Font, LineItem } from '../types.js';

/** ST_HighlightColor → RRGGBB. */
export const HIGHLIGHT: Readonly<Record<string, string>> = {
  black: '000000',
  blue: '0000FF',
  cyan: '00FFFF',
  green: '00FF00',
  magenta: 'FF00FF',
  red: 'FF0000',
  yellow: 'FFFF00',
  white: 'FFFFFF',
  darkBlue: '000080',
  darkCyan: '008080',
  darkGreen: '008000',
  darkMagenta: '800080',
  darkRed: '800000',
  darkYellow: '808000',
  darkGray: '808080',
  lightGray: 'C0C0C0',
};

/** A colour value (RRGGBB) or black for `auto`/invalid values. */
export function colorOf(c: string | undefined, fallback = '000000'): string {
  return c !== undefined && /^[0-9A-Fa-f]{6}$/.test(c) ? c.toUpperCase() : fallback;
}

function unit(font: Font, v: number, size: number): number {
  const upm = font.metrics.unitsPerEm > 0 ? font.metrics.unitsPerEm : 1000;
  return (v * size) / upm;
}

/**
 * Underline items for `[x1, x2]` at a baseline (y = 0) for style `style` (w:u values).
 * `font`/`size` give the position and thickness; `rise` shifts with superscript/subscript text.
 */
export function underlineItems(
  style: string,
  x1: number,
  x2: number,
  font: Font,
  size: number,
  rise: number,
  color: string,
): DisplayItem[] {
  if (style === 'none' || x2 <= x1) return [];
  const m = font.metrics;
  let t = unit(font, m.underlineThickness || m.unitsPerEm / 20, size);
  t = Math.max(t, 0.25);
  const y = unit(font, -(m.underlinePosition || -m.unitsPerEm / 10), size) - rise + t / 2;
  const heavy = /heavy|thick/i.test(style);
  const w = heavy ? t * 2 : t;
  const line = (yy: number, dash?: number[]): LineItem => {
    const l: LineItem = { kind: 'line', x1, y1: yy, x2, y2: yy, color, width: w };
    if (dash !== undefined) l.dash = dash;
    return l;
  };
  switch (style) {
    case 'double':
      return [line(y), line(y + 2 * w)];
    case 'dotted':
    case 'dottedHeavy':
      return [line(y, [w, w])];
    case 'dash':
    case 'dashedHeavy':
      return [line(y, [4 * w, 2 * w])];
    case 'dashLong':
    case 'dashLongHeavy':
      return [line(y, [8 * w, 3 * w])];
    case 'dotDash':
    case 'dashDotHeavy':
      return [line(y, [4 * w, 2 * w, w, 2 * w])];
    case 'dotDotDash':
    case 'dashDotDotHeavy':
      return [line(y, [4 * w, 2 * w, w, 2 * w, w, 2 * w])];
    case 'wave':
    case 'wavyHeavy':
    case 'wavyDouble': {
      const out: DisplayItem[] = [wave(x1, x2, y + w, w, color)];
      if (style === 'wavyDouble') out.push(wave(x1, x2, y + 3 * w, w, color));
      return out;
    }
    default:
      return [line(y)];
  }
}

function wave(x1: number, x2: number, y: number, w: number, color: string): DisplayItem {
  const step = Math.max(1.5, w * 3);
  const amp = Math.max(0.6, w);
  let d = `M ${r(x1)} ${r(y)}`;
  let up = true;
  for (let x = x1; x < x2; x += step) {
    const nx = Math.min(x2, x + step);
    d += ` Q ${r((x + nx) / 2)} ${r(y + (up ? -amp : amp))} ${r(nx)} ${r(y)}`;
    up = !up;
  }
  return { kind: 'path', d, stroke: color, strokeWidth: w };
}

function r(v: number): number {
  return Math.round(v * 1000) / 1000;
}

/** Strike-through items (single or double) for `[x1, x2]`. */
export function strikeItems(
  double: boolean,
  x1: number,
  x2: number,
  font: Font,
  size: number,
  rise: number,
  color: string,
): DisplayItem[] {
  if (x2 <= x1) return [];
  const m = font.metrics;
  const t = Math.max(0.25, unit(font, m.strikeoutSize || m.unitsPerEm / 20, size));
  const pos = m.strikeoutPosition || m.xHeight / 2 || m.unitsPerEm * 0.25;
  const y = -unit(font, pos, size) - rise + t / 2;
  if (!double) return [{ kind: 'line', x1, y1: y, x2, y2: y, color, width: t }];
  return [
    { kind: 'line', x1, y1: y - t, x2, y2: y - t, color, width: t },
    { kind: 'line', x1, y1: y + t, x2, y2: y + t, color, width: t },
  ];
}

/** Width of a border line, points (`w:sz` is in eighths of a point). */
export function borderWidth(b: BorderSpec | undefined): number {
  return b === undefined ? 0 : Math.max(0.25, b.size / 8);
}

/** Space plus width a border adds outside the text, points. */
export function borderPad(b: BorderSpec | undefined): number {
  return b === undefined ? 0 : b.space + borderWidth(b);
}

function borderLine(
  b: BorderSpec,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  nx: number,
  ny: number,
): DisplayItem[] {
  const w = borderWidth(b);
  const color = colorOf(b.color);
  const mk = (dx: number, dy: number, width: number, dash?: number[]): LineItem => {
    const l: LineItem = {
      kind: 'line',
      x1: x1 + dx,
      y1: y1 + dy,
      x2: x2 + dx,
      y2: y2 + dy,
      color,
      width,
    };
    if (dash !== undefined) l.dash = dash;
    return l;
  };
  switch (b.style) {
    case 'double':
      // Two thin lines filling the border width.
      return [mk((-nx * w) / 3, (-ny * w) / 3, w / 3), mk((nx * w) / 3, (ny * w) / 3, w / 3)];
    case 'dotted':
      return [mk(0, 0, w, [w, w])];
    case 'dashed':
    case 'dashSmallGap':
      return [mk(0, 0, w, [3 * w, b.style === 'dashed' ? 2 * w : w])];
    case 'dotDash':
      return [mk(0, 0, w, [3 * w, w, w, w])];
    case 'dotDotDash':
      return [mk(0, 0, w, [3 * w, w, w, w, w, w])];
    default:
      return [mk(0, 0, w)];
  }
}

/**
 * Paragraph shading and borders for a box spanning `[left, right]` horizontally (the text area:
 * indents) and `height` points vertically from y = 0 (the top of the first line, which already
 * includes the top border's space). `top`/`bottom` false omit those borders (a paragraph split
 * across pages, or grouped paragraphs with identical borders).
 */
export function paragraphDecorations(
  borders: Borders,
  shading: string | undefined,
  left: number,
  right: number,
  height: number,
  o: { top?: boolean; bottom?: boolean } = {},
): DisplayItem[] {
  const out: DisplayItem[] = [];
  const top = o.top !== false ? borders.top : undefined;
  const bottom = o.bottom !== false ? borders.bottom : undefined;
  const l = borders.left;
  const rt = borders.right;
  const x1 = left - (l !== undefined ? l.space : 0);
  const x2 = right + (rt !== undefined ? rt.space : 0);
  const fill =
    shading !== undefined && /^[0-9A-Fa-f]{6}$/.test(shading) ? shading.toUpperCase() : undefined;
  if (fill !== undefined) {
    const y1 = top !== undefined ? borderWidth(top) : 0;
    const y2 = height - (bottom !== undefined ? borderWidth(bottom) : 0);
    out.push({
      kind: 'rect',
      x: x1,
      y: y1,
      width: Math.max(0, x2 - x1),
      height: Math.max(0, y2 - y1),
      fill,
    });
  }
  const outerL = x1 - borderWidth(l);
  const outerR = x2 + borderWidth(rt);
  if (top !== undefined) {
    const y = borderWidth(top) / 2;
    out.push(...borderLine(top, outerL, y, outerR, y, 0, -1));
  }
  if (bottom !== undefined) {
    const y = height - borderWidth(bottom) / 2;
    out.push(...borderLine(bottom, outerL, y, outerR, y, 0, 1));
  }
  if (l !== undefined) {
    const x = x1 - borderWidth(l) / 2;
    out.push(...borderLine(l, x, 0, x, height, -1, 0));
  }
  if (rt !== undefined) {
    const x = x2 + borderWidth(rt) / 2;
    out.push(...borderLine(rt, x, 0, x, height, 1, 0));
  }
  return out;
}
