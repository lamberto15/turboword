// Fills and outlines: DrawingML a:spPr fill choices and a:ln (width, colour, dash, arrowheads)
// resolved to plain values, and the PathItems that draw a path with them.
//
// Approximations forced by the PageModel (no gradients, dashes on paths or transparency):
// gradients become the average of their stops, pattern fills the average of fore/background,
// dashes are materialised as separate stroke segments, alpha is blended toward white.

import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import type { DisplayItem, PathItem } from '../types.js';
import {
  averageColor,
  colorOf,
  colorRgbaOf,
  type GraphicsTheme,
  type RGBA,
  rgbaToHex,
} from './color.js';
import { dashPath, fmt, Path } from './path.js';

/** Arrowhead kinds of a:headEnd / a:tailEnd (and VML startarrow/endarrow). */
export type ArrowType = 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';

export interface ArrowSpec {
  type: ArrowType;
  /** Width and length as multiples of the line width (sm = 2, med = 3, lg = 5). */
  w: number;
  len: number;
}

/** A resolved outline. */
export interface LineStyle {
  color: string;
  /** Points. */
  width: number;
  /** Dash/gap lengths as multiples of the line width. */
  dash?: number[];
  head?: ArrowSpec;
  tail?: ArrowSpec;
}

/** Dash patterns of a:prstDash (multiples of the line width). */
export const PRESET_DASH: Record<string, number[] | undefined> = {
  solid: undefined,
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

const ARROW_SIZE: Record<string, number> = { sm: 2, med: 3, lg: 5 };

/** a:headEnd / a:tailEnd → ArrowSpec (undefined for none). */
export function arrowOf(el: XElement | undefined): ArrowSpec | undefined {
  const type = el?.attr('type');
  if (type === undefined || type === 'none') return undefined;
  const t = (['triangle', 'stealth', 'diamond', 'oval', 'arrow'] as const).find((k) => k === type);
  if (t === undefined) return undefined;
  return {
    type: t,
    w: ARROW_SIZE[el?.attr('w') ?? 'med'] ?? 3,
    len: ARROW_SIZE[el?.attr('len') ?? 'med'] ?? 3,
  };
}

/**
 * The fill of an element holding one of the DrawingML fill choices (spPr, a:bg, c:spPr, ...):
 * RRGGBB, `null` for an explicit a:noFill, `undefined` when none is given (use the style).
 * `group` is the fill a:grpFill inherits.
 */
export function fillOf(
  parent: XElement | undefined,
  theme: GraphicsTheme,
  phClr?: RGBA,
  group?: string | null,
): string | null | undefined {
  if (parent === undefined) return undefined;
  for (const el of parent.elements(NS.a)) {
    const local = el.name.slice(el.name.indexOf(':') + 1);
    switch (local) {
      case 'noFill':
        return null;
      case 'solidFill':
        return colorOf(el, theme, phClr) ?? '000000';
      case 'gradFill': {
        const stops = el
          .element(NS.a, 'gsLst')
          ?.elements(NS.a, 'gs')
          .map((gs) => colorRgbaOf(gs, theme, phClr))
          .filter((c): c is RGBA => c !== undefined);
        const avg = averageColor(stops ?? []);
        return avg === undefined ? undefined : rgbaToHex(avg);
      }
      case 'pattFill': {
        const fg = colorRgbaOf(el.element(NS.a, 'fgClr'), theme, phClr);
        const bg = colorRgbaOf(el.element(NS.a, 'bgClr'), theme, phClr);
        const avg = averageColor([fg, bg].filter((c): c is RGBA => c !== undefined));
        return avg === undefined ? undefined : rgbaToHex(avg);
      }
      case 'grpFill':
        return group;
      case 'blipFill':
        return undefined;
    }
  }
  return undefined;
}

/**
 * Resolves an a:ln over a style default (`base`, e.g. from the style matrix lnRef). Returns
 * `null` when the outline is off (a:noFill, or no colour anywhere).
 */
export function lineOf(
  ln: XElement | undefined,
  theme: GraphicsTheme,
  base?: LineStyle | null,
  phClr?: RGBA,
): LineStyle | null {
  let color = base?.color;
  let width = base?.width ?? 0.75;
  let dash = base?.dash;
  let head = base?.head;
  let tail = base?.tail;
  if (ln !== undefined) {
    const w = ln.attr('w');
    if (w !== undefined && Number.isFinite(Number(w))) width = Number(w) / 12700;
    const f = fillOf(ln, theme, phClr);
    if (f === null) return null;
    if (f !== undefined) color = f;
    const pd = ln.element(NS.a, 'prstDash')?.attr('val');
    if (pd !== undefined) dash = PRESET_DASH[pd];
    const cd = ln.element(NS.a, 'custDash');
    if (cd !== undefined) {
      const ds: number[] = [];
      for (const d of cd.elements(NS.a, 'ds')) {
        ds.push(Number(d.attr('d') ?? 0) / 100000, Number(d.attr('sp') ?? 0) / 100000);
      }
      dash = ds.length > 0 ? ds : undefined;
    }
    const he = ln.element(NS.a, 'headEnd');
    if (he !== undefined) head = arrowOf(he);
    const te = ln.element(NS.a, 'tailEnd');
    if (te !== undefined) tail = arrowOf(te);
  }
  if (color === undefined) return null;
  const out: LineStyle = { color, width };
  if (dash !== undefined) out.dash = dash;
  if (head !== undefined) out.head = head;
  if (tail !== undefined) out.tail = tail;
  return out;
}

/** Point and unit direction (pointing outward) at the start or end of a path. */
function endTangent(path: Path, atEnd: boolean): [number, number, number, number] | undefined {
  const subs = path.flatten();
  const sub = atEnd ? subs[subs.length - 1] : subs[0];
  if (sub === undefined || sub.pts.length < 4) return undefined;
  const p = sub.pts;
  const n = p.length / 2;
  const at = (i: number): [number, number] => [p[2 * i] ?? 0, p[2 * i + 1] ?? 0];
  const [x, y] = atEnd ? at(n - 1) : at(0);
  for (let k = 1; k < n; k++) {
    const [qx, qy] = atEnd ? at(n - 1 - k) : at(k);
    const d = Math.hypot(x - qx, y - qy);
    if (d > 1e-6) return [x, y, (x - qx) / d, (y - qy) / d];
  }
  return undefined;
}

/** The arrowhead outline at (x, y) pointing along (dx, dy); `open` heads are stroked only. */
function arrowPath(a: ArrowSpec, x: number, y: number, dx: number, dy: number, lw: number) {
  const width = Math.max(a.w * lw, 2.5);
  const len = Math.max(a.len * lw, 2.5);
  const nx = -dy;
  const ny = dx;
  const hw = width / 2;
  const pt = (along: number, side: number): [number, number] => [
    x - dx * along + nx * side,
    y - dy * along + ny * side,
  ];
  const p = new Path();
  switch (a.type) {
    case 'triangle':
      p.moveTo(x, y)
        .lineTo(...pt(len, hw))
        .lineTo(...pt(len, -hw))
        .close();
      return { path: p, open: false };
    case 'stealth':
      p.moveTo(x, y)
        .lineTo(...pt(len, hw))
        .lineTo(...pt(len * 0.7, 0))
        .lineTo(...pt(len, -hw))
        .close();
      return { path: p, open: false };
    case 'diamond':
      p.moveTo(...pt(-len / 2, 0))
        .lineTo(...pt(0, hw))
        .lineTo(...pt(len / 2, 0))
        .lineTo(...pt(0, -hw))
        .close();
      return { path: p, open: false };
    case 'oval': {
      const [cx, cy] = pt(0, 0);
      p.ellipse(cx - len / 2, cy - hw, len, width);
      return { path: p, open: false };
    }
    default:
      p.moveTo(...pt(len, hw))
        .lineTo(x, y)
        .lineTo(...pt(len, -hw));
      return { path: p, open: true };
  }
}

/** Stroke items for `path`: dashes materialised, arrowheads added. */
export function strokeItems(path: Path, line: LineStyle): DisplayItem[] {
  const out: DisplayItem[] = [];
  if (path.empty || !(line.width >= 0)) return out;
  const lw = line.width;
  const stroked =
    line.dash !== undefined
      ? dashPath(
          path,
          line.dash.map((d) => d * Math.max(lw, 0.5)),
        )
      : path;
  if (!stroked.empty) {
    out.push({ kind: 'path', d: stroked.toString(), stroke: line.color, strokeWidth: lw });
  }
  for (const [spec, atEnd] of [
    [line.head, false],
    [line.tail, true],
  ] as const) {
    if (spec === undefined) continue;
    const t = endTangent(path, atEnd);
    if (t === undefined) continue;
    const a = arrowPath(spec, t[0], t[1], t[2], t[3], lw);
    out.push(
      a.open
        ? { kind: 'path', d: a.path.toString(), stroke: line.color, strokeWidth: lw }
        : { kind: 'path', d: a.path.toString(), fill: line.color },
    );
  }
  return out;
}

/** Items drawing `path` filled with `fill` and outlined with `line` (either may be absent). */
export function paintPath(
  path: Path,
  fill: string | null | undefined,
  line: LineStyle | null | undefined,
  evenOdd = false,
): DisplayItem[] {
  if (path.empty) return [];
  const area = fill == null ? undefined : path.fillable();
  const simpleLine =
    line != null && line.dash === undefined && line.head === undefined && line.tail === undefined;
  if (area !== undefined && area === path && simpleLine && fill != null) {
    const item: PathItem = {
      kind: 'path',
      d: path.toString(),
      fill,
      stroke: line.color,
      strokeWidth: line.width,
    };
    if (evenOdd) item.evenOdd = true;
    return [item];
  }
  const out: DisplayItem[] = [];
  if (area !== undefined && fill != null && !area.empty) {
    const item: PathItem = { kind: 'path', d: area.toString(), fill };
    if (evenOdd) item.evenOdd = true;
    out.push(item);
  }
  if (line != null) out.push(...strokeItems(path, line));
  return out;
}

/** "M x y L ..." for a polyline (charts, metafiles). */
export function polylineData(pts: readonly number[], closed: boolean): string {
  let s = '';
  for (let i = 0; i + 1 < pts.length; i += 2) {
    s += `${i === 0 ? 'M' : ' L'} ${fmt(pts[i] ?? 0)} ${fmt(pts[i + 1] ?? 0)}`;
  }
  return closed && s !== '' ? `${s} Z` : s;
}
