// Shared drawing helpers for charts: scales, the drawing context (text placement, paint,
// automatic series colours), series/point paint, markers, polylines and data labels.

import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import type { LayoutEnv } from '../engine.js';
import type { DisplayItem } from '../types.js';
import {
  type ChartGroup,
  type ChartText,
  flag,
  formatValue,
  type GroupKind,
  type Series,
  type TextCtx,
  textStyle,
} from './chart-data.js';
import { type GraphicsTheme, shadeHex } from './color.js';
import { fillOf, type LineStyle, lineOf, paintPath } from './paint.js';
import { Path } from './path.js';
import { offsetItems, runProps, type ShapedText, shapeText } from './text.js';

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** pos(v) = a + (v - min) / (max - min) · (b - a). */
export interface Scale {
  min: number;
  max: number;
  a: number;
  b: number;
  step: number;
}

export const pos = (s: Scale, v: number): number =>
  s.a + ((v - s.min) / (s.max - s.min || 1)) * (s.b - s.a);
export const clamp = (v: number, lo: number, hi: number): number => Math.max(lo, Math.min(hi, v));

export const PAD = 7;
export const AUTO_SYMBOLS: readonly string[] = [
  'diamond',
  'square',
  'triangle',
  'x',
  'star',
  'circle',
  'plus',
  'dot',
  'dash',
];

/** Drawing state of one chart. */
export class Ctx {
  readonly out: DisplayItem[] = [];
  readonly #lh = new Map<string, number>();
  constructor(
    readonly env: Pick<LayoutEnv, 'fonts' | 'shaper'>,
    readonly theme: GraphicsTheme,
    readonly tc: TextCtx,
    readonly blanks: string,
  ) {}

  text(s: string, t: ChartText): ShapedText {
    return shapeText(s, runProps(t), this.env, t.color);
  }

  /** Places shaped text: x by `align`, y as top / middle / baseline. */
  put(st: ShapedText, x: number, y: number, align: 'l' | 'c' | 'r', v: 't' | 'm' | 'b'): void {
    const dx = align === 'l' ? x : align === 'c' ? x - st.width / 2 : x - st.width;
    const dy = v === 'b' ? y : v === 't' ? y + st.ascent : y + (st.ascent - st.descent) / 2;
    this.out.push(...offsetItems(st.items, dx, dy));
  }

  lineHeight(t: ChartText): number {
    const key = `${t.font}|${t.size}|${t.bold}|${t.italic}`;
    let h = this.#lh.get(key);
    if (h === undefined) {
      const st = this.text('Xg', t);
      h = st.ascent + st.descent;
      this.#lh.set(key, h);
    }
    return h;
  }

  paint(path: Path, fill: string | null | undefined, line: LineStyle | null | undefined): void {
    this.out.push(...paintPath(path, fill, line));
  }

  autoColor(i: number): string {
    const base = this.theme.colors[`accent${(i % 6) + 1}`] ?? '4472C4';
    const k = Math.floor(i / 6) % 3;
    return k === 0 ? base : shadeHex(base, k === 1 ? 'darkenLess' : 'lightenLess');
  }
}

export const isLineKind = (k: GroupKind): boolean =>
  k === 'line' || k === 'scatter' || k === 'radar';

/** Fill and outline of a series (or of one of its points). */
export function paintOf(
  ctx: Ctx,
  s: Series,
  kind: GroupKind,
  auto: string,
  pt?: number,
): { fill: string | null; line: LineStyle | null } {
  const dpt = pt === undefined ? undefined : s.points.get(pt)?.element(NS.c, 'spPr');
  const lineKind = isLineKind(kind);
  let fill = fillOf(dpt, ctx.theme);
  if (fill === undefined) fill = fillOf(s.spPr, ctx.theme);
  if (fill === undefined) fill = lineKind ? null : auto;
  const base: LineStyle | null = lineKind
    ? { color: auto, width: 2.25 }
    : kind === 'pie' || kind === 'doughnut'
      ? { color: 'FFFFFF', width: 0.75 }
      : null;
  let line = lineOf(s.spPr?.element(NS.a, 'ln'), ctx.theme, base);
  const dln = dpt?.element(NS.a, 'ln');
  if (dln !== undefined) line = lineOf(dln, ctx.theme, line);
  return { fill, line };
}

/** Marker shape centred on (x, y). */
export function marker(
  ctx: Ctx,
  symbol: string,
  size: number,
  x: number,
  y: number,
  fill: string | null,
  line: LineStyle | null,
): void {
  const r = clamp(size, 2, 72) / 2;
  const p = new Path();
  let open = false;
  switch (symbol) {
    case 'circle':
      p.ellipse(x - r, y - r, 2 * r, 2 * r);
      break;
    case 'dot':
      p.ellipse(x - r / 2, y - r / 2, r, r);
      break;
    case 'square':
      p.rect(x - r, y - r, 2 * r, 2 * r);
      break;
    case 'diamond':
      p.poly([x, y - r, x + r, y, x, y + r, x - r, y], true);
      break;
    case 'triangle':
      p.poly([x, y - r, x + r, y + r, x - r, y + r], true);
      break;
    case 'dash':
      p.rect(x - r, y - r / 4, 2 * r, r / 2);
      break;
    case 'x':
    case 'star':
    case 'plus':
      open = true;
      if (symbol !== 'plus')
        p.moveTo(x - r, y - r)
          .lineTo(x + r, y + r)
          .moveTo(x + r, y - r)
          .lineTo(x - r, y + r);
      if (symbol !== 'x') p.moveTo(x, y - r).lineTo(x, y + r);
      if (symbol === 'plus') p.moveTo(x - r, y).lineTo(x + r, y);
      break;
    default:
      return;
  }
  if (open) {
    const color = line?.color ?? fill ?? '000000';
    ctx.paint(p, null, { color, width: Math.max(0.75, line?.width ?? 0.75) });
  } else
    ctx.paint(
      p,
      fill,
      line === null ? null : { color: line.color, width: Math.min(line.width, 1) },
    );
}

/** Marker symbol and paint of a series in a line/scatter/radar group ('none' = no markers). */
/** A series marker: symbol ('none' = no markers), size in points and paint. */
export interface MarkerSpec {
  symbol: string;
  size: number;
  fill: string | null;
  line: LineStyle | null;
}

export function markerOf(ctx: Ctx, s: Series, dflt: boolean, auto: string): MarkerSpec {
  let symbol = s.symbol ?? (dflt ? 'auto' : 'none');
  if (symbol === 'auto') symbol = AUTO_SYMBOLS[s.idx % AUTO_SYMBOLS.length] ?? 'circle';
  let fill = fillOf(s.markerSpPr, ctx.theme);
  if (fill === undefined) fill = auto;
  const line = lineOf(s.markerSpPr?.element(NS.a, 'ln'), ctx.theme, { color: auto, width: 0.75 });
  return { symbol, size: s.markerSize ?? 5, fill, line };
}

/** Polyline (or Catmull-Rom smoothed curve) through points. */
export function linePath(p: Path, pts: [number, number][], smooth: boolean): void {
  const first = pts[0];
  if (first === undefined) return;
  p.moveTo(first[0], first[1]);
  for (let i = 1; i < pts.length; i++) {
    const cur = pts[i] as [number, number];
    if (!smooth) {
      p.lineTo(cur[0], cur[1]);
      continue;
    }
    const p0 = pts[i - 2] ?? pts[i - 1] ?? cur;
    const p1 = pts[i - 1] ?? cur;
    const p3 = pts[i + 1] ?? cur;
    p.cubicTo(
      p1[0] + (cur[0] - p0[0]) / 6,
      p1[1] + (cur[1] - p0[1]) / 6,
      cur[0] - (p3[0] - p1[0]) / 6,
      cur[1] - (p3[1] - p1[1]) / 6,
      cur[0],
      cur[1],
    );
  }
}

/** Splits a series' points at blanks per c:dispBlanksAs. */
export function runs(vals: (number | undefined)[], blanks: string): [number, number][][] {
  const out: [number, number][][] = [];
  let cur: [number, number][] = [];
  vals.forEach((v, i) => {
    if (v === undefined || !Number.isFinite(v)) {
      if (blanks === 'zero') cur.push([i, 0]);
      else if (blanks === 'gap' && cur.length > 0) {
        out.push(cur);
        cur = [];
      }
      return;
    }
    cur.push([i, v]);
  });
  if (cur.length > 0) out.push(cur);
  return out;
}

// ---- data labels --------------------------------------------------------------------------

export interface LabelOpts {
  val: boolean;
  pct: boolean;
  cat: boolean;
  ser: boolean;
  pos: string;
  style: ChartText;
}

export function labelOpts(ctx: Ctx, g: ChartGroup, s: Series): LabelOpts | undefined {
  const d = s.el.element(NS.c, 'dLbls') ?? g.el.element(NS.c, 'dLbls');
  if (d === undefined || flag(d, 'delete', false)) return undefined;
  const o: LabelOpts = {
    val: flag(d, 'showVal', false),
    pct: flag(d, 'showPercent', false),
    cat: flag(d, 'showCatName', false),
    ser: flag(d, 'showSerName', false),
    pos: d.element(NS.c, 'dLblPos')?.attr('val') ?? '',
    style: textStyle(d, ctx.tc),
  };
  return o.val || o.pct || o.cat || o.ser ? o : undefined;
}

export function labelText(o: LabelOpts, s: Series, i: number, v: number, total: number): string {
  const parts: string[] = [];
  if (o.ser) parts.push(s.name);
  if (o.cat) parts.push(s.cats[i] ?? String(i + 1));
  if (o.val) parts.push(formatValue(v, s.format));
  if (o.pct) parts.push(`${Math.round((v / (total || 1)) * 100)}%`);
  return parts.join(', ');
}

/** Rounds a tick value to the precision of its step (no 0.30000000000000004 labels). */
export function round(v: number, step: number): number {
  const dec = Math.max(0, Math.min(10, -Math.floor(Math.log10(step)) + 1));
  return Number(v.toFixed(dec));
}

/** Date serial → text for a date format code (d, dd, m, mm, mmm, yy, yyyy). */
export function formatDate(serial: number, code: string): string {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000));
  const months = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ];
  const f = code.replace(/\[[^\]]*\]|"/g, '').split(';')[0] ?? '';
  return f.replace(/yyyy|yy|mmmm|mmm|mm|m|dd|d/gi, (tok) => {
    const t = tok.toLowerCase();
    if (t === 'yyyy') return String(d.getUTCFullYear());
    if (t === 'yy') return String(d.getUTCFullYear() % 100).padStart(2, '0');
    if (t === 'mmm' || t === 'mmmm') return months[d.getUTCMonth()] ?? '';
    if (t === 'mm') return String(d.getUTCMonth() + 1).padStart(2, '0');
    if (t === 'm') return String(d.getUTCMonth() + 1);
    if (t === 'dd') return String(d.getUTCDate()).padStart(2, '0');
    return String(d.getUTCDate());
  });
}

export function catLabel(s: string, format: string | undefined): string {
  const v = Number(s);
  if (format === undefined || s.trim() === '' || !Number.isFinite(v)) return s;
  if (/[dy]/i.test(format) && /m/i.test(format) && !/general/i.test(format))
    return formatDate(v, format);
  return /general/i.test(format) ? s : formatValue(v, format);
}

/** Manual plot-area layout (fractions of the chart box). */
export function manualLayout(
  plotArea: XElement,
  w: number,
  h: number,
): { r: Rect; inner: boolean } | undefined {
  const m = plotArea.element(NS.c, 'layout')?.element(NS.c, 'manualLayout');
  if (m === undefined) return undefined;
  const v = (k: string): number | undefined => {
    const s = m.element(NS.c, k)?.attr('val');
    const n = Number(s);
    return s === undefined || !Number.isFinite(n) ? undefined : n;
  };
  const x = v('x');
  const y = v('y');
  const fw = v('w');
  const fh = v('h');
  if (x === undefined || y === undefined || fw === undefined || fh === undefined) return undefined;
  const edge = (k: string): boolean => m.element(NS.c, k)?.attr('val') === 'edge';
  const rw = edge('wMode') ? fw - x : fw;
  const rh = edge('hMode') ? fh - y : fh;
  const r = {
    x: clamp(x, 0, 1) * w,
    y: clamp(y, 0, 1) * h,
    w: clamp(rw, 0, 1) * w,
    h: clamp(rh, 0, 1) * h,
  };
  if (r.w < 4 || r.h < 4) return undefined;
  return { r, inner: m.element(NS.c, 'layoutTarget')?.attr('val') === 'inner' };
}
