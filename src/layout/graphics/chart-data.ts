// Chart model: reads a c:chartSpace into plain data for chart.ts — chart groups with their series
// (cached values only: c:numCache/c:strCache/c:numLit/c:strLit/c:multiLvlStrCache; the embedded
// workbook is never opened), axes, text styles (c:txPr / c:rich run properties over the chart
// space defaults) and number formatting for labels.

import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import { colorOf, type GraphicsTheme } from './color.js';

/** Chart group kinds drawn (3D variants are drawn as their 2D twins). */
export type GroupKind = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar';

export interface Series {
  el: XElement;
  idx: number;
  order: number;
  name: string;
  /** Values (y for scatter); undefined = blank cell. */
  values: (number | undefined)[];
  /** Scatter x values. */
  xs?: (number | undefined)[];
  cats: string[];
  /** formatCode of a numeric category cache (dates). */
  catFormat?: string;
  /** numCache formatCode of the values. */
  format?: string;
  spPr?: XElement;
  /** c:marker/c:symbol when given. */
  symbol?: string;
  markerSize?: number;
  markerSpPr?: XElement;
  smooth: boolean;
  /** c:dPt by point index. */
  points: Map<number, XElement>;
  explosion: number;
}

export interface ChartGroup {
  kind: GroupKind;
  el: XElement;
  /** bar: 'col' (vertical bars) or 'bar' (horizontal). */
  dir: 'col' | 'bar';
  grouping: string;
  varyColors: boolean;
  series: Series[];
  axIds: string[];
  gapWidth: number;
  overlap: number;
  holeSize: number;
  firstSliceAng: number;
  /** lineChart/c:marker: series show markers by default. */
  markers: boolean;
  style: string;
  hiLowLines?: XElement;
}

export interface AxisDef {
  id: string;
  kind: 'cat' | 'val' | 'date' | 'ser';
  el: XElement;
  pos: string;
  deleted: boolean;
  crossAx?: string;
  min?: number;
  max?: number;
  majorUnit?: number;
  reverse: boolean;
  gridlines?: XElement;
  minorGridlines?: XElement;
  tickLblPos: string;
  majorTick: string;
  format?: string;
  sourceLinked: boolean;
  crosses: string;
  crossesAt?: number;
  crossBetween?: string;
  title?: XElement;
}

/** Resolved text formatting of one chart text element. */
export interface ChartText {
  font: string;
  size: number;
  bold: boolean;
  italic: boolean;
  color: string;
}

const c = (el: XElement | undefined, local: string): XElement | undefined =>
  el?.element(NS.c, local);
const val = (el: XElement | undefined, local: string): string | undefined =>
  c(el, local)?.attr('val');
const num = (el: XElement | undefined, local: string, dflt: number): number => {
  const v = Number(val(el, local));
  return Number.isFinite(v) && val(el, local) !== undefined ? v : dflt;
};
/** ST_Boolean: absent element → `dflt`; present without val → true. */
export function flag(el: XElement | undefined, local: string, dflt: boolean): boolean {
  const e = c(el, local);
  if (e === undefined) return dflt;
  const v = e.attr('val');
  return v === undefined || v === '1' || v === 'true';
}

/** Points of a data reference (c:cat / c:val / c:tx / c:xVal / c:yVal). */
export function cachePoints(ref: XElement | undefined): {
  pts: (string | undefined)[];
  format?: string;
} {
  if (ref === undefined) return { pts: [] };
  const v = c(ref, 'v');
  if (v !== undefined) return { pts: [v.text] };
  const src = c(ref, 'numRef') ?? c(ref, 'strRef') ?? c(ref, 'multiLvlStrRef');
  const cache =
    c(src, 'numCache') ??
    c(src, 'strCache') ??
    c(c(src, 'multiLvlStrCache'), 'lvl') ??
    c(ref, 'numLit') ??
    c(ref, 'strLit');
  if (cache === undefined) return { pts: [] };
  const count = num(cache, 'ptCount', 0);
  const pts: (string | undefined)[] = [];
  for (const pt of cache.elements(NS.c, 'pt')) {
    const i = Number(pt.attr('idx'));
    if (!Number.isInteger(i) || i < 0 || i > 100000) continue;
    pts[i] = c(pt, 'v')?.text;
  }
  // multiLvlStrCache: c:lvl has the points directly, ptCount sits on the cache.
  const total = Math.min(100000, Math.max(count, num(cache.parent as XElement, 'ptCount', 0)));
  while (pts.length < total) pts.push(undefined);
  for (let i = 0; i < pts.length; i++) if (!(i in pts)) pts[i] = undefined;
  const format = c(cache, 'formatCode')?.text;
  return format === undefined ? { pts } : { pts, format };
}

const toNum = (s: string | undefined): number | undefined => {
  if (s === undefined || s.trim() === '') return undefined;
  const n = Number(s);
  return Number.isFinite(n) ? n : undefined;
};

function readSeries(ser: XElement, i: number, scatter: boolean): Series {
  const vals = cachePoints(c(ser, scatter ? 'yVal' : 'val'));
  const catRef = c(ser, scatter ? 'xVal' : 'cat');
  const cats = cachePoints(catRef);
  const marker = c(ser, 'marker');
  const s: Series = {
    el: ser,
    idx: num(ser, 'idx', i),
    order: num(ser, 'order', i),
    name: cachePoints(c(ser, 'tx')).pts[0] ?? `Series ${i + 1}`,
    values: vals.pts.map(toNum),
    cats: cats.pts.map((p) => p ?? ''),
    smooth: flag(ser, 'smooth', false),
    points: new Map(),
    explosion: num(ser, 'explosion', 0),
  };
  if (scatter) {
    const xs = cats.pts.map(toNum);
    s.xs = xs.some((x) => x === undefined) ? cats.pts.map((_, k) => k + 1) : xs;
  }
  if (vals.format !== undefined) s.format = vals.format;
  if (cats.format !== undefined) s.catFormat = cats.format;
  const spPr = c(ser, 'spPr');
  if (spPr !== undefined) s.spPr = spPr;
  if (marker !== undefined) {
    const sym = val(marker, 'symbol');
    if (sym !== undefined) s.symbol = sym;
    const size = val(marker, 'size');
    if (size !== undefined) s.markerSize = Number(size) || 5;
    const mp = c(marker, 'spPr');
    if (mp !== undefined) s.markerSpPr = mp;
  }
  for (const dpt of ser.elements(NS.c, 'dPt')) {
    const k = num(dpt, 'idx', -1);
    if (k >= 0) s.points.set(k, dpt);
  }
  return s;
}

const GROUPS: Readonly<Record<string, GroupKind>> = {
  barChart: 'bar',
  bar3DChart: 'bar',
  lineChart: 'line',
  line3DChart: 'line',
  stockChart: 'line',
  areaChart: 'area',
  area3DChart: 'area',
  pieChart: 'pie',
  pie3DChart: 'pie',
  ofPieChart: 'pie',
  doughnutChart: 'doughnut',
  scatterChart: 'scatter',
  radarChart: 'radar',
};

/** The supported chart groups of a c:plotArea, in document order. */
export function readGroups(plotArea: XElement): ChartGroup[] {
  const out: ChartGroup[] = [];
  for (const el of plotArea.elements(NS.c)) {
    const local = el.name.slice(el.name.indexOf(':') + 1);
    const kind = GROUPS[local];
    if (kind === undefined) continue;
    const scatter = kind === 'scatter';
    const series = el.elements(NS.c, 'ser').map((s, i) => readSeries(s, i, scatter));
    series.sort((a, b) => a.order - b.order);
    const grouping = val(el, 'grouping') ?? (kind === 'bar' ? 'clustered' : 'standard');
    const stacked = grouping === 'stacked' || grouping === 'percentStacked';
    const g: ChartGroup = {
      kind,
      el,
      dir: val(el, 'barDir') === 'bar' ? 'bar' : 'col',
      grouping,
      varyColors: flag(el, 'varyColors', kind === 'pie' || kind === 'doughnut'),
      series,
      axIds: el.elements(NS.c, 'axId').map((a) => a.attr('val') ?? ''),
      gapWidth: num(el, 'gapWidth', 150),
      overlap: num(el, 'overlap', stacked ? 100 : 0),
      holeSize: num(el, 'holeSize', 50),
      firstSliceAng: num(el, 'firstSliceAng', 0),
      markers: local === 'stockChart' ? false : flag(el, 'marker', kind !== 'area'),
      style: val(el, 'scatterStyle') ?? val(el, 'radarStyle') ?? '',
    };
    const hl = c(el, 'hiLowLines');
    if (hl !== undefined) g.hiLowLines = hl;
    out.push(g);
  }
  return out;
}

/** The axes of a c:plotArea by axId. */
export function readAxes(plotArea: XElement): Map<string, AxisDef> {
  const out = new Map<string, AxisDef>();
  for (const el of plotArea.elements(NS.c)) {
    const local = el.name.slice(el.name.indexOf(':') + 1);
    const kind =
      local === 'catAx'
        ? 'cat'
        : local === 'valAx'
          ? 'val'
          : local === 'dateAx'
            ? 'date'
            : local === 'serAx'
              ? 'ser'
              : undefined;
    if (kind === undefined) continue;
    const scaling = c(el, 'scaling');
    const nf = c(el, 'numFmt');
    const a: AxisDef = {
      id: val(el, 'axId') ?? '',
      kind,
      el,
      pos: val(el, 'axPos') ?? 'b',
      deleted: flag(el, 'delete', false),
      reverse: val(scaling, 'orientation') === 'maxMin',
      tickLblPos: val(el, 'tickLblPos') ?? 'nextTo',
      majorTick: val(el, 'majorTickMark') ?? 'cross',
      sourceLinked: nf?.attr('sourceLinked') === '1',
      crosses: val(el, 'crosses') ?? 'autoZero',
    };
    const opt = (k: string): number | undefined => {
      const v = Number(val(k === 'majorUnit' ? el : scaling, k));
      return Number.isFinite(v) && val(k === 'majorUnit' ? el : scaling, k) !== undefined
        ? v
        : undefined;
    };
    const mn = opt('min');
    const mx = opt('max');
    const mu = opt('majorUnit');
    if (mn !== undefined) a.min = mn;
    if (mx !== undefined) a.max = mx;
    if (mu !== undefined && mu > 0) a.majorUnit = mu;
    const cross = val(el, 'crossAx');
    if (cross !== undefined) a.crossAx = cross;
    const ca = Number(val(el, 'crossesAt'));
    if (val(el, 'crossesAt') !== undefined && Number.isFinite(ca)) a.crossesAt = ca;
    const cb = val(el, 'crossBetween');
    if (cb !== undefined) a.crossBetween = cb;
    const fmt = nf?.attr('formatCode');
    if (fmt !== undefined) a.format = fmt;
    const grid = c(el, 'majorGridlines');
    if (grid !== undefined) a.gridlines = grid;
    const minor = c(el, 'minorGridlines');
    if (minor !== undefined) a.minorGridlines = minor;
    const title = c(el, 'title');
    if (title !== undefined) a.title = title;
    out.set(a.id, a);
  }
  return out;
}

// ---- text ---------------------------------------------------------------------------------

/** Context for resolving chart text styles. */
export interface TextCtx {
  theme: GraphicsTheme;
  /** c:chartSpace/c:txPr default run properties. */
  base?: XElement;
}

/** First a:defRPr of a c:txPr or c:rich body. */
function defRPr(body: XElement | undefined): XElement | undefined {
  return body?.element(NS.a, 'p')?.element(NS.a, 'pPr')?.element(NS.a, 'defRPr');
}

function applyRPr(t: ChartText, rPr: XElement | undefined, ctx: TextCtx): void {
  if (rPr === undefined) return;
  const sz = Number(rPr.attr('sz'));
  if (rPr.attr('sz') !== undefined && sz > 0) t.size = sz / 100;
  const b = rPr.attr('b');
  if (b !== undefined) t.bold = b === '1' || b === 'true';
  const i = rPr.attr('i');
  if (i !== undefined) t.italic = i === '1' || i === 'true';
  const color = colorOf(rPr.element(NS.a, 'solidFill'), ctx.theme);
  if (color !== undefined) t.color = color;
  const face = rPr.element(NS.a, 'latin')?.attr('typeface');
  if (face !== undefined && face !== '') {
    const f = face.startsWith('+mj')
      ? ctx.theme.fonts.major
      : face.startsWith('+mn')
        ? ctx.theme.fonts.minor
        : face;
    if (f !== undefined) t.font = f;
  }
}

/**
 * Text style of a chart element owning c:txPr or c:tx/c:rich (title, axis, legend, dLbls):
 * role defaults ← chart space c:txPr ← own c:txPr ← rich text paragraph/run properties.
 */
export function textStyle(owner: XElement | undefined, ctx: TextCtx, title = false): ChartText {
  const t: ChartText = {
    font: ctx.theme.fonts.minor ?? 'Calibri',
    size: 10,
    bold: false,
    italic: false,
    color: '595959',
  };
  applyRPr(t, ctx.base, ctx);
  if (title) t.size = Math.max(t.size * 1.4, 14);
  applyRPr(t, defRPr(c(owner, 'txPr')), ctx);
  const rich = c(c(owner, 'tx'), 'rich');
  applyRPr(t, defRPr(rich), ctx);
  applyRPr(t, rich?.element(NS.a, 'p')?.element(NS.a, 'r')?.element(NS.a, 'rPr'), ctx);
  return t;
}

/** Lines of rich text (c:tx/c:rich a:p), or of a c:strRef cache; undefined when none. */
export function richLines(owner: XElement | undefined): string[] | undefined {
  const tx = c(owner, 'tx');
  if (tx === undefined) return undefined;
  const rich = c(tx, 'rich');
  if (rich === undefined) {
    const p = cachePoints(tx).pts[0];
    return p === undefined ? undefined : [p];
  }
  const lines: string[] = [];
  for (const p of rich.elements(NS.a, 'p')) {
    let s = '';
    for (const r of p.elements(NS.a)) {
      if (r.is(NS.a, 'r') || r.is(NS.a, 'fld')) s += r.element(NS.a, 't')?.text ?? '';
      else if (r.is(NS.a, 'br')) {
        lines.push(s);
        s = '';
      }
    }
    lines.push(s);
  }
  return lines;
}

// ---- numbers ------------------------------------------------------------------------------

/** Formats a value with a (lightly interpreted) number format code. */
export function formatValue(v: number, code: string | undefined): string {
  const f = (code ?? 'General').split(';')[0] ?? 'General';
  const dec = /\.(0+)/.exec(f)?.[1]?.length ?? 0;
  let s: string;
  if (f.includes('%')) s = `${(v * 100).toFixed(dec)}%`;
  else if (/0/.test(f) && !/general/i.test(f)) s = v.toFixed(dec);
  else s = String(Number(v.toPrecision(10)));
  if (f.includes('#,##')) {
    const [int = '', frac] = s.split('.');
    s = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (frac === undefined ? '' : `.${frac}`);
  }
  if (/\$/.test(f) && !s.includes('$')) s = s.startsWith('-') ? `-$${s.slice(1)}` : `$${s}`;
  return s;
}

/** A "nice" axis scale: min/max snapped to multiples of a 1/2/5 × 10^k step. */
export function niceScale(
  lo: number,
  hi: number,
  count: number,
  fixed: { min?: number; max?: number; unit?: number },
  includeZero: boolean,
): { min: number; max: number; step: number } {
  let a = Number.isFinite(lo) ? lo : 0;
  let b = Number.isFinite(hi) ? hi : 1;
  if (includeZero) {
    if (a > 0 && (fixed.min === undefined || fixed.min <= 0) && a < (b * 5) / 6) a = 0;
    if (b < 0 && b > (a * 5) / 6) b = 0;
  }
  if (fixed.min !== undefined) a = fixed.min;
  if (fixed.max !== undefined) b = fixed.max;
  if (b <= a) b = a === 0 ? 1 : a + Math.abs(a) * 0.1;
  const span = b - a;
  let step = fixed.unit ?? 0;
  if (!(step > 0)) {
    const rough = (span * (fixed.max === undefined ? 1.05 : 1)) / Math.max(1, count);
    const mag = 10 ** Math.floor(Math.log10(rough));
    const r = rough / mag;
    step = (r <= 1 ? 1 : r <= 2 ? 2 : r <= 2.5 ? 2.5 : r <= 5 ? 5 : 10) * mag;
  }
  const min = fixed.min ?? Math.floor(a / step + 1e-9) * step;
  let max = fixed.max ?? Math.ceil(b / step - 1e-9) * step;
  if (fixed.max === undefined && max - b < span * 0.02 && b > 0 && b !== 0) max += step;
  if (max <= min) max = min + step;
  // Bound the tick count (a tiny fixed majorUnit must not create thousands of ticks).
  if ((max - min) / step > 200) step = (max - min) / 200;
  return { min, max, step };
}
