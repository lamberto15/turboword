// DrawingML charts drawn from their cached values (c:chartSpace): chart area, title, legend,
// plot area with value/category axes (nice ticks, gridlines, tick marks, labels) and the chart
// groups — bar/column (clustered, stacked, percent), line (markers, smoothing, high-low lines for
// stock charts), area, pie/doughnut, scatter and radar. Bar + line combos share the category axis
// and map each group to its own value axis (secondary axes included).
//
// Approximations: 3D charts are drawn flat; text cannot rotate in the PageModel, so vertical axis
// titles are drawn horizontally above the axis and crowded category labels are wrapped or
// skipped instead of rotated; date axes format serial dates lightly; logarithmic scales are linear.

import type { ResolvedRunProps } from '../../resolve/types.js';
import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import type { LayoutEnv } from '../engine.js';
import type { DisplayItem } from '../types.js';
import {
  type AxisDef,
  type ChartGroup,
  type ChartText,
  flag,
  formatValue,
  niceScale,
  readAxes,
  readGroups,
  richLines,
  type TextCtx,
  textStyle,
} from './chart-data.js';
import {
  Ctx,
  catLabel,
  clamp,
  isLineKind,
  labelOpts,
  labelText,
  manualLayout,
  marker,
  markerOf,
  PAD,
  paintOf,
  type Rect,
  round,
} from './chart-draw.js';
import { drawCartesian } from './chart-plot.js';
import { type GraphicsTheme, graphicsTheme } from './color.js';
import { fillOf, type LineStyle, lineOf, paintPath } from './paint.js';
import { Path } from './path.js';
import { runProps, type ShapedText } from './text.js';

/** The c:chartSpace a c:chart reference points at (undefined when missing or not a chart). */
function chartSpace(chartRef: XElement, env: Pick<LayoutEnv, 'partOf'>): XElement | undefined {
  const id = chartRef.attrNS(NS.r, 'id');
  if (id === undefined) return undefined;
  try {
    const root = env.partOf(chartRef)?.related(id)?.xml().root;
    return root?.is(NS.c, 'chartSpace') === true ? root : undefined;
  } catch {
    return undefined;
  }
}

function textCtx(space: XElement, theme: GraphicsTheme): TextCtx {
  const base = space
    .element(NS.c, 'txPr')
    ?.element(NS.a, 'p')
    ?.element(NS.a, 'pPr')
    ?.element(NS.a, 'defRPr');
  return base === undefined ? { theme } : { theme, base };
}

// ---- title & legend -----------------------------------------------------------------------

function drawTitle(ctx: Ctx, chart: XElement, groups: ChartGroup[], area: Rect): Rect {
  const title = chart.element(NS.c, 'title');
  const series = groups.flatMap((g) => g.series);
  let lines = title === undefined ? undefined : richLines(title);
  if (lines === undefined) {
    if (title !== undefined)
      lines = [series.length === 1 ? (series[0]?.name ?? '') : 'Chart Title'];
    else if (!flag(chart, 'autoTitleDeleted', false) && series.length === 1) {
      lines = [series[0]?.name ?? ''];
    }
  }
  if (lines === undefined || lines.every((l) => l === '')) return area;
  const style = textStyle(title, ctx.tc, true);
  const lh = ctx.lineHeight(style);
  let y = area.y;
  for (const l of lines) {
    ctx.put(ctx.text(l, style), area.x + area.w / 2, y, 'c', 't');
    y += lh;
  }
  if (title !== undefined && flag(title, 'overlay', false)) return area;
  const used = y - area.y + 6;
  return { x: area.x, y: area.y + used, w: area.w, h: Math.max(10, area.h - used) };
}

interface LegendEntry {
  label: string;
  fill: string | null;
  line: LineStyle | null;
  isLine: boolean;
  symbol: string;
}

function legendEntries(ctx: Ctx, groups: ChartGroup[]): LegendEntry[] {
  const g0 = groups[0];
  const s0 = g0?.series[0];
  const byPoint =
    groups.length === 1 &&
    g0 !== undefined &&
    s0 !== undefined &&
    (g0.kind === 'pie' ||
      g0.kind === 'doughnut' ||
      (g0.varyColors && g0.series.length === 1 && !isLineKind(g0.kind)));
  const out: LegendEntry[] = [];
  if (byPoint) {
    const n = Math.max(s0.cats.length, s0.values.length);
    for (let i = 0; i < n; i++) {
      const p = paintOf(ctx, s0, g0.kind, ctx.autoColor(g0.varyColors ? i : s0.idx), i);
      out.push({ label: s0.cats[i] ?? String(i + 1), ...p, isLine: false, symbol: 'none' });
    }
    return out;
  }
  for (const g of groups) {
    for (const s of g.series) {
      const auto = ctx.autoColor(s.idx);
      const p = paintOf(ctx, s, g.kind, auto);
      const lineKind = isLineKind(g.kind) && g.style !== 'filled';
      const m = lineKind ? markerOf(ctx, s, g.kind === 'scatter' || g.markers, auto) : undefined;
      out.push({ label: s.name, ...p, isLine: lineKind, symbol: m?.symbol ?? 'none' });
    }
  }
  return out;
}

function drawLegend(ctx: Ctx, chart: XElement, groups: ChartGroup[], area: Rect): Rect {
  const legend = chart.element(NS.c, 'legend');
  if (legend === undefined) return area;
  const deleted = new Set<number>();
  for (const e of legend.elements(NS.c, 'legendEntry')) {
    if (flag(e, 'delete', false)) deleted.add(Number(e.element(NS.c, 'idx')?.attr('val')));
  }
  const entries = legendEntries(ctx, groups).filter((_, i) => !deleted.has(i));
  if (entries.length === 0) return area;
  const where = legend.element(NS.c, 'legendPos')?.attr('val') ?? 'r';
  const style = textStyle(legend, ctx.tc);
  const lh = ctx.lineHeight(style);
  const sw = (e: LegendEntry): number => (e.isLine ? 14 : 7);
  const shaped = entries.map((e) => ctx.text(e.label, style));
  const widths = entries.map((e, i) => sw(e) + 4 + (shaped[i]?.width ?? 0));
  const overlay = flag(legend, 'overlay', false);
  const drawEntry = (i: number, x: number, y: number): void => {
    const e = entries[i] as LegendEntry;
    const cy = y + lh / 2;
    if (e.isLine) {
      if (e.line !== null)
        ctx.paint(new Path().moveTo(x, cy).lineTo(x + 14, cy), null, {
          color: e.line.color,
          width: Math.min(e.line.width, 2.25),
        });
      const color = e.line?.color ?? e.fill ?? '000000';
      marker(ctx, e.symbol, 5, x + 7, cy, color, { color, width: 0.75 });
    } else ctx.paint(new Path().rect(x, cy - 3.5, 7, 7), e.fill ?? 'FFFFFF', e.line);
    ctx.put(shaped[i] as ShapedText, x + sw(e) + 4, cy, 'l', 'm');
  };
  if (where === 'b' || where === 't') {
    const rows: number[][] = [[]];
    let rowW = 0;
    widths.forEach((w, i) => {
      const row = rows[rows.length - 1] as number[];
      if (row.length > 0 && rowW + 9 + w > area.w) {
        rows.push([i]);
        rowW = w;
      } else {
        rowW += (row.length > 0 ? 9 : 0) + w;
        row.push(i);
      }
    });
    const h = rows.length * lh;
    let y = where === 'b' ? area.y + area.h - h : area.y;
    for (const row of rows) {
      const total = row.reduce((a, i) => a + (widths[i] ?? 0), 0) + 9 * (row.length - 1);
      let x = area.x + (area.w - total) / 2;
      for (const i of row) {
        drawEntry(i, x, y);
        x += (widths[i] ?? 0) + 9;
      }
      y += lh;
    }
    if (overlay) return area;
    return where === 'b'
      ? { x: area.x, y: area.y, w: area.w, h: Math.max(10, area.h - h - 6) }
      : { x: area.x, y: area.y + h + 6, w: area.w, h: Math.max(10, area.h - h - 6) };
  }
  const w = Math.min(Math.max(...widths), area.w / 2);
  const h = entries.length * (lh + 2);
  const x = where === 'l' ? area.x : area.x + area.w - w;
  let y = where === 'tr' ? area.y : area.y + Math.max(0, (area.h - h) / 2);
  for (let i = 0; i < entries.length; i++) {
    drawEntry(i, x, y);
    y += lh + 2;
  }
  if (overlay) return area;
  return where === 'l'
    ? { x: area.x + w + 8, y: area.y, w: Math.max(10, area.w - w - 8), h: area.h }
    : { x: area.x, y: area.y, w: Math.max(10, area.w - w - 8), h: area.h };
}

// ---- pie / doughnut -----------------------------------------------------------------------

function drawPie(ctx: Ctx, g: ChartGroup, area: Rect): void {
  const rings = g.kind === 'pie' ? g.series.slice(0, 1) : g.series;
  if (rings.length === 0) return;
  const cx = area.x + area.w / 2;
  const cy = area.y + area.h / 2;
  let maxExp = 0;
  for (const s of rings) {
    maxExp = Math.max(maxExp, s.explosion);
    for (const d of s.points.values()) {
      maxExp = Math.max(maxExp, Number(d.element(NS.c, 'explosion')?.attr('val') ?? 0) || 0);
    }
  }
  const r = Math.max(1, Math.min(area.w, area.h) / 2 - 2) / (1 + clamp(maxExp, 0, 400) / 100);
  const hole = g.kind === 'doughnut' ? clamp(g.holeSize, 10, 90) / 100 : 0;
  const ringW = (r * (1 - hole)) / rings.length;
  rings.forEach((s, k) => {
    const vals = s.values.map((v) => (v !== undefined && v > 0 ? v : 0));
    const total = vals.reduce((a, b) => a + b, 0);
    if (total <= 0) return;
    const rO = r - k * ringW;
    const rI = g.kind === 'doughnut' ? rO - ringW : 0;
    const labels = labelOpts(ctx, g, s);
    let ang = ((g.firstSliceAng - 90) * Math.PI) / 180;
    vals.forEach((v, i) => {
      if (v <= 0) return;
      const sweep = (v / total) * 2 * Math.PI;
      const mid = ang + sweep / 2;
      const dpt = s.points.get(i);
      const exp = (Number(dpt?.element(NS.c, 'explosion')?.attr('val') ?? s.explosion) || 0) / 100;
      const ox = cx + Math.cos(mid) * exp * r;
      const oy = cy + Math.sin(mid) * exp * r;
      const p = new Path();
      if (sweep >= 2 * Math.PI - 1e-9) {
        p.ellipse(ox - rO, oy - rO, 2 * rO, 2 * rO);
        if (rI > 0) p.ellipse(ox - rI, oy - rI, 2 * rI, 2 * rI);
      } else {
        if (rI > 0) p.moveTo(ox + rI * Math.cos(ang), oy + rI * Math.sin(ang));
        else p.moveTo(ox, oy);
        p.ellipseArc(ox, oy, rO, rO, ang, sweep);
        if (rI > 0) p.ellipseArc(ox, oy, rI, rI, ang + sweep, -sweep);
        p.close();
      }
      const paint = paintOf(ctx, s, g.kind, ctx.autoColor(g.varyColors ? i : s.idx), i);
      ctx.out.push(...paintPath(p, paint.fill, paint.line, rI > 0));
      if (labels !== undefined) {
        const out = labels.pos === 'outEnd';
        const lr = out ? rO + 4 : rI > 0 ? (rO + rI) / 2 : rO * 0.65;
        const st = ctx.text(labelText(labels, s, i, s.values[i] ?? 0, total), labels.style);
        const lx = ox + Math.cos(mid) * lr;
        const align = out ? (Math.cos(mid) >= 0 ? 'l' : 'r') : 'c';
        ctx.put(st, lx, oy + Math.sin(mid) * lr, align, 'm');
      }
      ang += sweep;
    });
  });
}

// ---- radar --------------------------------------------------------------------------------

function drawRadar(ctx: Ctx, g: ChartGroup, axes: Map<string, AxisDef>, area: Rect): void {
  const catAx = axes.get(g.axIds[0] ?? '');
  const valAx = axes.get(g.axIds[1] ?? '');
  const n = Math.max(1, ...g.series.map((s) => Math.max(s.values.length, s.cats.length)));
  const withCats = g.series.find((s) => s.cats.length > 0);
  const cats = (withCats?.cats ?? []).map((c) => catLabel(c, withCats?.catFormat));
  const labelStyle = textStyle(catAx?.el, ctx.tc);
  const lh = ctx.lineHeight(labelStyle);
  const showCats = catAx !== undefined && !catAx.deleted && catAx.tickLblPos !== 'none';
  const shaped = showCats
    ? Array.from({ length: n }, (_, i) => ctx.text(cats[i] ?? String(i + 1), labelStyle))
    : [];
  const maxW = Math.max(0, ...shaped.map((s) => s.width));
  const r = Math.max(5, Math.min(area.w / 2 - maxW - 6, area.h / 2 - lh - 4));
  const cx = area.x + area.w / 2;
  const cy = area.y + area.h / 2;
  const all = g.series.flatMap((s) => s.values.filter((v): v is number => v !== undefined));
  const fixed: { min?: number; max?: number; unit?: number } = {};
  if (valAx?.min !== undefined) fixed.min = valAx.min;
  if (valAx?.max !== undefined) fixed.max = valAx.max;
  if (valAx?.majorUnit !== undefined) fixed.unit = valAx.majorUnit;
  const sc = niceScale(Math.min(0, ...all), Math.max(0, ...all), 5, fixed, true);
  const angle = (i: number): number => -Math.PI / 2 + (i * 2 * Math.PI) / n;
  const at = (i: number, v: number): [number, number] => {
    const rr = ((clamp(v, sc.min, sc.max) - sc.min) / (sc.max - sc.min || 1)) * r;
    return [cx + Math.cos(angle(i)) * rr, cy + Math.sin(angle(i)) * rr];
  };
  const grid = lineOf(valAx?.gridlines?.element(NS.c, 'spPr')?.element(NS.a, 'ln'), ctx.theme, {
    color: 'D9D9D9',
    width: 0.75,
  });
  const ticks: number[] = [];
  for (let t = sc.min; t <= sc.max + sc.step * 1e-6; t += sc.step) ticks.push(t);
  if (valAx?.gridlines !== undefined) {
    for (const t of ticks) {
      if (t <= sc.min) continue;
      const p = new Path();
      p.poly(Array.from({ length: n }, (_, i) => at(i, t)).flat(), true);
      ctx.paint(p, null, grid);
    }
  }
  const spoke = lineOf(catAx?.el.element(NS.c, 'spPr')?.element(NS.a, 'ln'), ctx.theme, {
    color: 'D9D9D9',
    width: 0.75,
  });
  for (let i = 0; i < n; i++)
    ctx.paint(new Path().moveTo(cx, cy).lineTo(...at(i, sc.max)), null, spoke);
  shaped.forEach((st, i) => {
    const a = angle(i);
    const [x, y] = [cx + Math.cos(a) * (r + 4), cy + Math.sin(a) * (r + 4)];
    const align = Math.abs(Math.cos(a)) < 0.2 ? 'c' : Math.cos(a) > 0 ? 'l' : 'r';
    ctx.put(st, x, y, align, Math.sin(a) < -0.8 ? 'b' : Math.sin(a) > 0.8 ? 't' : 'm');
  });
  if (valAx !== undefined && !valAx.deleted && valAx.tickLblPos !== 'none') {
    const vs = textStyle(valAx.el, ctx.tc);
    for (const t of ticks) {
      const [, y] = at(0, t);
      ctx.put(ctx.text(formatValue(round(t, sc.step), valAx.format), vs), cx - 3, y, 'r', 'm');
    }
  }
  for (const s of g.series) {
    const auto = ctx.autoColor(s.idx);
    const paint = paintOf(ctx, s, g.style === 'filled' ? 'area' : 'radar', auto);
    const p = new Path().poly(
      Array.from({ length: n }, (_, i) => at(i, s.values[i] ?? 0)).flat(),
      true,
    );
    ctx.paint(p, g.style === 'filled' ? paint.fill : null, paint.line);
    const m = markerOf(ctx, s, g.style === 'marker', auto);
    if (m.symbol !== 'none' && g.style !== 'filled') {
      for (let i = 0; i < n; i++) {
        const v = s.values[i];
        if (v !== undefined) marker(ctx, m.symbol, m.size, ...at(i, v), m.fill, m.line);
      }
    }
  }
}

// ---- entry points -------------------------------------------------------------------------

/**
 * Display items for the chart a `c:chart` reference (r:id) points at, laid out in a box of
 * `width` × `height` points (relative to its top-left). Undefined when the chart part is missing,
 * is not a c:chartSpace (e.g. a chartex part) or has no supported chart group.
 */
export function chartItems(
  chartRef: XElement,
  width: number,
  height: number,
  env: LayoutEnv,
): DisplayItem[] | undefined {
  const space = chartSpace(chartRef, env);
  const chart = space?.element(NS.c, 'chart');
  const plotArea = chart?.element(NS.c, 'plotArea');
  if (space === undefined || chart === undefined || plotArea === undefined) return undefined;
  const groups = readGroups(plotArea);
  if (groups.length === 0 || !(width > 0) || !(height > 0)) return undefined;
  const theme = graphicsTheme(env.doc);
  const blanks = chart.element(NS.c, 'dispBlanksAs')?.attr('val') ?? 'gap';
  const ctx = new Ctx(env, theme, textCtx(space, theme), blanks);
  const box: Rect = { x: 0, y: 0, w: width, h: height };
  const spPr = space.element(NS.c, 'spPr');
  const bg = fillOf(spPr, theme);
  ctx.paint(
    new Path().rect(0, 0, width, height),
    bg === undefined ? 'FFFFFF' : bg,
    lineOf(spPr?.element(NS.a, 'ln'), theme, null),
  );
  let area: Rect = {
    x: PAD,
    y: PAD,
    w: Math.max(10, width - 2 * PAD),
    h: Math.max(10, height - 2 * PAD),
  };
  area = drawTitle(ctx, chart, groups, area);
  area = drawLegend(ctx, chart, groups, area);
  const plotSpPr = plotArea.element(NS.c, 'spPr');
  const pie = groups.find((g) => g.kind === 'pie' || g.kind === 'doughnut');
  const radar = groups.find((g) => g.kind === 'radar');
  if (pie !== undefined) drawPie(ctx, pie, manualLayout(plotArea, width, height)?.r ?? area);
  else if (radar !== undefined)
    drawRadar(ctx, radar, readAxes(plotArea), manualLayout(plotArea, width, height)?.r ?? area);
  else {
    const plotFill = fillOf(plotSpPr, theme);
    if (plotFill != null) {
      const m = manualLayout(plotArea, width, height);
      if (m !== undefined) ctx.paint(new Path().rect(m.r.x, m.r.y, m.r.w, m.r.h), plotFill, null);
    }
    drawCartesian(ctx, groups, readAxes(plotArea), plotArea, area, box);
  }
  return ctx.out;
}

/** Run properties of every text element of the chart (for FontManager.want before load). */
export function chartRunProps(
  chartRef: XElement,
  env: Pick<LayoutEnv, 'doc' | 'partOf'>,
): ResolvedRunProps[] {
  const space = chartSpace(chartRef, env);
  const chart = space?.element(NS.c, 'chart');
  if (space === undefined || chart === undefined) return [];
  const tc = textCtx(space, graphicsTheme(env.doc));
  const styles: ChartText[] = [textStyle(chart.element(NS.c, 'title'), tc, true)];
  styles.push(textStyle(chart.element(NS.c, 'legend'), tc));
  const plotArea = chart.element(NS.c, 'plotArea');
  if (plotArea !== undefined) {
    for (const a of readAxes(plotArea).values()) {
      styles.push(textStyle(a.el, tc));
      if (a.title !== undefined) styles.push(textStyle(a.title, tc));
    }
    for (const g of readGroups(plotArea)) {
      for (const d of [g.el, ...g.series.map((s) => s.el)]) {
        const lbls = d.element(NS.c, 'dLbls');
        if (lbls !== undefined) styles.push(textStyle(lbls, tc));
      }
    }
  }
  const seen = new Set<string>();
  const out: ResolvedRunProps[] = [];
  for (const t of styles) {
    const key = `${t.font}|${t.size}|${t.bold}|${t.italic}|${t.color}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(runProps(t));
  }
  return out;
}
