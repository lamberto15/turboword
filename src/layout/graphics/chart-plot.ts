// Cartesian chart plots: bar/column, line, area and scatter groups on category and value axes
// (primary and secondary), with nice ticks, gridlines, tick marks, labels and axis titles.

import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import {
  type AxisDef,
  type ChartGroup,
  type ChartText,
  formatValue,
  niceScale,
  richLines,
  textStyle,
} from './chart-data.js';
import {
  type Ctx,
  catLabel,
  clamp,
  labelOpts,
  labelText,
  linePath,
  manualLayout,
  marker,
  markerOf,
  paintOf,
  pos,
  type Rect,
  round,
  runs,
  type Scale,
} from './chart-draw.js';
import { type LineStyle, lineOf } from './paint.js';
import { Path } from './path.js';
import type { ShapedText } from './text.js';

interface AxisPlan {
  def: AxisDef | undefined;
  /** Positions along x ('h') or y ('v'). */
  dir: 'h' | 'v';
  /** Labels on the far side (right / top). */
  far: boolean;
  cat: boolean;
  scale: Scale;
  labels: boolean;
  style: ChartText;
  format?: string | undefined;
}

/** Data range of a group on its value axis (stacked sums, percent → ±1). */
function groupRange(g: ChartGroup, n: number): [number, number] {
  let lo = Infinity;
  let hi = -Infinity;
  const stacked = g.grouping === 'stacked' || g.grouping === 'percentStacked';
  if (g.grouping === 'percentStacked') {
    for (const s of g.series) {
      for (const v of s.values) {
        if (v === undefined) continue;
        if (v < 0) lo = -1;
        if (v > 0) hi = 1;
      }
    }
    return [Math.min(0, lo), Math.max(0, hi)];
  }
  if (stacked) {
    for (let i = 0; i < n; i++) {
      let p = 0;
      let q = 0;
      for (const s of g.series) {
        const v = s.values[i] ?? 0;
        if (v >= 0) p += v;
        else q += v;
      }
      hi = Math.max(hi, p);
      lo = Math.min(lo, q);
    }
    return [lo, hi];
  }
  for (const s of g.series) {
    for (const v of s.values) {
      if (v === undefined || !Number.isFinite(v)) continue;
      lo = Math.min(lo, v);
      hi = Math.max(hi, v);
    }
  }
  return [lo, hi];
}

/** Cumulative (and percent-normalised) values per series for stacked groups. */
function stackValues(g: ChartGroup, n: number): { lo: number[][]; hi: number[][] } {
  const stacked = g.grouping === 'stacked' || g.grouping === 'percentStacked';
  const pct = g.grouping === 'percentStacked';
  const totals = Array.from({ length: n }, (_, i) =>
    g.series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0),
  );
  const pos = new Array<number>(n).fill(0);
  const neg = new Array<number>(n).fill(0);
  const lo: number[][] = [];
  const hi: number[][] = [];
  for (const s of g.series) {
    const l: number[] = [];
    const h: number[] = [];
    for (let i = 0; i < n; i++) {
      let v = s.values[i] ?? 0;
      if (pct) v = totals[i] ? v / (totals[i] as number) : 0;
      if (!stacked) {
        l.push(0);
        h.push(v);
      } else if (v >= 0) {
        l.push(pos[i] as number);
        pos[i] = (pos[i] as number) + v;
        h.push(pos[i] as number);
      } else {
        l.push(neg[i] as number);
        neg[i] = (neg[i] as number) + v;
        h.push(neg[i] as number);
      }
    }
    lo.push(l);
    hi.push(h);
  }
  return { lo, hi };
}

export function drawCartesian(
  ctx: Ctx,
  groups: ChartGroup[],
  axes: Map<string, AxisDef>,
  plotArea: XElement,
  area: Rect,
  box: Rect,
): void {
  const scatter = groups.every((g) => g.kind === 'scatter');
  const horizontal = !scatter && groups.some((g) => g.kind === 'bar' && g.dir === 'bar');
  let n = 1;
  let cats: string[] = [];
  let catFormat: string | undefined;
  for (const g of groups) {
    for (const s of g.series) {
      n = Math.max(n, s.values.length, g.kind === 'scatter' ? 0 : s.cats.length);
      if (cats.length === 0 && s.cats.length > 0 && g.kind !== 'scatter') {
        cats = s.cats;
        catFormat = s.catFormat;
      }
    }
  }
  const hasBars = groups.some((g) => g.kind === 'bar');
  // Axis plans: one per distinct axis id used by the groups.
  const plans = new Map<string, AxisPlan>();
  const ranges = new Map<string, [number, number]>();
  const groupAxes = groups.map((g) => {
    const ids = [g.axIds[0] ?? 'cat', g.axIds[1] ?? 'val'] as const;
    ids.forEach((id, k) => {
      const def = axes.get(id);
      const cat = !scatter && k === 0;
      if (!plans.has(id)) {
        const dir: 'h' | 'v' = scatter
          ? k === 0
            ? 'h'
            : 'v'
          : (k === 0) !== horizontal
            ? 'h'
            : 'v';
        const usedFar = [...plans.values()].some((p) => p.dir === dir && !p.far);
        const far =
          def === undefined
            ? usedFar
            : dir === 'v'
              ? def.pos === 'r' || (usedFar && def.pos !== 'l')
              : def.pos === 't';
        const plan: AxisPlan = {
          def,
          dir,
          far,
          cat,
          scale: { min: 0, max: 1, a: 0, b: 1, step: 1 },
          labels: def !== undefined && !def.deleted && def.tickLblPos !== 'none',
          style: textStyle(def?.el, ctx.tc),
        };
        plans.set(id, plan);
      }
      if (!cat) {
        let r: [number, number];
        if (scatter) {
          const vals = g.series.flatMap((s) => (k === 0 ? (s.xs ?? []) : s.values));
          const f = vals.filter((v): v is number => v !== undefined && Number.isFinite(v));
          r = [Math.min(...f), Math.max(...f)];
        } else r = groupRange(g, n);
        const prev = ranges.get(id);
        ranges.set(id, prev === undefined ? r : [Math.min(prev[0], r[0]), Math.max(prev[1], r[1])]);
        const plan = plans.get(id) as AxisPlan;
        if (g.grouping === 'percentStacked' && plan.format === undefined) plan.format = '0%';
        if (plan.format === undefined) {
          const sf = g.series.find((s) => s.format !== undefined)?.format;
          const df = plan.def?.format;
          plan.format =
            df !== undefined && (plan.def?.sourceLinked !== true || sf === undefined) ? df : sf;
        }
      }
    });
    return ids;
  });
  const between =
    hasBars || [...plans.values()].some((p) => !p.cat && p.def?.crossBetween !== 'midCat');
  const lhOf = (p: AxisPlan): number => ctx.lineHeight(p.style);

  const computeScales = (plot: Rect): void => {
    for (const [id, p] of plans) {
      const len = p.dir === 'h' ? plot.w : plot.h;
      let a = p.dir === 'h' ? plot.x : plot.y + plot.h;
      let b = p.dir === 'h' ? plot.x + plot.w : plot.y;
      if (p.def?.reverse === true) [a, b] = [b, a];
      if (p.cat) {
        p.scale = between
          ? { min: 0, max: n, a, b, step: 1 }
          : { min: 0, max: Math.max(1, n - 1), a, b, step: 1 };
        continue;
      }
      const [lo, hi] = ranges.get(id) ?? [0, 1];
      const fixed: { min?: number; max?: number; unit?: number } = {};
      if (p.def?.min !== undefined) fixed.min = p.def.min;
      if (p.def?.max !== undefined) fixed.max = p.def.max;
      if (p.def?.majorUnit !== undefined) fixed.unit = p.def.majorUnit;
      if (p.format === '0%' && fixed.max === undefined && hi <= 1 && lo >= -1) {
        fixed.max = hi > 0 ? 1 : 0;
        if (fixed.min === undefined) fixed.min = lo < 0 ? -1 : 0;
      }
      const count = Math.max(2, Math.min(10, Math.floor(len / (p.dir === 'h' ? 45 : 30))));
      const sc = niceScale(lo, hi, count, fixed, !scatter || lo >= 0);
      p.scale = { min: sc.min, max: sc.max, a, b, step: sc.step };
    }
  };
  const ticksOf = (p: AxisPlan): number[] => {
    const out: number[] = [];
    const { min, max, step } = p.scale;
    for (let i = 0; i <= 200; i++) {
      const t = min + i * step;
      if (t > max + step * 1e-6) break;
      out.push(round(t, step));
    }
    return out;
  };
  const valueLabels = (p: AxisPlan): string[] => ticksOf(p).map((t) => formatValue(t, p.format));
  const catLabels = (): string[] =>
    Array.from({ length: n }, (_, i) => catLabel(cats[i] ?? String(i + 1), catFormat));

  // Plot rectangle: manual layout, else the area minus label/title space.
  const manual = manualLayout(plotArea, box.w, box.h);
  let plot: Rect;
  const reserve = (base: Rect): Rect => {
    computeScales({
      x: base.x + 30,
      y: base.y + 10,
      w: Math.max(10, base.w - 40),
      h: Math.max(10, base.h - 30),
    });
    let left = 0;
    let right = 0;
    let top = 0;
    let bottom = 0;
    for (const p of plans.values()) {
      const titleH =
        p.def?.title !== undefined && richLines(p.def.title) !== undefined
          ? ctx.lineHeight(textStyle(p.def.title, ctx.tc, false)) + 4
          : 0;
      if (p.dir === 'v') top = Math.max(top, titleH);
      if (!p.labels) {
        if (p.dir === 'h' && p.far) top += titleH;
        else if (p.dir === 'h') bottom += titleH;
        continue;
      }
      const texts = p.cat ? catLabels() : valueLabels(p);
      if (p.dir === 'v') {
        const w = Math.min(
          base.w * 0.4,
          Math.max(0, ...texts.map((t) => ctx.text(t, p.style).width)) + 8,
        );
        if (p.far) right += w;
        else left += w;
      } else {
        const h = lhOf(p) * (p.cat ? 2 : 1) + 6 + titleH;
        if (p.far) top += h;
        else bottom += h;
      }
    }
    top = Math.max(top, 6);
    right = Math.max(right, 10);
    bottom = Math.max(bottom, 4);
    left = Math.max(left, 4);
    return {
      x: base.x + left,
      y: base.y + top,
      w: Math.max(10, base.w - left - right),
      h: Math.max(10, base.h - top - bottom),
    };
  };
  if (manual?.inner === true) plot = manual.r;
  else plot = reserve(manual?.r ?? area);
  computeScales(plot);

  const xy = (catPos: number, valPos: number): [number, number] =>
    horizontal ? [valPos, catPos] : [catPos, valPos];
  const gridDefault: LineStyle = { color: 'D9D9D9', width: 0.75 };
  const lineOfEl = (el: XElement | undefined, base: LineStyle | null): LineStyle | null =>
    lineOf(el?.element(NS.c, 'spPr')?.element(NS.a, 'ln'), ctx.theme, base);
  // Where each axis line sits: on its crossing axis (autoZero / min / max / crossesAt).
  const crossPos = (p: AxisPlan): number => {
    const other = [...plans.values()].find(
      (q) =>
        q !== p && q.dir !== p.dir && (p.def?.crossAx === undefined || q.def?.id === p.def.crossAx),
    );
    if (other === undefined) return p.dir === 'h' ? plot.y + plot.h : plot.x;
    const s = other.scale;
    const c = p.def?.crosses ?? 'autoZero';
    let v: number;
    if (p.def?.crossesAt !== undefined) v = p.def.crossesAt;
    else if (c === 'max') v = s.max;
    else if (c === 'min' || other.cat) v = c === 'max' ? s.max : s.min;
    else v = clamp(0, s.min, s.max);
    return pos(s, clamp(v, s.min, s.max));
  };

  // Gridlines.
  for (const p of plans.values()) {
    for (const [grid, base] of [
      [p.def?.minorGridlines, { color: 'F2F2F2', width: 0.75 }],
      [p.def?.gridlines, gridDefault],
    ] as const) {
      if (grid === undefined) continue;
      const style = lineOfEl(grid, base);
      if (style === null) continue;
      const path = new Path();
      const minor = grid === p.def?.minorGridlines;
      const at: number[] = [];
      if (p.cat) for (let i = 0; i <= (between ? n : n - 1); i++) at.push(pos(p.scale, i));
      else {
        const ticks = ticksOf(p);
        for (const t of ticks) at.push(pos(p.scale, t));
        if (minor) {
          at.length = 0;
          for (let t = p.scale.min; t <= p.scale.max + 1e-9; t += p.scale.step / 5)
            at.push(pos(p.scale, t));
        }
      }
      for (const c of at) {
        if (p.dir === 'h') path.moveTo(c, plot.y).lineTo(c, plot.y + plot.h);
        else path.moveTo(plot.x, c).lineTo(plot.x + plot.w, c);
      }
      ctx.paint(path, null, style);
    }
  }

  // Series.
  const labelJobs: (() => void)[] = [];
  groups.forEach((g, gi) => {
    const [cid, vid] = groupAxes[gi] as readonly [string, string];
    const cp = plans.get(cid) as AxisPlan;
    const vp = plans.get(vid) as AxisPlan;
    const cs = cp.scale;
    const vs = vp.scale;
    const vclamp = (v: number): number =>
      pos(vs, clamp(v, Math.min(vs.min, vs.max), Math.max(vs.min, vs.max)));
    const stack = stackValues(g, n);
    const base = !cp.cat
      ? 0
      : (() => {
          const c = cp.def?.crosses ?? 'autoZero';
          if (cp.def?.crossesAt !== undefined) return cp.def.crossesAt;
          return c === 'max' ? vs.max : c === 'min' ? vs.min : clamp(0, vs.min, vs.max);
        })();
    const catAt = (i: number): number => pos(cs, between ? i + 0.5 : i);
    if (g.kind === 'bar') {
      const stacked = g.grouping === 'stacked' || g.grouping === 'percentStacked';
      const slots = stacked ? 1 : Math.max(1, g.series.length);
      const ov = clamp(g.overlap, -100, 100) / 100;
      const barW = 1 / (slots - (slots - 1) * ov + clamp(g.gapWidth, 0, 500) / 100);
      const start = (1 - barW * (slots - (slots - 1) * ov)) / 2;
      g.series.forEach((s, j) => {
        const labels = labelOpts(ctx, g, s);
        for (let i = 0; i < n; i++) {
          const raw = s.values[i];
          if (raw === undefined) continue;
          const v0 = stacked ? (stack.lo[j]?.[i] ?? 0) : base;
          const v1 = stack.hi[j]?.[i] ?? 0;
          const c0 = i + start + (stacked ? 0 : j) * barW * (1 - ov);
          const [x0, y0] = xy(pos(cs, c0), vclamp(v0));
          const [x1, y1] = xy(pos(cs, c0 + barW), vclamp(v1));
          const vary = g.varyColors && g.series.length === 1;
          const paint = paintOf(ctx, s, 'bar', ctx.autoColor(vary ? i : s.idx), i);
          const rx = Math.min(x0, x1);
          const ry = Math.min(y0, y1);
          ctx.paint(
            new Path().rect(rx, ry, Math.abs(x1 - x0), Math.abs(y1 - y0)),
            paint.fill,
            paint.line,
          );
          if (labels !== undefined) {
            labelJobs.push(() => {
              const st = ctx.text(labelText(labels, s, i, raw, 0), labels.style);
              const center =
                labels.pos === 'ctr' ||
                labels.pos === 'inEnd' ||
                labels.pos === 'inBase' ||
                (stacked && labels.pos === '');
              const [ex, ey] = xy(pos(cs, c0 + barW / 2), vclamp(v1));
              const [mx, my] = xy(pos(cs, c0 + barW / 2), (vclamp(v0) + vclamp(v1)) / 2);
              if (center) ctx.put(st, mx, my, 'c', 'm');
              else if (horizontal)
                ctx.put(st, ex + (v1 >= v0 ? 3 : -3), ey, v1 >= v0 ? 'l' : 'r', 'm');
              else
                ctx.put(st, ex, v1 >= v0 ? ey - 2 - st.descent : ey + 2, 'c', v1 >= v0 ? 'b' : 't');
            });
          }
        }
      });
      return;
    }
    if (g.kind === 'area') {
      g.series.forEach((s, j) => {
        const top: [number, number][] = [];
        const bot: [number, number][] = [];
        const stacked = g.grouping !== 'standard';
        for (let i = 0; i < n; i++) {
          top.push(xy(catAt(i), vclamp(stack.hi[j]?.[i] ?? 0)));
          bot.push(xy(catAt(i), vclamp(stacked ? (stack.lo[j]?.[i] ?? 0) : base)));
        }
        const p = new Path();
        linePath(p, top, false);
        for (let i = bot.length - 1; i >= 0; i--) p.lineTo(...(bot[i] as [number, number]));
        p.close();
        const paint = paintOf(ctx, s, 'area', ctx.autoColor(s.idx));
        ctx.paint(p, paint.fill, paint.line);
      });
      return;
    }
    // line / scatter
    const hiLow: [number, number][] = [];
    g.series.forEach((s, j) => {
      const auto = ctx.autoColor(s.idx);
      const paint = paintOf(ctx, s, g.kind, auto);
      const stacked = g.kind === 'line' && g.grouping !== 'standard';
      const vals = stacked
        ? s.values.map((v, i) => (v === undefined ? v : (stack.hi[j]?.[i] ?? 0)))
        : s.values;
      const toXY = (i: number, v: number): [number, number] =>
        g.kind === 'scatter'
          ? [pos(cp.scale, s.xs?.[i] ?? i + 1), vclamp(v)]
          : xy(catAt(i), vclamp(v));
      const smooth = s.smooth || g.style.startsWith('smooth');
      if (paint.line !== null && !(g.kind === 'scatter' && g.style === 'marker')) {
        const p = new Path();
        for (const run of runs(vals, ctx.blanks))
          linePath(
            p,
            run.map(([i, v]) => toXY(i, v)),
            smooth,
          );
        ctx.paint(p, null, paint.line);
      }
      const m = markerOf(
        ctx,
        s,
        g.kind === 'scatter' ? g.style !== 'line' && g.style !== 'smooth' : g.markers,
        auto,
      );
      const labels = labelOpts(ctx, g, s);
      vals.forEach((v, i) => {
        if (v === undefined || !Number.isFinite(v)) return;
        const [x, y] = toXY(i, v);
        if (m.symbol !== 'none') marker(ctx, m.symbol, m.size, x, y, m.fill, m.line);
        if (labels !== undefined) {
          labelJobs.push(() =>
            ctx.put(
              ctx.text(labelText(labels, s, i, s.values[i] ?? v, 0), labels.style),
              x + 4,
              y,
              'l',
              'm',
            ),
          );
        }
        if (g.hiLowLines !== undefined) {
          const h = hiLow[i];
          hiLow[i] = h === undefined ? [v, v] : [Math.min(h[0], v), Math.max(h[1], v)];
        }
      });
    });
    if (g.hiLowLines !== undefined) {
      const style = lineOfEl(g.hiLowLines, { color: '000000', width: 0.75 });
      const p = new Path();
      hiLow.forEach((h, i) => {
        if (h === undefined) return;
        p.moveTo(...xy(catAt(i), vclamp(h[0]))).lineTo(...xy(catAt(i), vclamp(h[1])));
      });
      if (style !== null) ctx.paint(p, null, style);
    }
  });

  // Axis lines, tick marks, labels, titles.
  for (const p of plans.values()) {
    if (p.def === undefined || p.def.deleted) continue;
    const at = crossPos(p);
    const style = lineOfEl(p.def.el, { color: '868686', width: 0.75 });
    const s = p.scale;
    if (style !== null) {
      const line = new Path();
      if (p.dir === 'h') line.moveTo(Math.min(s.a, s.b), at).lineTo(Math.max(s.a, s.b), at);
      else line.moveTo(at, Math.min(s.a, s.b)).lineTo(at, Math.max(s.a, s.b));
      const tick = p.def.majorTick;
      if (tick !== 'none') {
        const outward = p.dir === 'h' ? (p.far ? -1 : 1) : p.far ? 1 : -1;
        const t0 = tick === 'in' ? 0 : 4 * outward;
        const t1 = tick === 'out' ? 0 : -4 * outward;
        const marks = p.cat
          ? Array.from({ length: (between ? n : n - 1) + 1 }, (_, i) => pos(s, i))
          : ticksOf(p).map((t) => pos(s, t));
        for (const c of marks) {
          if (p.dir === 'h') line.moveTo(c, at + t0).lineTo(c, at + t1);
          else line.moveTo(at + t0, c).lineTo(at + t1, c);
        }
      }
      ctx.paint(line, null, style);
    }
    const lh = lhOf(p);
    // Labels sit next to the axis line, or at the low/high edge of the plot.
    const edge =
      p.def.tickLblPos === 'low'
        ? p.dir === 'h'
          ? plot.y + plot.h
          : plot.x
        : p.def.tickLblPos === 'high'
          ? p.dir === 'h'
            ? plot.y
            : plot.x + plot.w
          : p.dir === 'h'
            ? p.far
              ? Math.min(at, plot.y)
              : Math.max(at, plot.y + plot.h)
            : p.far
              ? Math.max(at, plot.x + plot.w)
              : Math.min(at, plot.x);
    let labelExtent = 0;
    if (p.labels) {
      const gap = 4 + (p.def.majorTick === 'out' || p.def.majorTick === 'cross' ? 2 : 0);
      if (p.cat) {
        const texts = catLabels();
        const band = Math.abs(s.b - s.a) / Math.max(1, between ? n : n - 1);
        const shaped = texts.map((t) => ctx.text(t, p.style));
        if (p.dir === 'h') {
          // Wrap long labels on spaces, then skip labels that still collide.
          const lines = texts.map((t, i) =>
            wrap(ctx, t, p.style, band - 2, shaped[i] as ShapedText),
          );
          const widest = Math.max(0, ...lines.flat().map((l) => l.width));
          const every = Math.max(1, Math.ceil((widest + 4) / Math.max(band, 1)));
          lines.forEach((ls, i) => {
            if (i % every !== 0) return;
            const c = pos(s, between ? i + 0.5 : i);
            ls.slice(0, 3).forEach((l, k) => {
              const y = p.far ? edge - gap - (ls.length - k) * lh : edge + gap + k * lh;
              ctx.put(l, c, y, 'c', 't');
            });
            labelExtent = Math.max(labelExtent, Math.min(3, ls.length) * lh);
          });
        } else {
          const every = Math.max(1, Math.ceil(lh / Math.max(band, 1)));
          shaped.forEach((st, i) => {
            if (i % every !== 0) return;
            const c = pos(s, between ? i + 0.5 : i);
            ctx.put(st, p.far ? edge + gap : edge - gap, c, p.far ? 'l' : 'r', 'm');
          });
        }
      } else {
        const ticks = ticksOf(p);
        ticks.forEach((t) => {
          const st = ctx.text(formatValue(t, p.format), p.style);
          const c = pos(s, t);
          if (p.dir === 'h') ctx.put(st, c, p.far ? edge - gap - lh : edge + gap, 'c', 't');
          else ctx.put(st, p.far ? edge + gap : edge - gap, c, p.far ? 'l' : 'r', 'm');
        });
        labelExtent = lh;
      }
    }
    const lines = p.def.title === undefined ? undefined : richLines(p.def.title);
    if (lines !== undefined && p.def.title !== undefined) {
      const ts = textStyle(p.def.title, ctx.tc, false);
      const st = ctx.text(lines.join(' '), ts);
      if (p.dir === 'h') {
        const y = p.far ? edge - 4 - labelExtent - ctx.lineHeight(ts) - 4 : edge + 6 + labelExtent;
        ctx.put(st, plot.x + plot.w / 2, y, 'c', 't');
      } else {
        // No rotated text in the PageModel: vertical axis titles go above the axis.
        ctx.put(st, p.far ? plot.x + plot.w : plot.x, plot.y - 4, p.far ? 'r' : 'l', 'b');
      }
    }
  }
  for (const job of labelJobs) job();
}

/** Greedy word wrap into lines no wider than `max` (single words may overflow). */
function wrap(
  ctx: Ctx,
  text: string,
  style: ChartText,
  max: number,
  whole: ShapedText,
): ShapedText[] {
  if (whole.width <= max || !text.includes(' ')) return [whole];
  const out: ShapedText[] = [];
  let cur = '';
  for (const word of text.split(/\s+/)) {
    const next = cur === '' ? word : `${cur} ${word}`;
    if (cur !== '' && ctx.text(next, style).width > max) {
      out.push(ctx.text(cur, style));
      cur = word;
    } else cur = next;
  }
  if (cur !== '') out.push(ctx.text(cur, style));
  return out;
}
