// Table layout: column widths, cell content, row heights, borders and shading → one strip per row.
//
// Word-saved documents store the grid Word computed (`w:tblGrid`), so its widths are used as
// they are ("fixed"). Only when there is no usable grid (generated documents) the columns are
// auto-fitted from the cells' minimum (longest word) and maximum (unbroken paragraph) content
// widths, roughly as Word's autofit does. Row heights follow `w:trHeight` (atLeast/exact);
// vertically merged cells (vMerge) span rows and grow the last spanned row when their content
// needs it. Borders are collapsed (the heavier of two adjacent borders wins, table borders fill
// in where the cells say nothing). Rows that may split (no cantSplit, no exact height) carry a
// `split` function the paginator calls at a page bottom; header rows are flagged for repetition.

import type { BorderSpec, Borders } from '../resolve/types.js';
import { stackBlocks } from './page/blocks.js';
import {
  type BreakRule,
  type LayoutCtx,
  measure,
  mergeStrip,
  type Strip,
  stackHeight,
  stackLayout,
} from './page/strip.js';
import type { DisplayItem, IRBlock, IRCell, IRParagraph, IRRow, IRTable } from './types.js';

const BIG = 100_000;

/** Width of the longest unbreakable word of a paragraph (autofit minimum). */
function longestWord(ctx: LayoutCtx, p: IRParagraph): number {
  const fonts = ctx.env.fonts;
  let best = 0;
  let cur = 0;
  for (const i of p.inlines) {
    if (i.kind === 'text') {
      const scale = (i.props.scale || 100) / 100;
      for (const ch of i.text) {
        const cp = ch.codePointAt(0)!;
        if (cp === 0x20 || cp === 0x09 || cp === 0x2002 || cp === 0x2003) {
          cur = 0;
          continue;
        }
        const font = fonts.fontFor(i.props, cp);
        const adv = (font.advance(font.glyphId(cp)) * i.props.size) / font.metrics.unitsPerEm;
        cur += adv * scale + i.props.spacing / 20;
        if (cur > best) best = cur;
      }
    } else if (i.kind === 'image' || i.kind === 'drawing') {
      cur += i.width;
      if (cur > best) best = cur;
    } else if (i.kind === 'tab' || i.kind === 'break') {
      cur = 0;
    }
  }
  const ind = p.props.indent;
  return best + Math.max(0, ind.left + Math.max(0, ind.firstLine) + ind.right) / 20;
}

/** Widest line of the paragraph laid out without width limit (autofit maximum). */
function widestLine(ctx: LayoutCtx, p: IRParagraph): number {
  const pl = measure(ctx, p, BIG);
  let w = 0;
  for (const l of pl.lines) w = Math.max(w, pl.indentLeft + l.width);
  return w + Math.max(0, p.props.indent.right) / 20;
}

/** [min, max] content width of blocks (explicit stack over nested tables). */
function contentRange(ctx: LayoutCtx, blocks: readonly IRBlock[]): [number, number] {
  let min = 0;
  let max = 0;
  for (const b of blocks) {
    if (b.kind === 'paragraph') {
      min = Math.max(min, longestWord(ctx, b));
      max = Math.max(max, widestLine(ctx, b));
    } else if (ctx.depth < 32) {
      ctx.depth++;
      try {
        const w = columnWidths(ctx, b, BIG).reduce((a, c) => a + c, 0);
        min = Math.max(min, w);
        max = Math.max(max, w);
      } finally {
        ctx.depth--;
      }
    }
  }
  return [min, Math.max(min, max)];
}

/** Number of grid columns the rows need. */
function gridCount(t: IRTable): number {
  let n = t.grid.length;
  for (const r of t.rows) {
    let c = r.gridBefore + r.gridAfter;
    for (const cell of r.cells) c += Math.max(1, cell.gridSpan);
    n = Math.max(n, c);
  }
  return n;
}

/** Target table width from `w:tblW`, or undefined for auto. */
function preferredWidth(t: IRTable, avail: number): number | undefined {
  const w = t.props.width;
  if (w === undefined) return undefined;
  if (w.type === 'dxa' && w.value > 0) return w.value / 20;
  if (w.type === 'pct' && w.value > 0) return (avail * w.value) / 5000;
  return undefined;
}

/** Column widths in points: the stored grid, or autofit when there is no usable grid. */
export function columnWidths(ctx: LayoutCtx, t: IRTable, avail: number): number[] {
  const n = gridCount(t);
  const grid = t.grid.filter((w) => Number.isFinite(w) && w >= 0);
  const sum = grid.reduce((a, c) => a + c, 0);
  if (grid.length === t.grid.length && sum > 0 && n > 0) {
    const out = grid.slice(0, n);
    const fill = out.length < n ? Math.max(0, avail - sum) / (n - out.length) : 0;
    while (out.length < n) out.push(fill);
    return out;
  }
  if (n === 0) return [];
  // autofit: min/max per column from single-span cells, spans spread evenly
  const mins = new Array<number>(n).fill(0);
  const maxs = new Array<number>(n).fill(0);
  const spans: { col: number; span: number; min: number; max: number }[] = [];
  for (const r of t.rows) {
    let col = r.gridBefore;
    for (const cell of r.cells) {
      const span = Math.max(1, cell.gridSpan);
      const pad = cell.margins.left + cell.margins.right;
      const [mn, mx] = contentRange(ctx, cell.blocks);
      if (span === 1 && col < n) {
        mins[col] = Math.max(mins[col]!, mn + pad);
        maxs[col] = Math.max(maxs[col]!, mx + pad);
      } else spans.push({ col, span, min: mn + pad, max: mx + pad });
      col += span;
    }
  }
  for (const s of spans) {
    let curMin = 0;
    let curMax = 0;
    const end = Math.min(n, s.col + s.span);
    for (let c = s.col; c < end; c++) {
      curMin += mins[c]!;
      curMax += maxs[c]!;
    }
    const k = end - s.col;
    for (let c = s.col; c < end && k > 0; c++) {
      if (s.min > curMin) mins[c]! += (s.min - curMin) / k;
      if (s.max > curMax) maxs[c]! += (s.max - curMax) / k;
    }
  }
  for (let c = 0; c < n; c++) maxs[c] = Math.max(maxs[c]!, mins[c]!);
  const sumMin = mins.reduce((a, c) => a + c, 0);
  const sumMax = maxs.reduce((a, c) => a + c, 0);
  const target = preferredWidth(t, avail);
  if (target !== undefined && sumMax <= target) {
    const k = sumMax > 0 ? target / sumMax : 0;
    return maxs.map((m) => (k > 0 ? m * k : target / n));
  }
  const limit = target ?? avail;
  if (sumMax <= limit) return maxs;
  if (sumMin >= limit) return mins;
  const f = (limit - sumMin) / (sumMax - sumMin);
  return mins.map((m, c) => m + (maxs[c]! - m) * f);
}

// ---- rows -------------------------------------------------------------------------------------

interface CellBox {
  x: number;
  w: number;
  cell: IRCell;
  /** Content strips (laid out for continuation cells too; they only count toward row height). */
  strips: Strip[];
  /** Content height including top/bottom margins. */
  content: number;
  /** Rows covered by a vMerge restart cell (≥ 1). */
  rowSpan: number;
  /** Restart cell of a vMerge continuation segment. */
  owner?: CellBox;
  /** First cell of the row. */
  first: boolean;
  last: boolean;
  /** Cell above (same start column) and to the left, for border conflicts. */
  above?: CellBox;
  left?: CellBox;
}

interface RowModel {
  row: IRRow;
  index: number;
  cells: CellBox[];
  height: number;
  /** Row heights of the rows a merged cell spans, from this row on (for vAlign over the span). */
  spanHeights: number[];
  /** Widths of the row's top border (and bottom border of the last row), part of its height. */
  topW: number;
  botW: number;
}

function heavier(a: BorderSpec | undefined, b: BorderSpec | undefined): BorderSpec | undefined {
  if (a === undefined) return b;
  if (b === undefined) return a;
  return b.size > a.size ? b : a;
}

function borderLines(
  b: BorderSpec | undefined,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  out: DisplayItem[],
): void {
  if (b === undefined) return;
  const width = Math.max(0.25, b.size / 8);
  const color = b.color === 'auto' || b.color === '' ? '000000' : b.color;
  const dash =
    b.style === 'dotted'
      ? [width, width * 2]
      : b.style.startsWith('dash')
        ? [width * 3, width * 3]
        : undefined;
  if (b.style === 'double') {
    const off = width;
    const vertical = x1 === x2;
    for (const s of [-off, off]) {
      out.push({
        kind: 'line',
        x1: vertical ? x1 + s : x1,
        y1: vertical ? y1 : y1 + s,
        x2: vertical ? x2 + s : x2,
        y2: vertical ? y2 : y2 + s,
        color,
        width: width / 2,
      });
    }
    return;
  }
  const line: DisplayItem = { kind: 'line', x1, y1, x2, y2, color, width };
  if (dash !== undefined) line.dash = dash;
  out.push(line);
}

/** Thickness a border adds to a row (Word sizes rows to content + horizontal borders). */
function borderWidth(b: BorderSpec | undefined): number {
  if (b === undefined) return 0;
  const w = Math.max(0.25, b.size / 8);
  return b.style === 'double' ? w * 3 : w;
}

/** The collapsed top border of a cell segment (none inside a vertical merge). */
function topBorderOf(box: CellBox, rowIndex: number, tb: Borders): BorderSpec | undefined {
  if (box.owner !== undefined) return undefined;
  const cb = box.cell.borders;
  if (rowIndex === 0 || box.above === undefined) return cb.top ?? tb.top;
  return heavier(cb.top, (box.above.owner ?? box.above).cell.borders.bottom) ?? tb.insideH;
}

/** Builds the strip of one row (or row part). */
function rowStrip(
  ctx: LayoutCtx,
  t: IRTable,
  m: RowModel,
  x0: number,
  isLastRow: boolean,
  breakBefore: BreakRule,
  allowSplit: boolean,
): Strip {
  const tb: Borders = t.props.borders;
  const strip: Strip = { height: m.height, spaceBefore: 0, spaceAfter: 0, items: [], breakBefore };
  const bg: DisplayItem[] = [];
  const lines: DisplayItem[] = [];
  for (const box of m.cells) {
    const owner = box.owner ?? box;
    const x = x0 + box.x;
    const shade = owner.cell.shading;
    if (shade !== undefined && shade !== 'auto' && shade !== '') {
      bg.push({ kind: 'rect', x, y: 0, width: box.w, height: m.height, fill: shade });
    }
    const cb = owner.cell.borders;
    borderLines(topBorderOf(box, m.index, tb), x, 0, x + box.w, 0, lines);
    const leftB = box.first
      ? (cb.left ?? tb.left)
      : (heavier(cb.left, box.left ? (box.left.owner ?? box.left).cell.borders.right : undefined) ??
        tb.insideV);
    borderLines(leftB, x, 0, x, m.height, lines);
    if (box.last) borderLines(cb.right ?? tb.right, x + box.w, 0, x + box.w, m.height, lines);
    if (isLastRow) borderLines(cb.bottom ?? tb.bottom, x, m.height, x + box.w, m.height, lines);
    if (box.owner === undefined) {
      // content, vertically aligned over the rows the cell spans
      const span = m.spanHeights.slice(0, box.rowSpan).reduce((a, c) => a + c, 0) || m.height;
      const mg = box.cell.margins;
      const used = box.content + m.topW + (isLastRow ? m.botW : 0);
      const free = Math.max(0, span - used);
      const off = box.cell.vAlign === 'center' ? free / 2 : box.cell.vAlign === 'bottom' ? free : 0;
      const { ys } = stackLayout(box.strips);
      box.strips.forEach((s, si) => {
        mergeStrip(strip, s, x + mg.left, m.topW + mg.top + off + ys[si]!);
      });
    }
  }
  strip.items.unshift(...bg);
  strip.items.push(...lines);
  const exact = m.row.height?.rule === 'exact';
  if (allowSplit && !m.row.cantSplit && !exact) {
    strip.split = (avail) => splitRow(ctx, t, m, x0, isLastRow, avail);
  }
  return strip;
}

/** Cuts a row so the first part is at most `avail` high (content split at strip boundaries). */
function splitRow(
  ctx: LayoutCtx,
  t: IRTable,
  m: RowModel,
  x0: number,
  isLastRow: boolean,
  avail: number,
): [Strip, Strip] | undefined {
  let anyFirst = false;
  let anyRest = false;
  const firstCells: CellBox[] = [];
  const restCells: CellBox[] = [];
  let h1 = 0;
  let h2 = 0;
  for (const box of m.cells) {
    if (box.owner !== undefined || box.rowSpan > 1) {
      firstCells.push(box);
      restCells.push({ ...box, strips: [], content: 0, owner: box.owner ?? box });
      continue;
    }
    const mg = box.cell.margins;
    const room = avail - mg.top - mg.bottom - m.topW;
    const { ys } = stackLayout(box.strips);
    let k = 0;
    while (k < box.strips.length && ys[k]! + box.strips[k]!.height <= room + 0.01) k++;
    const a = box.strips.slice(0, k);
    const b = box.strips.slice(k);
    if (a.length > 0) anyFirst = true;
    if (b.length > 0) anyRest = true;
    const ca = mg.top + stackHeight(a) + mg.bottom;
    const cbh = b.length > 0 ? mg.top + stackHeight(b) + mg.bottom : 0;
    h1 = Math.max(h1, Math.min(ca + m.topW, avail));
    h2 = Math.max(h2, cbh > 0 ? cbh + m.topW + (isLastRow ? m.botW : 0) : 0);
    firstCells.push({ ...box, strips: a, content: ca });
    restCells.push({ ...box, strips: b, content: cbh });
  }
  const minH = m.row.height?.value ?? 0;
  if (!anyRest) {
    // all content fits; only an atLeast height makes the row too tall
    if (minH <= avail) return undefined;
    h1 = avail;
    h2 = minH - avail;
  } else if (!anyFirst) return undefined;
  const first: RowModel = { ...m, cells: firstCells, height: h1, spanHeights: [h1] };
  const rest: RowModel = {
    ...m,
    index: m.index + 1, // draws its own top border
    cells: restCells,
    height: Math.max(h2, 0),
    spanHeights: [Math.max(h2, 0)],
  };
  return [
    rowStrip(ctx, t, first, x0, false, 'auto', false),
    rowStrip(ctx, t, rest, x0, isLastRow, 'auto', true),
  ];
}

/** Lays a table out in a box `avail` wide; one strip per row. */
export function tableStrips(
  ctx: LayoutCtx,
  t: IRTable,
  avail: number,
  keepWithPrevious: boolean,
): Strip[] {
  const widths = columnWidths(ctx, t, avail);
  const colX = [0];
  for (const w of widths) colX.push(colX[colX.length - 1]! + w);
  const total = colX[colX.length - 1]!;
  const indent = t.props.indent / 20;
  const x0 =
    t.props.alignment === 'center'
      ? (avail - total) / 2
      : t.props.alignment === 'right'
        ? avail - total
        : indent;
  const n = widths.length;
  const models: RowModel[] = [];
  const active = new Map<number, CellBox>(); // vMerge restart cell by start column
  let prevByCol = new Map<number, CellBox>();
  t.rows.forEach((row, ri) => {
    let col = row.gridBefore;
    const cells: CellBox[] = [];
    const byCol = new Map<number, CellBox>();
    row.cells.forEach((cell, ci) => {
      const span = Math.max(1, cell.gridSpan);
      const c0 = Math.min(col, n);
      const c1 = Math.min(col + span, n);
      const x = colX[c0]!;
      const w = Math.max(0, colX[c1]! - x);
      const inner = Math.max(1, w - cell.margins.left - cell.margins.right);
      const strips = stackBlocks(ctx, cell.blocks, inner, true);
      const box: CellBox = {
        x,
        w,
        cell,
        strips,
        content: cell.margins.top + stackHeight(strips) + cell.margins.bottom,
        rowSpan: 1,
        first: ci === 0,
        last: ci === row.cells.length - 1,
      };
      const above = prevByCol.get(col);
      if (above !== undefined) box.above = above;
      const left = cells[cells.length - 1];
      if (left !== undefined) box.left = left;
      const restart = active.get(col);
      if (cell.vMerge === 'continue' && restart !== undefined) {
        box.owner = restart;
        restart.rowSpan++;
      } else if (cell.vMerge === 'restart') active.set(col, box);
      else active.delete(col);
      cells.push(box);
      byCol.set(col, box);
      col += span;
    });
    prevByCol = byCol;
    models.push({ row, index: ri, cells, height: 0, spanHeights: [], topW: 0, botW: 0 });
  });
  // content heights (merged cells excluded) + horizontal borders, row height rules, then merged
  // cells grow their last row
  const tb = t.props.borders;
  models.forEach((m, ri) => {
    const last = ri === models.length - 1;
    for (const b of m.cells) {
      m.topW = Math.max(m.topW, borderWidth(topBorderOf(b, m.index, tb)));
      if (last)
        m.botW = Math.max(m.botW, borderWidth((b.owner ?? b).cell.borders.bottom ?? tb.bottom));
      if (b.owner !== undefined || b.rowSpan === 1) m.height = Math.max(m.height, b.content);
    }
    m.height += m.topW + m.botW;
    const rule = m.row.height;
    if (rule?.rule === 'exact') m.height = rule.value;
    else if (rule !== undefined) m.height = Math.max(m.height, rule.value);
  });
  models.forEach((m, ri) => {
    for (const b of m.cells) {
      if (b.owner !== undefined || b.rowSpan <= 1) continue;
      const end = Math.min(models.length, ri + b.rowSpan);
      let sum = 0;
      for (let r = ri; r < end; r++) sum += models[r]!.height;
      const last = models[end - 1]!;
      const need = b.content + m.topW + (end === models.length ? last.botW : 0);
      if (need > sum && last.row.height?.rule !== 'exact') last.height += need - sum;
    }
  });
  for (let ri = 0; ri < models.length; ri++) {
    models[ri]!.spanHeights = models.slice(ri).map((m) => m.height);
  }
  let headerCount = 0;
  while (headerCount < t.rows.length && t.rows[headerCount]!.header) headerCount++;
  const out: Strip[] = [];
  models.forEach((m, ri) => {
    let rule: BreakRule = 'auto';
    if (ri === 0 && keepWithPrevious) rule = 'avoid';
    else if (ri > 0 && ri <= headerCount) rule = 'avoid';
    const s = rowStrip(ctx, t, m, x0, ri === models.length - 1, rule, true);
    s.table = { key: t, header: ri < headerCount };
    out.push(s);
  });
  return out;
}
