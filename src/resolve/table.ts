// Tables: resolved table/cell properties, table-style conditional formatting and border conflicts.
//
// Conditional formatting (ECMA-376 §17.7.6): the table's `w:tblLook` switches regions on
// (header row, total row, first/last column; banding unless noHBand/noVBand), and each cell
// gets the regions it falls into, applied in Word's precedence order (low → high):
//   wholeTable, band1Vert/band2Vert, band1Horz/band2Horz, firstCol, lastCol, firstRow,
//   lastRow, neCell, nwCell, seCell, swCell.
// Banding skips the header row/first column when those are on, and counts in groups of
// `w:tblStyleRowBandSize`/`w:tblStyleColBandSize`.
//
// Borders: each cell edge comes from the cell's own `w:tcBorders`, else the highest conditional
// region defining it (regions behave like little tables: a header row's `left` is the left edge
// of its first cell, its `insideV` the edges between its cells), else the table borders (`top`
// on the first row, `insideH` between rows, ...; a row's `w:tblPrEx` first). Without cell
// spacing, adjacent cells share an edge and the conflict is settled as Word does (§17.4.66):
// the heavier border (width × style number) wins, then the earlier style, then the darker colour.

import type { Story } from '../wml/story.js';
import { Table } from '../wml/table.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { DocContext } from './context.js';
import { cached, store } from './context.js';
import type { BorderSide, ParaLayer } from './ppr.js';
import { finishBorders, mergePara } from './ppr.js';
import type { RunLayer } from './rpr.js';
import { mergeRun } from './rpr.js';
import type { TableStyleInfo } from './styles.js';
import type { Look, MarginSide, TblLayer, TcLayer } from './tblpr.js';
import { mergeTbl, readTblPr, readTcPr } from './tblpr.js';
import type {
  BorderSpec,
  Borders,
  CellMargins,
  ResolvedCellProps,
  ResolvedTableProps,
} from './types.js';

/** Position of one `w:tc` in its table. */
export interface CellPos {
  readonly tc: XElement;
  readonly tr: XElement;
  readonly row: number;
  /** Index among the row's cells. */
  readonly cellIndex: number;
  readonly rowCells: number;
  /** First grid column. */
  readonly col: number;
  readonly colSpan: number;
  /** Origin cell of a vertical merge (itself unless this is a `continue` cell). */
  origin: CellPos;
  /** Rows covered (origin cells only). */
  rowSpan: number;
}

/** Everything resolved once per table. */
export interface TableInfo {
  readonly tbl: XElement;
  readonly style: TableStyleInfo;
  /** Style whole-table tblPr overridden by the table's own tblPr. */
  readonly merged: TblLayer;
  readonly look: Look;
  readonly rowCount: number;
  readonly cells: ReadonlyMap<XElement, CellPos>;
  /** grid[row][col] → the cell occupying the slot (continuation cells included). */
  readonly grid: readonly (readonly (CellPos | undefined)[])[];
  /** `w:tblPrEx` of rows that have one. */
  readonly rowEx: ReadonlyMap<XElement, TblLayer>;
  /** @internal Per-cell memo of regions and style layers. */
  readonly memo: Map<XElement, CellMemo>;
}

interface CellMemo {
  regions: readonly string[];
  direct: TcLayer;
  pPr?: ParaLayer;
  rPr?: RunLayer;
}

const DEFAULT_LOOK: Look = {
  firstRow: false,
  lastRow: false,
  firstColumn: false,
  lastColumn: false,
  noHBand: false,
  noVBand: false,
};

/** The resolved info of a `w:tbl` (cached per table and story-part version). */
export function tableInfo(ctx: DocContext, tbl: XElement, story: Story): TableInfo {
  const hit = cached(ctx.tables, tbl);
  if (hit !== undefined) return hit;
  const tblPr = tbl.element(NS.w, 'tblPr');
  const styleId = tblPr?.element(NS.w, 'tblStyle')?.attrNS(NS.w, 'val');
  const style = ctx.tableStyle(styleId);
  const merged = mergeTbl(style.whole.tbl, readTblPr(ctx, tblPr));
  const rows = new Table(tbl, story).rows();
  const cells = new Map<XElement, CellPos>();
  const grid: (CellPos | undefined)[][] = [];
  const rowEx = new Map<XElement, TblLayer>();
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!;
    const ex = row.element.element(NS.w, 'tblPrEx');
    if (ex !== undefined) rowEx.set(row.element, readTblPr(ctx, ex));
    const line: (CellPos | undefined)[] = [];
    grid.push(line);
    const rowCells = row.cells();
    let col = row.gridBefore;
    for (let ci = 0; ci < rowCells.length; ci++) {
      const cell = rowCells[ci]!;
      const span = cell.gridSpan;
      const pos: CellPos = {
        tc: cell.element,
        tr: row.element,
        row: r,
        cellIndex: ci,
        rowCells: rowCells.length,
        col,
        colSpan: span,
        origin: undefined as unknown as CellPos,
        rowSpan: 1,
      };
      pos.origin = pos;
      if (cell.vMerge === 'continue') {
        const above = grid[r - 1]?.[col];
        if (above !== undefined) {
          pos.origin = above.origin;
          pos.origin.rowSpan = Math.max(pos.origin.rowSpan, r - pos.origin.row + 1);
        }
      }
      for (let c = col; c < col + span && c < 1024; c++) line[c] = pos;
      cells.set(cell.element, pos);
      col += span;
    }
  }
  const info: TableInfo = {
    tbl,
    style,
    merged,
    look: merged.look ?? DEFAULT_LOOK,
    rowCount: rows.length,
    cells,
    grid,
    rowEx,
    memo: new Map(),
  };
  return store(ctx.tables, tbl, info);
}

/** Region precedence, low → high. */
const REGION_ORDER = [
  'wholeTable',
  'band1Vert',
  'band2Vert',
  'band1Horz',
  'band2Horz',
  'firstCol',
  'lastCol',
  'firstRow',
  'lastRow',
  'neCell',
  'nwCell',
  'seCell',
  'swCell',
] as const;

/** Conditional regions a cell falls into, low → high precedence. */
export function cellRegions(info: TableInfo, pos: CellPos): string[] {
  const look = info.look;
  const isFirstRow = look.firstRow && pos.row === 0;
  const isLastRow = look.lastRow && pos.row === info.rowCount - 1;
  const isFirstCol = look.firstColumn && pos.cellIndex === 0;
  const isLastCol = look.lastColumn && pos.cellIndex === pos.rowCells - 1;
  const on = new Set<string>(['wholeTable']);
  if (!look.noVBand && !isFirstCol && !isLastCol) {
    const c = pos.cellIndex - (look.firstColumn ? 1 : 0);
    if (c >= 0)
      on.add(Math.floor(c / (info.merged.colBand ?? 1)) % 2 === 0 ? 'band1Vert' : 'band2Vert');
  }
  if (!look.noHBand && !isFirstRow && !isLastRow) {
    const r = pos.row - (look.firstRow ? 1 : 0);
    if (r >= 0)
      on.add(Math.floor(r / (info.merged.rowBand ?? 1)) % 2 === 0 ? 'band1Horz' : 'band2Horz');
  }
  if (isFirstCol) on.add('firstCol');
  if (isLastCol) on.add('lastCol');
  if (isFirstRow) on.add('firstRow');
  if (isLastRow) on.add('lastRow');
  if (isFirstRow && isLastCol) on.add('neCell');
  if (isFirstRow && isFirstCol) on.add('nwCell');
  if (isLastRow && isLastCol) on.add('seCell');
  if (isLastRow && isFirstCol) on.add('swCell');
  return REGION_ORDER.filter((r) => on.has(r));
}

function memoOf(ctx: DocContext, info: TableInfo, pos: CellPos): CellMemo {
  let m = info.memo.get(pos.tc);
  if (m === undefined) {
    m = { regions: cellRegions(info, pos), direct: readTcPr(ctx, pos.tc.element(NS.w, 'tcPr')) };
    info.memo.set(pos.tc, m);
  }
  return m;
}

/** Table-style paragraph and run layers for the content of a cell (whole table + regions). */
export function cellStyleLayers(
  ctx: DocContext,
  info: TableInfo,
  tc: XElement,
): { pPr: ParaLayer; rPr: RunLayer } {
  const pos = info.cells.get(tc);
  if (pos === undefined) return { pPr: info.style.whole.pPr, rPr: info.style.whole.rPr };
  const m = memoOf(ctx, info, pos);
  if (m.pPr === undefined || m.rPr === undefined) {
    let pPr = info.style.whole.pPr;
    let rPr = info.style.whole.rPr;
    for (const name of m.regions) {
      const reg = info.style.regions.get(name);
      if (reg === undefined) continue;
      pPr = mergePara(pPr, reg.pPr);
      rPr = mergeRun(rPr, reg.rPr);
    }
    m.pPr = pPr;
    m.rPr = rPr;
  }
  return { pPr: m.pPr, rPr: m.rPr };
}

/** The `w:tbl` and `w:tc` around a paragraph (nearest table; none inside text boxes). */
export function enclosingCell(p: XElement): { tc: XElement; tbl: XElement } | undefined {
  let n = p.parent;
  while (n !== null && n.nodeType === 1) {
    if (n.uri === NS.w) {
      if (n.local === 'tc') {
        const tbl = n.closest(NS.w, 'tbl');
        return tbl === undefined ? undefined : { tc: n, tbl };
      }
      if (n.local === 'txbxContent') return undefined;
    }
    n = n.parent;
  }
  return undefined;
}

// ---- table and cell properties -----------------------------------------------------------

/** Word's default left/right cell margin when nothing sets one (0.08", §17.4.42). */
const DEFAULT_SIDE_MARGIN = 115;

/** Resolved properties of a table. */
export function tableProps(info: TableInfo): ResolvedTableProps {
  const m = info.merged;
  const out: ResolvedTableProps = {
    alignment: m.jc ?? 'left',
    indent: m.ind ?? 0,
    layout: m.layout ?? 'autofit',
    cellMargins: {
      top: m.cellMar?.top ?? 0,
      left: m.cellMar?.left ?? DEFAULT_SIDE_MARGIN,
      bottom: m.cellMar?.bottom ?? 0,
      right: m.cellMar?.right ?? DEFAULT_SIDE_MARGIN,
    },
    cellSpacing: m.cellSpacing ?? 0,
    borders: finishBorders(m.borders),
    look: { ...info.look },
    bidiVisual: m.bidiVisual ?? false,
  };
  if (info.style.id !== undefined) out.styleId = info.style.id;
  if (m.width !== undefined) out.width = { ...m.width };
  return out;
}

type Edge = 'top' | 'left' | 'bottom' | 'right';

const OPPOSITE: Readonly<Record<Edge, Edge>> = {
  top: 'bottom',
  bottom: 'top',
  left: 'right',
  right: 'left',
};

const ROW_REGIONS: ReadonlySet<string> = new Set(['firstRow', 'lastRow', 'band1Horz', 'band2Horz']);
const COL_REGIONS: ReadonlySet<string> = new Set(['firstCol', 'lastCol', 'band1Vert', 'band2Vert']);

function lastRowOf(pos: CellPos): number {
  return pos.origin.row + pos.origin.rowSpan - 1;
}

/** Which table-level border applies to a cell edge (outer side or inside line). */
function tableKey(info: TableInfo, pos: CellPos, edge: Edge): BorderSide {
  switch (edge) {
    case 'top':
      return pos.origin.row === 0 ? 'top' : 'insideH';
    case 'bottom':
      return lastRowOf(pos) === info.rowCount - 1 ? 'bottom' : 'insideH';
    case 'left':
      return pos.cellIndex === 0 ? 'left' : 'insideV';
    case 'right':
      return pos.cellIndex === pos.rowCells - 1 ? 'right' : 'insideV';
  }
}

/** Which border of a conditional region applies to a cell edge. */
function regionKey(info: TableInfo, pos: CellPos, region: string, edge: Edge): BorderSide {
  if (ROW_REGIONS.has(region)) {
    if (edge === 'left') return pos.cellIndex === 0 ? 'left' : 'insideV';
    if (edge === 'right') return pos.cellIndex === pos.rowCells - 1 ? 'right' : 'insideV';
    return edge;
  }
  if (COL_REGIONS.has(region)) {
    if (edge === 'top') return pos.origin.row === 0 ? 'top' : 'insideH';
    if (edge === 'bottom') return lastRowOf(pos) === info.rowCount - 1 ? 'bottom' : 'insideH';
    return edge;
  }
  if (region === 'wholeTable') return tableKey(info, pos, edge);
  return edge;
}

/** A cell's own idea of one of its edges, before the conflict with its neighbour. */
function ownEdge(
  ctx: DocContext,
  info: TableInfo,
  pos: CellPos,
  edge: Edge,
): BorderSpec | null | undefined {
  const m = memoOf(ctx, info, pos);
  const direct = m.direct.borders?.[edge];
  if (direct !== undefined) return direct;
  for (let k = m.regions.length - 1; k >= 0; k--) {
    const region = m.regions[k]!;
    const b = info.style.regions.get(region)?.tc.borders;
    if (b === undefined) continue;
    const v = b[regionKey(info, pos, region, edge)];
    if (v !== undefined) return v;
  }
  const whole = info.style.whole.tc.borders?.[edge];
  if (whole !== undefined) return whole;
  const key = tableKey(info, pos, edge);
  const ex = info.rowEx.get(pos.tr)?.borders?.[key];
  if (ex !== undefined) return ex;
  return info.merged.borders?.[key];
}

/** The cell across `edge` (origin of a merge), if any. */
function neighbour(info: TableInfo, pos: CellPos, edge: Edge): CellPos | undefined {
  switch (edge) {
    case 'top':
      return info.grid[pos.origin.row - 1]?.[pos.col];
    case 'bottom':
      return info.grid[lastRowOf(pos) + 1]?.[pos.col];
    case 'left':
      return info.grid[pos.row]?.[pos.col - 1];
    case 'right':
      return info.grid[pos.row]?.[pos.col + pos.colSpan];
  }
}

/** Border style numbers used for the weight (ST_Border order, §17.4.66). */
const STYLE_ORDER: readonly string[] = [
  'single',
  'thick',
  'double',
  'dotted',
  'dashed',
  'dotDash',
  'dotDotDash',
  'triple',
  'thinThickSmallGap',
  'thickThinSmallGap',
  'thinThickThinSmallGap',
  'thinThickMediumGap',
  'thickThinMediumGap',
  'thinThickThinMediumGap',
  'thinThickLargeGap',
  'thickThinLargeGap',
  'thinThickThinLargeGap',
  'wave',
  'doubleWave',
  'dashSmallGap',
  'dashDotStroked',
  'threeDEmboss',
  'threeDEngrave',
  'outset',
  'inset',
];

function styleNumber(style: string): number {
  const i = STYLE_ORDER.indexOf(style);
  return i < 0 ? 1 : i + 1;
}

function rgb(color: string): [number, number, number] {
  if (color === 'auto' || color.length !== 6) return [0, 0, 0];
  return [
    Number.parseInt(color.slice(0, 2), 16),
    Number.parseInt(color.slice(2, 4), 16),
    Number.parseInt(color.slice(4, 6), 16),
  ];
}

/** Settles a shared edge between two cells' borders (§17.4.66 conflict rules). */
export function borderConflict(
  a: BorderSpec | null | undefined,
  b: BorderSpec | null | undefined,
): BorderSpec | null | undefined {
  if (!a) return b ?? a;
  if (!b) return a;
  const wa = a.size * styleNumber(a.style);
  const wb = b.size * styleNumber(b.style);
  if (wa !== wb) return wa > wb ? a : b;
  const sa = styleNumber(a.style);
  const sb = styleNumber(b.style);
  if (sa !== sb) return sa < sb ? a : b;
  const [ra, ga, ba] = rgb(a.color);
  const [rb, gb, bb] = rgb(b.color);
  const keys: readonly (readonly [number, number])[] = [
    [ra + ba + 2 * ga, rb + bb + 2 * gb],
    [ba + 2 * ga, bb + 2 * gb],
    [ga, gb],
  ];
  for (const [x, y] of keys) if (x !== y) return x < y ? a : b;
  return a;
}

function edgeOf(
  ctx: DocContext,
  info: TableInfo,
  pos: CellPos,
  edge: Edge,
  spacing: number,
): BorderSpec | null | undefined {
  // The bottom of a vertically merged cell is the bottom of its last row's cell.
  const self = edge === 'bottom' ? (info.grid[lastRowOf(pos)]?.[pos.col] ?? pos) : pos;
  const own = ownEdge(ctx, info, self, edge);
  if (spacing > 0) return own;
  const other = neighbour(info, pos, edge);
  if (other === undefined || other.origin === pos.origin) return own;
  const otherSelf =
    OPPOSITE[edge] === 'bottom' ? (info.grid[lastRowOf(other)]?.[other.col] ?? other) : other;
  return borderConflict(own, ownEdge(ctx, info, otherSelf, OPPOSITE[edge]));
}

function regionValue<K extends keyof TcLayer>(
  info: TableInfo,
  m: CellMemo,
  key: K,
): TcLayer[K] | undefined {
  const d = m.direct[key];
  if (d !== undefined) return d;
  for (let k = m.regions.length - 1; k >= 0; k--) {
    const v = info.style.regions.get(m.regions[k]!)?.tc[key];
    if (v !== undefined) return v;
  }
  return info.style.whole.tc[key];
}

/** Resolved properties of the cell `tc` of the table described by `info`. */
export function cellProps(ctx: DocContext, info: TableInfo, tc: XElement): ResolvedCellProps {
  const pos = info.cells.get(tc);
  const tableMar = (side: MarginSide, fallback: number): number =>
    (pos === undefined ? undefined : info.rowEx.get(pos.tr)?.cellMar?.[side]) ??
    info.merged.cellMar?.[side] ??
    fallback;
  if (pos === undefined) {
    return {
      margins: {
        top: tableMar('top', 0),
        left: tableMar('left', DEFAULT_SIDE_MARGIN),
        bottom: tableMar('bottom', 0),
        right: tableMar('right', DEFAULT_SIDE_MARGIN),
      },
      borders: {},
      vAlign: 'top',
      noWrap: false,
      conditional: [],
    };
  }
  const m = memoOf(ctx, info, pos);
  const mar = (side: MarginSide, fallback: number): number => {
    const d = m.direct.mar?.[side];
    if (d !== undefined) return d;
    for (let k = m.regions.length - 1; k >= 0; k--) {
      const v = info.style.regions.get(m.regions[k]!)?.tc.mar?.[side];
      if (v !== undefined) return v;
    }
    return info.style.whole.tc.mar?.[side] ?? tableMar(side, fallback);
  };
  const margins: CellMargins = {
    top: mar('top', 0),
    left: mar('left', DEFAULT_SIDE_MARGIN),
    bottom: mar('bottom', 0),
    right: mar('right', DEFAULT_SIDE_MARGIN),
  };
  const spacing = info.merged.cellSpacing ?? 0;
  const borders: Borders = {};
  for (const edge of ['top', 'left', 'bottom', 'right'] as const) {
    const b = edgeOf(ctx, info, pos, edge, spacing);
    if (b) borders[edge] = b;
  }
  const out: ResolvedCellProps = {
    margins,
    borders,
    vAlign: regionValue(info, m, 'vAlign') ?? 'top',
    noWrap: regionValue(info, m, 'noWrap') ?? false,
    conditional: m.regions.filter((r) => info.style.regions.has(r)),
  };
  if (m.direct.width !== undefined) out.width = { ...m.direct.width };
  const shd = regionValue(info, m, 'shd') ?? info.merged.shd;
  if (shd !== undefined && shd !== 'none') out.shading = shd;
  const dir = regionValue(info, m, 'textDirection');
  if (dir !== undefined) out.textDirection = dir;
  return out;
}
