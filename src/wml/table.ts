// Table views, the table's logical grid, and table editing.
//
// A `w:tbl` is not a rectangular matrix: rows can skip grid columns (`w:gridBefore`/`w:gridAfter`),
// cells span columns (`w:gridSpan`, or the legacy `w:hMerge`), and vertical merges are encoded
// as a `restart` cell followed by `continue` cells in the rows below (`w:vMerge`). `grid()` turns
// that into a matrix of slots so callers can ask "what covers row r, column c?" without
// re-implementing those rules.
//
// Rows and cells may be wrapped in row/cell-level content controls (`w:sdt/w:sdtContent`) or
// custom XML (`w:customXml`); those wrappers are transparent here, and so is
// `mc:AlternateContent` (same branch choice as text.ts). Tree walks use explicit stacks.
//
// Editing keeps the structural invariants Word relies on after every operation:
// - every `w:tc` ends with a `w:p` (looking through cell-level wrappers, like validate.ts);
// - every `w:tr` keeps at least one `w:tc` (a row or table that would become empty is removed);
// - in every row, gridBefore + the cells' gridSpans + gridAfter = the `w:tblGrid` column count
//   (column operations first normalize the table: legacy `w:hMerge` becomes `w:gridSpan`, short
//   rows get an explicit `w:gridAfter`, a missing or short `w:tblGrid` is completed);
// - every vMerge `continue` cell sits under a vMerge cell that starts at the same grid column.
// Bookmark, comment-range, permission and move-range markers whose partner lies outside removed
// content are moved out next to it, so no range is left half-open. Cell widths (`w:tcW` in
// twips) and a twips `w:tblW` follow grid-width changes.

import { UsageError } from '../core/errors.js';
import { ELEMENT_NODE, type XElement, type XNode } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { WordDocument } from './document.js';
import { DocumentIds, type HexIdKind, type NumericIdKind } from './ids.js';
import {
  type ChildOrder,
  ensureChild,
  ensureFirstChild,
  makeOrder,
  pruneEmpty,
  setWAttr,
  wElement,
  wPrefix,
} from './order.js';
import { Paragraph } from './paragraph.js';
import type { Story } from './story.js';
import { chooseAlternate } from './text.js';

/**
 * @internal Elements named `w:<locals>` among `parent`'s children, in document order, looking
 * through `w:sdt/w:sdtContent` and `w:customXml` wrappers and `mc:AlternateContent` (the branch
 * text.ts reads), at any nesting depth, but not into anything else.
 */
export function wrappedChildren(parent: XElement, locals: ReadonlySet<string>): XElement[] {
  const out: XElement[] = [];
  const stack: { kids: readonly XNode[]; i: number }[] = [{ kids: parent.children, i: 0 }];
  while (stack.length > 0) {
    const top = stack[stack.length - 1]!;
    if (top.i >= top.kids.length) {
      stack.pop();
      continue;
    }
    const k = top.kids[top.i++]!;
    if (k.nodeType !== ELEMENT_NODE) continue;
    if (k.uri === NS.w) {
      if (locals.has(k.local)) out.push(k);
      else if (k.local === 'sdt') {
        const content = k.element(NS.w, 'sdtContent');
        if (content !== undefined) stack.push({ kids: content.children, i: 0 });
      } else if (k.local === 'customXml') stack.push({ kids: k.children, i: 0 });
    } else if (k.uri === NS.mc && k.local === 'AlternateContent') {
      const branch = chooseAlternate(k);
      if (branch !== undefined) stack.push({ kids: branch.children, i: 0 });
    }
  }
  return out;
}

/**
 * Widest logical grid computed. Word itself allows 63 columns; the cap only stops a hostile
 * `w:gridSpan="99999999"` from allocating millions of slots.
 */
export const MAX_GRID_COLUMNS = 1024;

const clampCols = (n: number): number => Math.min(MAX_GRID_COLUMNS, n);

/** @internal Integer value of `{w}val` on `el`, or `fallback` when absent or not a number. */
export function intVal(el: XElement | undefined, fallback: number): number {
  const v = el?.attrNS(NS.w, 'val');
  if (v === undefined) return fallback;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : fallback;
}

const TR: ReadonlySet<string> = new Set(['tr']);
const TC: ReadonlySet<string> = new Set(['tc']);
const P: ReadonlySet<string> = new Set(['p']);
const TBL: ReadonlySet<string> = new Set(['tbl']);
const P_OR_TBL: ReadonlySet<string> = new Set(['p', 'tbl']);

/** `w:tblPr` children (CT_TblPr, §17.4.60). */
const TBLPR_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'tblStyle',
  'tblpPr',
  'tblOverlap',
  'bidiVisual',
  'tblStyleRowBandSize',
  'tblStyleColBandSize',
  'tblW',
  'jc',
  'tblCellSpacing',
  'tblInd',
  'tblBorders',
  'shd',
  'tblLayout',
  'tblCellMar',
  'tblLook',
  'tblCaption',
  'tblDescription',
  'tblPrChange',
]);

/** `w:trPr` children, in the order Word writes them (the schema allows any order). */
const TRPR_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'cnfStyle',
  'divId',
  'gridBefore',
  'gridAfter',
  'wBefore',
  'wAfter',
  'cantSplit',
  'trHeight',
  'tblHeader',
  'tblCellSpacing',
  'jc',
  'hidden',
  'ins',
  'del',
  'trPrChange',
]);

/** `w:tcPr` children (CT_TcPr, §17.4.70). */
const TCPR_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'cnfStyle',
  'tcW',
  'gridSpan',
  'hMerge',
  'vMerge',
  'tcBorders',
  'shd',
  'noWrap',
  'tcMar',
  'textDirection',
  'tcFitText',
  'vAlign',
  'hideMark',
  'headers',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tcPrChange',
]);

/** One slot of a table's logical grid (see {@link Table.grid}). */
export interface GridCell {
  /** The cell covering the slot: for merged regions, the top-left (origin) cell. */
  readonly cell: TableCell;
  /** Grid row of the covering cell's top-left slot. */
  readonly row: number;
  /** Grid column of the covering cell's top-left slot. */
  readonly col: number;
  /** Rows covered by the cell (vertical merges). */
  readonly rowSpan: number;
  /** Grid columns covered by the cell. */
  readonly colSpan: number;
  /** True for the cell's top-left slot only. */
  readonly origin: boolean;
}

/** The logical grid of a table: `rows` × `cols` slots. */
export interface TableGrid {
  readonly rows: number;
  readonly cols: number;
  /** The slot at (r, c), or undefined outside the grid or where no cell exists (`gridBefore`/`gridAfter`, short rows). */
  at(r: number, c: number): GridCell | undefined;
}

/** Options of {@link Table.cloneRow}. */
export interface CloneRowOptions {
  /**
   * Where the clone goes: after this row (index or view); -1 puts it first. Default: right
   * after the source row.
   */
  after?: number | TableRow;
}

/** Row height rule (`w:trHeight/@w:hRule`). */
export type RowHeightRule = 'atLeast' | 'exact' | 'auto';

/** Vertical alignment of a cell's content (`w:vAlign`). */
export type CellVerticalAlignment = 'top' | 'center' | 'bottom' | 'both';

/**
 * A table (`w:tbl`). Views are stateless: every call reads the DOM, so a view stays valid
 * across edits made through any other view.
 */
export class Table {
  /** The `w:tbl` element. */
  readonly element: XElement;
  /** The story the table lives in. */
  readonly story: Story;

  constructor(el: XElement, story: Story) {
    this.element = el;
    this.story = story;
  }

  /** Rows (`w:tr`), including rows wrapped in row-level content controls or custom XML. */
  rows(): TableRow[] {
    return wrappedChildren(this.element, TR).map((tr) => new TableRow(tr, this.story));
  }

  /** Number of `w:gridCol` entries in `w:tblGrid`. */
  get gridColumnCount(): number {
    return this.element.element(NS.w, 'tblGrid')?.elements(NS.w, 'gridCol').length ?? 0;
  }

  /** Grid column widths in twips (`w:tblGrid/w:gridCol/@w:w`; 0 where unset). */
  columnWidths(): number[] {
    return gridWidths(this.element);
  }

  /**
   * Computes the logical grid: `w:gridBefore`/`w:gridAfter`, `w:gridSpan`, legacy `w:hMerge`
   * and `w:vMerge` chains (continuation slots map to the origin cell above it). `cols` is the
   * larger of the `w:tblGrid` column count and the widest row. Recomputed on every call.
   */
  grid(): TableGrid {
    return computeGrid(this.rows(), this.gridColumnCount);
  }

  /** The origin cell covering grid slot (r, c), or undefined. */
  cell(r: number, c: number): TableCell | undefined {
    return this.grid().at(r, c)?.cell;
  }

  /** Plain text: rows joined by `\n`, cells by `\t` (cells as {@link TableCell.text}). */
  text(): string {
    return this.rows()
      .map((row) =>
        row
          .cells()
          .map((c) => c.text())
          .join('\t'),
      )
      .join('\n');
  }

  /** Table style id (`w:tblPr/w:tblStyle/@w:val`), undefined when none is set. */
  get style(): string | undefined {
    return this.element.element(NS.w, 'tblPr')?.element(NS.w, 'tblStyle')?.attrNS(NS.w, 'val');
  }

  /** Sets the table style id; undefined removes `w:tblStyle`. */
  set style(id: string | undefined) {
    const tbl = this.element;
    setPropVal(tbl.element(NS.w, 'tblPr'), () => ensureTblPr(tbl), 'tblStyle', TBLPR_ORDER, id);
  }

  // ---- rows ------------------------------------------------------------------------------

  /**
   * Inserts a deep copy of row `index` (after the source row by default) and returns it.
   * The copy gets fresh `w14:paraId`/`w14:textId`, drawing, content-control and revision ids;
   * footnotes and endnotes it references are duplicated so each reference has its own note;
   * bookmark, comment, permission and move-range markers are left out of it. A copied cell
   * that lands strictly inside a vertical merge continues it (its content becomes one empty
   * paragraph); a copied merge cell anywhere else becomes a plain cell.
   *
   * Ids come from {@link DocumentIds}, which re-scans a changed part before handing out ids;
   * when cloning many rows, wrap the loop in `DocumentIds.of(doc).hold()` / release to scan
   * once.
   */
  cloneRow(index: number, o?: CloneRowOptions): TableRow {
    const rows = wrappedChildren(this.element, TR);
    const src = rows[checkIndex(index, rows.length, 'row')]!;
    let after = index;
    if (o?.after !== undefined) {
      after = typeof o.after === 'number' ? o.after : rows.indexOf(o.after.element);
      const bad = typeof o.after === 'number' ? after < -1 : after < 0;
      if (!Number.isInteger(after) || bad || after >= rows.length) {
        throw new UsageError(`cloneRow(): 'after' must be a row of this table or -1`);
      }
    }
    const clone = src.clone(true);
    stripMarkers(clone);
    if (after < 0) rows[0]!.before(clone);
    else rows[after]!.after(clone);
    settleNewRow(this.element, clone);
    const ids = idSource(this.story, this.element);
    try {
      refreshIds(clone, this.story, ids, true);
    } finally {
      ids.release();
    }
    return new TableRow(clone, this.story);
  }

  /**
   * Inserts a new row so that it becomes row `index` (0 … row count) and returns it. Its row
   * properties, cell layout (spans), cell properties (widths, borders, shading, ...) and
   * paragraph formatting are copied from the row above (the first row when inserting at 0);
   * `w:tblHeader` is kept only inside the leading header rows. `cells` fills the new cells in
   * order with one run of text each, formatted like the first run of the cell it was copied
   * from. A cell that lands strictly inside a vertical merge continues it (and gets no text).
   */
  insertRow(index: number, cells?: readonly string[]): TableRow {
    const tbl = this.element;
    const rows = new LazyRows(tbl);
    if (!Number.isInteger(index) || index < 0 || index > rows.length) {
      throw new UsageError(`insertRow(): index ${index} out of range 0..${rows.length}`);
    }
    let tr: XElement;
    if (rows.length === 0) {
      tr = this.#bareRow();
      const grid = tbl.element(NS.w, 'tblGrid') ?? tbl.element(NS.w, 'tblPr');
      if (grid !== undefined) grid.after(tr);
      else tbl.append(tr);
    } else {
      const ref = rows.at(index > 0 ? index - 1 : 0)!;
      tr = wElement(tbl, 'tr');
      const ex = ref.tr.element(NS.w, 'tblPrEx');
      if (ex !== undefined) tr.append(cleanCopy(ex, REVISION_PROPS));
      const trPr = ref.tr.element(NS.w, 'trPr');
      if (trPr !== undefined) {
        const copy = cleanCopy(trPr, ROW_PROP_DROP);
        // A header row may only be followed by header rows: keep the flag inside that block.
        const next = rows.at(index);
        if (next === undefined || !readOnOff(next.tr.element(NS.w, 'trPr'), 'tblHeader')) {
          copy.element(NS.w, 'tblHeader')?.remove();
        }
        if (copy.firstElement() !== undefined || copy._attrs.length > 0) tr.append(copy);
      }
      for (const c of ref.cells) tr.append(newCellLike(c.tc, tbl, true));
      if (index > 0) ref.tr.after(tr);
      else ref.tr.before(tr);
    }
    const fresh = settleNewRow(tbl, tr);
    if (cells !== undefined) {
      const tcs = wrappedChildren(tr, TC);
      for (let i = 0; i < cells.length && i < tcs.length; i++) {
        const tc = tcs[i]!;
        const text = cells[i]!;
        if (text === '' || fresh.continued.has(tc)) continue;
        const p = wrappedChildren(tc, P)[0]!;
        appendFormattedRun(new Paragraph(p, this.story), text, fresh.runProps.get(tc));
      }
    }
    return new TableRow(tr, this.story);
  }

  /**
   * Deletes row `index`. Where the row held the `restart` of a vertical merge that goes on
   * below, the next cell of the merge becomes the restart and receives the merged cell's
   * content. Deleting the only row removes the whole table (see {@link remove}).
   */
  deleteRow(index: number): void {
    const tbl = this.element;
    const rows = new LazyRows(tbl);
    const row = rows.at(checkIndex(index, rows.length, 'row'))!;
    if (rows.length === 1) {
      this.remove();
      return;
    }
    const below = rows.at(index + 1);
    const after = rows.at(index + 2);
    for (const c of row.cells) {
      if (vOf(c.tc) !== 'restart' || below === undefined) continue;
      const next = startingAt(below, c.col);
      if (next === undefined || vOf(next.tc) !== 'continue') continue;
      const further = after === undefined ? undefined : startingAt(after, c.col);
      setV(
        next.tc,
        further !== undefined && vOf(further.tc) === 'continue' ? 'restart' : undefined,
      );
      replaceContent(c.tc, next.tc);
    }
    removeWithMarkers(row.tr);
    repairVMerge(tbl, index, index);
  }

  // ---- columns ---------------------------------------------------------------------------

  /**
   * Inserts a grid column so that it becomes column `index` (0 … column count) and returns the
   * cells created for it. With `width` (twips) the new column has that width and the table
   * grows; without it, the new column takes half of its neighbour's width (column `index`, or
   * the last column when appending), so the table keeps its width. Cells spanning across the
   * insertion point grow by one column; other rows get a new cell (formatted like its left
   * neighbour, else its right one) or, where the row skips that part of the grid, a larger
   * `w:gridBefore`/`w:gridAfter`.
   */
  insertColumn(index: number, o?: { width?: number }): TableCell[] {
    const tbl = this.element;
    if (o?.width !== undefined) checkTwips(o.width, 'insertColumn(): width');
    const rows = this.#normalize();
    const cols = this.gridColumnCount;
    if (!Number.isInteger(index) || index < 0 || index > cols) {
      throw new UsageError(`insertColumn(): index ${index} out of range 0..${cols}`);
    }
    const widths = gridWidths(tbl);
    const before = sum(widths);
    const changed = new Set<number>([index]);
    let width = o?.width;
    if (width === undefined) {
      const k = index < cols ? index : cols - 1;
      const w = widths[k] ?? 0;
      if (w >= 2) {
        width = Math.floor(w / 2);
        widths[k] = w - width;
        changed.add(k >= index ? k + 1 : k);
      } else {
        width = DEFAULT_COLUMN_WIDTH;
      }
    }
    widths.splice(index, 0, width);

    const spanned = new Set<XElement>();
    const created: XElement[] = [];
    for (const row of rows) {
      if (index < row.before || row.cells.length === 0) {
        setTrPrVal(row.tr, 'gridBefore', String(row.before + 1));
        continue;
      }
      if (index > row.end) {
        setTrPrVal(row.tr, 'gridAfter', String(row.after + 1));
        continue;
      }
      const crossing = row.cells.find((c) => c.col < index && index < c.col + c.span);
      if (crossing !== undefined) {
        setSpan(crossing.tc, crossing.span + 1);
        spanned.add(crossing.tc);
        continue;
      }
      const prev = row.cells.find((c) => c.col + c.span === index);
      const next = row.cells.find((c) => c.col === index);
      const tc = newCellLike((prev ?? next)!.tc, tbl, false);
      if (prev !== undefined) levelNode(prev.tc, row.tr).after(tc);
      else levelNode(next!.tc, row.tr).before(tc);
      created.push(tc);
      spanned.add(tc);
    }
    writeGridWidths(tbl, widths);
    syncCellWidths(tbl, widths, spanned, changed);
    adjustTblW(tbl, before, sum(widths));
    return created.map((tc) => new TableCell(tc, this.story));
  }

  /**
   * Deletes grid column `index`. Cells spanning it shrink by one column, cells only in it are
   * removed (with their content), rows that skip it get a smaller `w:gridBefore`/`w:gridAfter`.
   * A row whose only cell is removed is deleted; deleting the last column removes the table.
   */
  deleteColumn(index: number): void {
    const tbl = this.element;
    const rows = this.#normalize();
    const cols = this.gridColumnCount;
    checkIndex(index, cols, 'column');
    if (cols === 1) {
      this.remove();
      return;
    }
    const widths = gridWidths(tbl);
    const before = sum(widths);
    widths.splice(index, 1);
    const spanned = new Set<XElement>();
    const emptied: XElement[] = [];
    for (const row of rows) {
      if (index < row.before) {
        setTrPrVal(row.tr, 'gridBefore', row.before > 1 ? String(row.before - 1) : undefined);
        continue;
      }
      if (index >= row.end) {
        setTrPrVal(row.tr, 'gridAfter', row.after > 1 ? String(row.after - 1) : undefined);
        continue;
      }
      const c = coveringAt(row, index)!;
      if (c.span > 1) {
        setSpan(c.tc, c.span - 1);
        spanned.add(c.tc);
      } else if (row.cells.length === 1) {
        emptied.push(row.tr);
      } else {
        removeWithMarkers(c.tc);
      }
    }
    for (const tr of emptied) removeWithMarkers(tr);
    if (wrappedChildren(tbl, TR).length === 0) {
      this.remove();
      return;
    }
    writeGridWidths(tbl, widths);
    syncCellWidths(tbl, widths, spanned, new Set());
    adjustTblW(tbl, before, sum(widths));
    repairVMerge(tbl);
  }

  // ---- merging ---------------------------------------------------------------------------

  /**
   * Merges the rectangle of grid slots between (r1, c1) and (r2, c2), inclusive, into one cell
   * and returns it: one `w:gridSpan` cell per row, chained with `w:vMerge` restart/continue
   * when the rectangle spans several rows. Like a selection in Word, the rectangle grows to
   * take in every merged cell it touches. The content of the covered cells moves into the
   * top-left cell (empty cells are skipped); covered cells keep a single empty paragraph.
   * Throws UsageError when the rectangle leaves the grid or includes a slot without a cell
   * (`w:gridBefore`/`w:gridAfter`).
   */
  merge(r1: number, c1: number, r2: number, c2: number): TableCell {
    const tbl = this.element;
    for (const v of [r1, c1, r2, c2]) {
      if (!Number.isInteger(v)) throw new UsageError('merge(): coordinates must be integers');
    }
    let top = Math.min(r1, r2);
    let bottom = Math.max(r1, r2);
    let left = Math.min(c1, c2);
    let right = Math.max(c1, c2);
    let rows = model(tbl);
    const cols = Math.max(this.gridColumnCount, ...rows.map((r) => r.end + r.after));
    if (top < 0 || bottom >= rows.length || left < 0 || right >= cols) {
      throw new UsageError(
        `merge(): (${r1},${c1})-(${r2},${c2}) is outside the ${rows.length}x${cols} grid`,
      );
    }
    rows = this.#normalize();
    // Grow the rectangle until no cell (or vertical merge) crosses its border.
    for (let grown = true; grown; ) {
      grown = false;
      for (let r = top; r <= bottom; r++) {
        const row = rows[r]!;
        if (row.before > left || row.end <= right) {
          throw new UsageError(
            `merge(): row ${r} has no cell in some of columns ${left}..${right}`,
          );
        }
        for (const c of row.cells) {
          if (c.col + c.span <= left || c.col > right) continue;
          const [t, b] = chainExtent(rows, r, c.col);
          if (c.col < left || c.col + c.span - 1 > right || t < top || b > bottom) {
            left = Math.min(left, c.col);
            right = Math.max(right, c.col + c.span - 1);
            top = Math.min(top, t);
            bottom = Math.max(bottom, b);
            grown = true;
          }
        }
      }
    }
    const lines = rows
      .slice(top, bottom + 1)
      .map((row) => row.cells.filter((c) => c.col >= left && c.col + c.span - 1 <= right));
    const origin = lines[0]![0]!.tc;
    if (lines.length === 1 && lines[0]!.length === 1 && vOf(origin) === undefined) {
      return new TableCell(origin, this.story);
    }
    // Content first: everything non-empty moves into the origin, in reading order.
    const full = lines.flat().filter((c) => !isEmptyCell(c.tc));
    if (full.length > 0 && !(full.length === 1 && full[0]!.tc === origin)) {
      if (isEmptyCell(origin)) clearContent(origin);
      for (const c of full) if (c.tc !== origin) moveContent(c.tc, origin);
    }
    const widths = gridWidths(tbl);
    const span = right - left + 1;
    const width = sum(widths.slice(left, right + 1));
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i]!;
      const keep = line[0]!.tc;
      for (const c of line.slice(1)) removeWithMarkers(c.tc);
      setSpan(keep, span);
      syncWidth(keep, width);
      setV(keep, lines.length === 1 ? undefined : i === 0 ? 'restart' : 'continue');
      if (i > 0) resetToEmptyParagraph(keep);
    }
    ensureEndsWithParagraph(origin);
    repairVMerge(tbl);
    return new TableCell(origin, this.story);
  }

  /**
   * Splits the merged cell covering slot (r, c) back into one cell per grid slot and returns
   * the top-left one, which keeps the content. New cells copy the cell properties of the
   * piece they were split from. A cell that is not merged is left alone.
   */
  unmerge(r: number, c: number): TableCell {
    const tbl = this.element;
    const rows = this.#normalize();
    const row = rows[checkIndex(r, rows.length, 'row')]!;
    const hit = Number.isInteger(c) ? coveringAt(row, c) : undefined;
    if (hit === undefined) throw new UsageError(`unmerge(): no cell at (${r}, ${c})`);
    const [top, bottom] = chainExtent(rows, r, hit.col);
    const widths = gridWidths(tbl);
    let origin: XElement | undefined;
    for (let i = top; i <= bottom; i++) {
      const m = startingAt(rows[i]!, hit.col)!;
      origin ??= m.tc;
      setV(m.tc, undefined);
      if (m.span === 1) continue;
      setSpan(m.tc, 1);
      syncWidth(m.tc, widths[m.col] ?? 0);
      let at = levelNode(m.tc, rows[i]!.tr);
      for (let k = 1; k < m.span; k++) {
        const tc = newCellLike(m.tc, tbl, false);
        syncWidth(tc, widths[m.col + k] ?? 0);
        at.after(tc);
        at = tc;
      }
    }
    repairVMerge(tbl);
    return new TableCell(origin!, this.story);
  }

  // ---- grid ------------------------------------------------------------------------------

  /**
   * Rebuilds `w:tblGrid` from the cells: as many columns as the widest row needs, widths taken
   * from cell widths in twips where they determine a column (single-column cells first, then
   * what multi-column cells imply), else the previous grid width, else an even share. Short
   * rows get an explicit `w:gridAfter`. Returns the new widths.
   */
  recomputeGrid(): number[] {
    const tbl = this.element;
    const rows = model(tbl);
    const old = gridWidths(tbl);
    const cols = rows.length === 0 ? old.length : Math.max(...rows.map((r) => r.end + r.after));
    const widths = inferWidths(rows, cols, old, true);
    writeGridWidths(tbl, widths);
    for (const row of rows) {
      if (row.end + row.after < cols) setTrPrVal(row.tr, 'gridAfter', String(cols - row.end));
    }
    adjustTblW(tbl, sum(old), sum(widths));
    return widths;
  }

  /**
   * Sets grid column `index` to `twips`, updating the twips widths of the cells over it and a
   * twips table width.
   */
  setColumnWidth(index: number, twips: number): void {
    const tbl = this.element;
    checkIndex(index, this.gridColumnCount, 'column');
    checkTwips(twips, 'setColumnWidth()');
    const widths = gridWidths(tbl);
    const before = sum(widths);
    widths[index] = twips;
    writeGridWidths(tbl, widths);
    syncCellWidths(tbl, widths, new Set(), new Set([index]));
    adjustTblW(tbl, before, sum(widths));
  }

  /**
   * Removes the table from its story, together with a content control or custom XML wrapper
   * left empty. Range markers paired with something outside the table are kept in its place;
   * a cell it was in still ends with a paragraph, and a story it leaves empty gets one.
   */
  remove(): void {
    const parent = removeWithMarkers(this.element);
    if (parent === undefined) return;
    const tc = parent.is(NS.w, 'tc') ? parent : parent.closest(NS.w, 'tc');
    if (tc !== undefined) {
      ensureEndsWithParagraph(tc);
      return;
    }
    if (STORY_CONTAINERS.has(parent.local) && lastBlock(parent) === undefined) {
      const p = wElement(parent, 'p');
      const last = parent.lastElement();
      if (last?.is(NS.w, 'sectPr')) parent.insertBefore(p, last);
      else parent.append(p);
    }
  }

  // ---- internals -------------------------------------------------------------------------

  /** A row of plain cells over the whole grid (for a table without rows). */
  #bareRow(): XElement {
    const tbl = this.element;
    const tr = wElement(tbl, 'tr');
    const widths = gridWidths(tbl);
    for (const w of widths.length > 0 ? widths : [DEFAULT_COLUMN_WIDTH]) {
      tr.append(newCell(tbl, w));
    }
    return tr;
  }

  /**
   * Brings the table into the canonical shape column operations rely on (see the header):
   * legacy `w:hMerge` → `w:gridSpan`, a `w:tblGrid` covering every row, explicit `w:gridAfter`
   * on short rows. Only what is needed changes; the rendering stays the same.
   */
  #normalize(): MRow[] {
    const tbl = this.element;
    for (const tr of wrappedChildren(tbl, TR)) {
      let origin: XElement | undefined;
      let span = 0;
      for (const tc of wrappedChildren(tr, TC)) {
        const h = tc.element(NS.w, 'tcPr')?.element(NS.w, 'hMerge');
        if (h !== undefined && mergeState(h) === 'continue' && origin !== undefined) {
          span += spanOf(tc);
          setSpan(origin, span);
          if (!isEmptyCell(tc)) moveContent(tc, origin);
          removeWithMarkers(tc);
          continue;
        }
        if (h !== undefined) setTcPrVal(tc, 'hMerge', undefined);
        origin = tc;
        span = spanOf(tc);
      }
    }
    let rows = model(tbl);
    const have = this.gridColumnCount;
    const cols = Math.max(have, ...rows.map((r) => r.end + r.after));
    if (tbl.element(NS.w, 'tblGrid') === undefined || cols > have) {
      writeGridWidths(tbl, inferWidths(rows, cols, gridWidths(tbl), false));
    }
    let padded = false;
    for (const row of rows) {
      if (row.end + row.after < cols && row.cells.length > 0) {
        setTrPrVal(row.tr, 'gridAfter', String(cols - row.end));
        padded = true;
      }
    }
    if (padded) rows = model(tbl);
    return rows;
  }
}

/** A table row (`w:tr`). */
export class TableRow {
  /** The `w:tr` element. */
  readonly element: XElement;
  readonly story: Story;

  constructor(el: XElement, story: Story) {
    this.element = el;
    this.story = story;
  }

  /** Cells (`w:tc`), including cells wrapped in cell-level content controls or custom XML. */
  cells(): TableCell[] {
    return wrappedChildren(this.element, TC).map((tc) => new TableCell(tc, this.story));
  }

  /** The table the row belongs to. Throws UsageError for a detached row. */
  get table(): Table {
    const tbl = this.element.closest(NS.w, 'tbl');
    if (tbl === undefined) throw new UsageError('The row is not in a table');
    return new Table(tbl, this.story);
  }

  /** Position of the row in {@link Table.rows}. */
  get index(): number {
    return wrappedChildren(this.table.element, TR).indexOf(this.element);
  }

  /** Grid columns skipped before the first cell (`w:trPr/w:gridBefore`). */
  get gridBefore(): number {
    return clampCols(
      Math.max(0, intVal(this.element.element(NS.w, 'trPr')?.element(NS.w, 'gridBefore'), 0)),
    );
  }

  /** Grid columns left empty after the last cell (`w:trPr/w:gridAfter`). */
  get gridAfter(): number {
    return clampCols(
      Math.max(0, intVal(this.element.element(NS.w, 'trPr')?.element(NS.w, 'gridAfter'), 0)),
    );
  }

  /** Row height in twips (`w:trHeight/@w:val`), undefined when not set. */
  get height(): number | undefined {
    const v = this.element.element(NS.w, 'trPr')?.element(NS.w, 'trHeight')?.attrNS(NS.w, 'val');
    const n = v === undefined ? Number.NaN : Number.parseInt(v, 10);
    return Number.isFinite(n) ? n : undefined;
  }

  /**
   * Sets the row height in twips (an existing `w:hRule` is kept; without one Word treats the
   * height as a minimum). undefined removes `w:trHeight`.
   */
  set height(twips: number | undefined) {
    if (twips === undefined) {
      setTrPrVal(this.element, 'trHeight', undefined);
      return;
    }
    checkTwips(twips, 'height');
    const el = ensureChild(ensureTrPr(this.element), 'trHeight', TRPR_ORDER);
    setWAttr(el, 'val', String(twips));
  }

  /** `w:trHeight/@w:hRule`, undefined when not set. */
  get heightRule(): RowHeightRule | undefined {
    const v = this.element.element(NS.w, 'trPr')?.element(NS.w, 'trHeight')?.attrNS(NS.w, 'hRule');
    return v === 'atLeast' || v === 'exact' || v === 'auto' ? v : undefined;
  }

  /** Sets `w:hRule` (creating `w:trHeight` when needed); undefined removes the attribute. */
  set heightRule(rule: RowHeightRule | undefined) {
    const trPr = this.element.element(NS.w, 'trPr');
    if (rule === undefined) {
      const el = trPr?.element(NS.w, 'trHeight');
      if (el === undefined) return;
      el.removeAttrNS(NS.w, 'hRule');
      if (el.attrNS(NS.w, 'val') === undefined) setTrPrVal(this.element, 'trHeight', undefined);
      return;
    }
    const el = ensureChild(ensureTrPr(this.element), 'trHeight', TRPR_ORDER);
    setWAttr(el, 'hRule', rule);
  }

  /** True when the row repeats as a header row on every page (`w:tblHeader`). */
  get header(): boolean {
    return readOnOff(this.element.element(NS.w, 'trPr'), 'tblHeader');
  }

  /** Writes `<w:tblHeader/>` or removes it. Word repeats only the table's leading header rows. */
  set header(on: boolean) {
    setOnOffProp(this.element, 'tblHeader', on);
  }

  /** True when the row may not break across pages (`w:cantSplit`). */
  get cantSplit(): boolean {
    return readOnOff(this.element.element(NS.w, 'trPr'), 'cantSplit');
  }

  /** Writes `<w:cantSplit/>` or removes it. */
  set cantSplit(on: boolean) {
    setOnOffProp(this.element, 'cantSplit', on);
  }

  /** Deletes the row (see {@link Table.deleteRow}). */
  remove(): void {
    const table = this.table;
    table.deleteRow(wrappedChildren(table.element, TR).indexOf(this.element));
  }
}

/** A table cell (`w:tc`). */
export class TableCell {
  /** The `w:tc` element. */
  readonly element: XElement;
  readonly story: Story;

  constructor(el: XElement, story: Story) {
    this.element = el;
    this.story = story;
  }

  /** The row the cell belongs to. Throws UsageError for a detached cell. */
  get row(): TableRow {
    const tr = this.element.closest(NS.w, 'tr');
    if (tr === undefined) throw new UsageError('The cell is not in a table row');
    return new TableRow(tr, this.story);
  }

  /** Block-level paragraphs of the cell (looking through content controls / custom XML). */
  paragraphs(): Paragraph[] {
    return wrappedChildren(this.element, P).map((p) => new Paragraph(p, this.story));
  }

  /** Tables nested directly in the cell (not tables nested inside those). */
  tables(): Table[] {
    return wrappedChildren(this.element, TBL).map((t) => new Table(t, this.story));
  }

  /** Plain text: the cell's paragraphs (and nested tables' text) joined by `\n`. */
  text(): string {
    return wrappedChildren(this.element, P_OR_TBL)
      .map((e) =>
        e.local === 'p' ? new Paragraph(e, this.story).text : new Table(e, this.story).text(),
      )
      .join('\n');
  }

  /**
   * Replaces the cell's content with one paragraph holding `text` (TAB/LF handled as in
   * {@link Paragraph.appendRun}) and returns it. The paragraph keeps the formatting of the
   * cell's first paragraph and the run that of its first run. Range markers paired with
   * something outside the cell stay in the cell.
   */
  setText(text: string): Paragraph {
    const tc = this.element;
    const first = wrappedChildren(tc, P)[0];
    const pPr = first?.element(NS.w, 'pPr');
    const firstRun = first === undefined ? undefined : new Paragraph(first, this.story).runs()[0];
    const rPr = firstRun?.element.element(NS.w, 'rPr');
    const p = wElement(tc, 'p');
    if (pPr !== undefined) p.append(cleanCopy(pPr, PARA_PROP_DROP));
    const runProps = rPr === undefined ? undefined : cleanCopy(rPr, RUN_PROP_DROP);
    const kept = outsideMarkers(contentNodes(tc));
    clearContent(tc);
    if (kept.length > 0) tc.append(...kept);
    tc.append(p);
    const para = new Paragraph(p, this.story);
    if (text !== '') appendFormattedRun(para, text, runProps);
    return para;
  }

  /** Appends a paragraph at the end of the cell, optionally with text and a paragraph style. */
  appendParagraph(text?: string, o?: { style?: string }): Paragraph {
    const p = wElement(this.element, 'p');
    this.element.append(p);
    const para = new Paragraph(p, this.story);
    if (o?.style !== undefined) para.style = o.style;
    if (text !== undefined && text !== '') para.appendRun(text);
    return para;
  }

  /** Grid columns spanned (`w:tcPr/w:gridSpan`, default 1). */
  get gridSpan(): number {
    return spanOf(this.element);
  }

  /** Vertical merge state (`w:tcPr/w:vMerge`; a missing `w:val` means `continue`). */
  get vMerge(): 'restart' | 'continue' | undefined {
    return vOf(this.element);
  }

  /** Legacy horizontal merge state (`w:tcPr/w:hMerge`; a missing `w:val` means `continue`). */
  get hMerge(): 'restart' | 'continue' | undefined {
    return mergeState(this.#tcPr()?.element(NS.w, 'hMerge'));
  }

  /** Background fill (`w:shd/@w:fill`, `RRGGBB` or `auto`), undefined when not set. */
  get shading(): string | undefined {
    return this.#tcPr()?.element(NS.w, 'shd')?.attrNS(NS.w, 'fill');
  }

  /** Sets a solid background (`<w:shd w:val="clear" w:color="auto" w:fill="…"/>`); undefined removes it. */
  set shading(fill: string | undefined) {
    if (fill === undefined) {
      setTcPrVal(this.element, 'shd', undefined);
      return;
    }
    if (!/^([0-9A-Fa-f]{6}|auto)$/.test(fill)) {
      throw new UsageError(`shading: expected RRGGBB hex or 'auto', got '${fill}'`);
    }
    const shd = ensureChild(ensureFirstChild(this.element, 'tcPr'), 'shd', TCPR_ORDER);
    setWAttr(shd, 'val', 'clear');
    setWAttr(shd, 'color', 'auto');
    setWAttr(shd, 'fill', fill);
  }

  /** Vertical alignment (`w:vAlign`), undefined when not set (Word's default is top). */
  get vAlign(): CellVerticalAlignment | undefined {
    const v = this.#tcPr()?.element(NS.w, 'vAlign')?.attrNS(NS.w, 'val');
    return v === 'top' || v === 'center' || v === 'bottom' || v === 'both' ? v : undefined;
  }

  /** Sets `w:vAlign`; undefined removes it. */
  set vAlign(v: CellVerticalAlignment | undefined) {
    setTcPrVal(this.element, 'vAlign', v);
  }

  /** Preferred width in twips (`w:tcW` of type `dxa`), undefined when unset or not in twips. */
  get width(): number | undefined {
    return dxaWidth(this.#tcPr()?.element(NS.w, 'tcW'));
  }

  /**
   * Sets the preferred width in twips (`<w:tcW w:w="…" w:type="dxa"/>`); undefined removes it.
   * The grid is not changed: call {@link Table.recomputeGrid} after changing cell widths.
   */
  set width(twips: number | undefined) {
    if (twips === undefined) {
      setTcPrVal(this.element, 'tcW', undefined);
      return;
    }
    checkTwips(twips, 'width');
    setTcW(this.element, twips);
  }

  #tcPr(): XElement | undefined {
    return this.element.element(NS.w, 'tcPr');
  }
}

/**
 * A new `w:tbl` (not yet inserted) with `rows` × `cols` empty cells: a `w:tblGrid` (equal
 * widths summing to `width` twips, default 9360 = 6.5", or explicit `widths`), `w:tblStyle`
 * when `style` is given and single borders all round otherwise, a twips `w:tblW` and Word's
 * default `w:tblLook`. Elements are named with the prefix `story`'s container binds to
 * WordprocessingML. A table must not be the last block of a cell or story: insert it before
 * a paragraph.
 */
export function newTable(
  story: Story,
  rows: number,
  cols: number,
  o?: { widths?: readonly number[]; width?: number; style?: string },
): XElement {
  if (!Number.isInteger(rows) || rows < 1) throw new UsageError('newTable(): rows must be >= 1');
  if (!Number.isInteger(cols) || cols < 1 || cols > MAX_GRID_COLUMNS) {
    throw new UsageError(`newTable(): cols must be in 1..${MAX_GRID_COLUMNS}`);
  }
  let widths: number[];
  if (o?.widths !== undefined) {
    if (o.widths.length !== cols) {
      throw new UsageError(`newTable(): ${o.widths.length} widths for ${cols} columns`);
    }
    for (const w of o.widths) checkTwips(w, 'newTable(): width');
    widths = o.widths.slice();
  } else {
    const total = o?.width ?? DEFAULT_TABLE_WIDTH;
    checkTwips(total, 'newTable(): width');
    const each = Math.floor(total / cols);
    // Word puts the rounding remainder on the last columns (3080, 3081, 3081).
    widths = Array.from({ length: cols }, (_, i) => each + (i >= cols - (total % cols) ? 1 : 0));
  }
  const scope = story.element;
  const el = (local: string, attrs?: Record<string, string>): XElement =>
    wElement(scope, local, attrs);
  const tbl = el('tbl');
  const tblPr = el('tblPr');
  if (o?.style !== undefined) tblPr.append(el('tblStyle', { val: o.style }));
  tblPr.append(el('tblW', { w: String(sum(widths)), type: 'dxa' }));
  if (o?.style === undefined) {
    const borders = el('tblBorders');
    for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
      borders.append(el(side, { val: 'single', sz: '4', space: '0', color: 'auto' }));
    }
    tblPr.append(borders);
  }
  // Strict OOXML dropped the hex `w:val` of tblLook; transitional Word writes both forms.
  const strict = scope.lookupNamespace(wPrefix(scope)) === STRICT_W;
  tblPr.append(
    el('tblLook', {
      ...(strict ? {} : { val: '04A0' }),
      firstRow: '1',
      lastRow: '0',
      firstColumn: '1',
      lastColumn: '0',
      noHBand: '0',
      noVBand: '1',
    }),
  );
  const grid = el('tblGrid');
  for (const w of widths) grid.append(el('gridCol', { w: String(w) }));
  tbl.append(tblPr, grid);
  for (let r = 0; r < rows; r++) {
    const tr = el('tr');
    for (const w of widths) tr.append(newCell(scope, w));
    tbl.append(tr);
  }
  return tbl;
}

// ---- the editing model -------------------------------------------------------------------

const DEFAULT_TABLE_WIDTH = 9360;
const STRICT_W = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
const DEFAULT_COLUMN_WIDTH = 1440;

/** A cell placed in its row: first grid column, span, vertical merge state when modelled. */
interface MCell {
  readonly tc: XElement;
  readonly col: number;
  readonly span: number;
}

/** A row: its cells, `gridBefore`, `gridAfter`, and `end` = the column after its last cell. */
interface MRow {
  readonly tr: XElement;
  readonly before: number;
  readonly after: number;
  readonly end: number;
  readonly cells: readonly MCell[];
}

/** One row as editing sees it (legacy hMerge cells count as plain cells). */
function rowModel(tr: XElement): MRow {
  const trPr = tr.element(NS.w, 'trPr');
  const before = clampCols(Math.max(0, intVal(trPr?.element(NS.w, 'gridBefore'), 0)));
  const after = clampCols(Math.max(0, intVal(trPr?.element(NS.w, 'gridAfter'), 0)));
  let col = before;
  const cells: MCell[] = [];
  for (const tc of wrappedChildren(tr, TC)) {
    const span = spanOf(tc);
    cells.push({ tc, col, span });
    col += span;
  }
  return { tr, before, after, end: col, cells };
}

/** Every row of the table, modelled. */
function model(tbl: XElement): MRow[] {
  return wrappedChildren(tbl, TR).map(rowModel);
}

/**
 * The table's rows, each modelled on first access: row operations only look at the few rows
 * around the change, so they stay cheap on long tables. Merge states are always read live;
 * the cached geometry stays valid as long as cells and spans do not change.
 */
class LazyRows {
  readonly trs: readonly XElement[];
  readonly #cache: (MRow | undefined)[] = [];

  constructor(tbl: XElement) {
    this.trs = wrappedChildren(tbl, TR);
  }

  get length(): number {
    return this.trs.length;
  }

  at(i: number): MRow | undefined {
    const tr = this.trs[i];
    if (tr === undefined || i < 0) return undefined;
    let row = this.#cache[i];
    if (row === undefined) {
      row = rowModel(tr);
      this.#cache[i] = row;
    }
    return row;
  }
}

const startingAt = (row: MRow, col: number): MCell | undefined =>
  row.cells.find((c) => c.col === col);

const coveringAt = (row: MRow, col: number): MCell | undefined =>
  row.cells.find((c) => c.col <= col && col < c.col + c.span);

/** First and last row of the vertical merge through the cell starting at (r, col). */
function chainExtent(rows: readonly MRow[], r: number, col: number): [number, number] {
  let top = r;
  while (top > 0) {
    const here = startingAt(rows[top]!, col);
    const up = startingAt(rows[top - 1]!, col);
    if (here === undefined || vOf(here.tc) !== 'continue' || up === undefined) break;
    if (vOf(up.tc) === undefined) break;
    top--;
  }
  let bottom = r;
  while (bottom + 1 < rows.length) {
    const here = startingAt(rows[bottom]!, col);
    const down = startingAt(rows[bottom + 1]!, col);
    if (here === undefined || vOf(here.tc) === undefined) break;
    if (down === undefined || vOf(down.tc) !== 'continue') break;
    bottom++;
  }
  return [top, bottom];
}

/**
 * Makes every vMerge `continue` cell sit under a vMerge cell starting at the same column: an
 * orphan becomes the restart of what continues below it, or a plain cell. A fix never orphans
 * another cell, so after a local change only the rows `from`..`to` need checking.
 */
function repairVMerge(tbl: XElement, from = 0, to = Number.POSITIVE_INFINITY): void {
  const rows = new LazyRows(tbl);
  for (let r = Math.max(0, from); r < rows.length && r <= to; r++) {
    for (const c of rows.at(r)!.cells) {
      if (vOf(c.tc) !== 'continue') continue;
      const above = rows.at(r - 1);
      const up = above === undefined ? undefined : startingAt(above, c.col);
      if (up !== undefined && vOf(up.tc) !== undefined) continue;
      const below = rows.at(r + 1);
      const down = below === undefined ? undefined : startingAt(below, c.col);
      setV(c.tc, down !== undefined && vOf(down.tc) === 'continue' ? 'restart' : undefined);
    }
  }
}

/**
 * After a row was inserted or cloned: its cells continue a vertical merge exactly when they
 * sit strictly inside one with the same columns (merged cell above, continuation below);
 * otherwise they are plain, and a merge the row cuts through restarts below it.
 * Continued cells keep a single empty paragraph. Returns the continued cells and, per new
 * cell, the run properties of the first run in the cell it lines up with in the row it was
 * modelled on (for text added to inserted rows).
 */
function settleNewRow(
  tbl: XElement,
  tr: XElement,
): { continued: Set<XElement>; runProps: Map<XElement, XElement> } {
  const rows = new LazyRows(tbl);
  const i = rows.trs.indexOf(tr);
  const row = rows.at(i)!;
  const above = rows.at(i - 1);
  const below = rows.at(i + 1);
  const continued = new Set<XElement>();
  const runProps = new Map<XElement, XElement>();
  for (const c of row.cells) {
    const up = above === undefined ? undefined : startingAt(above, c.col);
    const down = below === undefined ? undefined : startingAt(below, c.col);
    const inside =
      up !== undefined &&
      up.span === c.span &&
      vOf(up.tc) !== undefined &&
      down !== undefined &&
      down.span === c.span &&
      vOf(down.tc) === 'continue';
    if (inside) {
      if (vOf(c.tc) !== 'continue' || !isEmptyCell(c.tc)) resetToEmptyParagraph(c.tc);
      setV(c.tc, 'continue');
      continued.add(c.tc);
    } else {
      setV(c.tc, undefined);
    }
    // Run formatting for inserted text: the first run of the neighbour cell the new cell
    // copied (same column in the reference row).
    const ref = i > 0 ? above : below;
    const src = ref === undefined ? undefined : startingAt(ref, c.col);
    const firstP = src === undefined ? undefined : wrappedChildren(src.tc, P)[0];
    const rPr = firstRunProps(firstP);
    if (rPr !== undefined) runProps.set(c.tc, cleanCopy(rPr, RUN_PROP_DROP));
  }
  repairVMerge(tbl, i, i + 1);
  return { continued, runProps };
}

function firstRunProps(p: XElement | undefined): XElement | undefined {
  return p?.descendants(NS.w, 'r')[0]?.element(NS.w, 'rPr');
}

// ---- cells, rows and their properties ----------------------------------------------------

function spanOf(tc: XElement): number {
  return clampCols(Math.max(1, intVal(tc.element(NS.w, 'tcPr')?.element(NS.w, 'gridSpan'), 1)));
}

function vOf(tc: XElement): 'restart' | 'continue' | undefined {
  return mergeState(tc.element(NS.w, 'tcPr')?.element(NS.w, 'vMerge'));
}

function mergeState(el: XElement | undefined): 'restart' | 'continue' | undefined {
  if (el === undefined) return undefined;
  return el.attrNS(NS.w, 'val') === 'restart' ? 'restart' : 'continue';
}

/** Sets or removes `w:<props>/w:<local>/@w:val` of `owner` (props created/pruned as needed). */
function setPropVal(
  props: XElement | undefined,
  ensure: () => XElement,
  local: string,
  order: ChildOrder,
  val: string | undefined,
): void {
  if (val === undefined) {
    const el = props?.element(NS.w, local);
    if (props !== undefined && el !== undefined) {
      el.remove();
      pruneEmpty(props);
    }
    return;
  }
  setWAttr(ensureChild(ensure(), local, order), 'val', val);
}

function setTcPrVal(tc: XElement, local: string, val: string | undefined): void {
  setPropVal(tc.element(NS.w, 'tcPr'), () => ensureFirstChild(tc, 'tcPr'), local, TCPR_ORDER, val);
}

function setTrPrVal(tr: XElement, local: string, val: string | undefined): void {
  setPropVal(tr.element(NS.w, 'trPr'), () => ensureTrPr(tr), local, TRPR_ORDER, val);
}

/** `w:trPr` of a row, created after `w:tblPrEx` (which the schema puts first) when missing. */
function ensureTrPr(tr: XElement): XElement {
  const found = tr.element(NS.w, 'trPr');
  if (found !== undefined) return found;
  const el = wElement(tr, 'trPr');
  const ex = tr.element(NS.w, 'tblPrEx');
  const first = tr.firstElement();
  if (ex !== undefined) ex.after(el);
  else if (first !== undefined) tr.insertBefore(el, first);
  else tr.append(el);
  return el;
}

function setSpan(tc: XElement, n: number): void {
  setTcPrVal(tc, 'gridSpan', n > 1 ? String(n) : undefined);
}

function setV(tc: XElement, v: 'restart' | 'continue' | undefined): void {
  if (v === vOf(tc) && v !== 'restart') return;
  if (v === undefined) {
    setTcPrVal(tc, 'vMerge', undefined);
    return;
  }
  const el = ensureChild(ensureFirstChild(tc, 'tcPr'), 'vMerge', TCPR_ORDER);
  if (v === 'restart') setWAttr(el, 'val', 'restart');
  else el.removeAttrNS(NS.w, 'val'); // Word writes continuation cells as <w:vMerge/>
}

/** ST_OnOff child `w:<local>` of `pr` is present and on. */
function readOnOff(pr: XElement | undefined, local: string): boolean {
  const el = pr?.element(NS.w, local);
  if (el === undefined) return false;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined || !(v === 'false' || v === '0' || v === 'off');
}

function setOnOffProp(tr: XElement, local: string, on: boolean): void {
  if (!on) {
    setTrPrVal(tr, local, undefined);
    return;
  }
  ensureChild(ensureTrPr(tr), local, TRPR_ORDER).removeAttrNS(NS.w, 'val');
}

/** Width in twips of a `w:tcW`/`w:tblW` whose type is `dxa` (the default), else undefined. */
function dxaWidth(el: XElement | undefined): number | undefined {
  if (el === undefined) return undefined;
  const type = el.attrNS(NS.w, 'type') ?? 'dxa';
  const w = el.attrNS(NS.w, 'w');
  if (type !== 'dxa' || w === undefined) return undefined;
  const n = Number.parseInt(w, 10);
  return Number.isFinite(n) ? n : undefined;
}

function setTcW(tc: XElement, twips: number): void {
  const el = ensureChild(ensureFirstChild(tc, 'tcPr'), 'tcW', TCPR_ORDER);
  setWAttr(el, 'w', String(twips));
  setWAttr(el, 'type', 'dxa');
}

/** Updates a cell's twips width (only when it has one: percentages and auto stay). */
function syncWidth(tc: XElement, twips: number): void {
  const tcW = tc.element(NS.w, 'tcPr')?.element(NS.w, 'tcW');
  if (dxaWidth(tcW) !== undefined && twips > 0) setWAttr(tcW!, 'w', String(twips));
}

/** A new empty cell `<w:tc><w:tcPr><w:tcW/></w:tcPr><w:p/></w:tc>` named in `scope`'s prefix. */
function newCell(scope: XElement, twips: number): XElement {
  const tc = wElement(scope, 'tc');
  const tcPr = wElement(scope, 'tcPr');
  tcPr.append(wElement(scope, 'tcW', { w: String(twips), type: 'dxa' }));
  tc.append(tcPr, wElement(scope, 'p'));
  return tc;
}

/** Revision bookkeeping that must not be copied into new content. */
const REVISION_PROPS: ReadonlySet<string> = new Set([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'rPrChange',
  'pPrChange',
  'trPrChange',
  'tcPrChange',
  'tblPrChange',
  'tblPrExChange',
  'cellIns',
  'cellDel',
  'cellMerge',
]);
const ROW_PROP_DROP: ReadonlySet<string> = new Set([...REVISION_PROPS, 'cnfStyle']);
const CELL_PROP_DROP: ReadonlySet<string> = new Set([
  ...REVISION_PROPS,
  'cnfStyle',
  'vMerge',
  'hMerge',
]);
const PARA_PROP_DROP: ReadonlySet<string> = new Set([...REVISION_PROPS, 'cnfStyle', 'sectPr']);
const RUN_PROP_DROP: ReadonlySet<string> = REVISION_PROPS;

/**
 * Deep copy of a properties element without the `w:` children named in `drop`, looking into
 * nested property elements too (the paragraph mark's `w:rPr` inside `w:pPr`).
 */
function cleanCopy(el: XElement, drop: ReadonlySet<string>): XElement {
  const copy = el.clone(true);
  for (const d of copy.descendants(NS.w)) {
    if (drop.has(d.local) && d.parent !== null) d.remove();
  }
  return copy;
}

/**
 * A new cell formatted like `tc`: its cell properties without merge state or revisions (and
 * without `w:gridSpan` unless `keepSpan`), plus one empty paragraph formatted like its first.
 */
function newCellLike(tc: XElement, scope: XElement, keepSpan: boolean): XElement {
  const cell = wElement(scope, 'tc');
  const tcPr = tc.element(NS.w, 'tcPr');
  if (tcPr !== undefined) {
    const copy = cleanCopy(tcPr, CELL_PROP_DROP);
    if (!keepSpan) copy.element(NS.w, 'gridSpan')?.remove();
    if (copy.firstElement() !== undefined || copy._attrs.length > 0) cell.append(copy);
  }
  cell.append(paragraphLike(tc, scope));
  return cell;
}

/** An empty paragraph with the (cleaned) properties of `tc`'s first paragraph. */
function paragraphLike(tc: XElement, scope: XElement): XElement {
  const p = wElement(scope, 'p');
  const pPr = wrappedChildren(tc, P)[0]?.element(NS.w, 'pPr');
  if (pPr !== undefined) p.append(cleanCopy(pPr, PARA_PROP_DROP));
  return p;
}

/** Appends a run with `text`, giving it a copy of `rPr` when there is one. */
function appendFormattedRun(para: Paragraph, text: string, rPr: XElement | undefined): void {
  const run = para.appendRun(text);
  if (rPr === undefined) return;
  run.element.element(NS.w, 'rPr')?.remove();
  run.element.prepend(rPr.clone(true));
}

/** The cell's children other than `w:tcPr` (its content), in order. */
function contentNodes(tc: XElement): XNode[] {
  return tc.children.filter((k) => !(k.nodeType === ELEMENT_NODE && k.is(NS.w, 'tcPr')));
}

function clearContent(tc: XElement): void {
  for (const k of contentNodes(tc)) k.remove();
}

/** Moves the content of `from` to the end of `to`. */
function moveContent(from: XElement, to: XElement): void {
  for (const k of contentNodes(from)) to.append(k);
}

/** Replaces the content of `to` with the content of `from` (which is left without content). */
function replaceContent(from: XElement, to: XElement): void {
  clearContent(to);
  moveContent(from, to);
  from.append(paragraphLike(to, from));
}

/** Content becomes one empty paragraph formatted like the cell's first paragraph. */
function resetToEmptyParagraph(tc: XElement): void {
  const p = paragraphLike(tc, tc);
  const kept = outsideMarkers(contentNodes(tc));
  clearContent(tc);
  if (kept.length > 0) tc.append(...kept);
  tc.append(p);
}

/** True when the cell holds nothing but paragraphs without content (property elements only). */
function isEmptyCell(tc: XElement): boolean {
  for (const k of tc.children) {
    if (k.nodeType !== ELEMENT_NODE) continue;
    if (k.is(NS.w, 'tcPr')) continue;
    if (!k.is(NS.w, 'p')) return false;
    for (const pk of k.children) {
      if (pk.nodeType === ELEMENT_NODE && !pk.is(NS.w, 'pPr')) return false;
    }
  }
  return true;
}

/** WordprocessingML elements that are block-level content of a cell or story. */
const BLOCKS: ReadonlySet<string> = new Set(['p', 'tbl', 'sdt', 'customXml', 'altChunk']);
const STORY_CONTAINERS: ReadonlySet<string> = new Set([
  'body',
  'hdr',
  'ftr',
  'footnote',
  'endnote',
  'comment',
]);

/** Last block of a container, looking through block-level content controls and custom XML. */
function lastBlock(container: XElement): XElement | undefined {
  let c: XElement = container;
  for (;;) {
    let found: XElement | undefined;
    const kids = c.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === ELEMENT_NODE && k.uri === NS.w && BLOCKS.has(k.local)) {
        found = k;
        break;
      }
    }
    if (found === undefined) return undefined;
    if (found.local === 'sdt') {
      const content = found.element(NS.w, 'sdtContent');
      if (content === undefined) return undefined;
      c = content;
    } else if (found.local === 'customXml') {
      c = found;
    } else {
      return found;
    }
  }
}

/** Appends an empty paragraph to a cell whose last block is not a paragraph. */
function ensureEndsWithParagraph(tc: XElement): void {
  if (lastBlock(tc)?.is(NS.w, 'p') === true) return;
  tc.append(wElement(tc, 'p'));
}

/** The child of `tr` that contains `tc` (the cell itself, or its cell-level wrapper). */
function levelNode(tc: XElement, tr: XElement): XElement {
  let n = tc;
  while (n.parent !== null && n.parent !== tr && n.parent.nodeType === ELEMENT_NODE) {
    n = n.parent;
  }
  return n;
}

// ---- grid widths -------------------------------------------------------------------------

const sum = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0);

function getOrInit<K, V>(m: Map<K, V>, key: K, make: () => V): V {
  let v = m.get(key);
  if (v === undefined) {
    v = make();
    m.set(key, v);
  }
  return v;
}

function gridWidths(tbl: XElement): number[] {
  const grid = tbl.element(NS.w, 'tblGrid');
  if (grid === undefined) return [];
  return grid.elements(NS.w, 'gridCol').map((g) => {
    const n = Number.parseInt(g.attrNS(NS.w, 'w') ?? '', 10);
    return Number.isFinite(n) && n > 0 ? n : 0;
  });
}

/** Rewrites the `w:gridCol` entries (reusing the existing elements), creating `w:tblGrid`. */
function writeGridWidths(tbl: XElement, widths: readonly number[]): void {
  let grid = tbl.element(NS.w, 'tblGrid');
  if (grid === undefined) {
    grid = wElement(tbl, 'tblGrid');
    ensureTblPr(tbl).after(grid);
  }
  const cols = grid.elements(NS.w, 'gridCol');
  const change = grid.element(NS.w, 'tblGridChange');
  for (let i = 0; i < widths.length; i++) {
    const w = String(widths[i]!);
    const g = cols[i];
    if (g !== undefined) setWAttr(g, 'w', w);
    else {
      const el = wElement(tbl, 'gridCol', { w });
      if (change !== undefined) grid.insertBefore(el, change);
      else grid.append(el);
    }
  }
  for (let i = widths.length; i < cols.length; i++) cols[i]!.remove();
}

/** `w:tblPr`, created (after leading range markers, which the schema allows first) when missing. */
function ensureTblPr(tbl: XElement): XElement {
  const found = tbl.element(NS.w, 'tblPr');
  if (found !== undefined) return found;
  const pr = wElement(tbl, 'tblPr');
  const first = tbl.children.find(
    (k) => k.nodeType === ELEMENT_NODE && !(k.uri === NS.w && RANGE_MARKERS.has(k.local)),
  );
  if (first !== undefined) tbl.insertBefore(pr, first);
  else tbl.append(pr);
  return pr;
}

/**
 * Twips widths of `cols` grid columns. Cell widths (`w:tcW` in twips) decide where they pin a
 * column down — single-column cells by majority, then multi-column cells with one unknown
 * column, then an even split of what multi-column cells leave — either before (`preferCells`)
 * or after the `existing` grid widths. Columns still unknown get the average known width.
 */
function inferWidths(
  rows: readonly MRow[],
  cols: number,
  existing: readonly number[],
  preferCells: boolean,
): number[] {
  const out: (number | undefined)[] = new Array<number | undefined>(cols).fill(undefined);
  const fromExisting = (): void => {
    for (let c = 0; c < cols; c++) {
      const e = existing[c];
      if (out[c] === undefined && e !== undefined && e > 0) out[c] = e;
    }
  };
  if (!preferCells) fromExisting();
  const cons: { col: number; span: number; w: number }[] = [];
  for (const row of rows) {
    for (const c of row.cells) {
      const w = dxaWidth(c.tc.element(NS.w, 'tcPr')?.element(NS.w, 'tcW'));
      if (w !== undefined && w > 0 && c.col + c.span <= cols) cons.push({ ...c, w });
    }
  }
  const votes = new Map<number, Map<number, number>>();
  for (const k of cons) {
    if (k.span !== 1 || out[k.col] !== undefined) continue;
    const m = getOrInit(votes, k.col, () => new Map<number, number>());
    m.set(k.w, (m.get(k.w) ?? 0) + 1);
  }
  for (const [col, m] of votes) {
    let best = 0;
    let bestN = 0;
    for (const [w, n] of m) {
      if (n > bestN) [best, bestN] = [w, n];
    }
    out[col] = best;
  }
  const unknownIn = (k: { col: number; span: number }): number[] => {
    const u: number[] = [];
    for (let c = k.col; c < k.col + k.span; c++) if (out[c] === undefined) u.push(c);
    return u;
  };
  const knownSum = (k: { col: number; span: number }): number => {
    let s = 0;
    for (let c = k.col; c < k.col + k.span; c++) s += out[c] ?? 0;
    return s;
  };
  for (let progress = true; progress; ) {
    progress = false;
    for (const k of cons) {
      const u = unknownIn(k);
      const rest = k.w - knownSum(k);
      if (u.length === 1 && rest > 0) {
        out[u[0]!] = rest;
        progress = true;
      }
    }
  }
  for (const k of cons) {
    const u = unknownIn(k);
    const rest = k.w - knownSum(k);
    if (u.length < 2 || rest < u.length) continue;
    const each = Math.floor(rest / u.length);
    u.forEach((c, i) => {
      out[c] = i === u.length - 1 ? rest - each * (u.length - 1) : each;
    });
  }
  if (preferCells) fromExisting();
  const known = out.filter((w): w is number => w !== undefined);
  const avg = known.length > 0 ? Math.round(sum(known) / known.length) : DEFAULT_COLUMN_WIDTH;
  return out.map((w) => w ?? avg);
}

/**
 * Sets the twips width of each cell that is in `cells` or covers a column in `cols` to the sum
 * of its grid columns.
 */
function syncCellWidths(
  tbl: XElement,
  widths: readonly number[],
  cells: ReadonlySet<XElement>,
  cols: ReadonlySet<number>,
): void {
  for (const row of model(tbl)) {
    for (const c of row.cells) {
      let hit = cells.has(c.tc);
      for (let k = c.col; !hit && k < c.col + c.span; k++) hit = cols.has(k);
      if (hit) syncWidth(c.tc, sum(widths.slice(c.col, c.col + c.span)));
    }
  }
}

/** Shifts a twips `w:tblW` by the change of the grid's total width. */
function adjustTblW(tbl: XElement, before: number, after: number): void {
  if (before === after) return;
  const tblW = tbl.element(NS.w, 'tblPr')?.element(NS.w, 'tblW');
  const w = dxaWidth(tblW);
  if (w === undefined || w === 0) return;
  setWAttr(tblW!, 'w', String(Math.max(0, w + after - before)));
}

function checkIndex(i: number, n: number, what: string): number {
  if (!Number.isInteger(i) || i < 0 || i >= n) {
    throw new UsageError(`${what} index ${i} out of range 0..${n - 1}`);
  }
  return i;
}

function checkTwips(v: number, what: string): void {
  if (!Number.isInteger(v) || v < 0 || v > 31680) {
    throw new UsageError(`${what}: expected whole twips in 0..31680, got ${v}`);
  }
}

// ---- removal and range markers -----------------------------------------------------------

/** Range markers: start/end pairs matched by `w:id`, plus proofing marks (never paired). */
const RANGE_PAIRS: ReadonlyMap<string, string> = new Map([
  ['bookmarkStart', 'bookmarkEnd'],
  ['bookmarkEnd', 'bookmarkStart'],
  ['commentRangeStart', 'commentRangeEnd'],
  ['commentRangeEnd', 'commentRangeStart'],
  ['permStart', 'permEnd'],
  ['permEnd', 'permStart'],
  ['moveFromRangeStart', 'moveFromRangeEnd'],
  ['moveFromRangeEnd', 'moveFromRangeStart'],
  ['moveToRangeStart', 'moveToRangeEnd'],
  ['moveToRangeEnd', 'moveToRangeStart'],
]);
const RANGE_MARKERS: ReadonlySet<string> = new Set([...RANGE_PAIRS.keys(), 'proofErr']);

function topOf(el: XElement): XElement {
  let n = el;
  while (n.parent !== null && n.parent.nodeType === ELEMENT_NODE) n = n.parent;
  return n;
}

function isInside(node: XElement, roots: readonly XNode[]): boolean {
  let n: XElement | null = node;
  while (n !== null) {
    if (roots.includes(n)) return true;
    const p: XElement | null =
      n.parent !== null && n.parent.nodeType === ELEMENT_NODE ? n.parent : null;
    n = p;
  }
  return false;
}

/**
 * Range markers inside `nodes` whose partner (same kind family and `w:id`) is elsewhere in the
 * part, detached from where they were, in document order.
 */
function outsideMarkers(nodes: readonly XNode[]): XElement[] {
  const found: XElement[] = [];
  for (const n of nodes) {
    if (n.nodeType !== ELEMENT_NODE) continue;
    if (n.uri === NS.w && RANGE_PAIRS.has(n.local)) found.push(n);
    for (const d of n.descendants(NS.w)) if (RANGE_PAIRS.has(d.local)) found.push(d);
  }
  if (found.length === 0) return [];
  const first = nodes.find((n): n is XElement => n.nodeType === ELEMENT_NODE);
  const root = first === undefined ? undefined : topOf(first);
  if (root === undefined) return [];
  const index = new Map<string, XElement[]>();
  for (const d of root.descendants(NS.w)) {
    if (!RANGE_PAIRS.has(d.local)) continue;
    const key = `${d.local}#${d.attrNS(NS.w, 'id') ?? ''}`;
    getOrInit(index, key, () => []).push(d);
  }
  const out: XElement[] = [];
  for (const m of found) {
    const partners = index.get(`${RANGE_PAIRS.get(m.local)}#${m.attrNS(NS.w, 'id') ?? ''}`);
    if (partners?.some((p) => !isInside(p, nodes)) === true) out.push(m);
  }
  for (const m of out) m.remove();
  return out;
}

/** A wrapper that only holds `except` (plus markers and its own properties) once it is removed. */
function onlyHolds(container: XElement, except: XElement): boolean {
  for (const k of container.children) {
    if (k === except || k.nodeType !== ELEMENT_NODE) continue;
    if (k.uri === NS.w && (RANGE_MARKERS.has(k.local) || k.local === 'customXmlPr')) continue;
    return false;
  }
  return true;
}

/**
 * Removes `el` (and any `w:sdt`/`w:customXml` wrappers that would be left empty), first moving
 * out range markers whose partners stay in the document. Returns the element it was removed
 * from, or undefined when it was detached.
 */
function removeWithMarkers(el: XElement): XElement | undefined {
  let target = el;
  for (;;) {
    const p = target.parent;
    if (p === null || p.nodeType !== ELEMENT_NODE) break;
    let wrapper: XElement | undefined;
    if (p.is(NS.w, 'sdtContent')) {
      const sdt = p.parent;
      if (sdt !== null && sdt.nodeType === ELEMENT_NODE && sdt.is(NS.w, 'sdt')) wrapper = sdt;
    } else if (p.is(NS.w, 'customXml')) {
      wrapper = p;
    }
    if (wrapper === undefined || !onlyHolds(p, target)) break;
    target = wrapper;
  }
  const parent = target.parent;
  if (parent === null || parent.nodeType !== ELEMENT_NODE) return undefined;
  const kept = outsideMarkers([target]);
  if (kept.length > 0) target.before(...kept);
  target.remove();
  return parent;
}

/** Removes bookmark, comment, permission and move-range markers from copied content. */
function stripMarkers(root: XElement): void {
  for (const d of root.descendants(NS.w)) {
    if (d.parent === null) continue;
    if (RANGE_PAIRS.has(d.local)) d.remove();
    else if (d.local === 'commentReference') {
      const run = d.parent;
      d.remove();
      if (run.nodeType === ELEMENT_NODE && run.is(NS.w, 'r')) {
        if (run.elements().every((k) => k.is(NS.w, 'rPr'))) run.remove();
      }
    }
  }
}

// ---- fresh ids for copied content --------------------------------------------------------

/** `w:` elements whose `w:id` is a revision id. */
const REVISION_ID_ELEMENTS: ReadonlySet<string> = new Set([
  ...REVISION_PROPS,
  'numberingChange',
  'sectPrChange',
]);

interface IdSource {
  hex(kind: HexIdKind): string;
  num(kind: NumericIdKind): number;
}

/**
 * Ids from the document's {@link DocumentIds} registry, held for the batch (the writes only use
 * ids it hands out, so it need not re-scan after each one); for a table with no document
 * behind its story, from a scan of the table's own tree. Nothing is scanned until the first id
 * is asked for. Call `release()` when done.
 */
function idSource(story: Story, scope: XElement): IdSource & { release(): void } {
  const doc: WordDocument | undefined = (story as { doc?: WordDocument }).doc;
  let registry: DocumentIds | undefined;
  let release: (() => void) | undefined;
  let local: LocalIds | undefined;
  const get = (): DocumentIds | LocalIds => {
    if (doc === undefined) {
      local ??= new LocalIds(topOf(scope));
      return local;
    }
    if (registry === undefined) {
      registry = DocumentIds.of(doc);
      release = registry.hold();
    }
    return registry;
  };
  return {
    hex: (k) => {
      const src = get();
      return src instanceof LocalIds ? src.hex(k) : src.nextHex(k);
    },
    num: (k) => {
      const src = get();
      return src instanceof LocalIds ? src.num(k) : src.next(k);
    },
    release: () => release?.(),
  };
}

/** Id source for content without a document: scan the tree once, then count up. */
class LocalIds {
  readonly #used = new Map<string, Set<string>>();
  readonly #max = new Map<string, number>();
  #seed = 0x2f6b9d1;

  constructor(root: XElement) {
    for (const el of [root, ...root.descendants()]) this.#scan(el);
  }

  #note(kind: string, v: string | undefined): void {
    if (v === undefined) return;
    getOrInit(this.#used, kind, () => new Set<string>()).add(v.toUpperCase());
    const n = Number.parseInt(v, 10);
    if (Number.isFinite(n)) this.#max.set(kind, Math.max(this.#max.get(kind) ?? 0, n));
  }

  #scan(el: XElement): void {
    this.#note('paraId', el.attrNS(NS.w14, 'paraId'));
    this.#note('textId', el.attrNS(NS.w14, 'textId'));
    if (el.uri === NS.wp && el.local === 'docPr') this.#note('docPr', el.attrNS('', 'id'));
    if (el.uri !== NS.w) return;
    const id = el.attrNS(NS.w, 'id');
    if (el.local === 'bookmarkStart') this.#note('bookmark', id);
    else if (el.local === 'comment' || el.local === 'commentRangeStart') this.#note('comment', id);
    else if (el.local === 'footnote' || el.local === 'footnoteReference')
      this.#note('footnote', id);
    else if (el.local === 'endnote' || el.local === 'endnoteReference') this.#note('endnote', id);
    else if (REVISION_ID_ELEMENTS.has(el.local)) this.#note('revision', id);
    else if (el.local === 'id' && el.parent?.nodeType === ELEMENT_NODE) {
      if (el.parent.is(NS.w, 'sdtPr')) this.#note('sdt', el.attrNS(NS.w, 'val'));
    }
  }

  num(kind: NumericIdKind): number {
    const n = (this.#max.get(kind) ?? 0) + 1;
    this.#max.set(kind, n);
    return n;
  }

  hex(kind: HexIdKind): string {
    const set = getOrInit(this.#used, kind, () => new Set<string>());
    for (;;) {
      // xorshift32: deterministic, and values stay below 0x80000000 as Word requires.
      let x = this.#seed;
      x ^= x << 13;
      x ^= x >>> 17;
      x ^= x << 5;
      this.#seed = x >>> 0;
      const v = (this.#seed & 0x7fffffff).toString(16).toUpperCase().padStart(8, '0');
      if (v !== '00000000' && !set.has(v)) {
        set.add(v);
        return v;
      }
    }
  }
}

/**
 * Gives copied content fresh document-wide ids: `w14:paraId`/`w14:textId`, drawing ids,
 * content-control ids and revision ids. With `notes`, footnotes/endnotes referenced from it are
 * duplicated and the references pointed at the copies.
 */
function refreshIds(root: XElement, story: Story, ids: IdSource, notes: boolean): void {
  for (const el of [root, ...root.descendants()]) {
    if (el.attrNS(NS.w14, 'paraId') !== undefined)
      el.setAttrNS(NS.w14, 'paraId', ids.hex('paraId'));
    if (el.attrNS(NS.w14, 'textId') !== undefined)
      el.setAttrNS(NS.w14, 'textId', ids.hex('textId'));
    if (el.uri === NS.wp && el.local === 'docPr' && el.attrNS('', 'id') !== undefined) {
      el.setAttrNS('', 'id', String(ids.num('docPr')));
      continue;
    }
    if (el.uri !== NS.w) continue;
    if (el.local === 'id' && el.parent?.nodeType === ELEMENT_NODE && el.parent.is(NS.w, 'sdtPr')) {
      setWAttr(el, 'val', String(ids.num('sdt')));
    } else if (REVISION_ID_ELEMENTS.has(el.local) && el.attrNS(NS.w, 'id') !== undefined) {
      setWAttr(el, 'id', String(ids.num('revision')));
    } else if (notes && (el.local === 'footnoteReference' || el.local === 'endnoteReference')) {
      duplicateNote(el, el.local === 'footnoteReference' ? 'footnote' : 'endnote', story, ids);
    }
  }
}

function duplicateNote(
  ref: XElement,
  kind: 'footnote' | 'endnote',
  story: Story,
  ids: IdSource,
): void {
  const doc: WordDocument | undefined = (story as { doc?: WordDocument }).doc;
  if (doc === undefined) return;
  const id = ref.attrNS(NS.w, 'id');
  const note = (kind === 'footnote' ? doc.footnotes() : doc.endnotes()).find((n) => n.id === id);
  const parent = note?.element.parent;
  if (note === undefined || parent === null || parent === undefined) return;
  const copy = note.element.clone(true);
  stripMarkers(copy);
  const fresh = String(ids.num(kind));
  setWAttr(copy, 'id', fresh);
  refreshIds(copy, note, ids, false);
  parent.append(copy);
  setWAttr(ref, 'id', fresh);
}

// ---- read-only grid ----------------------------------------------------------------------

/** A cell placed in the grid while it is being computed (spans grow as merges are found). */
interface Placed {
  cell: TableCell;
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
}

function computeGrid(rows: readonly TableRow[], gridCols: number): TableGrid {
  const slots: (Placed | undefined)[][] = [];
  let width = clampCols(gridCols);
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!;
    const line: (Placed | undefined)[] = [];
    slots.push(line);
    const above = r > 0 ? slots[r - 1]! : undefined;
    let col = row.gridBefore;
    for (const cell of row.cells()) {
      // Cells past the column cap are dropped from the grid (they stay in rows()/cells()).
      const span = Math.min(cell.gridSpan, MAX_GRID_COLUMNS - col);
      if (span <= 0) break;
      let target: Placed | undefined;
      // Vertical continuation: joins whatever cell covers the slot above (lenient, like Word).
      if (cell.vMerge === 'continue') {
        target = above?.[col];
        if (target !== undefined) target.rowSpan = Math.max(target.rowSpan, r - target.row + 1);
      }
      // Legacy horizontal merge: extends the previous cell of this same row.
      if (target === undefined && cell.hMerge === 'continue' && col > 0) {
        const prev = line[col - 1];
        if (prev !== undefined && prev.row === r) {
          prev.colSpan += span;
          target = prev;
        }
      }
      target ??= { cell, row: r, col, rowSpan: 1, colSpan: span };
      for (let c = col; c < col + span; c++) line[c] = target;
      col += span;
    }
    width = Math.max(width, clampCols(col + row.gridAfter));
  }
  // Freeze one view object per slot now that every span is final.
  const views: (GridCell | undefined)[][] = slots.map((line, r) => {
    const out: (GridCell | undefined)[] = [];
    for (let c = 0; c < line.length; c++) {
      const p = line[c];
      if (p === undefined) continue;
      out[c] = {
        cell: p.cell,
        row: p.row,
        col: p.col,
        rowSpan: p.rowSpan,
        colSpan: p.colSpan,
        origin: p.row === r && p.col === c,
      };
    }
    return out;
  });
  return {
    rows: rows.length,
    cols: width,
    at(r: number, c: number): GridCell | undefined {
      return views[r]?.[c];
    },
  };
}
