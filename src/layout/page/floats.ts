// Floating objects (wp:anchor): positioning, wrap exclusion zones and their display items.
//
// A float is positioned against a frame chosen by `relativeFrom` (page, margin, column,
// paragraph, line, character, the margins themselves, inside/outside by page parity), either by
// an offset from the frame's start or aligned inside it. Wrapping is rectangular: `topAndBottom`
// blocks the full column width over the float's height; `square`/`tight`/`through` block the
// float's box (plus its distances) so lines beside it get the wider remaining side. `none`
// (behind/in front of text) blocks nothing. Text boxes are laid out inside their box with
// DrawingML's default insets.

import { drawingItems, metafileImageItems } from '../graphics/index.js';
import type { DisplayItem, IRAnchor } from '../types.js';
import { stackBlocks } from './blocks.js';
import { type LayoutCtx, mergeStrip, moveItem, type Strip, stackLayout } from './strip.js';

/** Lines narrower than this beside a float are not used (text continues below it). */
export const MIN_WRAP_GAP = 36;
/** DrawingML default text box insets (91440 / 45720 EMU). */
const INSET_X = 7.2;
const INSET_Y = 3.6;

/** Geometry a float is positioned against. */
export interface FloatFrame {
  pageWidth: number;
  pageHeight: number;
  /** Page margins (left includes the gutter). */
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  columnX: number;
  columnWidth: number;
  /** Top of the anchoring paragraph and its left edge. */
  paraTop: number;
  paraLeft: number;
  /** 0-based physical page index (inside/outside alternate). */
  pageIndex: number;
}

/** A rectangle text must avoid; `full` blocks the whole column width. */
export interface Exclusion {
  x0: number;
  x1: number;
  y0: number;
  y1: number;
  full: boolean;
}

/** A positioned float. */
export interface PlacedFloat {
  anchor: IRAnchor;
  x: number;
  y: number;
  exclusion?: Exclusion;
}

function place(
  start: number,
  size: number,
  objSize: number,
  offset: number | undefined,
  align: string | undefined,
  odd: boolean,
): number {
  if (offset !== undefined && align === undefined) return start + offset;
  switch (align) {
    case 'center':
      return start + (size - objSize) / 2;
    case 'right':
    case 'bottom':
      return start + size - objSize;
    case 'inside':
      return odd ? start : start + size - objSize;
    case 'outside':
      return odd ? start + size - objSize : start;
    default:
      return start + (offset ?? 0);
  }
}

/** Top-left of `a` in page coordinates. */
export function positionFloat(a: IRAnchor, f: FloatFrame): { x: number; y: number } {
  const odd = f.pageIndex % 2 === 0; // physical page 1, 3, ... (recto)
  const h = a.horizontal;
  let hs = f.marginLeft;
  let hw = f.pageWidth - f.marginLeft - f.marginRight;
  switch (h.relativeTo) {
    case 'page':
      hs = 0;
      hw = f.pageWidth;
      break;
    case 'column':
      hs = f.columnX;
      hw = f.columnWidth;
      break;
    case 'character':
      hs = f.paraLeft;
      hw = f.columnWidth - (f.paraLeft - f.columnX);
      break;
    case 'leftMargin':
      hs = 0;
      hw = f.marginLeft;
      break;
    case 'rightMargin':
      hs = f.pageWidth - f.marginRight;
      hw = f.marginRight;
      break;
    case 'insideMargin':
      hs = odd ? 0 : f.pageWidth - f.marginRight;
      hw = odd ? f.marginLeft : f.marginRight;
      break;
    case 'outsideMargin':
      hs = odd ? f.pageWidth - f.marginRight : 0;
      hw = odd ? f.marginRight : f.marginLeft;
      break;
  }
  const v = a.vertical;
  let vs = f.marginTop;
  let vh = f.pageHeight - f.marginTop - f.marginBottom;
  switch (v.relativeTo) {
    case 'page':
      vs = 0;
      vh = f.pageHeight;
      break;
    case 'paragraph':
    case 'line':
      vs = f.paraTop;
      vh = f.pageHeight - f.marginBottom - f.paraTop;
      break;
    case 'topMargin':
    case 'insideMargin':
      vs = 0;
      vh = f.marginTop;
      break;
    case 'bottomMargin':
    case 'outsideMargin':
      vs = f.pageHeight - f.marginBottom;
      vh = f.marginBottom;
      break;
  }
  return {
    x: place(hs, hw, a.width, h.offset, h.align, odd),
    y: place(vs, vh, a.height, v.offset, v.align === 'inside' ? 'top' : v.align, odd),
  };
}

/** The exclusion zone of a float at (x, y), or undefined when text ignores it. */
export function exclusionOf(a: IRAnchor, x: number, y: number): Exclusion | undefined {
  if (a.wrap === 'none') return undefined;
  const d = a.distance;
  return {
    x0: x - d.left,
    x1: x + a.width + d.right,
    y0: y - d.top,
    y1: y + a.height + d.bottom,
    full: a.wrap === 'topAndBottom',
  };
}

/**
 * Horizontal space left for text between `top` and `bottom` in a column: the widest free
 * interval as {left, width} relative to the column (width 0 when blocked: too narrow or a
 * topAndBottom float) and `below`, the lowest bottom of the intersecting zones; null when no
 * exclusion intersects.
 */
export function freeSpace(
  exclusions: readonly Exclusion[],
  colX: number,
  colW: number,
  top: number,
  bottom: number,
): { left: number; width: number; below: number } | null {
  let intervals: [number, number][] = [[colX, colX + colW]];
  let hit = false;
  let below = top;
  for (const e of exclusions) {
    if (e.y1 <= top + 0.01 || e.y0 >= bottom - 0.01) continue;
    if (!e.full && (e.x1 <= colX || e.x0 >= colX + colW)) continue;
    hit = true;
    below = Math.max(below, e.y1);
    if (e.full) {
      intervals = [];
      continue;
    }
    const next: [number, number][] = [];
    for (const [a, b] of intervals) {
      if (e.x1 <= a || e.x0 >= b) next.push([a, b]);
      else {
        if (e.x0 > a) next.push([a, e.x0]);
        if (e.x1 < b) next.push([e.x1, b]);
      }
    }
    intervals = next;
  }
  if (!hit) return null;
  let best: [number, number] | undefined;
  for (const iv of intervals)
    if (best === undefined || iv[1] - iv[0] > best[1] - best[0]) best = iv;
  if (best === undefined || best[1] - best[0] < MIN_WRAP_GAP) return { left: 0, width: 0, below };
  return { left: best[0] - colX, width: best[1] - best[0], below };
}

/** Display items of a float placed at (x, y). */
export function floatItems(ctx: LayoutCtx, a: IRAnchor, x: number, y: number): DisplayItem[] {
  const out: DisplayItem[] = [];
  if (a.kind === 'image' && a.image !== undefined) {
    // EMF/WMF pictures become vector items (renderers cannot draw metafiles as images)
    const res = ctx.env.images.get(a.image);
    const vector =
      res === undefined ? undefined : metafileImageItems(res, a.width, a.height, ctx.env);
    if (vector !== undefined) {
      for (const it of vector) out.push(moveItem(it, x, y));
    } else out.push({ kind: 'image', x, y, width: a.width, height: a.height, image: a.image });
    return out;
  }
  try {
    const shape = drawingItems(a.source, a.width, a.height, ctx.env);
    if (shape !== undefined) {
      const tmp: Strip = {
        height: 0,
        spaceBefore: 0,
        spaceAfter: 0,
        items: [],
        breakBefore: 'auto',
      };
      mergeStrip(tmp, { ...tmp, items: shape }, x, y);
      out.push(...tmp.items);
    }
  } catch (e) {
    ctx.env.warn('LAYOUT_DRAWING', `Drawing not rendered: ${e instanceof Error ? e.message : e}`);
  }
  if (a.blocks !== undefined && a.blocks.length > 0) {
    const strips = stackBlocks(ctx, a.blocks, Math.max(1, a.width - 2 * INSET_X));
    const box: Strip = { height: 0, spaceBefore: 0, spaceAfter: 0, items: [], breakBefore: 'auto' };
    const { ys } = stackLayout(strips);
    strips.forEach((s, i) => {
      mergeStrip(box, s, x + INSET_X, y + INSET_Y + ys[i]!);
    });
    out.push(...box.items);
  }
  return out;
}
