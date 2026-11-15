// Strips: the unit the paginator moves around.
//
// Every block is turned into a list of strips before it is placed: one strip per paragraph line,
// one per table row. A strip is an indivisible horizontal slice (the split of a table row across
// pages produces two new strips) with its display items relative to its own top-left corner, the
// spacing around it, and whether a column/page break may come before it (keep-with-next, keep
// lines together and widow/orphan control all reduce to that one flag). The same strips stack
// inside table cells, headers/footers, notes and text boxes, so all of them share one code path.
//
// The layout context carries the injectable paragraph layouter (the real one is layout-A's
// `layoutParagraph`; tests may inject a fake), the measurement cache and the nesting depth.

import type { XElement } from '../../xml/dom.js';
import type { LayoutEnv, ParagraphLayout } from '../engine.js';
import type { DisplayItem, IRAnchor, IRBlock, IRInline, IRParagraph } from '../types.js';

/** A page-dependent field inline of the IR. */
export type FieldInline = Extract<IRInline, { kind: 'field' }>;

/** May a column/page break come before a strip: freely, preferably not, or must it (hard break). */
export type BreakRule = 'auto' | 'avoid' | 'page' | 'column';

/** A floating object anchored in a paragraph, with the paragraph's top-left relative to its strip. */
export interface AnchorRef {
  anchor: IRAnchor;
  /** Paragraph content box left/top relative to the strip's top-left. */
  dx: number;
  dy: number;
  /** Width of the column the paragraph was laid out in. */
  width: number;
}

/** One indivisible slice of content. */
export interface Strip {
  /** Content height (the line pitch or row height), excluding spaceBefore/spaceAfter. */
  height: number;
  /** Space above; dropped at the top of a column reached by a soft break. */
  spaceBefore: number;
  /** Space below; it need not fit at the bottom of a column. */
  spaceAfter: number;
  /** Relative to (strip left, strip content top). */
  items: DisplayItem[];
  breakBefore: BreakRule;
  /** A hard page/column break ends this strip (w:br). */
  breakAfter?: 'page' | 'column';
  /** Field placeholder items → their field (for per-page substitution). */
  fields?: Map<DisplayItem, FieldInline>;
  /** Footnotes referenced in this strip (reserved on the same page). */
  footnotes?: string[];
  /** Endnotes referenced in this strip (in order). */
  endnotes?: string[];
  /** Bookmarks starting in this strip (destinations point at its top). */
  bookmarks?: string[];
  /** Source paragraphs starting in this strip (DocumentLayout.paragraphPages). */
  starts?: XElement[];
  /** Heading starting here (DocumentLayout.outline). */
  outline?: { title: string; level: number };
  anchors?: AnchorRef[];
  /** Table row membership (header rows repeat after a break). */
  table?: { key: object; header: boolean };
  /** Splits a too-tall strip so the first part is at most `avail` high; undefined = cannot. */
  split?: (avail: number) => [Strip, Strip] | undefined;
  /** First strip of a paragraph: its style and contextualSpacing (spacing between paragraphs). */
  para?: { styleId: string | undefined; contextual: boolean };
  /** Last strip of a paragraph: its style and contextualSpacing. */
  paraEnd?: {
    styleId: string | undefined;
    contextual: boolean;
    keepNext: boolean;
    /** w:afterAutospacing (HTML spacing) and list membership, for Word's autospacing rules. */
    afterAuto: boolean;
    listed: boolean;
  };
}

/** Paragraph layouter signature (layout-A's layoutParagraph). */
export type ParagraphLayouter = (
  p: IRParagraph,
  width: number,
  env: LayoutEnv,
  o?: ParagraphConstraints,
) => ParagraphLayout;

/**
 * 4th argument of layoutParagraph (layout-A's ParagraphOptions), used for text wrapping beside
 * floats: extra points kept free on the physical left and right of line `line` (0-based), on
 * top of the paragraph's indents. Items keep their usual coordinates (column edge +
 * indentLeft), so a left inset moves them right, and `indentLeft + LineBox.width` is still where
 * the content ends. A layouter that ignores the insets is detected (a line too wide or not moved
 * past a left float) and the paginator narrows the whole paragraph instead.
 */
export interface ParagraphConstraints {
  insets?: (line: number) => { left: number; right: number };
}

/** Line insets with a cache key (the constraint set used for a measurement). */
export interface InsetSpec {
  insets: (line: number) => { left: number; right: number };
  key: string;
}

/** Shared state of one layout pass. */
export interface LayoutCtx {
  env: LayoutEnv;
  layoutParagraph: ParagraphLayouter;
  /** (paragraph) → (width + field values) → layout. */
  cache: WeakMap<IRParagraph, Map<string, ParagraphLayout>>;
  /** Table nesting depth while laying out cell content. */
  depth: number;
}

/** Table nesting deeper than this is not laid out (a warning is reported). */
export const MAX_TABLE_DEPTH = 32;

const fieldLists = new WeakMap<IRParagraph, FieldInline[]>();

/** The field inlines of a paragraph (cached). */
export function fieldsOf(p: IRParagraph): FieldInline[] {
  let list = fieldLists.get(p);
  if (list === undefined) {
    list = [];
    for (const i of p.inlines) if (i.kind === 'field') list.push(i);
    fieldLists.set(p, list);
  }
  return list;
}

/** Lays a paragraph out at `width`, cached by width and the current field fallbacks. */
export function measure(
  ctx: LayoutCtx,
  p: IRParagraph,
  width: number,
  insets?: InsetSpec,
): ParagraphLayout {
  const w = Math.max(1, width);
  const fields = fieldsOf(p);
  let key = w.toFixed(3);
  if (insets !== undefined) key += `\u0002${insets.key}`;
  for (const f of fields) key += `\u0001${f.fallback}`;
  let byKey = ctx.cache.get(p);
  if (byKey === undefined) {
    byKey = new Map();
    ctx.cache.set(p, byKey);
  }
  let pl = byKey.get(key);
  if (pl === undefined) {
    pl =
      insets === undefined
        ? ctx.layoutParagraph(p, w, ctx.env)
        : ctx.layoutParagraph(p, w, ctx.env, { insets: insets.insets });
    byKey.set(key, pl);
  }
  return pl;
}

// ---- display items -----------------------------------------------------------------------------

const round = (n: number): number => Math.round(n * 1000) / 1000;

/** Translates SVG path data (absolute commands; a leading relative moveto counts as absolute). */
export function translatePath(d: string, dx: number, dy: number): string {
  if (dx === 0 && dy === 0) return d;
  const tokens = d.match(/[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) ?? [];
  let out = '';
  let cmd = '';
  let argIndex = 0;
  let first = true;
  let firstPair = false;
  for (const t of tokens) {
    if (/^[A-Za-z]$/.test(t)) {
      cmd = t;
      argIndex = 0;
      firstPair = first && t === 'm';
      first = false;
      out += out === '' ? t : ` ${t}`;
      continue;
    }
    let v = Number(t);
    const abs = cmd === cmd.toUpperCase() || (firstPair && argIndex < 2);
    if (abs) {
      switch (cmd.toUpperCase()) {
        case 'H':
          v += dx;
          break;
        case 'V':
          v += dy;
          break;
        case 'A': {
          const k = argIndex % 7;
          if (k === 5) v += dx;
          else if (k === 6) v += dy;
          break;
        }
        default:
          v += argIndex % 2 === 0 ? dx : dy;
      }
    }
    argIndex++;
    out += ` ${round(v)}`;
  }
  return out;
}

/** A copy of `it` moved by (dx, dy). */
export function moveItem(it: DisplayItem, dx: number, dy: number): DisplayItem {
  switch (it.kind) {
    case 'line':
      return { ...it, x1: it.x1 + dx, y1: it.y1 + dy, x2: it.x2 + dx, y2: it.y2 + dy };
    case 'path':
      return { ...it, d: translatePath(it.d, dx, dy) };
    default:
      return { ...it, x: it.x + dx, y: it.y + dy };
  }
}

/** Appends `src`'s items (moved by dx, dy) and metadata to `dst`. */
export function mergeStrip(dst: Strip, src: Strip, dx: number, dy: number): void {
  for (const it of src.items) {
    const moved = moveItem(it, dx, dy);
    dst.items.push(moved);
    const f = src.fields?.get(it);
    if (f !== undefined) {
      dst.fields ??= new Map();
      dst.fields.set(moved, f);
    }
  }
  if (src.footnotes) dst.footnotes = [...(dst.footnotes ?? []), ...src.footnotes];
  if (src.endnotes) dst.endnotes = [...(dst.endnotes ?? []), ...src.endnotes];
  if (src.bookmarks) dst.bookmarks = [...(dst.bookmarks ?? []), ...src.bookmarks];
  if (src.starts) dst.starts = [...(dst.starts ?? []), ...src.starts];
  if (src.anchors) {
    const moved = src.anchors.map((a) => ({ ...a, dx: a.dx + dx, dy: a.dy + dy }));
    dst.anchors = [...(dst.anchors ?? []), ...moved];
  }
}

/**
 * Space between two stacked strips. Word does not add the space after a paragraph and the space
 * before the next one: the larger of the two separates them.
 */
export function gap(prev: Strip | undefined, next: Strip): number {
  return prev === undefined ? next.spaceBefore : Math.max(prev.spaceAfter, next.spaceBefore);
}

/** Content tops of `strips` stacked from y = 0, and the total height (last space after included). */
export function stackLayout(strips: readonly Strip[]): { ys: number[]; height: number } {
  const ys: number[] = [];
  let y = 0;
  let prev: Strip | undefined;
  for (const s of strips) {
    y += gap(prev, s);
    ys.push(y);
    y += s.height;
    prev = s;
  }
  return { ys, height: y + (prev?.spaceAfter ?? 0) };
}

/** Total stacked height of strips (collapsed spacing included). */
export function stackHeight(strips: readonly Strip[]): number {
  return stackLayout(strips).height;
}

/** Plain text of a paragraph (outline titles). */
export function paragraphText(p: IRParagraph): string {
  let s = '';
  for (const i of p.inlines) {
    if (i.kind === 'text') s += i.text;
    else if (i.kind === 'field') s += i.fallback;
    else if (i.kind === 'tab') s += ' ';
  }
  return s.replace(/\s+/g, ' ').trim();
}

/** Every paragraph in `blocks`, depth-first, table cells and text boxes included (no recursion). */
export function* allParagraphs(blocks: readonly IRBlock[]): Generator<IRParagraph> {
  const stack: IRBlock[] = [...blocks].reverse();
  while (stack.length > 0) {
    const b = stack.pop()!;
    if (b.kind === 'paragraph') {
      yield b;
      for (let i = b.anchors.length - 1; i >= 0; i--) {
        const inner = b.anchors[i]!.blocks;
        if (inner) for (let j = inner.length - 1; j >= 0; j--) stack.push(inner[j]!);
      }
    } else {
      for (let r = b.rows.length - 1; r >= 0; r--) {
        const cells = b.rows[r]!.cells;
        for (let c = cells.length - 1; c >= 0; c--) {
          const bl = cells[c]!.blocks;
          for (let j = bl.length - 1; j >= 0; j--) stack.push(bl[j]!);
        }
      }
    }
  }
}
