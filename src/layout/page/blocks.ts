// Blocks → strips: paragraphs become one strip per line (with keep/widow rules turned into break
// permissions), tables one strip per row (table.ts). `stackBlocks` is the in-box variant used for
// table cells, headers/footers, notes and text boxes, where hard breaks do not apply.

import { tableStrips } from '../table.js';
import type { DisplayItem, IRBlock, IRParagraph } from '../types.js';
import {
  type BreakRule,
  type FieldInline,
  type InsetSpec,
  type LayoutCtx,
  MAX_TABLE_DEPTH,
  measure,
  moveItem,
  paragraphText,
  type Strip,
} from './strip.js';

/**
 * Space before `p` given the previous paragraph (its last strip), zeroing the previous
 * paragraph's space after where Word drops it:
 * - contextual spacing between paragraphs of the same style;
 * - HTML autospacing (w:before/afterAutospacing) between two list paragraphs.
 */
function spacingBefore(
  prev: Strip | undefined,
  p: IRParagraph,
  styleId: string | undefined,
  contextual: boolean,
  before: number,
): number {
  const end = prev?.paraEnd;
  if (prev === undefined || end === undefined) return before;
  let sb = before;
  const listed = p.props.numbering !== undefined;
  if (end.listed && listed) {
    if (end.afterAuto) prev.spaceAfter = 0;
    if (p.props.spacing.beforeAutospacing) sb = 0;
  }
  if (end.styleId === styleId) {
    if (end.contextual) prev.spaceAfter = 0;
    if (contextual) sb = 0;
  }
  return sb;
}

/** Lines of a paragraph as strips. `first` = break rule before the first line. */
export function paragraphStrips(
  ctx: LayoutCtx,
  p: IRParagraph,
  width: number,
  prev: Strip | undefined,
  first: BreakRule,
  insets?: InsetSpec,
): Strip[] {
  const pl = measure(ctx, p, width, insets);
  const styleId = pl.styleId ?? p.props.styleId;
  const contextual = pl.contextualSpacing;
  const sb = spacingBefore(prev, p, styleId, contextual, pl.spaceBefore);
  // footnote vs endnote ids (LineBox.noteRefs only carries ids)
  const endnoteIds = new Set<string>();
  const footnoteIds = new Set<string>();
  for (const i of p.inlines) {
    if (i.kind === 'noteRef') (i.note === 'endnote' ? endnoteIds : footnoteIds).add(i.id);
  }
  const lines = pl.lines;
  const n = Math.max(1, lines.length);
  const out: Strip[] = [];
  for (let k = 0; k < n; k++) {
    const line = lines[k];
    const items: DisplayItem[] = [];
    let fields: Map<DisplayItem, FieldInline> | undefined;
    if (line !== undefined) {
      const slots = new Map<number, FieldInline>();
      for (const f of line.fields) slots.set(f.item, f.field);
      line.items.forEach((it, idx) => {
        const moved = moveItem(it, pl.indentLeft, line.baseline);
        items.push(moved);
        const f = slots.get(idx);
        if (f !== undefined) {
          fields ??= new Map();
          fields.set(moved, f);
        }
      });
    }
    let rule: BreakRule;
    if (k === 0) rule = first;
    else if (pl.keepLines) rule = 'avoid';
    else if (pl.widowControl && (k === 1 || k === n - 1)) rule = 'avoid';
    else rule = 'auto';
    const s: Strip = {
      height: line?.height ?? 0,
      spaceBefore: k === 0 ? sb : 0,
      spaceAfter: k === n - 1 ? pl.spaceAfter : 0,
      items,
      breakBefore: rule,
    };
    if (fields !== undefined) s.fields = fields;
    if (line?.breakAfter !== undefined) s.breakAfter = line.breakAfter;
    if (line !== undefined && line.noteRefs.length > 0) {
      for (const id of line.noteRefs) {
        if (endnoteIds.has(id) && !footnoteIds.has(id)) s.endnotes = [...(s.endnotes ?? []), id];
        else s.footnotes = [...(s.footnotes ?? []), id];
      }
    }
    if (line !== undefined && line.bookmarks.length > 0) s.bookmarks = [...line.bookmarks];
    if (k === 0) {
      s.para = { styleId, contextual };
      s.starts = [p.source];
      if (p.bookmarks.length > 0)
        s.bookmarks = [...new Set([...p.bookmarks, ...(s.bookmarks ?? [])])];
      const level = p.props.outlineLevel;
      if (level !== undefined && level >= 0 && level < 9) {
        const title = paragraphText(p);
        if (title !== '') s.outline = { title, level };
      }
      if (pl.decorations.length > 0) s.items.unshift(...pl.decorations);
      if (p.anchors.length > 0) {
        s.anchors = p.anchors.map((anchor) => ({ anchor, dx: 0, dy: 0, width }));
      }
    }
    if (k === n - 1) {
      s.paraEnd = {
        styleId,
        contextual,
        keepNext: pl.keepNext,
        afterAuto: p.props.spacing.afterAutospacing,
        listed: p.props.numbering !== undefined,
      };
    }
    out.push(s);
  }
  return out;
}

/**
 * Strips of `blocks` stacked in a box `width` wide (cells, headers/footers, notes, text boxes):
 * hard breaks are ignored, tables nest up to MAX_TABLE_DEPTH.
 */
export function stackBlocks(
  ctx: LayoutCtx,
  blocks: readonly IRBlock[],
  width: number,
  cell = false,
): Strip[] {
  const out: Strip[] = [];
  for (const b of blocks) {
    const prev = out[out.length - 1];
    if (b.kind === 'paragraph') {
      for (const s of paragraphStrips(ctx, b, width, prev, 'auto')) {
        if (s.breakBefore === 'page' || s.breakBefore === 'column') s.breakBefore = 'auto';
        delete s.breakAfter;
        out.push(s);
      }
    } else if (ctx.depth >= MAX_TABLE_DEPTH) {
      ctx.env.warn(
        'LAYOUT_TABLE_DEPTH',
        `Tables nested deeper than ${MAX_TABLE_DEPTH} levels are not laid out`,
      );
    } else {
      ctx.depth++;
      try {
        out.push(...tableStrips(ctx, b, width, prev?.paraEnd?.keepNext === true));
      } finally {
        ctx.depth--;
      }
    }
  }
  if (cell) {
    // Word drops HTML autospacing at the top and bottom of a table cell
    const first = blocks[0];
    const head = out[0];
    if (first?.kind === 'paragraph' && first.props.spacing.beforeAutospacing && head) {
      head.spaceBefore = 0;
    }
    const last = blocks[blocks.length - 1];
    const tail = out[out.length - 1];
    if (last?.kind === 'paragraph' && last.props.spacing.afterAutospacing && tail) {
      tail.spaceAfter = 0;
    }
  }
  return out;
}
