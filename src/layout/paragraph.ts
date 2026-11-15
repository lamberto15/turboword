// Paragraph layout: one IRParagraph broken into lines at a column width, the way Word does it.
//
//   inlines → pieces (caps, super/subscript, label, fields, objects; line/pieces.ts)
//           → atoms (itemize by face/script/bidi level, shape; line/shape.ts)
//           → lines (greedy breaking at UAX #14 opportunities with tab sizing; line/breaker.ts)
//           → items (alignment, justification, bidi reordering, decorations; line/emit.ts)
//           → line heights (Word's win-metrics rule and line spacing; line/metrics.ts)
//
// Coordinates. Line items are relative to the paragraph's left indent (column left edge +
// `indentLeft`) horizontally and to the line's baseline vertically; a hanging first line or a
// list label therefore has negative x. Paragraph decorations (borders, shading) are relative to
// the column's left edge and the top of the first line. `LineBox.width` is the right edge of the
// line's content measured from the left indent, so `indentLeft + width` is where it ends.
//
// Indents (twips in the resolved props, already physical left/right) are applied on the start
// side (left for LTR, right for RTL paragraphs) for the first-line/hanging offset; alignment
// `left`/`right` mean start/end, so they mirror in RTL paragraphs as in Word.
//
// The top border's space + width is added to the first line's height (and baseline), the bottom
// border's to the last line's, so the paragraph's height is the sum of its line heights.

import type { LayoutEnv, LineBox, ParagraphLayout } from './engine.js';
import {
  type BreakInput,
  breakLines,
  JUSTIFY_SHRINK,
  placeLine,
  type TabResolver,
} from './line/breaker.js';
import { borderPad, paragraphDecorations } from './line/decorations.js';
import { emitLine, type LineAlign } from './line/emit.js';
import { paragraphExtras } from './line/extras.js';
import { fontExtents, lineBox } from './line/metrics.js';
import { buildPieces, LABEL } from './line/pieces.js';
import { buildAtoms } from './line/shape.js';
import { type TabAlign, TabStops } from './line/tabs.js';
import { bidiLevels, lineBreakOpportunities } from './text/index.js';
import type { DisplayItem, IRParagraph } from './types.js';

export { paragraphDecorations } from './line/decorations.js';
export { fontExtents, lineBox } from './line/metrics.js';

/** Word's paragraph spacing for `w:beforeAutospacing`/`w:afterAutospacing` (HTML rule), points. */
export const AUTOSPACING = 14;

/** Characters (or ranges) that can only come from right-to-left text. */
const RTL_CHARS =
  /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff]|\ud802[\udc00-\udfff]|\ud83a[\udc00-\udfff]|\ud83b[\udc00-\udeff]/;

function alignmentOf(p: IRParagraph): { body: LineAlign; last: LineAlign } {
  switch (p.props.alignment) {
    case 'center':
      return { body: 'center', last: 'center' };
    case 'right':
      return { body: 'end', last: 'end' };
    case 'both':
    case 'mediumKashida':
    case 'highKashida':
    case 'lowKashida':
    case 'thaiDistribute':
      return { body: 'justify', last: 'start' };
    case 'distribute':
      return { body: 'distribute', last: 'distribute' };
    default:
      return { body: 'start', last: 'start' };
  }
}

/** Optional constraints of {@link layoutParagraph}. */
export interface ParagraphOptions {
  /**
   * Extra points kept free on the physical left and right of line `line` (0-based), on top of
   * the paragraph's indents (text wrapping beside square floats). Lines are filled, aligned,
   * justified and tab-aligned within the narrowed box; tab stops stay measured from the indent
   * origin. Item coordinates keep their usual origin (column edge + indentLeft), so a left
   * inset of 40 moves that line's items 40 pt to the right. In RTL paragraphs `right` is the
   * start side.
   */
  insets?: (line: number) => { left: number; right: number };
}

/** Breaks a paragraph into lines for a column of `width` points (tabs, justification, indents, line spacing, inline images/math/drawings). */
export function layoutParagraph(
  p: IRParagraph,
  width: number,
  env: LayoutEnv,
  o?: ParagraphOptions,
): ParagraphLayout {
  const props = p.props;
  const extras = paragraphExtras(p);
  const rtl = props.bidi;
  const ind = props.indent;
  const left = ind.left / 20;
  const right = ind.right / 20;
  const startIndent = rtl ? right : left;
  const endIndent = rtl ? left : right;
  const firstOffset = (ind.firstLine - ind.hanging) / 20;

  // ---- content -----------------------------------------------------------------------------
  const { pieces, text } = buildPieces(p, env);
  const needsBidi = rtl || RTL_CHARS.test(text) || pieces.some((pc) => pc.props.rtl);
  const levels = needsBidi && text !== '' ? bidiLevels(text, rtl) : undefined;
  const atoms = buildAtoms(pieces, levels, env);
  const opportunities = text === '' ? new Uint8Array(0) : lineBreakOpportunities(text);

  // ---- geometry and tabs ---------------------------------------------------------------------
  const end = width - endIndent;
  const insets = o?.insets;
  const geometries: { start: number; end: number }[] = [];
  const geometry = (line: number): { start: number; end: number } => {
    let g = geometries[line];
    if (g === undefined) {
      const ins = insets?.(line);
      const atStart = ins === undefined ? 0 : Math.max(0, rtl ? ins.right : ins.left);
      const atEnd = ins === undefined ? 0 : Math.max(0, rtl ? ins.left : ins.right);
      g = { start: startIndent + (line === 0 ? firstOffset : 0) + atStart, end: end - atEnd };
      geometries[line] = g;
    }
    return g;
  };
  const stops = new TabStops(
    props.tabs.map((t) => {
      const s: { pos: number; val: typeof t.val; leader?: NonNullable<typeof t.leader> } = {
        pos: t.pos / 20,
        val: t.val,
      };
      if (t.leader !== undefined) s.leader = t.leader;
      return s;
    }),
    env.defaultTabStop,
    firstOffset < 0 ? startIndent : undefined,
  );
  const tabs: TabResolver = {
    stops,
    explicit(i) {
      const piece = pieces[atoms[i]!.piece];
      if (piece?.kind !== 'tab' || piece.ptab === undefined) return undefined;
      const pt = piece.ptab;
      const lo = pt.relativeTo === 'indent' ? startIndent : 0;
      const hi = pt.relativeTo === 'indent' ? end : width;
      const pos = pt.alignment === 'left' ? lo : pt.alignment === 'center' ? (lo + hi) / 2 : hi;
      const out: { pos: number; align: TabAlign; leader?: NonNullable<typeof pt.leader> } = {
        pos,
        align: pt.alignment,
      };
      if (pt.leader !== undefined) out.leader = pt.leader;
      return out;
    },
  };

  // A right- or centre-aligned list label ends at / is centred on the first-line start.
  let firstLineShift = 0;
  if (p.label !== undefined && p.label.alignment !== 'left') {
    let w = 0;
    for (const a of atoms) {
      const pc = pieces[a.piece]!;
      if (pc.inline === LABEL && pc.kind === 'text' && pc.labelSuffix !== true) w += a.width;
    }
    firstLineShift = p.label.alignment === 'right' ? -w : -w / 2;
  }

  const justified = props.alignment === 'both' && (extras.compat?.compatibilityMode ?? 0) >= 15;
  const input: BreakInput = {
    atoms,
    opportunities,
    text,
    geometry,
    firstLineShift,
    tabs,
    decimal: extras.compat?.decimalSymbol ?? '.',
    shrink: justified ? JUSTIFY_SHRINK : 0,
  };
  const ranges = breakLines(input);
  const align = alignmentOf(p);
  const expandShiftReturn = extras.compat?.doNotExpandShiftReturn !== true;
  const ctx = { atoms, pieces, text, env, rtl, width, indentLeft: left };

  // Bookmarks by the atom where they start.
  const bookmarkAtom: number[] = [];
  if (p.bookmarks.length > 0) {
    const at = extras.bookmarkAt;
    for (let b = 0; b < p.bookmarks.length; b++) {
      const inline = at?.[b] ?? 0;
      let idx = atoms.findIndex((a) => {
        const pi = pieces[a.piece]!.inline;
        return pi !== LABEL && pi >= inline;
      });
      if (idx < 0) idx = atoms.length;
      bookmarkAtom.push(idx);
    }
  }

  const markExt = fontExtents(env.fonts.fontFor(props.markRun, 0x20), props.markRun.size);
  const lines: LineBox[] = [];
  for (let li = 0; li < ranges.length; li++) {
    const { from, to } = ranges[li]!;
    const isLast = li === ranges.length - 1;
    const lastAtom = to > from ? atoms[to - 1] : undefined;
    const endsWithBreak = lastAtom?.kind === 'break' ? lastAtom.breakType : undefined;
    let mode: LineAlign;
    if (isLast) mode = align.last;
    else if (endsWithBreak === 'line') mode = expandShiftReturn ? align.body : align.last;
    else if (endsWithBreak !== undefined) mode = align.last;
    else mode = align.body;
    const placed = placeLine(input, li, from, to);
    const out = emitLine(ctx, placed, mode, !isLast && endsWithBreak === undefined, justified);
    let ascent = out.ascent;
    let descent = out.descent;
    // The paragraph mark's font counts on the last line; an empty line has only that.
    if (isLast || to === from) {
      ascent = Math.max(ascent, markExt.ascent);
      descent = Math.max(descent, markExt.descent);
    }
    const box = lineBox(ascent, descent, props.spacing, out.objectAscent, out.objectDescent);
    const line: LineBox = {
      height: box.height,
      baseline: box.baseline,
      width: out.width,
      items: out.items,
      fields: out.fields,
      noteRefs: out.noteRefs,
      bookmarks: [],
    };
    if (endsWithBreak === 'page' || endsWithBreak === 'column') line.breakAfter = endsWithBreak;
    // Bar tabs: vertical lines over the full line height.
    for (const bar of stops.bars) {
      const x = (rtl ? width - bar : bar) - left;
      line.items.push({
        kind: 'line',
        x1: x,
        y1: -line.baseline,
        x2: x,
        y2: line.height - line.baseline,
        color: '000000',
        width: 0.5,
      });
    }
    lines.push(line);
  }
  for (let b = 0; b < bookmarkAtom.length; b++) {
    const idx = bookmarkAtom[b]!;
    let li = ranges.findIndex((r) => idx < r.to);
    if (li < 0) li = lines.length - 1;
    lines[Math.max(0, li)]!.bookmarks.push(p.bookmarks[b]!);
  }

  // ---- borders and shading -------------------------------------------------------------------
  const borders = props.borders;
  const topPad = borderPad(borders.top);
  const bottomPad = borderPad(borders.bottom);
  const firstLine = lines[0]!;
  const lastLine = lines[lines.length - 1]!;
  if (topPad > 0) {
    firstLine.height += topPad;
    firstLine.baseline += topPad;
  }
  if (bottomPad > 0) lastLine.height += bottomPad;
  let decorations: DisplayItem[] = [];
  if (
    props.shading !== undefined ||
    borders.top ||
    borders.bottom ||
    borders.left ||
    borders.right
  ) {
    const height = lines.reduce((s, l) => s + l.height, 0);
    const boxLeft = Math.min(left, rtl ? left : left + firstOffset);
    const boxRight = width - Math.min(right, rtl ? right + firstOffset : right);
    decorations = paragraphDecorations(borders, props.shading, boxLeft, boxRight, height);
  }

  const sp = props.spacing;
  return {
    source: p.source,
    lines,
    spaceBefore: sp.beforeAutospacing ? AUTOSPACING : sp.before / 20,
    spaceAfter: sp.afterAutospacing ? AUTOSPACING : sp.after / 20,
    keepNext: props.keepNext,
    keepLines: props.keepLines,
    widowControl: props.widowControl,
    pageBreakBefore: props.pageBreakBefore,
    contextualSpacing: props.contextualSpacing,
    styleId: props.styleId,
    decorations,
    indentLeft: left,
  };
}
