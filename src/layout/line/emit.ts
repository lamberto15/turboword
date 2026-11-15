// One placed line → display items (relative to the paragraph's left indent and the baseline).
//
// 1. Alignment and justification, in start coordinates: `end`/`center` shift the line; `justify`
//    spreads the free space over the U+0020 spaces after the line's last tab (Word never stretches
//    the part before a tab); `distribute` spreads it between all clusters after the last tab.
// 2. Visual order (UAX #9 L1/L2): tabs split the line into segments laid out in paragraph order;
//    within a segment atoms are reordered by their embedding levels (trailing whitespace takes the
//    paragraph level); RTL paragraphs are mirrored (x = width - start-coordinate).
// 3. Items: highlight/shading rectangles first (behind), then text items (one per run of
//    contiguous glyph clusters of one shaped item), inline objects and tab leaders, then
//    underline/strike-through lines, then link areas.

import type { FieldSlot, LayoutEnv } from '../engine.js';
import { visualOrder } from '../text/index.js';
import type { DisplayItem, Font, LinkItem, TextItem } from '../types.js';
import type { PlacedLine } from './breaker.js';
import { colorOf, HIGHLIGHT, strikeItems, underlineItems } from './decorations.js';
import { fontExtents } from './metrics.js';
import { type Piece, translate } from './pieces.js';
import { type Atom, linkOf, textColor } from './shape.js';
import type { TabLeader } from './tabs.js';

export type LineAlign = 'start' | 'end' | 'center' | 'justify' | 'distribute';

export interface EmitContext {
  atoms: readonly Atom[];
  pieces: readonly Piece[];
  text: string;
  env: LayoutEnv;
  /** Paragraph base direction. */
  rtl: boolean;
  /** Column width, points. */
  width: number;
  /** Physical left indent: item x origin. */
  indentLeft: number;
}

export interface EmittedLine {
  items: DisplayItem[];
  fields: FieldSlot[];
  /** Largest ascent/descent of the text (and of the fonts of tabs, breaks, objects' runs). */
  ascent: number;
  descent: number;
  /** Largest ascent/descent of inline objects (pictures, drawings, math); 0 when none. */
  objectAscent: number;
  objectDescent: number;
  width: number;
  noteRefs: string[];
}

const EPS = 0.01;

const LEADER_CHAR: Readonly<Record<TabLeader, number>> = {
  dot: 0x2e,
  hyphen: 0x2d,
  underscore: 0x5f,
  heavy: 0x5f,
  middleDot: 0xb7,
};

/** Font and size for metrics/decorations of a non-glyph atom (tab, break, object). */
function pieceFont(ctx: EmitContext, piece: Piece): { font: Font; size: number; rise: number } {
  const factor = piece.kind === 'text' ? piece.factor : 1;
  const rise = piece.kind === 'text' ? piece.rise : 0;
  return { font: ctx.env.fonts.fontFor(piece.props, 0x20), size: piece.props.size * factor, rise };
}

function emptyText(x: number, y: number, fontIndex: number, size: number, color: string): TextItem {
  return {
    kind: 'text',
    x,
    y,
    font: fontIndex,
    size,
    color,
    glyphs: [],
    advances: [],
    text: '',
    clusters: [],
  };
}

/**
 * Display items and content extents of a placed line.
 * `brokeAfter`: the line was broken (it is not the paragraph's last), which shows a final soft hyphen.
 */
export function emitLine(
  ctx: EmitContext,
  placed: PlacedLine,
  align: LineAlign,
  brokeAfter: boolean,
  shrink = false,
): EmittedLine {
  const { atoms, pieces, env } = ctx;
  const { from, to, lastContent } = placed;
  const n = to - from;
  const X = placed.xs.slice();
  const W = placed.ws.slice();
  const ex = new Array<number>(n).fill(0);

  // ---- 1. alignment / justification ------------------------------------------------------
  let lastTab = from - 1;
  for (let j = from; j <= lastContent; j++) if (atoms[j]!.kind === 'tab') lastTab = j;
  const free = placed.geometry.end - placed.contentEnd;
  let shift = 0;
  if (free < -EPS && shrink) {
    // A justified line that only fits with narrower spaces (Word 2013+): shrink them.
    const slots: number[] = [];
    for (let j = lastTab + 1; j < lastContent; j++) if (atoms[j]!.space) slots.push(j - from);
    for (const k of slots) ex[k] = free / slots.length;
  } else if (free > EPS) {
    if (align === 'end') shift = free;
    else if (align === 'center') shift = free / 2;
    else if (align === 'justify' || align === 'distribute') {
      const slots: number[] = [];
      for (let j = lastTab + 1; j < lastContent; j++) {
        const a = atoms[j]!;
        if (align === 'justify' ? a.space : a.kind === 'glyphs' || a.kind === 'object')
          slots.push(j - from);
      }
      if (slots.length > 0) for (const k of slots) ex[k] = free / slots.length;
      else if (align === 'distribute') shift = free / 2;
    }
  }
  let acc = shift;
  for (let k = 0; k < n; k++) {
    X[k]! += acc;
    W[k]! += ex[k]!;
    acc += ex[k]!;
  }

  // ---- 2. physical placement --------------------------------------------------------------
  const base = ctx.rtl ? 1 : 0;
  const levels = new Uint8Array(n);
  let mixed = ctx.rtl;
  for (let k = 0; k < n; k++) {
    const a = atoms[from + k]!;
    const lvl = from + k > lastContent || a.kind === 'tab' || a.kind === 'break' ? base : a.level;
    levels[k] = lvl;
    if (lvl !== 0) mixed = true;
  }
  const phys = new Array<number>(n);
  const visual: number[] = [];
  if (!mixed) {
    for (let k = 0; k < n; k++) {
      phys[k] = X[k]!;
      visual.push(k);
    }
  } else {
    const segs: { start: number; order: number[]; left: number }[] = [];
    let s = 0;
    const flush = (e: number): void => {
      if (e <= s) return;
      const lv = levels.subarray(s, e);
      const order = visualOrder(new Uint8Array(lv));
      const xs = X[s]!;
      const xe = X[e - 1]! + W[e - 1]!;
      segs.push({ start: s, order, left: ctx.rtl ? ctx.width - xe : xs });
    };
    for (let k = 0; k < n; k++) {
      if (atoms[from + k]!.kind !== 'tab') continue;
      flush(k);
      segs.push({ start: k, order: [0], left: ctx.rtl ? ctx.width - X[k]! - W[k]! : X[k]! });
      s = k + 1;
    }
    flush(n);
    segs.sort((a, b) => a.left - b.left);
    for (const seg of segs) {
      let x = seg.left;
      for (const o of seg.order) {
        const k = seg.start + o;
        phys[k] = x;
        x += W[k]!;
        visual.push(k);
      }
    }
  }

  // ---- 3. items -------------------------------------------------------------------------
  const behind: DisplayItem[] = [];
  const content: DisplayItem[] = [];
  const over: DisplayItem[] = [];
  const links: DisplayItem[] = [];
  const fieldItems: { item: number; field: FieldSlot['field'] }[] = [];
  const dx = ctx.indentLeft;

  // Text runs.
  let cur: TextItem | undefined;
  let curRun: Atom['run'];
  let curEnd = 0;
  const closeText = (): void => {
    cur = undefined;
    curRun = undefined;
  };
  for (const k of visual) {
    const j = from + k;
    const a = atoms[j]!;
    if (a.kind === 'object') {
      closeText();
      const piece = pieces[a.piece]!;
      if (piece.kind === 'object')
        for (const it of piece.items) content.push(translate(it, phys[k]! - dx, 0));
      continue;
    }
    if (a.kind === 'tab') {
      closeText();
      const leader = placed.leaders.get(j);
      if (leader !== undefined && W[k]! > 0) {
        const li = leaderItem(ctx, pieces[a.piece]!, leader, phys[k]! - dx, W[k]!);
        if (li !== undefined) content.push(li);
      }
      continue;
    }
    if (a.kind !== 'glyphs' || a.run === undefined) {
      closeText();
      continue;
    }
    const run = a.run;
    const piece = pieces[a.piece]!;
    let glyphs: number[];
    let advs: number[];
    let chars = ctx.text.slice(a.start, a.end);
    glyphs = [];
    advs = [];
    if (chars !== '\u00ad') {
      const rev = run.reversed && run.rtl;
      for (let q = rev ? a.g1 - 1 : a.g0; rev ? q >= a.g0 : q < a.g1; q += rev ? -1 : 1) {
        glyphs.push(run.glyphs[q]!);
        advs.push(run.advances[q]!);
      }
      if (advs.length > 0) advs[advs.length - 1]! += ex[k]!;
    }
    if (a.soft) {
      // A soft hyphen shows (as '-') only where the line breaks after it.
      if (brokeAfter && j === lastContent) {
        const gid = run.font.glyphId(0x2d);
        const upm = run.font.metrics.unitsPerEm || 1000;
        glyphs.push(gid);
        advs.push(((run.font.advance(gid) * run.size) / upm) * (run.scale / 100));
        chars = `${chars.slice(0, -1)}-`;
      } else {
        chars = chars.slice(0, -1);
      }
      if (glyphs.length === 0) continue;
    }
    const x = phys[k]! - dx;
    const field = piece.kind === 'text' ? piece.field : undefined;
    if (cur === undefined || curRun !== run || Math.abs(curEnd - x) > EPS) {
      cur = emptyText(x, -run.rise, env.fonts.indexOf(run.font), run.size, run.color);
      if (run.scale !== 100) cur.scale = run.scale;
      if (run.fakeBold) cur.fakeBold = true;
      if (run.fakeItalic) cur.fakeItalic = true;
      curRun = run;
      curEnd = x;
      content.push(cur);
      // A field's placeholder is its own shaped item (pieces never share runs): one slot.
      if (field !== undefined && !fieldItems.some((f) => f.field === field)) {
        fieldItems.push({ item: content.length - 1, field });
      }
    }
    const at = cur.text.length;
    cur.text += chars;
    for (let q = 0; q < glyphs.length; q++) {
      cur.glyphs.push(glyphs[q]!);
      cur.advances.push(advs[q]!);
      cur.clusters.push(at);
      curEnd += advs[q]!;
    }
    if (glyphs.length === 0) curEnd += W[k]!;
  }

  // Decorations per piece, over contiguous content atoms.
  let g0 = 0;
  const vis = visual.filter((k) => from + k <= lastContent && atoms[from + k]!.kind !== 'break');
  while (g0 < vis.length) {
    const first = vis[g0]!;
    const pi = atoms[from + first]!.piece;
    let g1 = g0 + 1;
    while (
      g1 < vis.length &&
      atoms[from + vis[g1]!]!.piece === pi &&
      Math.abs(phys[vis[g1 - 1]!]! + W[vis[g1 - 1]!]! - phys[vis[g1]!]!) < EPS
    ) {
      g1++;
    }
    decorate(
      ctx,
      pieces[pi]!,
      vis.slice(g0, g1).map((k) => ({ atom: atoms[from + k]!, x: phys[k]! - dx, w: W[k]! })),
      behind,
      over,
      links,
    );
    g0 = g1;
  }

  // ---- extents and note references --------------------------------------------------------
  let ascent = 0;
  let descent = 0;
  let objectAscent = 0;
  let objectDescent = 0;
  const noteRefs: string[] = [];
  for (let j = from; j < to; j++) {
    const a = atoms[j]!;
    const piece = pieces[a.piece]!;
    let asc: number;
    let desc: number;
    if (a.kind === 'glyphs' && a.run !== undefined) {
      const e = fontExtents(a.run.font, a.run.size);
      asc = e.ascent + a.run.rise;
      desc = e.descent - a.run.rise;
    } else if (a.kind === 'object' && piece.kind === 'object') {
      objectAscent = Math.max(objectAscent, piece.ascent);
      objectDescent = Math.max(objectDescent, piece.descent);
      // The run the object sits in still has its font's height (text extents).
      const f = pieceFont(ctx, piece);
      const e = fontExtents(f.font, f.size);
      asc = e.ascent;
      desc = e.descent;
    } else {
      const f = pieceFont(ctx, piece);
      const e = fontExtents(f.font, f.size);
      asc = e.ascent + f.rise;
      desc = e.descent - f.rise;
    }
    if (asc > ascent) ascent = asc;
    if (desc > descent) descent = desc;
    if (piece.kind === 'text' && piece.note !== undefined && !noteRefs.includes(piece.note)) {
      noteRefs.push(piece.note);
    }
  }

  const items = [...behind, ...content, ...over, ...links];
  const fields: FieldSlot[] = fieldItems.map((f) => ({
    item: f.item + behind.length,
    field: f.field,
  }));
  // Width used = the content's right edge measured from the left indent (so indentLeft + width
  // is where the line's content ends in the column; hanging and first-line indents included).
  let right = 0;
  for (let k = 0; k <= lastContent - from; k++) right = Math.max(right, phys[k]! + W[k]! - dx);
  return { items, fields, ascent, descent, objectAscent, objectDescent, width: right, noteRefs };
}

/** A run of leader characters filling a tab of width `w` at `x` (right-aligned, small gaps). */
function leaderItem(
  ctx: EmitContext,
  piece: Piece,
  leader: TabLeader,
  x: number,
  w: number,
): TextItem | undefined {
  const cp = LEADER_CHAR[leader];
  const size = piece.props.size;
  const font = ctx.env.fonts.fontFor(piece.props, cp);
  const gid = font.glyphId(cp);
  const upm = font.metrics.unitsPerEm || 1000;
  const adv = (font.advance(gid) * size) / upm;
  if (adv <= 0) return undefined;
  const count = Math.floor((w - adv * 0.5) / adv);
  if (count <= 0) return undefined;
  const ch = String.fromCodePoint(cp);
  const item = emptyText(
    x + w - count * adv - adv * 0.25,
    0,
    ctx.env.fonts.indexOf(font),
    size,
    textColor(piece.props),
  );
  for (let i = 0; i < count; i++) {
    item.glyphs.push(gid);
    item.advances.push(adv);
    item.clusters.push(i);
    item.text += ch;
  }
  return item;
}

/** Highlight/shading, underline, strike-through and link items for one piece's contiguous atoms. */
function decorate(
  ctx: EmitContext,
  piece: Piece,
  parts: { atom: Atom; x: number; w: number }[],
  behind: DisplayItem[],
  over: DisplayItem[],
  links: DisplayItem[],
): void {
  if (parts.length === 0) return;
  const props = piece.props;
  const glyphAtom = parts.find((p) => p.atom.run !== undefined)?.atom;
  const pf = pieceFont(ctx, piece);
  const font = glyphAtom?.run?.font ?? pf.font;
  const size = glyphAtom?.run?.size ?? pf.size;
  const rise = glyphAtom?.run?.rise ?? pf.rise;
  const x1 = parts[0]!.x;
  const last = parts[parts.length - 1]!;
  const x2 = last.x + last.w;
  const ext = fontExtents(font, size);
  const top = -ext.ascent - rise;
  const height = ext.ascent + ext.descent;
  const fill = props.highlight !== undefined ? HIGHLIGHT[props.highlight] : undefined;
  if (props.shading !== undefined && /^[0-9A-Fa-f]{6}$/.test(props.shading)) {
    behind.push({
      kind: 'rect',
      x: x1,
      y: top,
      width: x2 - x1,
      height,
      fill: props.shading.toUpperCase(),
    });
  }
  if (fill !== undefined)
    behind.push({ kind: 'rect', x: x1, y: top, width: x2 - x1, height, fill });
  const color = textColor(props);
  const u = props.underline;
  if (u !== undefined && u !== 'none' && piece.kind !== 'object') {
    const ucolor = colorOf(props.underlineColor, color);
    if (u === 'words') {
      let s = -1;
      for (let i = 0; i <= parts.length; i++) {
        const p = parts[i];
        const gap = p === undefined || p.atom.white;
        if (!gap && s < 0) s = i;
        if (gap && s >= 0) {
          const e = parts[i - 1]!;
          over.push(...underlineItems('single', parts[s]!.x, e.x + e.w, font, size, rise, ucolor));
          s = -1;
        }
      }
    } else {
      over.push(...underlineItems(u, x1, x2, font, size, rise, ucolor));
    }
  }
  if ((props.strike || props.doubleStrike) && piece.kind !== 'object') {
    over.push(...strikeItems(props.doubleStrike, x1, x2, font, size, rise, color));
  }
  const link = linkOf(piece);
  if (link !== undefined) {
    const li: LinkItem = {
      kind: 'link',
      x: x1,
      y: piece.kind === 'object' ? -piece.ascent : top,
      width: x2 - x1,
      height: piece.kind === 'object' ? piece.ascent + piece.descent : height,
    };
    if (link.url !== undefined) li.url = link.url;
    if (link.anchor !== undefined) li.dest = link.anchor;
    links.push(li);
  }
}
