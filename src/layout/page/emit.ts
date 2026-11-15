// Turns the flow's pages into the PageModel: absolute display items, headers/footers laid out per
// page with their page fields filled in, footnote areas, floats in z-order, page vertical
// alignment, and the document-level maps (outline, destinations, paragraph → page).
//
// Item order on a page (painter's order): floats behind text, headers/footers, body columns
// with their separators and footnotes, floats in front of text.

import type { XElement } from '../../xml/dom.js';
import type { DisplayItem, OutlineEntry, PageLayout, TextItem } from '../types.js';
import { stackBlocks } from './blocks.js';
import { fieldValue, type PageFacts } from './fields.js';
import { floatItems, positionFloat } from './floats.js';
import { contentBottom, type Flow, type PageState, type Placed } from './flow.js';
import {
  allParagraphs,
  type FieldInline,
  fieldsOf,
  type LayoutCtx,
  moveItem,
  type Strip,
  stackHeight,
  stackLayout,
} from './strip.js';

const SEPARATOR_WIDTH = 144;

/** Field facts of a finished flow. */
export function pageFacts(pages: readonly PageState[]): PageFacts {
  const bookmarks = new Map<string, number>();
  for (const p of pages) {
    for (const b of p.bands) {
      for (const c of b.columns) {
        for (const pl of c.placed) {
          for (const name of pl.strip.bookmarks ?? [])
            if (!bookmarks.has(name)) bookmarks.set(name, p.index);
        }
      }
    }
  }
  return {
    pages: pages.map((p) => ({
      value: p.value,
      format: p.section.pageNumberFormat,
      sectionIndex: p.sectionIndex,
    })),
    bookmarks,
  };
}

/**
 * Fields of the body (and notes) with the value they must show where they landed. Their
 * fallbacks feed the next layout pass (fixed point).
 */
export function bodyFieldValues(
  pages: readonly PageState[],
  facts: PageFacts,
): Map<FieldInline, string> {
  const out = new Map<FieldInline, string>();
  const visit = (s: Strip, index: number): void => {
    if (s.fields === undefined) return;
    for (const f of s.fields.values()) {
      const v = fieldValue(f, index, facts);
      if (v !== undefined && !out.has(f)) out.set(f, v);
    }
  };
  for (const p of pages) {
    for (const b of p.bands) {
      for (const c of b.columns) {
        for (const pl of c.placed) visit(pl.strip, p.index);
        for (const s of c.noteStrips ?? []) visit(s, p.index);
      }
    }
  }
  return out;
}

/** A copy of a placeholder text item showing `text` (re-shaped with the same face). */
function substitute(ctx: LayoutCtx, it: TextItem, text: string): TextItem {
  const font = ctx.env.fonts.fonts()[it.font];
  if (font === undefined) return it;
  const glyphs = ctx.env.shaper.shape(font, text, { rtl: false, kerning: false });
  const k = ((it.size / font.metrics.unitsPerEm) * (it.scale ?? 100)) / 100;
  const out: TextItem = {
    ...it,
    text,
    glyphs: glyphs.map((g) => g.id),
    advances: glyphs.map((g) => g.advance * k),
    clusters: glyphs.map((g) => g.cluster),
  };
  delete out.dy;
  return out;
}

type FieldText = (f: FieldInline) => string | undefined;

function emitStrip(
  ctx: LayoutCtx,
  out: DisplayItem[],
  s: Strip,
  x: number,
  y: number,
  fieldText: FieldText,
): void {
  for (const it of s.items) {
    let src = it;
    const f = s.fields?.get(it);
    if (f !== undefined && it.kind === 'text') {
      const v = fieldText(f);
      if (v !== undefined && v !== it.text) src = substitute(ctx, it, v);
    }
    out.push(moveItem(src, x, y));
  }
}

/** Header/footer strips with their fields showing this page's values. */
function hfStrips(
  ctx: LayoutCtx,
  blocks: NonNullable<PageState['header']>,
  width: number,
  fieldText: FieldText,
): Strip[] {
  const fields: FieldInline[] = [];
  for (const p of allParagraphs(blocks)) fields.push(...fieldsOf(p));
  const saved = fields.map((f) => f.fallback);
  for (const f of fields) {
    const v = fieldText(f);
    if (v !== undefined) f.fallback = v;
  }
  try {
    return stackBlocks(ctx, blocks, width);
  } finally {
    fields.forEach((f, i) => {
      f.fallback = saved[i]!;
    });
  }
}

/** Per-strip vertical shift implementing the section's page vAlign (single-band pages). */
function vAlignShift(page: PageState, sepH: number): (c: number, i: number) => number {
  const v = page.section.vAlign;
  if (v === 'top' || page.bands.length !== 1) return () => 0;
  const band = page.bands[0]!;
  let used = band.top;
  let notes = 0;
  for (const c of band.columns) {
    used = Math.max(used, contentBottom(c));
    const ns = c.noteStrips ?? [];
    notes = Math.max(notes, ns.length > 0 ? stackHeight(ns) + sepH : 0);
  }
  const extra = band.bottom - notes - used;
  if (extra <= 0) return () => 0;
  if (v === 'center') return () => extra / 2;
  if (v === 'bottom') return () => extra;
  // both: spread the free space over the gaps between paragraphs
  const gaps = band.columns.map((c) => {
    const at: number[] = [];
    let n = 0;
    c.placed.forEach((p, i) => {
      if (i > 0 && p.strip.starts !== undefined) n++;
      at.push(n);
    });
    return { at, n };
  });
  return (c, i) => {
    const g = gaps[c];
    if (g === undefined || g.n === 0) return 0;
    return (extra * (g.at[i] ?? 0)) / g.n;
  };
}

export interface EmitResult {
  pages: PageLayout[];
  outline: OutlineEntry[];
  destinations: Map<string, { page: number; y: number }>;
  paragraphPages: Map<XElement, number>;
}

/** Builds the PageModel from a finished flow. */
export function emit(ctx: LayoutCtx, flow: Flow): EmitResult {
  const pages = flow.pages;
  const facts = pageFacts(pages);
  const outline: OutlineEntry[] = [];
  const destinations = new Map<string, { page: number; y: number }>();
  const paragraphPages = new Map<XElement, number>();
  const result: PageLayout[] = [];
  for (const page of pages) {
    const fieldText: FieldText = (f) => fieldValue(f, page.index, facts);
    const behind: DisplayItem[] = [];
    const front: DisplayItem[] = [];
    const hf: DisplayItem[] = [];
    const body: DisplayItem[] = [];
    const hfWidth = Math.max(1, page.width - page.marginLeft - page.marginRight);
    const g = page.section.margins;
    // headers and footers
    const emitHf = (blocks: PageState['header'], footer: boolean): void => {
      if (blocks === undefined || blocks.length === 0) return;
      const strips = hfStrips(ctx, blocks, hfWidth, fieldText);
      const stack = stackLayout(strips);
      const y0 = footer ? page.height - g.footer - stack.height : g.header;
      strips.forEach((s, si) => {
        const y = y0 + stack.ys[si]!;
        emitStrip(ctx, hf, s, page.marginLeft, y, fieldText);
        for (const a of s.anchors ?? []) {
          const pos = positionFloat(a.anchor, {
            pageWidth: page.width,
            pageHeight: page.height,
            marginLeft: page.marginLeft,
            marginRight: page.marginRight,
            marginTop: page.bodyTop,
            marginBottom: page.height - page.bodyBottom,
            columnX: page.marginLeft,
            columnWidth: hfWidth,
            paraTop: y + a.dy,
            paraLeft: page.marginLeft + a.dx,
            pageIndex: page.index,
          });
          (a.anchor.behindText ? behind : front).push(...floatItems(ctx, a.anchor, pos.x, pos.y));
        }
      });
    };
    emitHf(page.header, false);
    emitHf(page.footer, true);
    // body
    const shift = vAlignShift(page, flow.sepH);
    const single = page.bands.length === 1;
    const pageNotes: Strip[] = [];
    let noteWidthX = page.marginLeft;
    for (const band of page.bands) {
      band.columns.forEach((c, ci) => {
        c.placed.forEach((pl: Placed, i) => {
          const y = pl.y + shift(ci, i);
          emitStrip(ctx, body, pl.strip, c.x, y, fieldText);
          const s = pl.strip;
          for (const src of s.starts ?? [])
            if (!paragraphPages.has(src)) paragraphPages.set(src, page.index);
          for (const name of s.bookmarks ?? []) {
            if (!destinations.has(name)) destinations.set(name, { page: page.index, y });
          }
          if (s.outline !== undefined && pl.repeat !== true) {
            outline.push({ ...s.outline, page: page.index, y });
          }
        });
        const notes = c.noteStrips ?? [];
        if (notes.length > 0) {
          if (single) {
            let y = c.bottom - stackHeight(notes) - flow.sepH;
            const cont = c.carryIn.length > 0;
            body.push({
              kind: 'line',
              x1: c.x,
              y1: y + flow.sepH / 2,
              x2: c.x + (cont ? c.width : Math.min(SEPARATOR_WIDTH, c.width)),
              y2: y + flow.sepH / 2,
              color: '000000',
              width: 0.5,
            });
            y += flow.sepH;
            const stack = stackLayout(notes);
            notes.forEach((s, si) => {
              emitStrip(ctx, body, s, c.x, y + stack.ys[si]!, fieldText);
            });
          } else {
            pageNotes.push(...notes);
            noteWidthX = c.x;
          }
        }
      });
      if (band.separator && band.columns.length > 1) {
        let bottom = band.top;
        for (const c of band.columns) bottom = Math.max(bottom, contentBottom(c));
        for (let i = 1; i < band.columns.length; i++) {
          const a = band.columns[i - 1]!;
          const b = band.columns[i]!;
          const x = (a.x + a.width + b.x) / 2;
          body.push({
            kind: 'line',
            x1: x,
            y1: band.top,
            x2: x,
            y2: bottom,
            color: '000000',
            width: 0.5,
          });
        }
      }
    }
    if (pageNotes.length > 0) {
      let y = page.bodyBottom - stackHeight(pageNotes) - flow.sepH;
      body.push({
        kind: 'line',
        x1: noteWidthX,
        y1: y + flow.sepH / 2,
        x2: noteWidthX + SEPARATOR_WIDTH,
        y2: y + flow.sepH / 2,
        color: '000000',
        width: 0.5,
      });
      y += flow.sepH;
      const stack = stackLayout(pageNotes);
      pageNotes.forEach((s, si) => {
        emitStrip(ctx, body, s, noteWidthX, y + stack.ys[si]!, fieldText);
      });
    }
    // floats, by z-order
    const floats = [...page.floats].sort((a, b) => a.anchor.zIndex - b.anchor.zIndex);
    for (const f of floats) {
      (f.anchor.behindText ? behind : front).push(...floatItems(ctx, f.anchor, f.x, f.y));
    }
    result.push({
      index: page.index,
      number: page.number,
      section: page.sectionIndex,
      width: page.width,
      height: page.height,
      items: [...behind, ...hf, ...body, ...front],
    });
  }
  return { pages: result, outline, destinations, paragraphPages };
}
