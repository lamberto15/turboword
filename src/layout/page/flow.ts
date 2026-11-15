// The page flow: places strips into columns, columns into bands, bands into pages.
//
// A page has one band per section that shares it (continuous section breaks start a new band
// below the previous one); a band has the section's columns. Strips are placed one at a time.
// When a strip does not fit, the flow looks back in the current column for the last strip a
// break may come before (keep-with-next, keep-lines and widow/orphan control are all expressed as
// `breakBefore: 'avoid'`) and moves everything from there on to the next column; when no such
// point exists the break is forced where the strip is, as Word does when a keep chain is longer
// than a page. Splittable table rows are split at the bottom instead; header rows of a table
// are repeated at the top of every column it continues into.
//
// Footnotes referenced by a strip are reserved in the same column (the strip moves on when its
// notes do not fit); notes that cannot fit even in an empty column continue below the
// continuation separator of the next one. Floats are positioned when the strip of their
// anchoring paragraph is placed (and re-positioned if it moves); their exclusion zones push
// later strips below them (topAndBottom) or narrow later paragraphs (square).
//
// Space before a paragraph is dropped at the top of a column reached by a soft (automatic)
// break and kept after a hard break (page break, section break, the first page), as Word does.

import { tableStrips } from '../table.js';
import type { IRBlock, IRDocument, IRParagraph, IRSection } from '../types.js';
import { paragraphStrips, stackBlocks } from './blocks.js';
import { pageNumberText } from './fields.js';
import {
  type Exclusion,
  exclusionOf,
  type FloatFrame,
  freeSpace,
  MIN_WRAP_GAP,
  type PlacedFloat,
  positionFloat,
} from './floats.js';
import {
  allParagraphs,
  gap,
  type InsetSpec,
  type LayoutCtx,
  measure,
  moveItem,
  type Strip,
  stackHeight,
  stackLayout,
} from './strip.js';

const EPS = 0.01;
/** Hard stop against runaway layouts (a page per strip at worst). */
const MAX_PAGES = 100_000;
/** Width of the footnote separator line (Word: about 2 inches). */
const SEPARATOR_WIDTH = 144;

/** A strip placed in a column. */
export interface Placed {
  strip: Strip;
  /** Content top (after the applied space before), page coordinates. */
  y: number;
  /** A repeated table header row. */
  repeat?: boolean;
  floats?: PlacedFloat[];
}

export interface Column {
  x: number;
  width: number;
  top: number;
  bottom: number;
  placed: Placed[];
  /** Reached by a hard break (space before is kept at its top). */
  topHard: boolean;
  /** Footnote ids referenced in this column, in order. */
  notes: string[];
  /** Note strips continued from the previous column. */
  carryIn: Strip[];
  /** Note strips shown at the bottom (decided when the column is left). */
  noteStrips?: Strip[];
}

export interface Band {
  top: number;
  bottom: number;
  columns: Column[];
  section: IRSection;
  separator: boolean;
}

/** A page under construction. */
export interface PageState {
  index: number;
  /** The section that started the page (geometry, headers, page numbering). */
  section: IRSection;
  sectionIndex: number;
  value: number;
  number: string;
  width: number;
  height: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  bodyTop: number;
  bodyBottom: number;
  header?: IRBlock[];
  footer?: IRBlock[];
  bands: Band[];
  floats: PlacedFloat[];
  firstOfSection: boolean;
}

interface Geometry {
  width: number;
  height: number;
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

function geometry(sec: IRSection): Geometry {
  const m = sec.margins;
  return {
    width: sec.pageWidth > 0 ? sec.pageWidth : 612,
    height: sec.pageHeight > 0 ? sec.pageHeight : 792,
    left: m.left + m.gutter,
    right: m.right,
    top: m.top,
    bottom: m.bottom,
    header: m.header,
    footer: m.footer,
  };
}

function columnsOf(sec: IRSection, g: Geometry, top: number, bottom: number): Column[] {
  const left = g.left;
  const width = Math.max(1, g.width - g.left - g.right);
  const c = sec.columns;
  const n = Math.max(1, c.count);
  const out: Column[] = [];
  const mk = (x: number, w: number): Column => ({
    x,
    width: Math.max(1, w),
    top,
    bottom,
    placed: [],
    topHard: true,
    notes: [],
    carryIn: [],
  });
  if (n > 1 && c.widths !== undefined && c.widths.length >= n) {
    let x = left;
    for (let i = 0; i < n; i++) {
      const cw = c.widths[i]!;
      out.push(mk(x, cw.width));
      x += cw.width + cw.space;
    }
  } else {
    const w = (width - c.space * (n - 1)) / n;
    for (let i = 0; i < n; i++) out.push(mk(left + i * (w + c.space), w));
  }
  return out;
}

/** Content bottom of a column (last strip's bottom, its space after excluded). */
export function contentBottom(c: Column): number {
  const last = c.placed[c.placed.length - 1];
  return last === undefined ? c.top : last.y + last.strip.height;
}

/** Content top of `s` if it were placed next in `c`. */
function nextY(c: Column, s: Strip): number {
  const last = c.placed[c.placed.length - 1];
  if (last === undefined) return c.top + (c.topHard ? s.spaceBefore : 0);
  return last.y + last.strip.height + gap(last.strip, s);
}

/** Index to break before in `c` so that `s` can move on; undefined = column is (effectively) empty. */
function findBreak(c: Column, s: Strip): number | undefined {
  let lead = 0;
  while (lead < c.placed.length && c.placed[lead]!.repeat === true) lead++;
  if (lead === c.placed.length) return undefined;
  if (s.breakBefore !== 'avoid') return c.placed.length;
  for (let k = c.placed.length - 1; k >= Math.max(1, lead + 1); k--) {
    if (c.placed[k]!.strip.breakBefore === 'auto') return k;
  }
  return c.placed.length;
}

/** Lays the document's sections out into pages (one pass, current field fallbacks). */
export class Flow {
  readonly ctx: LayoutCtx;
  readonly ir: IRDocument;
  readonly pages: PageState[] = [];
  /** Height of the footnote separator paragraph. */
  readonly sepH: number;
  #page: PageState | undefined;
  #band: Band | undefined;
  #col = 0;
  #section: IRSection;
  #sectionIndex = 0;
  #firstOfSection = true;
  #queue: Strip[] = [];
  #pendingBreak: 'page' | 'column' | undefined;
  #carry: Strip[] = [];
  #last: Strip | undefined;
  #tables = new Map<object, Strip[]>();
  #notes = new Map<string, Strip[]>();
  #hf = new Map<IRBlock[], Map<number, number>>();

  constructor(ctx: LayoutCtx, ir: IRDocument) {
    this.ctx = ctx;
    this.ir = ir;
    this.#section = ir.sections[0]!;
    let size = 11;
    for (const s of ir.sections) {
      const p = s.blocks.find((b): b is IRParagraph => b.kind === 'paragraph');
      if (p !== undefined) {
        size = p.props.runDefaults.size || size;
        break;
      }
    }
    this.sepH = size * 1.22;
  }

  /** Runs the flow over every section (and the endnotes) and returns the pages. */
  run(): PageState[] {
    const secs = this.ir.sections;
    secs.forEach((sec, i) => {
      this.#startSection(sec, i);
      for (const b of sec.blocks) this.addBlock(b);
    });
    this.#endnotes();
    this.#drain();
    if (this.#page === undefined) this.#newPage('hard');
    this.#closeColumn(this.#column());
    return this.pages;
  }

  // ---- blocks ---------------------------------------------------------------------------------

  /** Converts a block into strips and places them. */
  addBlock(b: IRBlock): void {
    if (this.#page === undefined) this.#newPage('hard');
    const prev = this.#last;
    const keep = prev?.paraEnd?.keepNext === true;
    const c = this.#column();
    let strips: Strip[];
    if (b.kind === 'paragraph') {
      const rule = b.props.pageBreakBefore ? 'page' : keep ? 'avoid' : 'auto';
      const { width, left, insets } = this.#paragraphBox(b, c);
      strips = paragraphStrips(this.ctx, b, width, prev, rule, insets);
      if (left !== 0) {
        for (const s of strips) {
          s.items = s.items.map((it) => {
            const moved = moveItem(it, left, 0);
            const f = s.fields?.get(it);
            if (f !== undefined) {
              s.fields!.delete(it);
              s.fields!.set(moved, f);
            }
            return moved;
          });
          if (s.anchors) for (const a of s.anchors) a.dx += left;
        }
      }
    } else {
      strips = tableStrips(this.ctx, b, c.width, keep);
      const headers = strips.filter((s) => s.table?.header === true);
      if (headers.length > 0) this.#tables.set(b, headers);
    }
    if (strips.length === 0) return;
    this.#last = strips[strips.length - 1];
    this.#queue.push(...strips);
    this.#drain();
  }

  /**
   * How a paragraph wraps beside floats: per-line insets when the paragraph layouter supports
   * them (ParagraphConstraints), else the whole paragraph narrowed to the free side.
   */
  #paragraphBox(p: IRParagraph, c: Column): { width: number; left: number; insets?: InsetSpec } {
    const page = this.#page!;
    const excl: Exclusion[] = [];
    for (const f of page.floats) if (f.exclusion) excl.push(f.exclusion);
    const top = contentBottom(c);
    if (p.anchors.length > 0) {
      for (const a of p.anchors) {
        const pos = positionFloat(a, this.#frame(c, top, c.x));
        const e = exclusionOf(a, pos.x, pos.y);
        if (e !== undefined) excl.push(e);
      }
    }
    const plain = { width: c.width, left: 0 };
    if (excl.length === 0) return plain;
    const full = measure(this.ctx, p, c.width);
    const h = full.lines.reduce((sum, l) => sum + l.height, 0);
    const whole = freeSpace(excl, c.x, c.width, top, top + Math.max(h, 1));
    if (whole === null) return plain;
    // per line: the free side at the line's estimated position (full-width line heights)
    const pitch = full.lines.length > 0 ? h / full.lines.length : 12;
    const insetOf = (i: number): { left: number; right: number } | undefined => {
      const l = full.lines[i];
      let y = top;
      for (let j = 0; j < i; j++) y += full.lines[j]?.height ?? pitch;
      const fs = freeSpace(excl, c.x, c.width, y, y + (l?.height ?? pitch));
      if (fs === null) return { left: 0, right: 0 };
      if (fs.width < MIN_WRAP_GAP) return undefined;
      return { left: fs.left, right: c.width - fs.left - fs.width };
    };
    const est = Math.max(full.lines.length * 2, 4);
    const table: { left: number; right: number }[] = [];
    let blocked = false;
    for (let i = 0; i < est; i++) {
      const ins = insetOf(i);
      if (ins === undefined) blocked = true;
      table.push(ins ?? { left: 0, right: 0 });
    }
    if (!blocked) {
      const spec: InsetSpec = {
        insets: (i) => table[Math.min(i, table.length - 1)] ?? { left: 0, right: 0 },
        key: table.map((t) => `${t.left.toFixed(2)},${t.right.toFixed(2)}`).join(';'),
      };
      const pl = measure(this.ctx, p, c.width, spec);
      // a layouter that ignores the insets leaves lines too wide or not moved past a left float
      const honoured = pl.lines.every((l, i) => {
        const t = spec.insets(i);
        const fits = pl.indentLeft + l.width <= c.width - t.right + 0.5 || l.items.length <= 1;
        let minX = Number.POSITIVE_INFINITY;
        for (const it of l.items) if (it.kind === 'text') minX = Math.min(minX, it.x);
        return fits && (t.left <= 0.5 || minX >= t.left - 0.5 || minX === Number.POSITIVE_INFINITY);
      });
      if (honoured && pl !== full) return { width: c.width, left: 0, insets: spec };
    }
    if (whole.width < MIN_WRAP_GAP) return plain;
    return { width: whole.width, left: whole.left };
  }

  #endnotes(): void {
    const ids: string[] = [];
    const seen = new Set<string>();
    for (const sec of this.ir.sections) {
      for (const p of allParagraphs(sec.blocks)) {
        for (const i of p.inlines) {
          if (i.kind === 'noteRef' && i.note === 'endnote' && !seen.has(i.id)) {
            seen.add(i.id);
            ids.push(i.id);
          }
        }
      }
    }
    const notes = ids.map((id) => this.ir.endnotes.get(id)).filter((n) => n !== undefined);
    if (notes.length === 0) return;
    if (this.#page === undefined) this.#newPage('hard');
    const sep: Strip = {
      height: this.sepH,
      spaceBefore: 0,
      spaceAfter: 0,
      items: [
        {
          kind: 'line',
          x1: 0,
          y1: this.sepH / 2,
          x2: Math.min(SEPARATOR_WIDTH, this.#column().width),
          y2: this.sepH / 2,
          color: '000000',
          width: 0.5,
        },
      ],
      breakBefore: 'auto',
    };
    this.#last = sep;
    this.#queue.push(sep);
    this.#drain();
    for (const blocks of notes) for (const b of blocks) this.addBlock(b);
  }

  // ---- placement ------------------------------------------------------------------------------

  #column(): Column {
    return this.#band!.columns[this.#col]!;
  }

  #pageHasContent(): boolean {
    for (const b of this.#page!.bands)
      for (const c of b.columns) if (c.placed.length > 0) return true;
    return false;
  }

  #exclusions(): Exclusion[] {
    const out: Exclusion[] = [];
    for (const f of this.#page!.floats) if (f.exclusion) out.push(f.exclusion);
    return out;
  }

  /** y for `s` in `c`, pushed below floats that leave no room beside them. */
  #yFor(c: Column, s: Strip): number {
    let y = nextY(c, s);
    const excl = this.#exclusions();
    if (excl.length === 0) return y;
    for (let guard = 0; guard < 32; guard++) {
      const fs = freeSpace(excl, c.x, c.width, y, y + Math.max(s.height, 0.5));
      if (fs === null || fs.width >= MIN_WRAP_GAP || fs.below <= y) break;
      y = fs.below;
    }
    return y;
  }

  #noteStrips(id: string, width: number): Strip[] {
    const key = `${id}\u0001${width.toFixed(2)}`;
    let s = this.#notes.get(key);
    if (s === undefined) {
      const blocks = this.ir.footnotes.get(id);
      s = blocks === undefined ? [] : stackBlocks(this.ctx, blocks, width);
      this.#notes.set(key, s);
    }
    return s;
  }

  /** Footnote area height of `c`, with the notes of `s` added. */
  #noteHeight(c: Column, s?: Strip): number {
    let h = stackHeight(c.carryIn);
    let any = c.carryIn.length > 0;
    for (const id of c.notes) {
      h += stackHeight(this.#noteStrips(id, c.width));
      any = true;
    }
    if (s?.footnotes) {
      for (const id of s.footnotes) {
        if (c.notes.includes(id)) continue;
        h += stackHeight(this.#noteStrips(id, c.width));
        any = true;
      }
    }
    return any ? h + this.sepH : 0;
  }

  #frame(c: Column, paraTop: number, paraLeft: number): FloatFrame {
    const p = this.#page!;
    return {
      pageWidth: p.width,
      pageHeight: p.height,
      marginLeft: p.marginLeft,
      marginRight: p.marginRight,
      marginTop: p.bodyTop,
      marginBottom: p.height - p.bodyBottom,
      columnX: c.x,
      columnWidth: c.width,
      paraTop,
      paraLeft,
      pageIndex: p.index,
    };
  }

  /** Positions the floats anchored in `s` against a paragraph top `y` and registers them. */
  #placeFloats(c: Column, s: Strip, y: number): PlacedFloat[] | undefined {
    if (s.anchors === undefined || s.anchors.length === 0) return undefined;
    const page = this.#page!;
    return s.anchors.map((ref) => {
      const pos = positionFloat(ref.anchor, this.#frame(c, y + ref.dy, c.x + ref.dx));
      const pf: PlacedFloat = { anchor: ref.anchor, x: pos.x, y: pos.y };
      const e = exclusionOf(ref.anchor, pos.x, pos.y);
      if (e !== undefined) pf.exclusion = e;
      page.floats.push(pf);
      return pf;
    });
  }

  #dropFloats(floats: PlacedFloat[] | undefined): void {
    if (floats === undefined) return;
    const page = this.#page!;
    page.floats = page.floats.filter((f) => !floats.includes(f));
  }

  #commit(c: Column, s: Strip, y: number, repeat: boolean, floats?: PlacedFloat[]): void {
    const placed: Placed = { strip: s, y };
    if (repeat) placed.repeat = true;
    if (s.footnotes && !repeat) {
      for (const id of s.footnotes) if (!c.notes.includes(id)) c.notes.push(id);
    }
    if (floats !== undefined) placed.floats = floats;
    c.placed.push(placed);
  }

  #unplace(c: Column, k: number): Strip[] {
    const removed = c.placed.splice(k);
    for (const p of removed) this.#dropFloats(p.floats);
    c.notes = [];
    for (const p of c.placed) {
      if (p.strip.footnotes)
        for (const id of p.strip.footnotes) if (!c.notes.includes(id)) c.notes.push(id);
    }
    return removed.filter((p) => p.repeat !== true).map((p) => p.strip);
  }

  #drain(): void {
    const q = this.#queue;
    let repeatCount = 0; // leading queue entries that are repeated header rows
    while (q.length > 0) {
      if (this.#pendingBreak !== undefined) {
        const kind = this.#pendingBreak;
        this.#pendingBreak = undefined;
        if (kind === 'page') this.#newPage('hard');
        else this.#newColumn('hard');
        repeatCount = this.#repeatHeaders();
        continue;
      }
      const s = q[0]!;
      const repeat = repeatCount > 0;
      if (!repeat && s.breakBefore === 'page' && this.#pageHasContent()) {
        this.#newPage('hard');
        repeatCount = this.#repeatHeaders();
        continue;
      }
      if (!repeat && s.breakBefore === 'column' && this.#column().placed.length > 0) {
        this.#newColumn('hard');
        repeatCount = this.#repeatHeaders();
        continue;
      }
      const c = this.#column();
      // floats anchored here are positioned against where the paragraph would start; its own
      // lines then avoid them like any later text
      const floats = repeat ? undefined : this.#placeFloats(c, s, nextY(c, s));
      const y = this.#yFor(c, s);
      const limit = c.bottom - this.#noteHeight(c, s);
      if (y + s.height <= limit + EPS) {
        q.shift();
        if (repeatCount > 0) repeatCount--;
        this.#commit(c, s, y, repeat, floats);
        if (s.breakAfter !== undefined && !repeat) this.#pendingBreak = s.breakAfter;
        continue;
      }
      this.#dropFloats(floats);
      if (s.split !== undefined && !repeat) {
        const parts = limit - y > 1 ? s.split(limit - y) : undefined;
        if (parts !== undefined) {
          q.splice(0, 1, parts[0], parts[1]);
          continue;
        }
      }
      const k = findBreak(c, s);
      if (k === undefined || repeat) {
        // nothing to move: place it anyway (overflows the column)
        q.shift();
        if (repeatCount > 0) repeatCount--;
        this.#commit(c, s, y, repeat, repeat ? undefined : this.#placeFloats(c, s, nextY(c, s)));
        if (s.breakAfter !== undefined && !repeat) this.#pendingBreak = s.breakAfter;
        continue;
      }
      const moved = this.#unplace(c, k);
      q.unshift(...moved);
      this.#newColumn('soft');
      repeatCount = this.#repeatHeaders();
    }
  }

  /** Re-inserts a continuing table's header rows at the queue front; returns how many. */
  #repeatHeaders(): number {
    const f = this.#queue[0];
    if (f?.table === undefined || f.table.header) return 0;
    const hs = this.#tables.get(f.table.key);
    if (hs === undefined || hs.length === 0) return 0;
    this.#queue.unshift(...hs);
    return hs.length;
  }

  // ---- columns, pages, sections ---------------------------------------------------------------

  /** Decides the note strips shown in `c`; what does not fit continues in the next column. */
  #closeColumn(c: Column): void {
    if (c.noteStrips !== undefined) return;
    const all: Strip[] = [...c.carryIn];
    for (const id of c.notes) all.push(...this.#noteStrips(id, c.width));
    if (all.length === 0) {
      c.noteStrips = [];
      return;
    }
    const room = c.bottom - contentBottom(c) - this.sepH;
    const { ys } = stackLayout(all);
    let i = 0;
    while (i < all.length && (i === 0 || ys[i]! + all[i]!.height <= room + EPS)) i++;
    const shown = all.slice(0, i);
    c.noteStrips = shown;
    this.#carry = all.slice(i);
  }

  #newColumn(kind: 'soft' | 'hard'): void {
    const band = this.#band!;
    this.#closeColumn(this.#column());
    if (this.#col + 1 < band.columns.length) {
      this.#col++;
      const c = this.#column();
      c.topHard = kind === 'hard';
      c.carryIn = this.#carry;
      this.#carry = [];
    } else this.#newPage(kind);
  }

  #hfHeight(blocks: IRBlock[] | undefined, width: number): number {
    if (blocks === undefined || blocks.length === 0) return 0;
    let byWidth = this.#hf.get(blocks);
    if (byWidth === undefined) {
      byWidth = new Map();
      this.#hf.set(blocks, byWidth);
    }
    let h = byWidth.get(width);
    if (h === undefined) {
      h = stackHeight(stackBlocks(this.ctx, blocks, width));
      byWidth.set(width, h);
    }
    return h;
  }

  #newPage(kind: 'soft' | 'hard', blank = false): void {
    if (this.#page !== undefined && this.#band !== undefined) this.#closeColumn(this.#column());
    if (this.pages.length >= MAX_PAGES) {
      throw new Error(`Layout exceeded ${MAX_PAGES} pages`);
    }
    const sec = this.#section;
    const prev = this.pages[this.pages.length - 1];
    const first = this.#firstOfSection;
    const value =
      first && sec.pageNumberStart !== undefined ? sec.pageNumberStart : (prev?.value ?? 0) + 1;
    const g = geometry(sec);
    const type =
      first && sec.titlePage
        ? 'first'
        : this.ir.evenAndOddHeaders && value % 2 === 0
          ? 'even'
          : 'default';
    const header = blank ? undefined : sec.headers[type];
    const footer = blank ? undefined : sec.footers[type];
    const hfWidth = Math.max(1, g.width - g.left - g.right);
    const hh = this.#hfHeight(header, hfWidth);
    const fh = this.#hfHeight(footer, hfWidth);
    const top = g.top < 0 ? -g.top : Math.max(g.top, g.header + hh);
    const bottomMargin = g.bottom < 0 ? -g.bottom : Math.max(g.bottom, g.footer + fh);
    let bodyBottom = g.height - bottomMargin;
    if (bodyBottom - top < 12) bodyBottom = top + 12;
    const page: PageState = {
      index: this.pages.length,
      section: sec,
      sectionIndex: this.#sectionIndex,
      value,
      number: pageNumberText(value, sec.pageNumberFormat),
      width: g.width,
      height: g.height,
      marginLeft: g.left,
      marginRight: g.right,
      marginTop: Math.abs(g.top),
      marginBottom: Math.abs(g.bottom),
      bodyTop: top,
      bodyBottom,
      bands: [],
      floats: [],
      firstOfSection: first,
    };
    if (header !== undefined) page.header = header;
    if (footer !== undefined) page.footer = footer;
    const band: Band = {
      top,
      bottom: bodyBottom,
      columns: columnsOf(sec, g, top, bodyBottom),
      section: sec,
      separator: sec.columns.separator,
    };
    band.columns[0]!.topHard = kind === 'hard';
    band.columns[0]!.carryIn = this.#carry;
    this.#carry = [];
    page.bands.push(band);
    this.pages.push(page);
    this.#page = page;
    this.#band = band;
    this.#col = 0;
    this.#firstOfSection = false;
    this.#pendingBreak = undefined;
  }

  /** Ends the current band (balancing its columns) and starts one for `sec` below it. */
  #continuousBand(sec: IRSection): void {
    const page = this.#page!;
    const old = this.#band!;
    let bottom = old.columns.length > 1 ? balance(old) : bandBottom(old);
    for (const c of old.columns) {
      if (old.columns.length > 1) {
        c.notes = [];
        for (const p of c.placed) {
          for (const id of p.strip.footnotes ?? []) if (!c.notes.includes(id)) c.notes.push(id);
        }
      }
      this.#closeColumn(c);
    }
    // notes of the closed band stay reserved at the page bottom
    let notes = 0;
    for (const c of old.columns)
      notes = Math.max(
        notes,
        stackHeight(c.noteStrips ?? []) + (c.noteStrips?.length ? this.sepH : 0),
      );
    const limit = page.bodyBottom - notes;
    if (bottom >= limit - 1) {
      this.#newPage('hard');
      return;
    }
    bottom = Math.max(bottom, old.top);
    const g = geometry(sec);
    const band: Band = {
      top: bottom,
      bottom: limit,
      columns: columnsOf(sec, { ...g, width: page.width }, bottom, limit),
      section: sec,
      separator: sec.columns.separator,
    };
    page.bands.push(band);
    this.#band = band;
    this.#col = 0;
  }

  #startSection(sec: IRSection, i: number): void {
    const prevSection = this.#section;
    this.#section = sec;
    this.#sectionIndex = i;
    if (i === 0 || this.#page === undefined) {
      this.#firstOfSection = true;
      this.#drain();
      this.#newPage('hard');
      return;
    }
    this.#drain();
    const page = this.#page;
    const sameSize =
      Math.abs(geometry(sec).width - page.width) < 0.5 &&
      Math.abs(geometry(sec).height - page.height) < 0.5;
    let type = sec.type;
    if ((type === 'continuous' || type === 'nextColumn') && !sameSize) type = 'nextPage';
    if (type === 'nextColumn') {
      const band = this.#band!;
      if (
        band.columns.length === Math.max(1, sec.columns.count) &&
        this.#col + 1 < band.columns.length
      ) {
        this.#newColumn('hard');
        return;
      }
      type = 'continuous';
    }
    if (type === 'continuous') {
      this.#continuousBand(sec);
      return;
    }
    this.#firstOfSection = true;
    if (type === 'evenPage' || type === 'oddPage') {
      const next =
        sec.pageNumberStart !== undefined
          ? sec.pageNumberStart
          : this.pages[this.pages.length - 1]!.value + 1;
      const wantEven = type === 'evenPage';
      if ((next % 2 === 0) !== wantEven) {
        // a blank page of the previous section first
        this.#section = prevSection;
        this.#firstOfSection = false;
        this.#newPage('hard', true);
        this.#section = sec;
        this.#firstOfSection = true;
      }
    }
    this.#newPage('hard');
  }
}

/** Bottom of a single-column band's content, its last space after included. */
function bandBottom(b: Band): number {
  let y = b.top;
  for (const c of b.columns) {
    const last = c.placed[c.placed.length - 1];
    if (last !== undefined) y = Math.max(y, last.y + last.strip.height + last.strip.spaceAfter);
  }
  return y;
}

/**
 * Distributes `strips` over `n` columns of height `h` (same break rules as the flow); returns
 * the strips of each column or undefined when they do not fit.
 */
function distribute(
  strips: readonly Strip[],
  n: number,
  h: number,
  firstHard: boolean,
): Strip[][] | undefined {
  const cols: { strips: Strip[]; ys: number[] }[] = [{ strips: [], ys: [] }];
  const q = [...strips];
  while (q.length > 0) {
    const s = q[0]!;
    const col = cols[cols.length - 1]!;
    const k0 = col.strips.length;
    const lastS = col.strips[k0 - 1];
    const y =
      lastS === undefined
        ? cols.length === 1 && firstHard
          ? s.spaceBefore
          : 0
        : col.ys[k0 - 1]! + lastS.height + gap(lastS, s);
    if (y + s.height <= h + EPS || k0 === 0) {
      q.shift();
      col.strips.push(s);
      col.ys.push(y);
      continue;
    }
    let k = k0;
    if (s.breakBefore === 'avoid') {
      k = k0;
      for (let j = k0 - 1; j >= 1; j--) {
        if (col.strips[j]!.breakBefore === 'auto') {
          k = j;
          break;
        }
      }
    }
    const moved = col.strips.splice(k);
    col.ys.splice(k);
    q.unshift(...moved);
    if (cols.length === n) return undefined;
    cols.push({ strips: [], ys: [] });
  }
  return cols.map((c) => c.strips);
}

/** Balances a multi-column band's columns; returns the band's new content bottom. */
function balance(b: Band): number {
  const strips: Strip[] = [];
  for (const c of b.columns) for (const p of c.placed) if (p.repeat !== true) strips.push(p.strip);
  if (strips.length === 0) return b.top;
  const n = b.columns.length;
  const firstHard = b.columns[0]!.topHard;
  const full = b.bottom - b.top;
  let total = 0;
  let tallest = 0;
  for (const s of strips) {
    total += s.spaceBefore + s.height + s.spaceAfter;
    tallest = Math.max(tallest, s.height);
  }
  if (distribute(strips, n, full, firstHard) === undefined)
    return bandBottom(b) === b.top ? b.top : maxBottom(b);
  let lo = Math.max(tallest, total / n) - 1;
  let hi = full;
  for (let i = 0; i < 20 && hi - lo > 0.25; i++) {
    const mid = (lo + hi) / 2;
    if (distribute(strips, n, mid, firstHard) !== undefined) hi = mid;
    else lo = mid;
  }
  const cols = distribute(strips, n, hi, firstHard)!;
  b.columns.forEach((c, i) => {
    const list = cols[i] ?? [];
    c.placed = [];
    let y = c.top;
    let prev: Strip | undefined;
    list.forEach((s, j) => {
      const yy = j === 0 ? y + (i === 0 && firstHard ? s.spaceBefore : 0) : y + gap(prev, s);
      c.placed.push({ strip: s, y: yy });
      y = yy + s.height;
      prev = s;
    });
  });
  return maxBottom(b);
}

function maxBottom(b: Band): number {
  let y = b.top;
  for (const c of b.columns) {
    const last = c.placed[c.placed.length - 1];
    if (last !== undefined) y = Math.max(y, last.y + last.strip.height + last.strip.spaceAfter);
  }
  return y;
}
