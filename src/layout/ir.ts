// The layout IR: everything layout needs from the document, read once, in points.
//
// buildIR walks the document with the typed views (sections, stories, tables) and the resolver
// (resolved paragraph/run/table/cell formatting, list labels) and never touches the DOM again
// afterwards: layout works on the IR alone. It only reads; the document is never modified.
//
// Block containers (a story, a table cell, a text box) are processed from a work queue rather than
// by recursion: each container gets a pre-allocated `IRBlock[]` that the job fills in place, so
// nested tables and text boxes need no call-stack depth. Container nesting is capped
// (MAX_NESTING) so the IR stays bounded for layout code that walks it recursively.
//
// Sections: every body-level paragraph carrying `w:pPr/w:sectPr` ends a section (the same rule
// and the same content-control/custom-XML transparency as `sectionsOf`); headers and footers are
// each section's *effective* ones (inherited "link to previous"), shared parts built once.

import type { Part } from '../opc/part.js';
import { resolveCell, resolveTable } from '../resolve/index.js';
import type { WordDocument } from '../wml/document.js';
import { type HeaderFooterType, type Section, sectionsOf } from '../wml/section.js';
import type { Story } from '../wml/story.js';
import { Table, wrappedChildren } from '../wml/table.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { FontManager, LayoutEnv } from './engine.js';
import { wantGraphicsFonts } from './graphics/index.js';
import type { LineCompat } from './line/extras.js';
import type { DrawingContext } from './line/ir-drawing.js';
import { FieldState } from './line/ir-fields.js';
import { buildParagraph } from './line/ir-inlines.js';
import { computeNoteMarks, type NoteMarks } from './line/ir-notes.js';
import {
  imagePixelSize,
  intAttr,
  onOff,
  onOffValue,
  settingsRoot,
  stylesRoot,
  twipsToPt,
  wVal,
} from './line/ir-util.js';
import type {
  ImageResource,
  IRBlock,
  IRCell,
  IRDocument,
  IRRow,
  IRSection,
  IRTable,
  LayoutMetadata,
} from './types.js';

export type {
  IRAnchorExtras,
  IRInlineExtras,
  IRParagraphExtras,
  LineCompat,
  PtabInfo,
} from './line/extras.js';

/** Deepest container nesting (tables in cells, text boxes in text boxes) kept in the IR. */
export const MAX_NESTING = 32;

/** Word's defaults when a section omits its page size / margins (US Letter, 1" margins). */
const DEFAULT_PAGE = { width: 12240, height: 15840 };
const DEFAULT_MARGIN = {
  top: 1440,
  right: 1440,
  bottom: 1440,
  left: 1440,
  header: 720,
  footer: 720,
  gutter: 0,
};

const BLOCKS: ReadonlySet<string> = /* @__PURE__ */ new Set(['p', 'tbl', 'bookmarkStart']);
const HF_TYPES: readonly HeaderFooterType[] = ['default', 'first', 'even'];

interface Job {
  container: XElement;
  out: IRBlock[];
  story: Story;
  part: Part;
  depth: number;
}

interface Builder {
  doc: WordDocument;
  warn: (code: string, message: string) => void;
  images: Map<string, ImageResource>;
  notes: NoteMarks;
  compat: LineCompat;
  jobs: Job[];
}

/** Registers the image `relId` of `part` points to; returns its IR id (the media part name). */
function registerImage(b: Builder, part: Part, relId: string): string | undefined {
  const rel = part.rels.get(relId);
  if (rel === undefined || rel.external || rel.partName === undefined) return undefined;
  const target = b.doc.package.part(rel.partName);
  if (target === undefined) {
    b.warn('LAYOUT_IMAGE_MISSING', `Image ${rel.partName} referenced by ${part.name} is missing`);
    return undefined;
  }
  const id = target.name;
  if (!b.images.has(id)) {
    const bytes = target.bytes();
    const res: ImageResource = {
      id,
      contentType: target.contentType ?? 'application/octet-stream',
      bytes,
    };
    const px = imagePixelSize(bytes);
    if (px !== undefined) {
      res.pixelWidth = px.width;
      res.pixelHeight = px.height;
    }
    b.images.set(id, res);
  }
  return id;
}

function drawingContext(b: Builder, job: Job): DrawingContext {
  return {
    image: (relId) => registerImage(b, job.part, relId),
    textbox: (content) => {
      const out: IRBlock[] = [];
      if (job.depth + 1 > MAX_NESTING) {
        b.warn('LAYOUT_NESTING', `Text box nested deeper than ${MAX_NESTING} levels dropped`);
        return out;
      }
      b.jobs.push({
        container: content,
        out,
        story: job.story,
        part: job.part,
        depth: job.depth + 1,
      });
      return out;
    },
  };
}

function buildTable(b: Builder, el: XElement, job: Job): IRTable {
  const table = new Table(el, job.story);
  const grid = (el.element(NS.w, 'tblGrid')?.elements(NS.w, 'gridCol') ?? []).map((c) =>
    twipsToPt(Math.max(0, intAttr(c, NS.w, 'w', 0))),
  );
  const rows: IRRow[] = [];
  const deep = job.depth + 1 > MAX_NESTING;
  if (deep)
    b.warn('LAYOUT_NESTING', `Table content nested deeper than ${MAX_NESTING} levels dropped`);
  for (const row of table.rows()) {
    const trPr = row.element.element(NS.w, 'trPr');
    const ir: IRRow = {
      cantSplit: onOff(trPr?.element(NS.w, 'cantSplit')),
      header: onOff(trPr?.element(NS.w, 'tblHeader')),
      gridBefore: row.gridBefore,
      gridAfter: row.gridAfter,
      cells: [],
    };
    const trHeight = trPr?.element(NS.w, 'trHeight');
    if (trHeight !== undefined) {
      const rule = trHeight.attrNS(NS.w, 'hRule');
      const value = twipsToPt(Math.max(0, wVal(trHeight, 0)));
      if (rule !== 'auto' && value > 0)
        ir.height = { value, rule: rule === 'exact' ? 'exact' : 'atLeast' };
    }
    for (const cell of row.cells()) {
      // Legacy horizontal merge: a continuation cell widens the previous cell instead.
      const prev = ir.cells[ir.cells.length - 1];
      if (cell.hMerge === 'continue' && prev !== undefined) {
        prev.gridSpan += cell.gridSpan;
        continue;
      }
      const cp = resolveCell(cell);
      const out: IRCell = {
        gridSpan: cell.gridSpan,
        blocks: [],
        margins: {
          top: twipsToPt(cp.margins.top),
          right: twipsToPt(cp.margins.right),
          bottom: twipsToPt(cp.margins.bottom),
          left: twipsToPt(cp.margins.left),
        },
        borders: cp.borders,
        vAlign: cp.vAlign,
      };
      const vm = cell.vMerge;
      if (vm !== undefined) out.vMerge = vm;
      if (cp.shading !== undefined) out.shading = cp.shading;
      if (!deep) {
        b.jobs.push({
          container: cell.element,
          out: out.blocks,
          story: job.story,
          part: job.part,
          depth: job.depth + 1,
        });
      }
      ir.cells.push(out);
    }
    rows.push(ir);
  }
  return { kind: 'table', source: el, props: resolveTable(table), grid, rows };
}

/** Processes one container: its paragraphs and tables, in order, into `job.out`. */
function runJob(b: Builder, job: Job): void {
  const fields = new FieldState();
  const pending: string[] = [];
  const ctx = {
    story: job.story,
    part: job.part,
    fields,
    notes: b.notes,
    drawing: drawingContext(b, job),
    compat: b.compat,
    warn: b.warn,
  };
  for (const el of wrappedChildren(job.container, BLOCKS)) {
    if (el.local === 'p') job.out.push(buildParagraph(el, ctx, pending));
    else if (el.local === 'tbl') job.out.push(buildTable(b, el, job));
    else {
      const name = el.attrNS(NS.w, 'name');
      if (name !== undefined && name !== '') pending.push(name);
    }
  }
}

function drain(b: Builder): void {
  // FIFO; every job writes into its own pre-allocated array, so the order does not matter.
  for (let i = 0; i < b.jobs.length; i++) runJob(b, b.jobs[i]!);
  b.jobs.length = 0;
}

function sectionShell(s: Section): IRSection {
  const el = s.element;
  const size = s.pageSize;
  const m = s.margins;
  const cols = el?.element(NS.w, 'cols');
  const count = s.columns;
  const columns: IRSection['columns'] = {
    count,
    space: twipsToPt(intAttr(cols, NS.w, 'space', 720)),
    separator: onOffValue(cols?.attrNS(NS.w, 'sep'), false),
  };
  const explicit = cols?.elements(NS.w, 'col') ?? [];
  const equal = cols?.attrNS(NS.w, 'equalWidth');
  if (explicit.length > 0 && !onOffValue(equal, true)) {
    columns.widths = explicit.map((c) => ({
      width: twipsToPt(intAttr(c, NS.w, 'w', 0)),
      space: twipsToPt(intAttr(c, NS.w, 'space', 0)),
    }));
  }
  const pgNum = el?.element(NS.w, 'pgNumType');
  const vAlign = el?.element(NS.w, 'vAlign')?.attrNS(NS.w, 'val');
  const out: IRSection = {
    index: s.index,
    type: s.type,
    pageWidth: twipsToPt(size?.width ?? DEFAULT_PAGE.width),
    pageHeight: twipsToPt(size?.height ?? DEFAULT_PAGE.height),
    margins: {
      top: twipsToPt(m?.top ?? DEFAULT_MARGIN.top),
      right: twipsToPt(m?.right ?? DEFAULT_MARGIN.right),
      bottom: twipsToPt(m?.bottom ?? DEFAULT_MARGIN.bottom),
      left: twipsToPt(m?.left ?? DEFAULT_MARGIN.left),
      header: twipsToPt(m?.header ?? DEFAULT_MARGIN.header),
      footer: twipsToPt(m?.footer ?? DEFAULT_MARGIN.footer),
      gutter: twipsToPt(m?.gutter ?? DEFAULT_MARGIN.gutter),
    },
    columns,
    titlePage: s.titlePage,
    pageNumberFormat: pgNum?.attrNS(NS.w, 'fmt') ?? 'decimal',
    headers: {},
    footers: {},
    blocks: [],
    vAlign: vAlign === 'center' || vAlign === 'both' || vAlign === 'bottom' ? vAlign : 'top',
  };
  const start = s.pageNumberStart;
  if (start !== undefined) out.pageNumberStart = start;
  return out;
}

function metadataOf(doc: WordDocument): LayoutMetadata {
  const out: LayoutMetadata = {};
  const p = doc.properties;
  const set = (k: keyof LayoutMetadata, v: string | undefined): void => {
    if (v !== undefined && v.trim() !== '') out[k] = v;
  };
  set('title', p.title);
  set('author', p.creator);
  set('subject', p.subject);
  set('keywords', p.keywords);
  set('creator', p.application);
  const lang =
    p.language ??
    stylesRoot(doc)
      ?.element(NS.w, 'docDefaults')
      ?.element(NS.w, 'rPrDefault')
      ?.element(NS.w, 'rPr')
      ?.element(NS.w, 'lang')
      ?.attrNS(NS.w, 'val');
  set('lang', lang);
  return out;
}

/** Builds the layout IR from the document (resolved formatting, list labels, images, notes). */
export function buildIR(
  doc: WordDocument,
  o?: { warn?: (code: string, message: string) => void },
): IRDocument {
  const settings = settingsRoot(doc);
  const compat: LineCompat = {};
  if (onOff(settings?.element(NS.w, 'compat')?.element(NS.w, 'doNotExpandShiftReturn'))) {
    compat.doNotExpandShiftReturn = true;
  }
  const decimal = settings?.element(NS.w, 'decimalSymbol')?.attrNS(NS.w, 'val');
  if (decimal !== undefined && decimal !== '') compat.decimalSymbol = decimal;
  for (const cs of settings?.element(NS.w, 'compat')?.elements(NS.w, 'compatSetting') ?? []) {
    if (cs.attrNS(NS.w, 'name') !== 'compatibilityMode') continue;
    const mode = Number.parseInt(cs.attrNS(NS.w, 'val') ?? '', 10);
    if (Number.isFinite(mode)) compat.compatibilityMode = mode;
  }

  const b: Builder = {
    doc,
    warn: o?.warn ?? (() => {}),
    images: new Map(),
    notes: computeNoteMarks(doc),
    compat,
    jobs: [],
  };

  // ---- sections and body blocks ----------------------------------------------------------
  const sections = sectionsOf(doc);
  const irSections = sections.map(sectionShell);
  const body = doc.body;
  const bodyJob: Job = {
    container: body.element,
    out: [],
    story: body,
    part: doc.mainPart,
    depth: 0,
  };
  {
    const fields = new FieldState();
    const pending: string[] = [];
    const ctx = {
      story: body,
      part: doc.mainPart,
      fields,
      notes: b.notes,
      drawing: drawingContext(b, bodyJob),
      compat,
      warn: b.warn,
    };
    let si = 0;
    for (const el of wrappedChildren(body.element, BLOCKS)) {
      const target = irSections[Math.min(si, irSections.length - 1)]!.blocks;
      if (el.local === 'p') {
        target.push(buildParagraph(el, ctx, pending));
        if (el.element(NS.w, 'pPr')?.element(NS.w, 'sectPr') !== undefined) si++;
      } else if (el.local === 'tbl') {
        target.push(buildTable(b, el, bodyJob));
      } else {
        const name = el.attrNS(NS.w, 'name');
        if (name !== undefined && name !== '') pending.push(name);
      }
    }
  }

  // ---- headers and footers (each part built once) ----------------------------------------
  const built = new Map<Part, IRBlock[]>();
  const hfBlocks = (story: Story): IRBlock[] => {
    let out = built.get(story.part);
    if (out === undefined) {
      out = [];
      built.set(story.part, out);
      b.jobs.push({ container: story.element, out, story, part: story.part, depth: 0 });
    }
    return out;
  };
  for (const s of sections) {
    const ir = irSections[s.index]!;
    for (const t of HF_TYPES) {
      const h = s.header(t);
      if (h !== undefined) ir.headers[t] = hfBlocks(h.story);
      const f = s.footer(t);
      if (f !== undefined) ir.footers[t] = hfBlocks(f.story);
    }
  }

  // ---- notes -------------------------------------------------------------------------------
  const footnotes = new Map<string, IRBlock[]>();
  const endnotes = new Map<string, IRBlock[]>();
  for (const [stories, map] of [
    [doc.footnotes(), footnotes],
    [doc.endnotes(), endnotes],
  ] as const) {
    for (const story of stories) {
      if (story.id === undefined || map.has(story.id)) continue;
      const out: IRBlock[] = [];
      map.set(story.id, out);
      b.jobs.push({ container: story.element, out, story, part: story.part, depth: 0 });
    }
  }

  drain(b);

  return {
    sections: irSections,
    evenAndOddHeaders: onOff(settings?.element(NS.w, 'evenAndOddHeaders')),
    defaultTabStop: twipsToPt(wVal(settings?.element(NS.w, 'defaultTabStop'), 720)) || 36,
    footnotes,
    endnotes,
    images: b.images,
    metadata: metadataOf(doc),
  };
}

/**
 * Registers with `fonts` every run formatting the IR uses (text, tabs, fields, note marks, list
 * labels, paragraph marks; headers, footers, notes, table cells and text boxes included), so one
 * `fonts.load()` afterwards makes every face available to layoutParagraph. With `graphics`, the
 * faces of drawings and math zones are requested too (graphics' wantGraphicsFonts).
 */
export function wantFonts(
  ir: IRDocument,
  fonts: FontManager,
  graphics?: Pick<LayoutEnv, 'doc' | 'partOf' | 'fonts'>,
): void {
  const stack: IRBlock[][] = [];
  for (const s of ir.sections) {
    stack.push(s.blocks);
    for (const hf of [s.headers, s.footers]) {
      if (hf.default) stack.push(hf.default);
      if (hf.first) stack.push(hf.first);
      if (hf.even) stack.push(hf.even);
    }
  }
  for (const blocks of ir.footnotes.values()) stack.push(blocks);
  for (const blocks of ir.endnotes.values()) stack.push(blocks);
  const seen = new Set<IRBlock[]>();
  while (stack.length > 0) {
    const blocks = stack.pop()!;
    if (seen.has(blocks)) continue;
    seen.add(blocks);
    for (const block of blocks) {
      if (block.kind === 'table') {
        for (const row of block.rows) for (const cell of row.cells) stack.push(cell.blocks);
        continue;
      }
      fonts.want(block.props.markRun);
      if (block.label) fonts.want(block.label.props);
      for (const inline of block.inlines) {
        if ('props' in inline) fonts.want(inline.props);
        if (graphics !== undefined && (inline.kind === 'drawing' || inline.kind === 'math')) {
          wantGraphicsFonts(inline.source, inline.props, graphics);
        }
      }
      for (const a of block.anchors) {
        if (a.blocks) stack.push(a.blocks);
        else if (graphics !== undefined && a.kind === 'drawing') {
          wantGraphicsFonts(a.source, block.props.runDefaults, graphics);
        }
      }
    }
  }
}
