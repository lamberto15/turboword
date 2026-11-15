// The layout pipeline: document → IR → fonts → pages, with the fixed point for page fields.
//
//   buildIR(doc) → want() every run's faces → await fonts.load() → LayoutEnv → Flow (pages)
//   → body page fields (PAGE in text, NUMPAGES, SECTIONPAGES, PAGEREF) evaluated where they
//   landed → if any shows a different text than it was measured with, lay out again with the
//   new texts (at most `maxPasses` passes) → emit the PageModel.
//
// Headers and footers never feed the fixed point: they are laid out per page at emission with
// that page's values (their height is taken from the first measurement).

import { warning } from '../core/warnings.js';
import type { Part } from '../opc/part.js';
import type { WordDocument } from '../wml/document.js';
import type { XDocument, XElement } from '../xml/dom.js';
import type { FontManager, LayoutEnv } from './engine.js';
import { createFontManager } from './fonts/index.js';
import { buildIR } from './ir.js';
import { bodyFieldValues, emit, pageFacts } from './page/emit.js';
import { Flow } from './page/flow.js';
import { allParagraphs, type LayoutCtx, type ParagraphLayouter } from './page/strip.js';
import { layoutParagraph } from './paragraph.js';
import { basicShaper } from './text/index.js';
import type { DocumentLayout, IRBlock, IRDocument, IRSection, LayoutOptions } from './types.js';

/** Internal knobs (tests inject a paragraph layouter; the public API never sets them). */
export interface LayoutInternals {
  maxPasses?: number;
  layoutParagraph?: ParagraphLayouter;
}

/** Registers every face the IR's runs may use. */
export function wantFonts(ir: IRDocument, fonts: FontManager): void {
  const roots: IRBlock[][] = [];
  for (const s of ir.sections) {
    roots.push(s.blocks);
    for (const hf of [s.headers, s.footers]) {
      if (hf.default) roots.push(hf.default);
      if (hf.first) roots.push(hf.first);
      if (hf.even) roots.push(hf.even);
    }
  }
  for (const n of ir.footnotes.values()) roots.push(n);
  for (const n of ir.endnotes.values()) roots.push(n);
  for (const blocks of roots) {
    for (const p of allParagraphs(blocks)) {
      fonts.want(p.props.markRun);
      if (p.label) fonts.want(p.label.props);
      for (const i of p.inlines) if ('props' in i) fonts.want(i.props);
    }
  }
}

/** Maps an element to the part whose XML it belongs to. */
function partResolver(doc: WordDocument): (el: XElement) => Part | undefined {
  let map: Map<XDocument, Part> | undefined;
  return (el) => {
    const owner = el.ownerDocument();
    if (owner === undefined) return undefined;
    if (map === undefined) {
      map = new Map();
      for (const part of doc.package.parts()) {
        if (part.contentType?.includes('wordprocessingml') !== true) continue;
        try {
          map.set(part.xml(), part);
        } catch {
          // a broken part simply resolves nothing
        }
      }
    }
    return map.get(owner);
  };
}

/** A US Letter section without content (documents always lay out to at least one page). */
function emptySection(): IRSection {
  return {
    index: 0,
    type: 'nextPage',
    pageWidth: 612,
    pageHeight: 792,
    margins: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36, gutter: 0 },
    columns: { count: 1, space: 36, separator: false },
    titlePage: false,
    pageNumberFormat: 'decimal',
    headers: {},
    footers: {},
    blocks: [],
    vAlign: 'top',
  };
}

/** Lays out an already built IR (the pipeline's synchronous half). */
export function layoutIR(ir: IRDocument, env: LayoutEnv, o: LayoutInternals = {}): DocumentLayout {
  if (ir.sections.length === 0) ir.sections.push(emptySection());
  const ctx: LayoutCtx = {
    env,
    layoutParagraph: o.layoutParagraph ?? layoutParagraph,
    cache: new WeakMap(),
    depth: 0,
  };
  const passes = Math.max(1, Math.floor(o.maxPasses ?? 3));
  let flow = new Flow(ctx, ir);
  flow.run();
  for (let pass = 1; pass < passes; pass++) {
    let changed = false;
    for (const [f, v] of bodyFieldValues(flow.pages, pageFacts(flow.pages))) {
      if (f.fallback !== v) {
        f.fallback = v;
        changed = true;
      }
    }
    if (!changed) break;
    flow = new Flow(ctx, ir);
    flow.run();
  }
  const out = emit(ctx, flow);
  return {
    pages: out.pages,
    fonts: env.fonts.fonts(),
    images: ir.images,
    outline: out.outline,
    destinations: out.destinations,
    metadata: ir.metadata,
    paragraphPages: out.paragraphPages,
  };
}

/** The whole pipeline (see the file header). */
export async function paginateDocument(
  doc: WordDocument,
  o: LayoutOptions & LayoutInternals = {},
): Promise<DocumentLayout> {
  const seen = new Set<string>();
  const warn = (code: string, message: string): void => {
    const key = `${code}\u0001${message}`;
    if (seen.has(key)) return;
    seen.add(key);
    o.onWarning?.(warning(code, message));
  };
  const ir = buildIR(doc, { warn });
  const fonts = createFontManager(doc, o);
  wantFonts(ir, fonts);
  await fonts.load();
  const env: LayoutEnv = {
    doc,
    partOf: partResolver(doc),
    fonts,
    shaper: o.shaper ?? basicShaper,
    defaultTabStop: ir.defaultTabStop,
    images: ir.images,
    warn,
  };
  const internals: LayoutInternals = {};
  if (o.maxPasses !== undefined) internals.maxPasses = o.maxPasses;
  if (o.layoutParagraph !== undefined) internals.layoutParagraph = o.layoutParagraph;
  return layoutIR(ir, env, internals);
}
