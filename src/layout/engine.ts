// CONTRACT (layout agents): the internal seams between the layout modules.
//
//   fonts/   (fonts agent)     createFontManager, loadFont, standard-14 fallback, substitutions
//   text/    (shaping agent)   lineBreakOpportunities (UAX #14), bidi (UAX #9), basicShaper, scripts
//   ir.ts, paragraph.ts (layout-A agent)  buildIR, layoutParagraph
//   graphics/ (graphics agent) drawingItems (DrawingML shapes, charts, EMF/WMF), layoutMath (OMML)
//   table.ts, paginate.ts, index.ts (layout-B agent)  tables, pages, headers/footers, notes, floats
//
// Units: points everywhere in layout. Coordinates in LineBox/MathBox/drawing items are relative:
// x from the box's left edge, y from the baseline (LineBox, MathBox) or the top (drawings).

import type { Part } from '../opc/part.js';
import type { ResolvedRunProps } from '../resolve/types.js';
import type { WordDocument } from '../wml/document.js';
import type { XElement } from '../xml/dom.js';
import type { DisplayItem, Font, ImageResource, IRInline, Shaper } from './types.js';

/** Loads and hands out fonts during one layout. */
export interface FontManager {
  /** Registers the faces a run may need (each rFonts slot × its bold/italic). */
  want(props: ResolvedRunProps): void;
  /** Resolves everything wanted: resolvers, embedded fonts, substitutions, standard-14 fallback. */
  load(): Promise<void>;
  /**
   * The face for a character of a run: picks the rFonts slot by Word's rules (ascii/hAnsi by
   * range, eastAsia by w:hint and script, cs when rtl/cs), then falls back per glyph when the face
   * lacks the character. Only valid after load().
   */
  fontFor(props: ResolvedRunProps, codePoint: number): Font;
  /** Every face handed out so far, in first-use order (becomes DocumentLayout.fonts). */
  fonts(): Font[];
  /** Index of `font` in fonts(). */
  indexOf(font: Font): number;
}

/** Shared services for measuring and placing content. */
export interface LayoutEnv {
  doc: WordDocument;
  /** The part an element belongs to (to resolve r:id/r:embed of drawings, charts, images). */
  partOf(el: XElement): Part | undefined;
  fonts: FontManager;
  shaper: Shaper;
  /** settings defaultTabStop, points. */
  defaultTabStop: number;
  images: Map<string, ImageResource>;
  warn(code: string, message: string): void;
}

/** A page-dependent field left in a line for the paginator to fill in. */
export interface FieldSlot {
  /** Index of the placeholder TextItem in LineBox.items. */
  item: number;
  field: Extract<IRInline, { kind: 'field' }>;
}

/** One laid-out line of a paragraph. */
export interface LineBox {
  /** Total line pitch (ascent + descent + line spacing rules applied), points. */
  height: number;
  /** Baseline distance from the top of the line box. */
  baseline: number;
  /** Width actually used by content. */
  width: number;
  /** Items positioned relative to (line left edge, baseline). */
  items: DisplayItem[];
  fields: FieldSlot[];
  /** Footnote/endnote ids referenced in this line (the paginator reserves note space). */
  noteRefs: string[];
  /** Bookmarks starting in this line. */
  bookmarks: string[];
  /** A page or column break ends this line. */
  breakAfter?: 'page' | 'column';
}

/** A paragraph broken into lines at a given width. */
export interface ParagraphLayout {
  source: XElement;
  lines: LineBox[];
  /** Points (after autospacing and contextual spacing are decided by the caller). */
  spaceBefore: number;
  spaceAfter: number;
  keepNext: boolean;
  keepLines: boolean;
  widowControl: boolean;
  pageBreakBefore: boolean;
  contextualSpacing: boolean;
  styleId: string | undefined;
  /** Paragraph borders/shading drawn behind the lines (relative to the paragraph box top-left). */
  decorations: DisplayItem[];
  /** Left offset of the lines from the column edge (indent). */
  indentLeft: number;
}

/** Inline math laid out (OMML). */
export interface MathBox {
  width: number;
  ascent: number;
  descent: number;
  /** Relative to (left edge, baseline). */
  items: DisplayItem[];
}
