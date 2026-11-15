// CONTRACT (layout agents): fonts, shaping, the layout IR and the PageModel.
//
// Pipeline: WordDocument → buildIR (resolve) → collect font requests → load fonts (async) →
// measure/break lines → tables → paginate (fixed point for PAGEREF/TOC/NUMPAGES) → DocumentLayout
// (the PageModel, plain JSON-able data in points, origin top-left, y grows downward) → renderers.
// Layout never reads XML after buildIR; renderers never read the IR or the document.

import type { WarningHandler } from '../core/warnings.js';
import type {
  ResolvedParagraphProps,
  ResolvedRunProps,
  ResolvedTableProps,
} from '../resolve/types.js';
import type { SectionType } from '../wml/section.js';
import type { XElement } from '../xml/dom.js';

// ---- fonts ------------------------------------------------------------------------------------

export interface FontMetrics {
  unitsPerEm: number;
  /** hhea ascender (or OS/2 typo/win per USE_TYPO_METRICS and Word's rules). */
  ascender: number;
  /** Negative below the baseline. */
  descender: number;
  lineGap: number;
  /** OS/2 usWinAscent / usWinDescent (Word uses these for line height). */
  winAscent: number;
  winDescent: number;
  capHeight: number;
  xHeight: number;
  underlinePosition: number;
  underlineThickness: number;
  strikeoutPosition: number;
  strikeoutSize: number;
  italicAngle: number;
}

/** A parsed sfnt (TrueType/OpenType) face. */
export interface Font {
  readonly family: string;
  readonly subfamily: string;
  readonly postscriptName: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly metrics: FontMetrics;
  /**
   * TrueType outlines (glyf; PDF subsets them), CFF (embedded whole), or one of the PDF base-14
   * fonts (metrics only, no bytes, never embedded) used as the last-resort fallback because core
   * ships no font files.
   */
  readonly outlines: 'truetype' | 'cff' | 'standard14';
  /** Glyph id for a code point (0 = .notdef). */
  glyphId(codePoint: number): number;
  /** Advance width in font units. */
  advance(glyphId: number): number;
  /** Pair kerning adjustment in font units (kern table or GPOS pair adjustment). */
  kerning(left: number, right: number): number;
  /** The face's sfnt bytes (a single face even when loaded from a collection); empty for standard14. */
  readonly bytes: Uint8Array;
  readonly numGlyphs: number;
}

export interface FontRequest {
  family: string;
  bold: boolean;
  italic: boolean;
  /** Which w:rFonts slot asked (affects fallback choice). */
  script: 'latin' | 'eastAsia' | 'cs' | 'symbol';
}

/**
 * Supplies font bytes or a parsed Font for a request; undefined = not available here (the next
 * resolver, then the metric-compatible substitution table, then the last-resort fallback is tried).
 */
export type FontResolver = (
  req: FontRequest,
) => Font | Uint8Array | undefined | Promise<Font | Uint8Array | undefined>;

// ---- shaping ----------------------------------------------------------------------------------

export interface ShapedGlyph {
  id: number;
  /** Advance in font units. */
  advance: number;
  /** Offsets in font units. */
  dx: number;
  dy: number;
  /** Index into the shaped text of the first character of this glyph's cluster. */
  cluster: number;
}

export interface ShapeOptions {
  rtl: boolean;
  script?: string;
  lang?: string;
  /** Apply pair kerning. */
  kerning: boolean;
}

/** Turns text into positioned glyphs. The built-in shaper does cmap + pair kerning; a HarfBuzz plugin can replace it. */
export interface Shaper {
  shape(font: Font, text: string, o: ShapeOptions): ShapedGlyph[];
}

// ---- layout IR (built from the document by buildIR) --------------------------------------------

export interface IRDocument {
  sections: IRSection[];
  /** settings evenAndOddHeaders. */
  evenAndOddHeaders: boolean;
  /** settings defaultTabStop, points. */
  defaultTabStop: number;
  footnotes: Map<string, IRBlock[]>;
  endnotes: Map<string, IRBlock[]>;
  images: Map<string, ImageResource>;
  metadata: LayoutMetadata;
}

export interface IRSection {
  index: number;
  type: SectionType;
  /** Points. */
  pageWidth: number;
  pageHeight: number;
  margins: {
    top: number;
    right: number;
    bottom: number;
    left: number;
    header: number;
    footer: number;
    gutter: number;
  };
  columns: {
    count: number;
    space: number;
    separator: boolean;
    widths?: { width: number; space: number }[];
  };
  titlePage: boolean;
  pageNumberStart?: number;
  /** pgNumType/@w:fmt (decimal, upperRoman, ...). */
  pageNumberFormat: string;
  headers: { default?: IRBlock[]; first?: IRBlock[]; even?: IRBlock[] };
  footers: { default?: IRBlock[]; first?: IRBlock[]; even?: IRBlock[] };
  blocks: IRBlock[];
  /** Vertical alignment of the page (w:vAlign). */
  vAlign: 'top' | 'center' | 'both' | 'bottom';
}

export type IRBlock = IRParagraph | IRTable;

export interface IRParagraph {
  kind: 'paragraph';
  /** The w:p this came from (for PageInfoProvider.pageOf and PAGEREF targets). */
  source: XElement;
  props: ResolvedParagraphProps;
  /** List label, already rendered. */
  label?: {
    text: string;
    props: ResolvedRunProps;
    suffix: 'tab' | 'space' | 'nothing';
    alignment: 'left' | 'center' | 'right';
  };
  inlines: IRInline[];
  /** Floating objects anchored in this paragraph. */
  anchors: IRAnchor[];
  /** Bookmarks starting in this paragraph (names), for PAGEREF and PDF destinations. */
  bookmarks: string[];
}

export type IRInline =
  | { kind: 'text'; text: string; props: ResolvedRunProps; source?: XElement; link?: IRLink }
  | { kind: 'tab'; props: ResolvedRunProps }
  | { kind: 'break'; type: 'line' | 'page' | 'column'; clear?: 'left' | 'right' | 'all' }
  | {
      kind: 'image';
      image: string;
      width: number;
      height: number;
      props: ResolvedRunProps;
      link?: IRLink;
    }
  | { kind: 'drawing'; source: XElement; width: number; height: number; props: ResolvedRunProps }
  | { kind: 'math'; source: XElement; props: ResolvedRunProps }
  /** Page-dependent field, evaluated per page by the paginator (PAGE, NUMPAGES, SECTIONPAGES, PAGEREF). */
  | {
      kind: 'field';
      type: string;
      arg?: string;
      format?: string;
      fallback: string;
      props: ResolvedRunProps;
    }
  | {
      kind: 'noteRef';
      note: 'footnote' | 'endnote';
      id: string;
      props: ResolvedRunProps;
      text: string;
    };

export interface IRLink {
  url?: string;
  /** Internal bookmark target. */
  anchor?: string;
}

export interface IRAnchor {
  kind: 'image' | 'textbox' | 'drawing';
  source: XElement;
  /** Points. */
  width: number;
  height: number;
  image?: string;
  /** Text box content. */
  blocks?: IRBlock[];
  horizontal: { relativeTo: string; offset?: number; align?: string };
  vertical: { relativeTo: string; offset?: number; align?: string };
  wrap: 'square' | 'tight' | 'through' | 'topAndBottom' | 'none';
  behindText: boolean;
  /** Distance from text, points. */
  distance: { top: number; right: number; bottom: number; left: number };
  zIndex: number;
}

export interface IRTable {
  kind: 'table';
  source: XElement;
  props: ResolvedTableProps;
  /** tblGrid widths, points. */
  grid: number[];
  rows: IRRow[];
}

export interface IRRow {
  height?: { value: number; rule: 'atLeast' | 'exact' };
  cantSplit: boolean;
  /** Repeat as header row on each page. */
  header: boolean;
  gridBefore: number;
  gridAfter: number;
  cells: IRCell[];
}

export interface IRCell {
  gridSpan: number;
  vMerge?: 'restart' | 'continue';
  blocks: IRBlock[];
  /** Points. */
  margins: { top: number; right: number; bottom: number; left: number };
  borders: import('../resolve/types.js').Borders;
  shading?: string;
  vAlign: 'top' | 'center' | 'bottom';
}

export interface ImageResource {
  id: string;
  /** image/png, image/jpeg, image/gif, image/bmp, image/tiff, image/x-emf, image/x-wmf, image/svg+xml. */
  contentType: string;
  bytes: Uint8Array;
  /** Pixel size when known. */
  pixelWidth?: number;
  pixelHeight?: number;
}

export interface LayoutMetadata {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  creator?: string;
  /** Document language (BCP 47) for PDF /Lang. */
  lang?: string;
}

// ---- PageModel --------------------------------------------------------------------------------

export interface DocumentLayout {
  pages: PageLayout[];
  /** Fonts referenced by TextItem.font. */
  fonts: Font[];
  images: Map<string, ImageResource>;
  outline: OutlineEntry[];
  /** Named destinations (bookmarks) → position. */
  destinations: Map<string, { page: number; y: number }>;
  metadata: LayoutMetadata;
  /** Where each source paragraph ended up (PageInfoProvider support). */
  paragraphPages: Map<XElement, number>;
}

export interface PageLayout {
  /** 0-based physical index. */
  index: number;
  /** Displayed page number text (honours restarts and formats). */
  number: string;
  section: number;
  width: number;
  height: number;
  items: DisplayItem[];
}

export type DisplayItem = TextItem | ImageItem | RectItem | LineItem | PathItem | LinkItem;

export interface TextItem {
  kind: 'text';
  /** Baseline origin of the first glyph, points. */
  x: number;
  y: number;
  /** Index into DocumentLayout.fonts. */
  font: number;
  /** Points. */
  size: number;
  /** RRGGBB. */
  color: string;
  glyphs: number[];
  /** Advance of each glyph in points (kerning, spacing and justification included). */
  advances: number[];
  /** Per-glyph vertical offsets in points (superscript handled via y). Optional. */
  dy?: number[];
  /** The characters (for text extraction / ToUnicode). */
  text: string;
  /** For each glyph, index into `text` of its cluster start. */
  clusters: number[];
  /** Horizontal scale percent (default 100). */
  scale?: number;
  /** Synthesize bold/italic when the face lacks it. */
  fakeBold?: boolean;
  fakeItalic?: boolean;
}

export interface ImageItem {
  kind: 'image';
  x: number;
  y: number;
  width: number;
  height: number;
  image: string;
  /** Crop in fractions (a:srcRect / 100000). */
  crop?: { left: number; top: number; right: number; bottom: number };
  rotation?: number;
}

export interface RectItem {
  kind: 'rect';
  x: number;
  y: number;
  width: number;
  height: number;
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
}

export interface LineItem {
  kind: 'line';
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  width: number;
  dash?: number[];
}

export interface PathItem {
  kind: 'path';
  /** SVG path data in points. */
  d: string;
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  /** Even-odd fill rule. */
  evenOdd?: boolean;
}

export interface LinkItem {
  kind: 'link';
  x: number;
  y: number;
  width: number;
  height: number;
  url?: string;
  /** Internal destination name (bookmark). */
  dest?: string;
}

export interface OutlineEntry {
  title: string;
  /** 0 = top level. */
  level: number;
  page: number;
  y: number;
}

// ---- options ----------------------------------------------------------------------------------

export interface LayoutOptions {
  /** Font sources, tried in order. */
  fonts?: FontResolver | readonly FontResolver[];
  /** Extra substitutions (family → family) tried before the built-in metric-compatible table. */
  substitutions?: Readonly<Record<string, string>>;
  /** Use fonts embedded in the docx (default true). */
  useEmbeddedFonts?: boolean;
  shaper?: Shaper;
  /** Max layout passes for page-number fixed points (default 3). */
  maxPasses?: number;
  onWarning?: WarningHandler;
}
