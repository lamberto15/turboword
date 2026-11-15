// CONTRACT (resolve agent): resolved formatting. Lengths are twips (1/20 pt) as in the XML;
// font sizes are points. Everything is fully resolved: docDefaults → table style (conditional
// formatting) → paragraph style chain (basedOn) → numbering level → direct paragraph props →
// character style chain → direct run props; theme fonts/colors substituted; toggle properties
// (b, i, caps, ...) applied with Word's XOR semantics.

export type Alignment =
  | 'left'
  | 'center'
  | 'right'
  | 'both'
  | 'distribute'
  | 'mediumKashida'
  | 'highKashida'
  | 'lowKashida'
  | 'thaiDistribute';

export interface BorderSpec {
  /** `w:val` (single, double, dotted, ...); 'none'/'nil' never appear here (absent instead). */
  style: string;
  /** Width in eighths of a point (`w:sz`). */
  size: number;
  /** Space in points (`w:space`). */
  space: number;
  /** RRGGBB or 'auto'. */
  color: string;
}

export interface Borders {
  top?: BorderSpec;
  left?: BorderSpec;
  bottom?: BorderSpec;
  right?: BorderSpec;
  between?: BorderSpec;
  bar?: BorderSpec;
  insideH?: BorderSpec;
  insideV?: BorderSpec;
}

export interface TabStop {
  /** Position in twips from the paragraph's start indent origin. */
  pos: number;
  val: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'num' | 'clear';
  leader?: 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface ResolvedRunProps {
  /** Font per script slot after theme resolution (`w:rFonts`). */
  fonts: {
    ascii?: string;
    hAnsi?: string;
    eastAsia?: string;
    cs?: string;
    /**
     * `w:rFonts/@w:hint`: which slot ambiguous characters (shared by several scripts, e.g.
     * quotes or CJK punctuation) use; absent = 'default'.
     */
    hint?: 'default' | 'eastAsia' | 'cs';
  };
  /** Points. */
  size: number;
  /** Points (complex scripts). */
  sizeCs: number;
  bold: boolean;
  italic: boolean;
  boldCs: boolean;
  italicCs: boolean;
  /** `w:u/@w:val` when underlined (single, double, words, ...). */
  underline?: string;
  underlineColor?: string;
  strike: boolean;
  doubleStrike: boolean;
  /** RRGGBB or 'auto'. */
  color: string;
  /** Highlight color name (yellow, ...). */
  highlight?: string;
  /** Shading fill RRGGBB. */
  shading?: string;
  caps: boolean;
  smallCaps: boolean;
  /** Hidden text (`w:vanish`). */
  hidden: boolean;
  vertAlign: 'baseline' | 'superscript' | 'subscript';
  /** Character spacing, twips. */
  spacing: number;
  /** Raised/lowered position, half-points (`w:position`). */
  position: number;
  /** Horizontal scale percent (`w:w`), default 100. */
  scale: number;
  /** Kerning threshold in points (`w:kern`), 0 = off. */
  kern: number;
  rtl: boolean;
  /** Complex-script formatting applies (`w:cs`). */
  cs: boolean;
  lang?: string;
  langEastAsia?: string;
  langBidi?: string;
  /** Character style id after resolution. */
  styleId?: string;
}

export interface ResolvedParagraphProps {
  styleId?: string;
  alignment: Alignment;
  /** Twips; `start`/`end` already mapped to left/right using `bidi`. */
  indent: { left: number; right: number; firstLine: number; hanging: number };
  spacing: {
    before: number;
    after: number;
    beforeAutospacing: boolean;
    afterAutospacing: boolean;
    /** Twips for exact/atLeast; 240ths of a line for auto. */
    line: number;
    lineRule: 'auto' | 'exact' | 'atLeast';
  };
  contextualSpacing: boolean;
  keepNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
  widowControl: boolean;
  suppressAutoHyphens: boolean;
  bidi: boolean;
  /** 0-8 (Heading 1 = 0); undefined = body text. */
  outlineLevel?: number;
  /** Effective list membership (numId 0 = explicitly none → undefined). */
  numbering?: { numId: number; level: number };
  /** Effective tab stops, sorted by position, `clear` stops removed. */
  tabs: TabStop[];
  borders: Borders;
  /** Shading fill RRGGBB. */
  shading?: string;
  /** Raw `w:framePr` when present (frames are laid out specially). */
  framePr?: Readonly<Record<string, string>>;
  /** Run properties of the paragraph's base (style + defaults), the start point for each run. */
  runDefaults: ResolvedRunProps;
  /** Run properties of the paragraph mark (empty paragraph height, list label font). */
  markRun: ResolvedRunProps;
}

export interface CellMargins {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface ResolvedTableProps {
  styleId?: string;
  /** `w:tblW`: twips or percent*50 when `type` is 'pct'. */
  width?: { value: number; type: 'dxa' | 'pct' | 'auto' | 'nil' };
  alignment: 'left' | 'center' | 'right';
  /** Twips. */
  indent: number;
  layout: 'fixed' | 'autofit';
  cellMargins: CellMargins;
  /** Twips. */
  cellSpacing: number;
  borders: Borders;
  look: {
    firstRow: boolean;
    lastRow: boolean;
    firstColumn: boolean;
    lastColumn: boolean;
    noHBand: boolean;
    noVBand: boolean;
  };
  bidiVisual: boolean;
}

export interface ResolvedCellProps {
  width?: { value: number; type: 'dxa' | 'pct' | 'auto' | 'nil' };
  margins: CellMargins;
  /** After conflict resolution with table borders and conditional formatting. */
  borders: Borders;
  shading?: string;
  vAlign: 'top' | 'center' | 'bottom';
  noWrap: boolean;
  textDirection?: string;
  /**
   * Table-style conditional formatting regions (`w:tblStylePr/@w:type`: firstRow, band1Horz,
   * nwCell, ...) that apply to this cell and that the style defines, low → high precedence.
   * Their paragraph/run props are already part of resolveParagraph/resolveRun for the cell.
   */
  conditional: string[];
}

export interface ListInfo {
  numId: number;
  abstractNumId: number;
  level: number;
  /** The rendered label, e.g. "1.2.a" or "•". */
  label: string;
  /** `w:numFmt` of the level (decimal, lowerLetter, bullet, ...). */
  format: string;
  isBullet: boolean;
  /** What follows the label: tab, space or nothing. */
  suffix: 'tab' | 'space' | 'nothing';
  /** Label alignment (`w:lvlJc`). */
  alignment: 'left' | 'center' | 'right';
  /** Run props for the label (level rPr over the paragraph mark). */
  runProps: ResolvedRunProps;
  /** Counter values of levels 0..level used to build the label. */
  counters: number[];
}

export interface ThemeInfo {
  fonts: {
    major: { latin?: string; eastAsia?: string; cs?: string };
    minor: { latin?: string; eastAsia?: string; cs?: string };
  };
  /** Theme color name (accent1, dk1, hyperlink, ...) → RRGGBB. */
  colors: Record<string, string>;
}
