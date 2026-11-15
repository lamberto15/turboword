// turboword: lossless editing of Word .docx files.
//
// The public entry point. Every export is named explicitly (no `export *`), so the API surface
// is exactly what is listed here and internal helpers never leak by accident.
//
//   readDocx / loadDocx / createDocx  →  WordDocument  →  writeDocx / docxBlob / docxStream
//
// The root entry holds the document model, editing, text search/replace, revisions, formatting
// resolution and protection. Heavier features live in subpaths so they tree-shake away:
// turboword/compose, /fields, /forms, /crypto, /layout, /pdf, /svg.
//
// Layers below the document model (OPC package, XML DOM, errors, limits) are exported too, for
// work the typed views do not cover yet: every part, relationship and node stays reachable.

// ---- errors, limits, warnings -------------------------------------------------------------
export { DocxFormatError, TurbowordError, UsageError, XmlParseError } from './core/errors.js';
export { DEFAULT_LIMITS, type Limits } from './core/limits.js';
export type { Warning, WarningHandler } from './core/warnings.js';

// ---- OPC package --------------------------------------------------------------------------
export { ContentTypes } from './opc/content-types.js';
export { OpcPackage, type OpenOptions, type SaveOptions } from './opc/package.js';
export { Part } from './opc/part.js';
export { CT, REL, relTypeIs } from './opc/rel-types.js';
export { type AddRelationshipOptions, type Relationship, Relationships } from './opc/rels.js';
// ---- protection -----------------------------------------------------------------------------
export {
  type ProtectionHashAlgorithm,
  type ProtectionMode,
  type ProtectOptions,
  protectDocument,
  protectionOf,
  unprotectDocument,
  verifyProtectionPassword,
} from './protect/index.js';
// ---- formatting resolution (style cascade, theme, list labels) ------------------------------
export {
  type Alignment,
  type BorderSpec,
  type Borders,
  type CellMargins,
  formatNumber,
  type ListInfo,
  listInfo,
  listLabel,
  type ResolvedCellProps,
  type ResolvedParagraphProps,
  type ResolvedRunProps,
  type ResolvedTableProps,
  resolveCell,
  resolveParagraph,
  resolveRun,
  resolveTable,
  type TabStop,
  type ThemeInfo,
  themeOf,
} from './resolve/index.js';
// ---- text: find/replace across runs, revisions --------------------------------------------
export {
  acceptAllRevisions,
  type FindOptions,
  type FlatText,
  findText,
  flatten,
  type Replacement,
  type ReplaceOptions,
  type RunInput,
  rejectAllRevisions,
  replaceText,
  type SearchScope,
  type TextMatch,
} from './text/index.js';
export {
  Bookmark,
  type BookmarkRange,
  Bookmarks,
  MAX_BOOKMARK_NAME,
} from './wml/bookmarks.js';
// ---- documents ----------------------------------------------------------------------------
export { createDocx } from './wml/create.js';
export {
  DOCX_MIME,
  type DocxSource,
  docxBlob,
  docxStream,
  loadDocx,
  type ReadOptions,
  readDocx,
  WordDocument,
  type WriteOptions,
  writeDocx,
} from './wml/document.js';
export {
  copyHeaderFooter,
  createHeaderFooter,
  HeaderFooter,
  type HeaderFooterKind,
} from './wml/header-footer.js';
export { Hyperlink } from './wml/hyperlinks.js';
export { DocumentIds, type HexIdKind, type NumericIdKind } from './wml/ids.js';
export {
  type AddedImage,
  addImagePart,
  type ImageInfo,
  type ImageRef,
  imageInfo,
  inlineDrawing,
} from './wml/media.js';
export {
  type HighlightColor,
  type HyperlinkProps,
  type ImageOptions,
  type LineSpacing,
  Paragraph,
  type ParagraphAlignment,
  Run,
  type RunProps,
  type UnderlineStyle,
} from './wml/paragraph.js';
export {
  type AppProperties,
  type CoreDateProperty,
  type CoreTextProperty,
  type CustomPropertyValue,
  DocumentProperties,
} from './wml/properties.js';
export {
  evenAndOddHeaders,
  type HeaderFooterType,
  insertSectionBreak,
  type PageMargins,
  type PageOrientation,
  type PageSize,
  removeSectionBreak,
  Section,
  type SectionType,
  setEvenAndOddHeaders,
} from './wml/section.js';
export { type Block, Story, type StoryKind } from './wml/story.js';
export {
  type CellVerticalAlignment,
  type CloneRowOptions,
  type GridCell,
  newTable,
  type RowHeightRule,
  Table,
  TableCell,
  type TableGrid,
  TableRow,
} from './wml/table.js';
export { paragraphText, type VisibleVisitor, walkVisible } from './wml/text.js';
export { type ValidationIssue, validatePackage } from './wml/validate.js';

// ---- XML DOM ------------------------------------------------------------------------------
export {
  domEqual,
  type XAttr,
  XComment,
  XDocument,
  XElement,
  type XNode,
  XProcessingInstruction,
  XText,
} from './xml/dom.js';
export { NS } from './xml/ns.js';
export { type ParseOptions, parseXml } from './xml/parse.js';
export { serializeXml } from './xml/serialize.js';

// ---- zip errors ---------------------------------------------------------------------------
export { ZipFormatError, ZipLimitError, ZipPathError } from './zip/index.js';
