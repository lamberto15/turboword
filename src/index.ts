// turboword: lossless editing of Word .docx files.
//
// The public entry point. Every export is named explicitly (no `export *`), so the API surface
// is exactly what is listed here and internal helpers never leak by accident.
//
//   readDocx / loadDocx / createDocx  →  WordDocument  →  writeDocx / docxBlob / docxStream
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
export { HeaderFooter, type HeaderFooterKind } from './wml/header-footer.js';
export { Paragraph, Run, type RunProps } from './wml/paragraph.js';
export {
  type AppProperties,
  type CoreDateProperty,
  type CoreTextProperty,
  type CustomPropertyValue,
  DocumentProperties,
} from './wml/properties.js';
export {
  type HeaderFooterType,
  type PageMargins,
  type PageSize,
  Section,
  type SectionType,
} from './wml/section.js';
export { type Block, Story, type StoryKind } from './wml/story.js';
export { type GridCell, Table, TableCell, type TableGrid, TableRow } from './wml/table.js';
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
