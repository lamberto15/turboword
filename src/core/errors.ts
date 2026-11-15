// Error classes. Every error turboword throws on purpose is a TurbowordError with a stable `code`,
// so callers can branch on `err.code` instead of parsing messages.

/** Base class of every error turboword throws on purpose. */
export class TurbowordError extends Error {
  /** Stable machine-readable code, e.g. `'ZIP_LIMIT'`, `'XML_PARSE'`, `'DOCX_FORMAT'`. */
  readonly code: string;

  constructor(code: string, message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = 'TurbowordError';
    this.code = code;
  }
}

/** The input is not a usable .docx package (not a zip, no main document part, ...). */
export class DocxFormatError extends TurbowordError {
  constructor(message: string, options?: { cause?: unknown }) {
    super('DOCX_FORMAT', message, options);
    this.name = 'DocxFormatError';
  }
}

/** An XML part is not well-formed (or exceeds `maxXmlDepth`). */
export class XmlParseError extends TurbowordError {
  /** Part the XML came from, when known. */
  readonly part: string | undefined;
  /** UTF-16 offset in the part's text where the problem was found. */
  readonly offset: number;

  constructor(message: string, offset: number, part?: string) {
    super('XML_PARSE', part === undefined ? message : `${message} (in ${part})`);
    this.name = 'XmlParseError';
    this.part = part;
    this.offset = offset;
  }
}

/** An operation was used incorrectly (bad argument, detached node, ...). */
export class UsageError extends TurbowordError {
  constructor(message: string) {
    super('USAGE', message);
    this.name = 'UsageError';
  }
}
