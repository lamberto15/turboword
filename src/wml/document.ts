// WordDocument and the read/write entry points.
//
// A WordDocument is a thin view over an OpcPackage: it locates the main document part through
// the package's officeDocument relationship (never by a hard-coded name: Word itself does not
// care what the part is called) and hands out story/section/header views over the live DOM.
// Nothing is copied into a parallel model, so whatever turboword does not understand survives a
// save untouched, and an untouched document saves to its exact input bytes.

import { DocxFormatError, TurbowordError, UsageError } from '../core/errors.js';
import type { Limits } from '../core/limits.js';
import type { WarningHandler } from '../core/warnings.js';
import { OpcPackage, type OpenOptions, type SaveOptions } from '../opc/package.js';
import type { Part } from '../opc/part.js';
import { relTypeIs } from '../opc/rel-types.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { Bookmarks } from './bookmarks.js';
import { type HeaderFooter, headerFooterParts } from './header-footer.js';
import { type Hyperlink, hyperlinksOf } from './hyperlinks.js';
import { type AddedImage, addImagePart, type ImageRef, imagesOf } from './media.js';
import { DocumentProperties } from './properties.js';
import { type Section, sectionsOf } from './section.js';
import { Story, type StoryKind } from './story.js';
import { validatePackage } from './validate.js';

/** MIME type of a .docx file. */
export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** Options of {@link readDocx} and {@link loadDocx}. */
export interface ReadOptions {
  /** Safety limits for untrusted input (zip entries, sizes, ratios, XML depth). */
  limits?: Partial<Limits>;
  /** Receives non-fatal problems found while reading. */
  onWarning?: WarningHandler;
}

/** Options of {@link writeDocx}, {@link docxBlob} and {@link docxStream}. */
export interface WriteOptions extends SaveOptions {
  /**
   * Run {@link validatePackage} first and throw a `TurbowordError` with code `'VALIDATION'`
   * listing the errors (warnings do not fail). Default false.
   */
  validate?: boolean;
}

/** Anything {@link loadDocx} can read bytes from. */
export type DocxSource =
  | Uint8Array
  | ArrayBuffer
  | Blob
  | ReadableStream<Uint8Array>
  | AsyncIterable<Uint8Array>;

/** `w:type` values of footnotes/endnotes that are separators, not notes. */
const NOTE_SEPARATORS: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'separator',
  'continuationSeparator',
  'continuationNotice',
]);

/** First element child of a part's DOM, or undefined (empty or non-XML part). */
function rootOf(part: Part): XElement | undefined {
  for (const k of part.xml().children) if (k.nodeType === 1) return k;
  return undefined;
}

/** The main document part of `pkg` (target of the package-level officeDocument relationship). */
function findMainPart(pkg: OpcPackage): Part {
  const rel = pkg.rels.all().find((r) => !r.external && relTypeIs(r.type, 'officeDocument'));
  if (rel === undefined || rel.partName === undefined) {
    throw new DocxFormatError('Not a .docx: no officeDocument relationship in _rels/.rels');
  }
  const part = pkg.part(rel.partName);
  if (part === undefined) {
    throw new DocxFormatError(`Main document part ${rel.partName} is missing`);
  }
  return part;
}

/** A Word document: typed views over an {@link OpcPackage}. */
export class WordDocument {
  /** The underlying package (every part, relationship and content type). */
  readonly package: OpcPackage;
  /** The main document part (usually `word/document.xml`). */
  readonly mainPart: Part;
  /** The main body story (`w:body`). */
  readonly body: Story;
  #properties: DocumentProperties | undefined;
  #bookmarks: Bookmarks | undefined;

  /**
   * Wraps an opened package. Throws DocxFormatError when it has no main document part or the
   * part is not a WordprocessingML document (`w:document` with a `w:body`).
   */
  constructor(pkg: OpcPackage) {
    this.package = pkg;
    this.mainPart = findMainPart(pkg);
    const root = rootOf(this.mainPart);
    if (root === undefined || !root.is(NS.w, 'document')) {
      const what = root === undefined ? 'no root element' : `<${root.name}>`;
      throw new DocxFormatError(
        `Main part ${this.mainPart.name} is not a WordprocessingML document (${what})`,
      );
    }
    const body = root.element(NS.w, 'body');
    if (body === undefined) {
      throw new DocxFormatError(`Main part ${this.mainPart.name} has no <w:body>`);
    }
    this.body = new Story(this, 'body', this.mainPart, body);
  }

  /** Core, extended and custom document properties (created on first access). */
  get properties(): DocumentProperties {
    this.#properties ??= new DocumentProperties(this.package);
    return this.#properties;
  }

  /** The document's bookmarks (list, get, add, remove, rename). */
  get bookmarks(): Bookmarks {
    this.#bookmarks ??= new Bookmarks(this);
    return this.#bookmarks;
  }

  /**
   * Adds an image part (reused when the same bytes are already in the package) with a
   * relationship from the main document part. To show it, use `Paragraph.appendImage`, or
   * `inlineDrawing` with the returned relationship id. Throws UsageError for an unrecognized
   * format unless `contentType` is given.
   */
  addImage(bytes: Uint8Array, o?: { contentType?: string }): AddedImage {
    return addImagePart(this, this.mainPart, bytes, o);
  }

  /**
   * Every image reference in every story (body, headers/footers, notes, comments, including
   * text boxes and VML pictures), with the image part and displayed size.
   */
  images(): ImageRef[] {
    return imagesOf(this.stories());
  }

  /** Every hyperlink (`w:hyperlink`) in every story, in story order. */
  hyperlinks(): Hyperlink[] {
    return hyperlinksOf(this.stories());
  }

  /** The body's sections, in order. */
  sections(): Section[] {
    return sectionsOf(this);
  }

  /** Every header and footer part, in relationship order. */
  headersFooters(): HeaderFooter[] {
    return headerFooterParts(this);
  }

  /**
   * Every story, in this order: body, headers and footers, footnotes, endnotes, comments.
   * `kinds` keeps only stories of those kinds (same order).
   */
  stories(kinds?: readonly StoryKind[]): Story[] {
    const want = (k: StoryKind): boolean => kinds === undefined || kinds.includes(k);
    const out: Story[] = [];
    if (want('body')) out.push(this.body);
    if (want('header') || want('footer')) {
      for (const hf of this.headersFooters()) if (want(hf.kind)) out.push(hf.story);
    }
    if (want('footnote')) out.push(...this.footnotes());
    if (want('endnote')) out.push(...this.endnotes());
    if (want('comment')) out.push(...this.comments());
    return out;
  }

  /** One story per footnote (`w:footnote`), separators excluded, with its `w:id`. */
  footnotes(): Story[] {
    return this.#notes('footnotes', 'footnote');
  }

  /** One story per endnote (`w:endnote`), separators excluded, with its `w:id`. */
  endnotes(): Story[] {
    return this.#notes('endnotes', 'endnote');
  }

  /** One story per comment (`w:comment`), with its `w:id`. */
  comments(): Story[] {
    const out: Story[] = [];
    for (const part of this.mainPart.relatedByType('comments')) {
      const root = rootOf(part);
      if (root === undefined) continue;
      for (const c of root.elements(NS.w, 'comment')) {
        out.push(new Story(this, 'comment', part, c, c.attrNS(NS.w, 'id')));
      }
    }
    return out;
  }

  #notes(rel: 'footnotes' | 'endnotes', kind: 'footnote' | 'endnote'): Story[] {
    const out: Story[] = [];
    for (const part of this.mainPart.relatedByType(rel)) {
      const root = rootOf(part);
      if (root === undefined) continue;
      for (const n of root.elements(NS.w, kind)) {
        const type = n.attrNS(NS.w, 'type');
        if (type !== undefined && NOTE_SEPARATORS.has(type)) continue;
        out.push(new Story(this, kind, part, n, n.attrNS(NS.w, 'id')));
      }
    }
    return out;
  }

  /** Visible text of the body (paragraphs joined with LF). */
  text(): string {
    return this.body.text();
  }

  /** An independent copy (copy-on-write: unchanged parts are shared until modified). */
  fork(): WordDocument {
    return new WordDocument(this.package.fork());
  }

  /** Removes parts no relationship reaches; returns their names. */
  gc(): string[] {
    return this.package.gc();
  }
}

function openOptions(o: ReadOptions | undefined): OpenOptions {
  const out: OpenOptions = {};
  if (o?.limits !== undefined) out.limits = o.limits;
  if (o?.onWarning !== undefined) out.onWarning = o.onWarning;
  return out;
}

/** ArrayBuffer test that also works across realms (iframes, workers' transferred buffers, jsdom). */
function isArrayBuffer(x: unknown): x is ArrayBuffer {
  return x instanceof ArrayBuffer || Object.prototype.toString.call(x) === '[object ArrayBuffer]';
}

function toBytes(data: Uint8Array | ArrayBuffer | ArrayBufferView): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (ArrayBuffer.isView(data))
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  return new Uint8Array(data);
}

/**
 * Opens a .docx from its bytes. Synchronous and lazy: parts are inflated and parsed on first
 * use. Throws DocxFormatError when the package has no WordprocessingML main document.
 */
export function readDocx(data: Uint8Array | ArrayBuffer, o?: ReadOptions): WordDocument {
  return new WordDocument(OpcPackage.open(toBytes(data), openOptions(o)));
}

function concat(chunks: Uint8Array[], total: number): Uint8Array {
  if (chunks.length === 1) return chunks[0]!;
  const out = new Uint8Array(total);
  let at = 0;
  for (const c of chunks) {
    out.set(c, at);
    at += c.length;
  }
  return out;
}

function chunkBytes(chunk: unknown): Uint8Array {
  if (chunk instanceof Uint8Array) return chunk;
  if (isArrayBuffer(chunk) || ArrayBuffer.isView(chunk)) return toBytes(chunk);
  throw new UsageError('loadDocx(): stream chunks must be Uint8Array (binary, not text)');
}

/** Reads every byte of `src` (see {@link DocxSource}). */
async function collect(src: DocxSource): Promise<Uint8Array> {
  if (isArrayBuffer(src) || ArrayBuffer.isView(src)) {
    return toBytes(src);
  }
  const chunks: Uint8Array[] = [];
  let total = 0;
  if (typeof (src as ReadableStream<Uint8Array>).getReader === 'function') {
    // A reader rather than `for await`: Safari's ReadableStream is not async-iterable.
    const reader = (src as ReadableStream<Uint8Array>).getReader();
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        const b = chunkBytes(value);
        chunks.push(b);
        total += b.length;
      }
    } finally {
      reader.releaseLock();
    }
    return concat(chunks, total);
  }
  if (typeof (src as Blob).arrayBuffer === 'function') {
    return new Uint8Array(await (src as Blob).arrayBuffer());
  }
  if (typeof (src as AsyncIterable<Uint8Array>)[Symbol.asyncIterator] === 'function') {
    for await (const chunk of src as AsyncIterable<Uint8Array>) {
      const b = chunkBytes(chunk);
      chunks.push(b);
      total += b.length;
    }
    return concat(chunks, total);
  }
  throw new UsageError(
    'loadDocx(): expected Uint8Array, ArrayBuffer, Blob, ReadableStream or AsyncIterable',
  );
}

/** Opens a .docx from bytes, a Blob/File, a ReadableStream or an async iterable of chunks. */
export async function loadDocx(src: DocxSource, o?: ReadOptions): Promise<WordDocument> {
  return readDocx(await collect(src), o);
}

function saveOptions(o: WriteOptions | undefined): SaveOptions {
  const out: SaveOptions = {};
  if (o?.level !== undefined) out.level = o.level;
  if (o?.deflate !== undefined) out.deflate = o.deflate;
  if (o?.zipDate !== undefined) out.zipDate = o.zipDate;
  return out;
}

/**
 * Saves the document to .docx bytes. Untouched parts are copied with their original compressed
 * bytes; an unmodified document returns its exact input bytes.
 */
export function writeDocx(doc: WordDocument, o?: WriteOptions): Uint8Array {
  if (o?.validate === true) {
    const errors = validatePackage(doc).filter((i) => i.severity === 'error');
    if (errors.length > 0) {
      const lines = errors.map((e) => `  ${e.code}${e.part ? ` (${e.part})` : ''}: ${e.message}`);
      throw new TurbowordError(
        'VALIDATION',
        `Package has ${errors.length} validation error(s):\n${lines.join('\n')}`,
      );
    }
  }
  return doc.package.save(saveOptions(o));
}

/** {@link writeDocx} as a Blob with the .docx MIME type. */
export function docxBlob(doc: WordDocument, o?: WriteOptions): Blob {
  return new Blob([writeDocx(doc, o) as Uint8Array<ArrayBuffer>], { type: DOCX_MIME });
}

/** Chunk size of {@link docxStream}. */
const STREAM_CHUNK = 64 * 1024;

/**
 * {@link writeDocx} as a ReadableStream. The bytes are produced up front (so errors throw
 * here, not mid-stream) and handed out in 64 KiB chunks.
 */
export function docxStream(doc: WordDocument, o?: WriteOptions): ReadableStream<Uint8Array> {
  const bytes = writeDocx(doc, o);
  let at = 0;
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (at >= bytes.length) {
        controller.close();
        return;
      }
      const end = Math.min(at + STREAM_CHUNK, bytes.length);
      controller.enqueue(bytes.subarray(at, end));
      at = end;
    },
  });
}
