// An OPC package: the zip of parts behind every .docx.
//
// Opening only indexes the archive (central directory) and parses [Content_Types].xml; parts are
// inflated and parsed when first used. Saving copies every clean entry with its original
// compressed payload and CRC, re-compresses only dirty parts, and returns the exact input bytes
// when nothing changed at all. Entry order follows the source, then added parts in insertion
// order, so output is deterministic. Part names are matched case-insensitively (see uri.ts).
//
// fork() makes an independent copy that shares the source bytes and zip entries (immutable) and
// deep-copies parsed DOMs, so one template can be opened once and edited many times.

import { DocxFormatError, UsageError } from '../core/errors.js';
import { type Limits, resolveLimits } from '../core/limits.js';
import { type WarningHandler, warning } from '../core/warnings.js';
import { ELEMENT_NODE, type XDocument, type XElement, type XNode } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import {
  type CompressOptions,
  type LazyEntry,
  makeEntry,
  type RawEntry,
  rawEntry,
  shouldStore,
  unzipLazy,
  writeZip,
  ZipFormatError,
  ZipLimitError,
} from '../zip/index.js';
import { ContentTypes } from './content-types.js';
import { Part } from './part.js';
import { CT } from './rel-types.js';
import { Relationships } from './rels.js';
import { normalizePartName, partExtension, partKey, relsPathFor, sourceOfRels } from './uri.js';
import { childElement, emptyDocument } from './xml-util.js';

export interface OpenOptions {
  limits?: Partial<Limits>;
  onWarning?: WarningHandler;
}

export interface SaveOptions {
  /** Deflate level for re-serialized/new parts (0-9). Default 6. */
  level?: number;
  deflate?: 'native' | 'fflate';
  /** Timestamp for re-serialized/new entries. Default 1980-01-01. */
  zipDate?: Date;
}

export type { CompressOptions };

const CONTENT_TYPES = '[Content_Types].xml';
const CONTENT_TYPES_KEY = CONTENT_TYPES.toLowerCase();

/** One entry of the source archive, in central-directory order. */
interface SourceEntry {
  readonly name: string;
  readonly key: string;
  readonly entry: LazyEntry;
}

/**
 * Reads the zip index. Malformed input becomes a DocxFormatError; limit and unsafe-path
 * violations (security signals) propagate unchanged.
 */
function readArchive(data: Uint8Array, limits: Limits): Map<string, LazyEntry> {
  try {
    return unzipLazy(data, limits);
  } catch (e) {
    const malformed = !(e instanceof ZipLimitError) || e instanceof ZipFormatError;
    if (!malformed) throw e;
    const why = e instanceof Error ? e.message : String(e);
    throw new DocxFormatError(`Not a valid .docx package: ${why}`, { cause: e });
  }
}

const isElement = (n: XNode): n is XElement => n.nodeType === ELEMENT_NODE;

/** Throws UsageError for a part name that cannot be a zip entry of a package. */
function assertPartName(name: string): void {
  if (name === '' || name.includes('\\') || name.includes('\0')) {
    throw new UsageError(`Invalid part name: ${JSON.stringify(name)}`);
  }
  for (const seg of name.split('/')) {
    if (seg === '' || seg === '.' || seg === '..') {
      throw new UsageError(`Invalid part name: ${JSON.stringify(name)}`);
    }
  }
}

export class OpcPackage {
  /** Input bytes (undefined for create()). Never mutated; shared by forks. */
  readonly #source: Uint8Array | undefined;
  /** Source entries kept (case-insensitive duplicates dropped). Never mutated; shared by forks. */
  readonly #order: readonly SourceEntry[];
  readonly #limits: Limits;
  /** partKey → part, for source and added parts. */
  readonly #parts = new Map<string, Part>();
  /** Parts added since open, in insertion order. */
  #added: Part[] = [];
  /** partKeys of removed source entries. */
  #removed = new Set<string>();
  #contentTypes: ContentTypes | undefined;
  #rels: Relationships | undefined;

  private constructor(
    source: Uint8Array | undefined,
    order: readonly SourceEntry[],
    limits: Limits,
  ) {
    this.#source = source;
    this.#order = order;
    this.#limits = limits;
  }

  /** Opens a package. Parts are parsed lazily; nothing is inflated up front. */
  static open(data: Uint8Array, options: OpenOptions = {}): OpcPackage {
    const limits = resolveLimits(options.limits);
    const entries = readArchive(data, limits);
    const onWarning = options.onWarning;
    const order: SourceEntry[] = [];
    const firstName = new Map<string, string>();
    const pkg = new OpcPackage(data, order, limits);
    for (const [name, entry] of entries) {
      const key = partKey(name);
      const first = firstName.get(key);
      if (first !== undefined) {
        onWarning?.(
          warning(
            'DUPLICATE_PART',
            `Zip entry ${name} duplicates ${first} (part names are case-insensitive); ignored`,
            name,
          ),
        );
        continue;
      }
      firstName.set(key, name);
      order.push({ name, key, entry });
      if (name.endsWith('/')) continue; // directory entry (unzipLazy skips them): not a part
      const part = new Part(pkg, normalizePartName(name));
      part._entry = entry;
      pkg.#parts.set(key, part);
    }
    if (!pkg.#parts.has(CONTENT_TYPES_KEY)) {
      throw new DocxFormatError(`Not a valid .docx package: missing ${CONTENT_TYPES}`);
    }
    const root = pkg.contentTypes.xml.children.find(isElement);
    if (root === undefined || root.local !== 'Types') {
      throw new DocxFormatError(`Not a valid .docx package: ${CONTENT_TYPES} has no <Types> root`);
    }
    if (onWarning !== undefined) pkg.#checkContentTypes(onWarning);
    return pkg;
  }

  /** An empty package ([Content_Types].xml with rels/xml defaults, empty _rels/.rels). */
  static create(): OpcPackage {
    const pkg = new OpcPackage(undefined, [], resolveLimits());
    const types = emptyDocument(NS.ct, 'Types');
    const root = types.root;
    root.append(
      childElement(root, 'Default', [
        ['Extension', 'rels'],
        ['ContentType', CT.relationships],
      ]),
      childElement(root, 'Default', [
        ['Extension', 'xml'],
        ['ContentType', CT.xml],
      ]),
    );
    pkg.addPart(CONTENT_TYPES, CT.xml, types);
    pkg.addPart('_rels/.rels', CT.relationships, emptyDocument(NS.rel, 'Relationships'));
    return pkg;
  }

  get contentTypes(): ContentTypes {
    const part = this.#parts.get(CONTENT_TYPES_KEY);
    if (part === undefined) throw new DocxFormatError(`Package has no ${CONTENT_TYPES}`);
    const doc = part.xml();
    if (this.#contentTypes?.xml !== doc) this.#contentTypes = new ContentTypes(doc);
    return this.#contentTypes;
  }

  /** Package-level relationships (_rels/.rels). */
  get rels(): Relationships {
    this.#rels ??= new Relationships('', (create) => this._relsDoc('', create));
    return this.#rels;
  }

  /** Part by name (case-insensitive; a leading '/' is ignored). Includes .rels and [Content_Types].xml. */
  part(name: string): Part | undefined {
    return this.#parts.get(partKey(name));
  }

  /** All parts: source order, then added ones. */
  parts(): Part[] {
    const out: Part[] = [];
    for (const s of this.#order) {
      if (this.#removed.has(s.key)) continue;
      const part = this.#parts.get(s.key);
      if (part !== undefined) out.push(part);
    }
    out.push(...this.#added);
    return out;
  }

  /** Adds a part (sets a content-type override unless the extension default already matches). Throws if it exists. */
  addPart(name: string, contentType: string, data: Uint8Array | XDocument): Part {
    const n = normalizePartName(name);
    assertPartName(n);
    const key = partKey(n);
    if (this.#parts.has(key)) throw new UsageError(`Part ${n} already exists`);
    const part = new Part(this, n);
    part._new = true;
    if (data instanceof Uint8Array) part._data = data;
    else part._doc = data;
    this.#parts.set(key, part);
    this.#added.push(part);
    if (key !== CONTENT_TYPES_KEY) {
      const ct = this.contentTypes;
      const ext = partExtension(n);
      if (ext === undefined || ct.getDefault(ext) !== contentType) ct.setOverride(n, contentType);
    }
    return part;
  }

  /** Removes a part, its override and its rels part. Relationships pointing at it are NOT removed. */
  removePart(name: string): boolean {
    const key = partKey(name);
    const part = this.#parts.get(key);
    if (part === undefined) return false;
    if (key === CONTENT_TYPES_KEY) throw new UsageError(`${CONTENT_TYPES} cannot be removed`);
    this.#parts.delete(key);
    const i = this.#added.indexOf(part);
    if (i >= 0) this.#added.splice(i, 1);
    else this.#removed.add(key);
    this.contentTypes.removeOverride(part.name);
    const relsName = relsPathFor(part.name);
    if (this.#parts.has(partKey(relsName))) this.removePart(relsName);
    return true;
  }

  /** First free name from a template with `{n}` (n = 1, 2, ...), e.g. 'word/header{n}.xml'. */
  uniqueName(template: string): string {
    if (!template.includes('{n}')) throw new UsageError(`Name template lacks {n}: ${template}`);
    for (let n = 1; ; n++) {
      const name = normalizePartName(template.replaceAll('{n}', String(n)));
      if (!this.#parts.has(partKey(name))) return name;
    }
  }

  /** Any part modified, added or removed since open. */
  get dirty(): boolean {
    if (this.#removed.size > 0 || this.#added.length > 0) return true;
    for (const part of this.#parts.values()) if (part.dirty) return true;
    return false;
  }

  /**
   * Writes the package. Clean entries are copied with their original compressed bytes and CRC;
   * when nothing changed, returns the exact input bytes.
   */
  save(options: SaveOptions = {}): Uint8Array {
    if (this.#source !== undefined && !this.dirty) return this.#source;
    const base: CompressOptions = {};
    if (options.level !== undefined) base.level = options.level;
    if (options.deflate !== undefined) base.deflate = options.deflate;
    if (options.zipDate !== undefined) base.date = options.zipDate;
    const fresh = (part: Part): RawEntry =>
      makeEntry(part.name, part.bytes(), { ...base, store: shouldStore(part.name) });

    const out: RawEntry[] = [];
    for (const s of this.#order) {
      if (this.#removed.has(s.key)) continue;
      const part = this.#parts.get(s.key);
      // no part = directory entry
      out.push(part === undefined || !part.dirty ? rawEntry(s.entry) : fresh(part));
    }
    for (const part of this.#added) out.push(fresh(part));
    return writeZip(out);
  }

  /** Independent copy sharing the unchanged compressed entries (copy-on-write). */
  fork(): OpcPackage {
    const f = new OpcPackage(this.#source, this.#order, this.#limits);
    for (const [key, part] of this.#parts) f.#parts.set(key, part._fork(f));
    f.#added = this.#added.map((p) => f.#parts.get(partKey(p.name))!);
    f.#removed = new Set(this.#removed);
    return f;
  }

  /** Removes parts not reachable from the package relationships; returns their names. */
  gc(): string[] {
    const reachable = new Set<string>();
    const queue: Part[] = [];
    const visit = (rels: Relationships): void => {
      for (const rel of rels.all()) {
        if (rel.partName === undefined) continue;
        const key = partKey(rel.partName);
        const part = this.#parts.get(key);
        if (part === undefined || reachable.has(key)) continue;
        reachable.add(key);
        queue.push(part);
      }
    };
    visit(this.rels);
    for (let part = queue.pop(); part !== undefined; part = queue.pop()) visit(part.rels);

    const doomed: string[] = [];
    for (const part of this.parts()) {
      const key = partKey(part.name);
      if (key === CONTENT_TYPES_KEY || reachable.has(key)) continue;
      const source = sourceOfRels(part.name);
      if (source !== undefined && (source === '' || reachable.has(partKey(source)))) continue;
      doomed.push(part.name);
    }
    for (const name of doomed) this.removePart(name);
    return doomed;
  }

  /** @internal Limits the package was opened with. */
  get _limits(): Limits {
    return this.#limits;
  }

  /**
   * @internal DOM of the rels part of `source` ('' = package). With `create`, a missing rels part
   * is added (and the `rels` Default content type ensured).
   */
  _relsDoc(source: string, create: boolean): XDocument | undefined {
    const relsName = relsPathFor(source);
    const existing = this.part(relsName);
    if (existing !== undefined) return existing.xml();
    if (!create) return undefined;
    const ct = this.contentTypes;
    if (ct.getDefault('rels') === undefined) ct.setDefault('rels', CT.relationships);
    const doc = emptyDocument(NS.rel, 'Relationships');
    this.addPart(relsName, CT.relationships, doc);
    return doc;
  }

  /** Reports dangling overrides and parts without a content type. */
  #checkContentTypes(onWarning: WarningHandler): void {
    const ct = this.contentTypes;
    const defaults = ct.defaults();
    const overridden = new Set<string>();
    for (const name of ct.overrides().keys()) {
      const key = partKey(name);
      overridden.add(key);
      if (!this.#parts.has(key)) {
        onWarning(
          warning('DANGLING_OVERRIDE', `Content type override for missing part ${name}`, name),
        );
      }
    }
    for (const part of this.parts()) {
      const key = partKey(part.name);
      if (key === CONTENT_TYPES_KEY || overridden.has(key)) continue;
      const ext = partExtension(part.name);
      if (ext !== undefined && defaults.has(ext)) continue;
      onWarning(
        warning('MISSING_CONTENT_TYPE', `Part ${part.name} has no content type`, part.name),
      );
    }
  }
}
