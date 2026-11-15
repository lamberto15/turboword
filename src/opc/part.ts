// One part of an OPC package.
//
// A part starts as a reference to its zip entry: nothing is inflated or parsed until `bytes()` or
// `xml()` is called. Its content is, in order of precedence: the cached DOM when it was mutated,
// replacement bytes from `setBytes()`, the DOM given to `setXml()`/`addPart()`, the source entry.
// "Dirty" (= must be rewritten on save) means new, replaced, or DOM mutated; merely parsing a part
// never makes it dirty, which is what lets an untouched package save back to its input bytes.

import { UsageError } from '../core/errors.js';
import type { XDocument } from '../xml/dom.js';
import { parseXmlBytes } from '../xml/parse.js';
import { serializeXmlBytes } from '../xml/serialize.js';
import type { LazyEntry } from '../zip/index.js';
import type { OpcPackage } from './package.js';
import { Relationships } from './rels.js';

const XML_TYPE = /(?:^|[/+])xml$/i;
const XML_NAME = /\.(?:xml|rels)$/i;

export class Part {
  readonly #pkg: OpcPackage;
  readonly #name: string;
  #rels: Relationships | undefined;

  /** @internal Source zip entry (undefined for new or replaced parts). */
  _entry: LazyEntry | undefined = undefined;
  /** @internal Replacement bytes (setBytes / addPart with bytes). */
  _data: Uint8Array | undefined = undefined;
  /** @internal Cached or supplied DOM. */
  _doc: XDocument | undefined = undefined;
  /** @internal Added since open. */
  _new: boolean = false;
  /** @internal Content replaced by setBytes/setXml. */
  _replaced: boolean = false;

  /** @internal */
  constructor(pkg: OpcPackage, name: string) {
    this.#pkg = pkg;
    this.#name = name;
  }

  get package(): OpcPackage {
    return this.#pkg;
  }

  /** Part name = zip entry name (no leading slash), original casing. */
  get name(): string {
    return this.#name;
  }

  get contentType(): string | undefined {
    return this.#pkg.contentTypes.get(this.#name);
  }

  /** XML content type (`…xml`, `…+xml`) or a `.xml`/`.rels` name. */
  get isXml(): boolean {
    const ct = this.contentType;
    return (ct !== undefined && XML_TYPE.test(ct)) || XML_NAME.test(this.#name);
  }

  /** Current content (the DOM serialized when it was modified). */
  bytes(): Uint8Array {
    const doc = this._doc;
    if (
      doc !== undefined &&
      (doc.dirty || (this._data === undefined && this._entry === undefined))
    ) {
      return serializeXmlBytes(doc);
    }
    if (this._data !== undefined) return this._data;
    if (this._entry !== undefined) return this._entry.bytes();
    return new Uint8Array(0);
  }

  /** Replaces the content (drops a cached DOM). */
  setBytes(data: Uint8Array): void {
    this._data = data;
    this._doc = undefined;
    this._entry = undefined;
    this._replaced = true;
  }

  /** Parsed DOM (lazy, cached). Mutating it marks the part dirty. */
  xml(): XDocument {
    let doc = this._doc;
    if (doc === undefined) {
      const src = this._data ?? this._entry?.bytes();
      if (src === undefined) throw new UsageError(`Part ${this.#name} has no content`);
      doc = parseXmlBytes(src, { part: this.#name, maxDepth: this.#pkg._limits.maxXmlDepth });
      doc.dirty = false;
      this._doc = doc;
    }
    return doc;
  }

  /** Replaces the content with a DOM. */
  setXml(doc: XDocument): void {
    this._doc = doc;
    this._data = undefined;
    this._entry = undefined;
    this._replaced = true;
  }

  /** Modified since open (bytes replaced, DOM mutated, or new). */
  get dirty(): boolean {
    return this._new || this._replaced || (this._doc?.dirty ?? false);
  }

  /** This part's relationships (an empty set when it has no rels part; created on first add). */
  get rels(): Relationships {
    this.#rels ??= new Relationships(this.#name, (create) =>
      this.#pkg._relsDoc(this.#name, create),
    );
    return this.#rels;
  }

  /** Target part of relationship `id` (internal targets only). */
  related(id: string): Part | undefined {
    const rel = this.rels.get(id);
    return rel?.partName === undefined ? undefined : this.#pkg.part(rel.partName);
  }

  /** Target parts of relationships matching `kind` (see Relationships.byType). */
  relatedByType(kind: string): Part[] {
    const out: Part[] = [];
    for (const rel of this.rels.byType(kind)) {
      const part = rel.partName === undefined ? undefined : this.#pkg.part(rel.partName);
      if (part !== undefined) out.push(part);
    }
    return out;
  }

  /**
   * @internal Copy for a forked package: shares the entry and replacement bytes, deep-copies a
   * cached DOM (a dirty DOM stays dirty in the copy).
   */
  _fork(pkg: OpcPackage): Part {
    const p = new Part(pkg, this.#name);
    p._entry = this._entry;
    p._data = this._data;
    p._new = this._new;
    p._replaced = this._replaced;
    const doc = this._doc;
    if (doc !== undefined) {
      const copy = doc.clone();
      if (doc.dirty) copy.markDirty();
      p._doc = copy;
    }
    return p;
  }
}
