// View over a relationships part (`_rels/.rels`, `word/_rels/document.xml.rels`, ...).
//
// Like ContentTypes, the DOM is the only source of truth: edits made directly to the DOM are seen,
// and every change goes through DOM methods, which marks the rels part dirty. A part without a
// rels part has an empty set; the rels part (and its `rels` Default content type) is created by
// the package on the first `add()`.
//
// Reads go through an index (relationships by id and by target part, the highest `rIdN`) that is
// rebuilt when the rels DOM changed since it was built (XDocument.version), and kept up to date
// by add() and remove() themselves, so a part with thousands of relationships (a merge of many
// documents) does not cost a full scan per lookup.

import { UsageError } from '../core/errors.js';
import { ELEMENT_NODE, type XDocument, type XElement } from '../xml/dom.js';
import { relTypeIs } from './rel-types.js';
import { partDir, partKey, resolveTarget } from './uri.js';
import { childElement, childrenNamed } from './xml-util.js';

export interface Relationship {
  readonly id: string;
  readonly type: string;
  /** Target as written. */
  readonly target: string;
  readonly external: boolean;
  /** Resolved part name for internal targets (no leading slash); undefined when external. */
  readonly partName: string | undefined;
}

export interface AddRelationshipOptions {
  external?: boolean;
  /** Explicit id (must be unused); default = next free `rIdN`. */
  id?: string;
}

/** `rId<N>` ids; longer numbers are ignored so `N + 1` stays an exact integer. */
const RID = /^rId(\d{1,9})$/;

/** Lookup index over one rels DOM at one version. */
interface Index {
  doc: XDocument;
  version: number;
  list: Relationship[];
  byId: Map<string, Relationship>;
  /** partKey of the resolved target → first internal relationship to it. */
  byPart: Map<string, Relationship>;
  max: number;
}

export class Relationships {
  readonly #source: string;
  readonly #getDoc: (create: boolean) => XDocument | undefined;
  #index: Index | undefined;

  /**
   * @internal `sourcePart` = name of the part these relationships belong to ('' = package);
   * `getDoc(create)` returns the rels part DOM, creating the part when `create` is true.
   */
  constructor(sourcePart: string, getDoc: (create: boolean) => XDocument | undefined) {
    this.#source = sourcePart;
    this.#getDoc = getDoc;
  }

  /** Name of the part these relationships belong to ('' for the package). */
  get source(): string {
    return this.#source;
  }

  all(): Relationship[] {
    return this.#lookup()?.list.slice() ?? [];
  }

  get(id: string): Relationship | undefined {
    return this.#lookup()?.byId.get(id);
  }

  /** The first internal relationship whose target is part `partName`, if any. */
  toPart(partName: string): Relationship | undefined {
    return this.#lookup()?.byPart.get(partKey(partName));
  }

  /** Relationships whose type matches `kind` via relTypeIs (a full URI or a last segment such as 'header'). */
  byType(kind: string): Relationship[] {
    return this.all().filter((r) => relTypeIs(r.type, kind));
  }

  /** Adds a relationship; `target` is relative to the source part's directory (or absolute URL when external). */
  add(type: string, target: string, options: AddRelationshipOptions = {}): Relationship {
    let id = options.id;
    if (id !== undefined) {
      if (this.get(id) !== undefined) throw new UsageError(`Relationship id ${id} already exists`);
    } else {
      id = this.nextId();
    }
    const doc = this.#getDoc(true);
    if (doc === undefined) throw new UsageError(`Cannot create relationships for ${this.#source}`);
    const external = options.external === true;
    const attrs: [string, string][] = [
      ['Id', id],
      ['Type', type],
      ['Target', target],
    ];
    if (external) attrs.push(['TargetMode', 'External']);
    const idx = this.#lookup();
    const root = doc.root;
    root.append(childElement(root, 'Relationship', attrs));
    const rel: Relationship = {
      id,
      type,
      target,
      external,
      partName: this.#resolve(target, external),
    };
    if (idx !== undefined && idx.doc === doc) {
      // Keep the index current instead of rebuilding it on the next read.
      addToIndex(idx, rel);
      idx.version = doc.version;
    }
    return rel;
  }

  remove(id: string): boolean {
    let found = false;
    for (const el of this.#elements()) {
      if (el.attr('Id') === id) {
        el.remove();
        found = true;
      }
    }
    if (found) this.#index = undefined;
    return found;
  }

  /** Next unused `rIdN`. */
  nextId(): string {
    const idx = this.#lookup();
    if (idx === undefined) return 'rId1';
    let n = idx.max + 1;
    while (idx.byId.has(`rId${n}`)) n++;
    return `rId${n}`;
  }

  /** The index of the current rels DOM (rebuilt when the DOM changed), undefined without one. */
  #lookup(): Index | undefined {
    const doc = this.#getDoc(false);
    if (doc === undefined) return undefined;
    const idx = this.#index;
    if (idx !== undefined && idx.doc === doc && idx.version === doc.version) return idx;
    const fresh: Index = {
      doc,
      version: doc.version,
      list: [],
      byId: new Map(),
      byPart: new Map(),
      max: 0,
    };
    for (const el of this.#elements()) {
      const rel = this.#toRelationship(el);
      if (rel !== undefined) addToIndex(fresh, rel);
      else {
        // An element without Target still reserves its id.
        const id = el.attr('Id');
        if (id !== undefined) reserveId(fresh, id);
      }
    }
    this.#index = fresh;
    return fresh;
  }

  #elements(): XElement[] {
    const doc = this.#getDoc(false);
    if (doc === undefined) return [];
    const root = doc.children.find((k): k is XElement => k.nodeType === ELEMENT_NODE);
    return root === undefined ? [] : childrenNamed(root, 'Relationship');
  }

  #resolve(target: string, external: boolean): string | undefined {
    return external ? undefined : resolveTarget(partDir(this.#source), target);
  }

  #toRelationship(el: XElement): Relationship | undefined {
    const id = el.attr('Id');
    const target = el.attr('Target');
    if (id === undefined || target === undefined) return undefined;
    const external = el.attr('TargetMode')?.toLowerCase() === 'external';
    return {
      id,
      type: el.attr('Type') ?? '',
      target,
      external,
      partName: this.#resolve(target, external),
    };
  }
}

function reserveId(idx: Index, id: string): void {
  const m = RID.exec(id);
  if (m !== null) idx.max = Math.max(idx.max, Number(m[1]));
}

function addToIndex(idx: Index, rel: Relationship): void {
  idx.list.push(rel);
  if (!idx.byId.has(rel.id)) idx.byId.set(rel.id, rel);
  reserveId(idx, rel.id);
  if (rel.partName !== undefined) {
    const key = partKey(rel.partName);
    if (!idx.byPart.has(key)) idx.byPart.set(key, rel);
  }
}
