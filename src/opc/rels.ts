// View over a relationships part (`_rels/.rels`, `word/_rels/document.xml.rels`, ...).
//
// Like ContentTypes, it keeps no copy of the data: every call reads the rels part's DOM (so edits
// made directly to the DOM are seen) and every change goes through DOM methods, which marks the
// rels part dirty. A part without a rels part has an empty set; the rels part (and its `rels`
// Default content type) is created by the package on the first `add()`.

import { UsageError } from '../core/errors.js';
import { ELEMENT_NODE, type XDocument, type XElement } from '../xml/dom.js';
import { relTypeIs } from './rel-types.js';
import { partDir, resolveTarget } from './uri.js';
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

export class Relationships {
  readonly #source: string;
  readonly #getDoc: (create: boolean) => XDocument | undefined;

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
    const out: Relationship[] = [];
    for (const el of this.#elements()) {
      const rel = this.#toRelationship(el);
      if (rel !== undefined) out.push(rel);
    }
    return out;
  }

  get(id: string): Relationship | undefined {
    for (const el of this.#elements()) {
      if (el.attr('Id') === id) return this.#toRelationship(el);
    }
    return undefined;
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
    const root = doc.root;
    root.append(childElement(root, 'Relationship', attrs));
    return { id, type, target, external, partName: this.#resolve(target, external) };
  }

  remove(id: string): boolean {
    let found = false;
    for (const el of this.#elements()) {
      if (el.attr('Id') === id) {
        el.remove();
        found = true;
      }
    }
    return found;
  }

  /** Next unused `rIdN`. */
  nextId(): string {
    const ids = new Set<string>();
    let max = 0;
    for (const el of this.#elements()) {
      const id = el.attr('Id');
      if (id === undefined) continue;
      ids.add(id);
      const m = RID.exec(id);
      if (m !== null) max = Math.max(max, Number(m[1]));
    }
    let n = max + 1;
    while (ids.has(`rId${n}`)) n++;
    return `rId${n}`;
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
