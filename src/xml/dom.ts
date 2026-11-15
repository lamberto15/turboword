// A small, lossless XML DOM for OOXML parts.
//
// - Qualified names are kept exactly as written (`w:p`); matching uses the canonical namespace
//   URI (Strict OOXML URIs alias their transitional twins, see ns.ts), so prefixes never matter.
// - Every node survives a round-trip: comments, processing instructions, whitespace text, unknown
//   elements, namespace declarations and attribute order.
// - Attributes are a flat array of triples `[qname, uri, value, ...]`; namespace declarations are
//   attributes too (uri = NS.xmlns), so they are written back where they were.
// - Every mutation goes through a method and marks the owning XDocument dirty, which is how an
//   OPC part knows it must be re-serialized on save.
// - No recursion anywhere (deep documents cannot overflow the stack): traversal, clone and
//   equality use explicit stacks.

import { UsageError } from '../core/errors.js';
import { canonicalNs, conventionalPrefix, conventionalUri, hasLocal, localOf, NS } from './ns.js';

export const ELEMENT_NODE = 1;
export const TEXT_NODE = 3;
export const PI_NODE = 7;
export const COMMENT_NODE = 8;
export const DOCUMENT_NODE = 9;

export type XNode = XElement | XText | XComment | XProcessingInstruction;
export type XParent = XElement | XDocument;

/** One attribute, as returned by {@link XElement.attributes}. */
export interface XAttr {
  /** Qualified name as written (`w:val`, `xmlns:w`, `Id`). */
  readonly name: string;
  /** Canonical namespace URI ('' for unprefixed attributes; NS.xmlns for declarations). */
  readonly uri: string;
  readonly value: string;
}

/** Walks up to the owning document and marks it dirty (no-op for detached nodes). */
function touch(node: XParent | null): void {
  let n: XParent | null = node;
  while (n !== null && n.nodeType === ELEMENT_NODE) n = n.parent;
  if (n !== null) {
    n.dirty = true;
    n.version++;
  }
}

/** Removes `node` from its parent's child list (no dirty marking). */
function unlink(node: XNode): void {
  const p = node.parent;
  if (p === null) return;
  const kids = p._children;
  const i = kids.indexOf(node);
  if (i >= 0) kids.splice(i, 1);
  node.parent = null;
}

abstract class XLeaf {
  parent: XParent | null = null;

  /** Detaches this node from its parent. */
  remove(): this {
    const p = this.parent;
    if (p !== null) {
      unlink(this as unknown as XNode);
      touch(p);
    }
    return this;
  }

  /** Index among the parent's children, or -1 when detached. */
  index(): number {
    return this.parent === null ? -1 : this.parent._children.indexOf(this as unknown as XNode);
  }

  /** Replaces this node with `nodes` (in order). */
  replaceWith(...nodes: XNode[]): void {
    const p = this.parent;
    if (p === null) throw new UsageError('replaceWith() on a detached node');
    const self = this as unknown as XNode;
    if (nodes.length === 1 && nodes[0] === self) return;
    for (const n of nodes) insertCheck(p, n);
    for (const n of nodes) {
      if (n !== self) detachForInsert(n);
    }
    const i = p._children.indexOf(self);
    p._children.splice(i, 1, ...nodes);
    if (!nodes.includes(self)) this.parent = null;
    for (const n of nodes) n.parent = p;
    touch(p);
  }

  /** Inserts `nodes` right before this node. */
  before(...nodes: XNode[]): void {
    const p = this.parent;
    if (p === null) throw new UsageError('before() on a detached node');
    for (const n of nodes) p.insertBefore(n, this as unknown as XNode);
  }

  /** Inserts `nodes` right after this node (in order). */
  after(...nodes: XNode[]): void {
    const p = this.parent;
    if (p === null) throw new UsageError('after() on a detached node');
    let ref: XNode = this as unknown as XNode;
    for (const n of nodes) {
      p.insertAfter(n, ref);
      ref = n;
    }
  }

  /** Next sibling element, if any. */
  nextElement(): XElement | undefined {
    const p = this.parent;
    if (p === null) return undefined;
    const kids = p._children;
    for (let i = kids.indexOf(this as unknown as XNode) + 1; i < kids.length; i++) {
      const k = kids[i]!;
      if (k.nodeType === ELEMENT_NODE) return k;
    }
    return undefined;
  }

  /** Previous sibling element, if any. */
  previousElement(): XElement | undefined {
    const p = this.parent;
    if (p === null) return undefined;
    const kids = p._children;
    for (let i = kids.indexOf(this as unknown as XNode) - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === ELEMENT_NODE) return k;
    }
    return undefined;
  }

  /** The document this node belongs to, if attached to one. */
  ownerDocument(): XDocument | undefined {
    let n: XParent | null = this.parent;
    while (n !== null && n.nodeType === ELEMENT_NODE) n = n.parent;
    return n ?? undefined;
  }
}

/** Character data (text and CDATA sections; CDATA delimiters are not preserved). */
export class XText extends XLeaf {
  readonly nodeType: 3 = TEXT_NODE;
  #text: string;

  constructor(text: string) {
    super();
    this.#text = text;
  }

  get text(): string {
    return this.#text;
  }

  set text(value: string) {
    if (value === this.#text) return;
    this.#text = value;
    touch(this.parent);
  }

  clone(): XText {
    return new XText(this.#text);
  }
}

/** `<!-- text -->` */
export class XComment extends XLeaf {
  readonly nodeType: 8 = COMMENT_NODE;
  #text: string;

  constructor(text: string) {
    super();
    this.#text = text;
  }

  get text(): string {
    return this.#text;
  }

  set text(value: string) {
    if (value === this.#text) return;
    this.#text = value;
    touch(this.parent);
  }

  clone(): XComment {
    return new XComment(this.#text);
  }
}

/** `<?target data?>` (the XML declaration is not one: see {@link XDocument.declaration}). */
export class XProcessingInstruction extends XLeaf {
  readonly nodeType: 7 = PI_NODE;
  readonly target: string;
  #data: string;

  constructor(target: string, data: string) {
    super();
    this.target = target;
    this.#data = data;
  }

  get data(): string {
    return this.#data;
  }

  set data(value: string) {
    if (value === this.#data) return;
    this.#data = value;
    touch(this.parent);
  }

  clone(): XProcessingInstruction {
    return new XProcessingInstruction(this.target, this.#data);
  }
}

function detachForInsert(node: XNode): void {
  const old = node.parent;
  if (old !== null) {
    unlink(node);
    touch(old);
  }
}

/** Throws when inserting `node` under `parent` would create a cycle or is otherwise invalid. */
function insertCheck(parent: XParent, node: XNode): void {
  if (node.nodeType !== ELEMENT_NODE) return;
  let p: XParent | null = parent;
  while (p !== null) {
    if (p === node) throw new UsageError('Cannot insert an element into itself or a descendant');
    p = p.nodeType === ELEMENT_NODE ? p.parent : null;
  }
}

/** Shared child-list behaviour of XElement and XDocument. */
function insertAtIndex(parent: XParent, index: number, node: XNode): void {
  insertCheck(parent, node);
  if (node.parent === parent) {
    const cur = parent._children.indexOf(node);
    parent._children.splice(cur, 1);
    if (cur < index) index--;
  } else {
    detachForInsert(node);
  }
  const kids = parent._children;
  if (index < 0 || index > kids.length) index = kids.length;
  kids.splice(index, 0, node);
  node.parent = parent;
  touch(parent);
}

export class XElement extends XLeaf {
  readonly nodeType: 1 = ELEMENT_NODE;
  /** Qualified name as written, e.g. `w:p`. */
  name: string;
  /** Local name, e.g. `p`. */
  local: string;
  /** Canonical namespace URI ('' when the element is in no namespace). */
  uri: string;
  /** @internal Flat attribute triples `[qname, canonicalUri, value, ...]` in source order. */
  _attrs: string[];
  /** @internal Child nodes in order. */
  _children: XNode[];

  /**
   * Low-level constructor: `name` is the qualified name, `uri` its namespace URI (canonicalized
   * here). Prefer {@link XDocument.createElement}, which picks the prefix for you.
   */
  constructor(name: string, uri: string, attrs?: string[], children?: XNode[]) {
    super();
    this.name = name;
    this.local = localOf(name);
    this.uri = canonicalNs(uri);
    this._attrs = attrs ?? [];
    this._children = children ?? [];
  }

  /** Renames the element (qualified name and namespace). */
  setName(name: string, uri: string): void {
    this.name = name;
    this.local = localOf(name);
    this.uri = canonicalNs(uri);
    touch(this.parent);
  }

  /** True when this element is `{uri}local`. */
  is(uri: string, local: string): boolean {
    return this.local === local && this.uri === uri;
  }

  // ---- children -------------------------------------------------------------------------

  /** Child nodes (read-only view; mutate through the methods). */
  get children(): readonly XNode[] {
    return this._children;
  }

  /** Child elements, optionally only those named `{uri}local` (`local` omitted: any local name). */
  elements(uri?: string, local?: string): XElement[] {
    const out: XElement[] = [];
    for (const k of this._children) {
      if (k.nodeType !== ELEMENT_NODE) continue;
      if (uri !== undefined && k.uri !== uri) continue;
      if (local !== undefined && k.local !== local) continue;
      out.push(k);
    }
    return out;
  }

  /** First child element named `{uri}local`. */
  element(uri: string, local: string): XElement | undefined {
    for (const k of this._children) {
      if (k.nodeType === ELEMENT_NODE && k.local === local && k.uri === uri) return k;
    }
    return undefined;
  }

  /** First child element of any name. */
  firstElement(): XElement | undefined {
    for (const k of this._children) if (k.nodeType === ELEMENT_NODE) return k;
    return undefined;
  }

  /** Last child element of any name. */
  lastElement(): XElement | undefined {
    const kids = this._children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === ELEMENT_NODE) return k;
    }
    return undefined;
  }

  /**
   * Descendant elements in document order (pre-order, excluding this element), optionally only
   * those named `{uri}local`. Returns a snapshot, so the tree may be mutated while iterating it.
   */
  descendants(uri?: string, local?: string): XElement[] {
    const out: XElement[] = [];
    const stack: XNode[][] = [this._children];
    const pos: number[] = [0];
    while (stack.length > 0) {
      const top = stack.length - 1;
      const kids = stack[top]!;
      const i = pos[top]!;
      if (i >= kids.length) {
        stack.pop();
        pos.pop();
        continue;
      }
      pos[top] = i + 1;
      const k = kids[i]!;
      if (k.nodeType !== ELEMENT_NODE) continue;
      if ((uri === undefined || k.uri === uri) && (local === undefined || k.local === local)) {
        out.push(k);
      }
      if (k._children.length > 0) {
        stack.push(k._children);
        pos.push(0);
      }
    }
    return out;
  }

  /** Nearest ancestor named `{uri}local`. */
  closest(uri: string, local: string): XElement | undefined {
    let p = this.parent;
    while (p !== null && p.nodeType === ELEMENT_NODE) {
      if (p.local === local && p.uri === uri) return p;
      p = p.parent;
    }
    return undefined;
  }

  /** Concatenated text of all descendant text nodes. */
  get text(): string {
    let out = '';
    const stack: XNode[][] = [this._children];
    const pos: number[] = [0];
    while (stack.length > 0) {
      const top = stack.length - 1;
      const kids = stack[top]!;
      const i = pos[top]!;
      if (i >= kids.length) {
        stack.pop();
        pos.pop();
        continue;
      }
      pos[top] = i + 1;
      const k = kids[i]!;
      if (k.nodeType === TEXT_NODE) out += k.text;
      else if (k.nodeType === ELEMENT_NODE && k._children.length > 0) {
        stack.push(k._children);
        pos.push(0);
      }
    }
    return out;
  }

  /** Replaces all children with a single text node (none for ''). */
  set text(value: string) {
    for (const k of this._children) k.parent = null;
    this._children = [];
    if (value !== '') {
      const t = new XText(value);
      t.parent = this;
      this._children.push(t);
    }
    touch(this);
  }

  append(...nodes: XNode[]): this {
    for (const n of nodes) insertAtIndex(this, this._children.length, n);
    return this;
  }

  prepend(...nodes: XNode[]): this {
    for (let i = 0; i < nodes.length; i++) insertAtIndex(this, i, nodes[i]!);
    return this;
  }

  /** Inserts `node` before `ref` (a child of this element), or at the end when `ref` is null. */
  insertBefore(node: XNode, ref: XNode | null): this {
    if (ref === null) return this.append(node);
    if (ref.parent !== this) throw new UsageError('insertBefore(): ref is not a child');
    if (node === ref) return this;
    insertCheck(this, node);
    detachForInsert(node);
    insertAtIndex(this, this._children.indexOf(ref), node);
    return this;
  }

  /** Inserts `node` after `ref` (a child of this element), or at the start when `ref` is null. */
  insertAfter(node: XNode, ref: XNode | null): this {
    if (ref === null) return this.prepend(node);
    if (ref.parent !== this) throw new UsageError('insertAfter(): ref is not a child');
    if (node === ref) return this;
    insertCheck(this, node);
    detachForInsert(node);
    insertAtIndex(this, this._children.indexOf(ref) + 1, node);
    return this;
  }

  /** Inserts `node` at child index `index` (clamped to the end). */
  insertAt(index: number, node: XNode): this {
    insertAtIndex(this, index, node);
    return this;
  }

  /** Removes all children. */
  clear(): this {
    if (this._children.length === 0) return this;
    for (const k of this._children) k.parent = null;
    this._children = [];
    touch(this);
    return this;
  }

  // ---- attributes -----------------------------------------------------------------------

  /** Attribute by qualified name as written (`'w:val'`). */
  attr(name: string): string | undefined {
    const a = this._attrs;
    for (let i = 0; i < a.length; i += 3) if (a[i] === name) return a[i + 2];
    return undefined;
  }

  /** Attribute by namespace and local name (`attrNS(NS.w, 'val')`); `uri` '' = unprefixed. */
  attrNS(uri: string, local: string): string | undefined {
    const a = this._attrs;
    for (let i = 0; i < a.length; i += 3) {
      if (a[i + 1] === uri && hasLocal(a[i]!, local)) return a[i + 2];
    }
    return undefined;
  }

  /** True when the attribute exists. */
  hasAttr(name: string): boolean {
    return this.attr(name) !== undefined;
  }

  /** All attributes (including namespace declarations) in order. */
  attributes(): XAttr[] {
    const a = this._attrs;
    const out: XAttr[] = [];
    for (let i = 0; i < a.length; i += 3)
      out.push({ name: a[i]!, uri: a[i + 1]!, value: a[i + 2]! });
    return out;
  }

  /**
   * Sets an attribute by qualified name. A prefixed name must be resolvable: declared in scope or
   * a conventional OOXML prefix (`w`, `r`, `w14`, ...), whose declaration is added on save.
   * `xmlns` / `xmlns:p` declare namespaces.
   */
  setAttr(name: string, value: string): this {
    const a = this._attrs;
    for (let i = 0; i < a.length; i += 3) {
      if (a[i] === name) {
        if (a[i + 2] !== value) {
          a[i + 2] = value;
          touch(this);
        }
        return this;
      }
    }
    a.push(name, this.#attrUri(name), value);
    touch(this);
    return this;
  }

  /**
   * Sets an attribute by namespace and local name. Updates an existing `{uri}local` attribute
   * whatever its prefix; otherwise uses the prefix bound to `uri` in scope, else the
   * conventional one. `uri` '' = unprefixed.
   */
  setAttrNS(uri: string, local: string, value: string): this {
    const cu = canonicalNs(uri);
    const a = this._attrs;
    for (let i = 0; i < a.length; i += 3) {
      if (a[i + 1] === cu && hasLocal(a[i]!, local)) {
        if (a[i + 2] !== value) {
          a[i + 2] = value;
          touch(this);
        }
        return this;
      }
    }
    let name = local;
    if (cu !== '') {
      const prefix = this.lookupPrefix(cu) ?? conventionalPrefix(cu);
      if (prefix === undefined || prefix === '') {
        throw new UsageError(`No prefix for namespace ${uri}: declare it first`);
      }
      name = `${prefix}:${local}`;
    }
    a.push(name, cu, value);
    touch(this);
    return this;
  }

  /** Removes an attribute by qualified name; returns whether it existed. */
  removeAttr(name: string): boolean {
    const a = this._attrs;
    for (let i = 0; i < a.length; i += 3) {
      if (a[i] === name) {
        a.splice(i, 3);
        touch(this);
        return true;
      }
    }
    return false;
  }

  /** Removes an attribute by namespace and local name; returns whether it existed. */
  removeAttrNS(uri: string, local: string): boolean {
    const a = this._attrs;
    for (let i = 0; i < a.length; i += 3) {
      if (a[i + 1] === uri && hasLocal(a[i]!, local)) {
        a.splice(i, 3);
        touch(this);
        return true;
      }
    }
    return false;
  }

  #attrUri(name: string): string {
    if (name === 'xmlns' || name.startsWith('xmlns:')) return NS.xmlns;
    const c = name.indexOf(':');
    if (c < 0) return '';
    const prefix = name.slice(0, c);
    const uri = this.lookupNamespace(prefix) ?? conventionalUri(prefix);
    if (uri === undefined) throw new UsageError(`Undeclared namespace prefix in ${name}`);
    return canonicalNs(uri);
  }

  // ---- namespaces -----------------------------------------------------------------------

  /** URI bound to `prefix` ('' = default namespace) in scope, as declared (not canonicalized). */
  lookupNamespace(prefix: string): string | undefined {
    if (prefix === 'xml') return NS.xml;
    if (prefix === 'xmlns') return NS.xmlns;
    const decl = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
    let e: XParent | null = this;
    while (e !== null && e.nodeType === ELEMENT_NODE) {
      const a = e._attrs;
      for (let i = 0; i < a.length; i += 3) if (a[i] === decl) return a[i + 2];
      e = e.parent;
    }
    return undefined;
  }

  /**
   * A prefix bound to namespace `uri` (compared canonically) in scope and not shadowed by a
   * closer declaration; '' when it is the default namespace; undefined when none.
   */
  lookupPrefix(uri: string): string | undefined {
    const cu = canonicalNs(uri);
    if (cu === NS.xml) return 'xml';
    const shadowed = new Set<string>();
    let e: XParent | null = this;
    while (e !== null && e.nodeType === ELEMENT_NODE) {
      const a = e._attrs;
      for (let i = 0; i < a.length; i += 3) {
        if (a[i + 1] !== NS.xmlns) continue;
        const n = a[i]!;
        const p = n === 'xmlns' ? '' : n.slice(6);
        if (shadowed.has(p)) continue;
        if (canonicalNs(a[i + 2]!) === cu) return p;
        shadowed.add(p);
      }
      e = e.parent;
    }
    return undefined;
  }

  /** Clone (deep by default), detached. */
  clone(deep = true): XElement {
    const root = new XElement(this.name, this.uri, this._attrs.slice());
    root.name = this.name;
    root.local = this.local;
    if (!deep) return root;
    const stack: [XElement, XElement][] = [[this, root]];
    while (stack.length > 0) {
      const [src, dst] = stack.pop()!;
      const kids = src._children;
      const out: XNode[] = new Array(kids.length);
      for (let i = 0; i < kids.length; i++) {
        const k = kids[i]!;
        let c: XNode;
        if (k.nodeType === ELEMENT_NODE) {
          const e = new XElement(k.name, k.uri, k._attrs.slice());
          stack.push([k, e]);
          c = e;
        } else c = k.clone();
        c.parent = dst;
        out[i] = c;
      }
      dst._children = out;
    }
    return root;
  }
}

export class XDocument {
  readonly nodeType: 9 = DOCUMENT_NODE;
  /** Always null (lets code walk `.parent` up to the document uniformly). */
  readonly parent: null = null;
  /** XML declaration exactly as written (`<?xml version="1.0" ...?>`), or null when absent. */
  declaration: string | null;
  /** The source started with a byte-order mark (re-emitted on save). */
  bom: boolean;
  /** Set by every mutation of the tree; cleared by whoever persists the document. */
  dirty = false;
  /**
   * Mutation counter: bumped by every change made through the DOM methods (never reset). Caches
   * derived from a part (resolved styles, list counters, layout) compare it to stay valid.
   */
  version = 0;
  /** @internal Top-level nodes: prolog/epilog comments, PIs, whitespace text and the root. */
  _children: XNode[];

  constructor(declaration: string | null = null, children?: XNode[], bom = false) {
    this.declaration = declaration;
    this.bom = bom;
    this._children = [];
    if (children) for (const c of children) this.append(c);
    this.dirty = false;
  }

  get children(): readonly XNode[] {
    return this._children;
  }

  /** The document element. Throws when there is none. */
  get root(): XElement {
    for (const k of this._children) if (k.nodeType === ELEMENT_NODE) return k;
    throw new UsageError('XML document has no root element');
  }

  /** Marks the document modified (for changes made outside the DOM methods). */
  markDirty(): void {
    this.dirty = true;
    this.version++;
  }

  append(...nodes: XNode[]): this {
    for (const n of nodes) {
      if (n.nodeType === ELEMENT_NODE && this._children.some((k) => k.nodeType === ELEMENT_NODE)) {
        throw new UsageError('XML document already has a root element');
      }
      insertAtIndex(this, this._children.length, n);
    }
    return this;
  }

  /** Inserts `node` before `ref` (a top-level node), or at the end when `ref` is null. */
  insertBefore(node: XNode, ref: XNode | null): this {
    if (ref === null) return this.append(node);
    if (ref.parent !== this) throw new UsageError('insertBefore(): ref is not a child');
    detachForInsert(node);
    insertAtIndex(this, this._children.indexOf(ref), node);
    return this;
  }

  /** Inserts `node` after `ref` (a top-level node), or at the start when `ref` is null. */
  insertAfter(node: XNode, ref: XNode | null): this {
    if (ref === null) {
      insertAtIndex(this, 0, node);
      return this;
    }
    if (ref.parent !== this) throw new UsageError('insertAfter(): ref is not a child');
    detachForInsert(node);
    insertAtIndex(this, this._children.indexOf(ref) + 1, node);
    return this;
  }

  /**
   * New detached element `{uri}local`, prefixed as the root element binds `uri` (else the
   * conventional prefix; else unprefixed when `uri` is the root's default namespace).
   * `attrs` keys are qualified names (`{ 'w:val': 'Heading1' }`), resolved like
   * {@link XElement.setAttr} against this document's root.
   */
  createElement(uri: string, local: string, attrs?: Readonly<Record<string, string>>): XElement {
    const name = this.qualify(uri, local);
    const el = new XElement(name, uri);
    if (attrs) {
      const root = this.#rootOrNull();
      for (const k in attrs) {
        const v = attrs[k]!;
        el._attrs.push(k, attrUri(root, k), v);
      }
    }
    return el;
  }

  /** Qualified name for `{uri}local` in this document's root scope. */
  qualify(uri: string, local: string): string {
    const cu = canonicalNs(uri);
    if (cu === '') return local;
    const root = this.#rootOrNull();
    const p = root?.lookupPrefix(cu) ?? conventionalPrefix(cu);
    if (p === undefined) throw new UsageError(`No prefix for namespace ${uri}: declare it first`);
    return p === '' ? local : `${p}:${local}`;
  }

  #rootOrNull(): XElement | null {
    for (const k of this._children) if (k.nodeType === ELEMENT_NODE) return k;
    return null;
  }

  /** Deep copy (not dirty). */
  clone(): XDocument {
    const d = new XDocument(this.declaration, undefined, this.bom);
    for (const k of this._children) {
      const c = k.nodeType === ELEMENT_NODE ? k.clone(true) : k.clone();
      c.parent = d;
      d._children.push(c);
    }
    return d;
  }
}

function attrUri(scope: XElement | null, name: string): string {
  if (name === 'xmlns' || name.startsWith('xmlns:')) return NS.xmlns;
  const c = name.indexOf(':');
  if (c < 0) return '';
  const prefix = name.slice(0, c);
  const uri = scope?.lookupNamespace(prefix) ?? conventionalUri(prefix);
  if (uri === undefined) throw new UsageError(`Undeclared namespace prefix in ${name}`);
  return canonicalNs(uri);
}

/**
 * Structural equality of two nodes or documents: same node kinds, names, namespace URIs,
 * attributes (order-insensitive) and text, recursively. Adjacent text nodes are compared merged.
 */
export function domEqual(a: XNode | XDocument, b: XNode | XDocument): boolean {
  const stack: [XNode | XDocument, XNode | XDocument][] = [[a, b]];
  while (stack.length > 0) {
    const [x, y] = stack.pop()!;
    if (x.nodeType !== y.nodeType) return false;
    switch (x.nodeType) {
      case TEXT_NODE:
      case COMMENT_NODE:
        if (x.text !== (y as XText | XComment).text) return false;
        continue;
      case PI_NODE: {
        const p = y as XProcessingInstruction;
        if (x.target !== p.target || x.data !== p.data) return false;
        continue;
      }
      case ELEMENT_NODE: {
        const e = y as XElement;
        if (x.name !== e.name || x.uri !== e.uri) return false;
        if (x._attrs.length !== e._attrs.length) return false;
        for (let i = 0; i < x._attrs.length; i += 3) {
          if (e.attr(x._attrs[i]!) !== x._attrs[i + 2]) return false;
        }
        break;
      }
      default:
        break;
    }
    const xs = mergedChildren((x as XParent)._children);
    const ys = mergedChildren((y as XParent)._children);
    if (xs.length !== ys.length) return false;
    for (let i = 0; i < xs.length; i++) stack.push([xs[i]!, ys[i]!]);
  }
  return true;
}

function mergedChildren(kids: readonly XNode[]): XNode[] {
  let merged: XNode[] | undefined;
  for (let i = 0; i < kids.length; i++) {
    const k = kids[i]!;
    if (k.nodeType === TEXT_NODE && i > 0 && kids[i - 1]!.nodeType === TEXT_NODE) {
      merged ??= kids.slice(0, i);
      const last = merged[merged.length - 1] as XText;
      merged[merged.length - 1] = new XText(last.text + k.text);
    } else if (merged) merged.push(k);
  }
  return merged ?? (kids as XNode[]);
}
