// XML serializer for the DOM of dom.ts.
//
// Writes exactly what the DOM holds, in Word's own style, so that a part Word wrote comes back
// byte-for-byte in the common case: the XML declaration verbatim, attributes in stored order and
// double quotes, empty elements as `<x/>`, no pretty-printing, whitespace around the root raw.
// Spellings that the DOM does not record are normalized (character references, CDATA delimiters,
// `<x></x>`, single quotes); the result is always DOM-equal to the input.
//
// Whole-document serialization also fixes namespaces up, because edits can introduce elements or
// attributes whose prefix is not declared (`w14:paraId` in a part that never used w14):
// - a prefix not declared in scope is declared on the root, with the node's URI (or its Strict
//   twin when the root already uses Strict URIs); if that URI is an extension namespace
//   (ns.ts IGNORABLE_NS), the prefix is also appended to the root's `mc:Ignorable`;
// - a prefix (or the default namespace) bound to a different URI in scope is re-declared on the
//   element itself;
// - prefixes not declared on the root are dropped from the root's `mc:Ignorable` and
//   `mc:ProcessContent` (consumers reject undeclared prefixes there).
// The fix-up is virtual: the DOM is not modified, only the output. An `xmlns:p=""` attribute
// (illegal in XML 1.0; the parser rejects it, but the DOM can hold one) is never written.
//
// No recursion: the tree is walked with an explicit stack.

import {
  COMMENT_NODE,
  ELEMENT_NODE,
  PI_NODE,
  TEXT_NODE,
  type XDocument,
  type XElement,
  type XNode,
  type XProcessingInstruction,
} from './dom.js';
import { escapeAttr, escapeText } from './escape.js';
import { canonicalNs, hasLocal, IGNORABLE_NS, isStrictNs, NS } from './ns.js';

// Hot constants copied into module scope (read once, not through a live import binding).
const ELEMENT = ELEMENT_NODE;
const TEXT = TEXT_NODE;
const COMMENT = COMMENT_NODE;
const PI = PI_NODE;
const XMLNS: string = NS.xmlns;

/**
 * Serializes a whole document. Before writing, declares on the root element every namespace
 * prefix used by an element/attribute but not declared in scope (and lists extension namespaces
 * from ns.ts IGNORABLE_NS in `mc:Ignorable`). Writes the original XML declaration (encoding
 * rewritten to UTF-8 if it said otherwise) and nothing is pretty-printed.
 */
export function serializeXml(doc: XDocument): string {
  const out: string[] = [];
  if (doc.declaration !== null) out.push(utf8Declaration(doc.declaration));
  for (const k of doc._children) {
    switch (k.nodeType) {
      case ELEMENT:
        writeElement(k, true, out);
        break;
      case TEXT: {
        // Whitespace around the root is written raw (Word's CRLF after the declaration).
        const t = k.text;
        out.push(isWhitespace(t) ? t : escapeText(t));
        break;
      }
      case COMMENT:
        out.push(`<!--${k.text}-->`);
        break;
      case PI:
        out.push(writePi(k));
        break;
    }
  }
  return out.join('');
}

let encoder: TextEncoder | undefined;

/** `serializeXml` encoded as UTF-8 (with a BOM when `doc.bom`). */
export function serializeXmlBytes(doc: XDocument): Uint8Array {
  encoder ??= new TextEncoder();
  const body = encoder.encode(serializeXml(doc));
  if (!doc.bom) return body;
  const out = new Uint8Array(body.length + 3);
  out[0] = 0xef;
  out[1] = 0xbb;
  out[2] = 0xbf;
  out.set(body, 3);
  return out;
}

/** Serializes one node (and its subtree) as a fragment, without any namespace fix-up. */
export function serializeNode(node: XNode): string {
  switch (node.nodeType) {
    case ELEMENT: {
      const out: string[] = [];
      writeElement(node, false, out);
      return out.join('');
    }
    case TEXT:
      return escapeText(node.text);
    case COMMENT:
      return `<!--${node.text}-->`;
    case PI:
      return writePi(node);
  }
}

// ---- writing ------------------------------------------------------------------------------

function writePi(pi: XProcessingInstruction): string {
  return pi.data === '' ? `<?${pi.target}?>` : `<?${pi.target} ${pi.data}?>`;
}

/** `xmlns:p=""`: an illegal prefix undeclaration, omitted from the output. */
function isEmptyPrefixDecl(a: readonly string[], i: number): boolean {
  return a[i + 2] === '' && a[i + 1] === XMLNS && a[i] !== 'xmlns';
}

function writeAttrs(a: readonly string[]): string {
  let s = '';
  for (let i = 0; i < a.length; i += 3) {
    if (!isEmptyPrefixDecl(a, i)) s += ` ${a[i]}="${escapeAttr(a[i + 2]!)}"`;
  }
  return s;
}

function isWhitespace(t: string): boolean {
  for (let i = 0; i < t.length; i++) {
    const c = t.charCodeAt(i);
    if (c !== 32 && c !== 10 && c !== 13 && c !== 9) return false;
  }
  return true;
}

const ENCODING = /(\sencoding\s*=\s*)(["'])([^"']*)\2/;

/** The declaration with a non-UTF-8 `encoding` rewritten to UTF-8 (we always write UTF-8). */
function utf8Declaration(decl: string): string {
  const m = ENCODING.exec(decl);
  if (m === null || /^utf-?8$/i.test(m[3]!)) return decl;
  return `${decl.slice(0, m.index)}${m[1]}${m[2]}UTF-8${m[2]}${decl.slice(m.index + m[0].length)}`;
}

/**
 * Appends element `root` and its subtree to `out`; with `fix`, namespace fix-up rooted at it.
 * Output goes to an array of pieces joined once at the end: building one string with `+=`
 * allocates a cons-string per piece that survives until the end (most of the cost in GC).
 */
function writeElement(root: XElement, fix: boolean, out: string[]): void {
  const fx = fix ? new Fixup(root) : null;
  const headAt = out.length;
  out.push(''); // the root's start tag, written last (the fix-up may add declarations)
  if (root._children.length > 0) {
    const els: XElement[] = [root];
    const pos: number[] = [0];
    const frames: Frame[] = fx === null ? [] : [fx.rootFrame];
    while (els.length > 0) {
      const top = els.length - 1;
      const el = els[top]!;
      const kids = el._children;
      const i = pos[top]!;
      if (i >= kids.length) {
        out.push('</', el.name, '>');
        els.pop();
        pos.pop();
        if (fx !== null) frames.pop();
        continue;
      }
      pos[top] = i + 1;
      const k = kids[i]!;
      switch (k.nodeType) {
        case TEXT:
          out.push(escapeText(k.text));
          break;
        case ELEMENT: {
          out.push('<', k.name);
          const a = k._attrs;
          for (let j = 0; j < a.length; j += 3) {
            if (!isEmptyPrefixDecl(a, j)) out.push(' ', a[j]!, '="', escapeAttr(a[j + 2]!), '"');
          }
          if (fx !== null) {
            const frame = fx.element(k, frames[frames.length - 1]!);
            if (fx.extra !== '') out.push(fx.extra);
            if (k._children.length > 0) frames.push(frame);
          }
          if (k._children.length === 0) out.push('/>');
          else {
            out.push('>');
            els.push(k);
            pos.push(0);
          }
          break;
        }
        case COMMENT:
          out.push('<!--', k.text, '-->');
          break;
        case PI:
          out.push(writePi(k));
          break;
      }
    }
  }
  const attrs = fx === null ? writeAttrs(root._attrs) : fx.rootAttrs();
  out[headAt] = `<${root.name}${attrs}${root._children.length > 0 ? '>' : '/>'}`;
}

// ---- namespace fix-up ---------------------------------------------------------------------

/** Namespace scope while writing (copy-on-declare, like the parser's). */
interface Frame {
  /** prefix → canonical URI. */
  readonly ns: Map<string, string>;
  /** Canonical default namespace ('' when none). */
  readonly def: string;
  /** Memo: qualified names already verified to be bound to this URI in this frame. */
  readonly ok: Map<string, string>;
}

/** New frame with the `xmlns*` declarations among attribute triples `a` applied over `outer`. */
function enter(outer: Frame, a: readonly string[]): Frame {
  const ns = new Map(outer.ns);
  let def = outer.def;
  for (let i = 0; i < a.length; i += 3) {
    if (a[i + 1] !== XMLNS) continue;
    const name = a[i]!;
    const v = canonicalNs(a[i + 2]!);
    if (name === 'xmlns') def = v;
    else {
      const p = name.slice(6);
      if (p === 'xml' || p === 'xmlns') continue;
      if (v === '') ns.delete(p);
      else ns.set(p, v);
    }
  }
  return { ns, def, ok: new Map() };
}

function declares(a: readonly string[]): boolean {
  for (let i = 1; i < a.length; i += 3) if (a[i] === XMLNS) return true;
  return false;
}

/** Strict OOXML twin of a transitional URI (inverse of ns.ts canonicalNs), or undefined. */
function strictTwin(uri: string): string | undefined {
  // ns.ts exposes only Strict → transitional: derive the candidate from the URI pattern
  // (`…/X/2006/y-z` → `http://purl.oclc.org/ooxml/X/yZ`) and keep it only if ns.ts confirms it.
  const m = /^http:\/\/schemas\.openxmlformats\.org\/(.+)\/2006\/(.+)$/.exec(uri);
  if (m === null) return undefined;
  const local = m[2]!.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
  const cand = `http://purl.oclc.org/ooxml/${m[1]}/${local}`;
  return isStrictNs(cand) && canonicalNs(cand) === uri ? cand : undefined;
}

class Fixup {
  readonly root: XElement;
  /** Frame inside the root element. */
  readonly rootFrame: Frame;
  /** Extra declarations for the element passed to the last {@link element} call. */
  extra = '';
  /** Prefixes to declare on the root, in first-use order: prefix → canonical URI. */
  readonly #added = new Map<string, string>();
  /** The root declares Strict OOXML URIs: new declarations use Strict twins too. */
  readonly #strict: boolean;
  /** Declarations the root element itself needs (only ever a default `xmlns`). */
  readonly #rootExtra: string;

  constructor(root: XElement) {
    this.root = root;
    let strict = false;
    const a = root._attrs;
    for (let i = 0; i < a.length; i += 3) {
      if (a[i + 1] === XMLNS && isStrictNs(a[i + 2]!)) strict = true;
    }
    this.#strict = strict;
    // A fresh outermost frame per serialization: frames carry a memo, which must not be shared.
    const base: Frame = { ns: new Map([['xml', NS.xml]]), def: '', ok: new Map() };
    this.rootFrame = this.element(root, base);
    this.#rootExtra = this.extra;
  }

  /** URI to write in a declaration of canonical `uri`. */
  #declared(uri: string): string {
    return this.#strict ? (strictTwin(uri) ?? uri) : uri;
  }

  /**
   * Checks the namespaces used by `el` (whose parent's frame is `outer`); returns the frame for
   * its content and leaves any declarations it needs in {@link extra}.
   */
  element(el: XElement, outer: Frame): Frame {
    this.extra = '';
    const a = el._attrs;
    let f = declares(a) ? enter(outer, a) : outer;
    f = this.#bind(el, f, el.name, el.uri, true);
    for (let i = 0; i < a.length; i += 3) {
      const uri = a[i + 1]!;
      if (uri !== '' && uri !== XMLNS) f = this.#bind(el, f, a[i]!, uri, false);
    }
    return f;
  }

  /** Ensures `qname`, in canonical namespace `uri`, is bound right in `f`; returns the frame. */
  #bind(el: XElement, f: Frame, qname: string, uri: string, isElement: boolean): Frame {
    if (f.ok.get(qname) === uri) return f;
    const ci = qname.indexOf(':');
    if (ci < 0) {
      if (!isElement) return f; // unprefixed attributes are in no namespace
      if (f.def === uri) {
        f.ok.set(qname, uri);
        return f;
      }
      if (el.attr('xmlns') !== undefined) return f; // conflicting own declaration: leave it
      this.extra += ` xmlns="${escapeAttr(this.#declared(uri))}"`;
      return { ns: f.ns, def: uri, ok: new Map() };
    }
    const prefix = qname.slice(0, ci);
    const bound = f.ns.get(prefix) ?? this.#added.get(prefix);
    if (bound === uri) {
      f.ok.set(qname, uri);
      return f;
    }
    if (uri === '' || prefix === 'xml' || prefix === 'xmlns') return f; // cannot be declared
    if (bound === undefined && this.root.attr(`xmlns:${prefix}`) === undefined) {
      this.#added.set(prefix, uri);
      f.ok.set(qname, uri);
      return f;
    }
    // Bound to another namespace here, or undeclared below a root that declares it (through an
    // `xmlns:p=""` in the DOM): re-declare on the element, unless it declares the prefix itself.
    const own = el.attr(`xmlns:${prefix}`);
    if (own !== undefined && own !== '') return f;
    this.extra += ` xmlns:${prefix}="${escapeAttr(this.#declared(uri))}"`;
    const ns = new Map(f.ns);
    ns.set(prefix, uri);
    return { ns, def: f.def, ok: new Map() };
  }

  /** The root's attributes (after the walk): mc lists fixed, missing declarations added. */
  rootAttrs(): string {
    const a = this.root._attrs;
    const declared = new Set<string>(this.#added.keys());
    for (let i = 0; i < a.length; i += 3) {
      if (a[i + 1] === XMLNS && a[i] !== 'xmlns' && a[i + 2] !== '') declared.add(a[i]!.slice(6));
    }
    const ignorable: string[] = [];
    for (const [p, u] of this.#added) if (IGNORABLE_NS.has(u)) ignorable.push(p);

    let s = '';
    let hasIgnorable = false;
    for (let i = 0; i < a.length; i += 3) {
      if (isEmptyPrefixDecl(a, i)) continue;
      const name = a[i]!;
      let v = a[i + 2]!;
      if (a[i + 1] === NS.mc) {
        // A list emptied by the fix-up is dropped; one that was empty already is kept.
        if (hasLocal(name, 'Ignorable')) {
          hasIgnorable = true;
          const fixed = fixIgnorable(v, declared, ignorable);
          if (fixed === '' && v !== '') continue;
          v = fixed;
        } else if (hasLocal(name, 'ProcessContent')) {
          const fixed = fixProcessContent(v, declared);
          if (fixed === '' && v !== '') continue;
          v = fixed;
        }
      }
      s += ` ${name}="${escapeAttr(v)}"`;
    }
    s += this.#rootExtra;
    for (const [p, u] of this.#added) s += ` xmlns:${p}="${escapeAttr(this.#declared(u))}"`;
    if (!hasIgnorable && ignorable.length > 0) {
      let mc: string | undefined;
      for (const [p, u] of this.rootFrame.ns) if (u === NS.mc) mc = p;
      for (const [p, u] of this.#added) if (u === NS.mc) mc ??= p;
      if (mc === undefined) {
        mc = 'mc';
        for (let n = 1; declared.has(mc); n++) mc = `mc${n}`;
        s += ` xmlns:${mc}="${NS.mc}"`;
      }
      s += ` ${mc}:Ignorable="${ignorable.join(' ')}"`;
    }
    return s;
  }
}

function splitList(v: string): string[] {
  return v.split(/[ \t\r\n]+/).filter((t) => t !== '');
}

/** `mc:Ignorable` keeping only declared prefixes, plus `add`; `v` itself when unchanged. */
function fixIgnorable(v: string, declared: ReadonlySet<string>, add: readonly string[]): string {
  const tokens = splitList(v);
  const keep = tokens.filter((p) => declared.has(p));
  for (const p of add) if (!keep.includes(p)) keep.push(p);
  return keep.length === tokens.length && keep.every((p, i) => p === tokens[i])
    ? v
    : keep.join(' ');
}

/** `mc:ProcessContent` (qualified names) keeping only declared prefixes; `v` when unchanged. */
function fixProcessContent(v: string, declared: ReadonlySet<string>): string {
  const tokens = splitList(v);
  const keep = tokens.filter((q) => {
    const c = q.indexOf(':');
    return c < 0 || declared.has(q.slice(0, c));
  });
  return keep.length === tokens.length ? v : keep.join(' ');
}
