// reconcileNamespaces: makes imported markup's namespace prefixes valid in the target part.
//
// Elements and attributes already carry their canonical namespace URI (dom.ts), so only the
// *prefixes* need care. For every imported element and attribute:
// - the prefix becomes the one the target part's root binds to that URI; when the root binds
//   none, one is declared on the root (the conventional prefix when free, else the source's,
//   else a numbered variant), with the target's flavour of URI (Strict twin when the target is
//   Strict OOXML);
// - local `xmlns:*` declarations inside the imported markup are dropped (every prefix now
//   resolves at the root, so they are redundant at best and stale at worst);
// - prefix lists inside markup-compatibility attributes are rewritten to the target's prefixes:
//   `mc:Ignorable`, `mc:MustUnderstand`, `mc:ProcessContent`, `mc:PreserveElements`,
//   `mc:PreserveAttributes` and `mc:Choice/@Requires`. Their source prefixes are resolved in the
//   SOURCE scope (source root + declarations inside the imported markup). Prefixes that do not
//   resolve there are dropped (consumers reject undeclared prefixes in these lists).
// - a namespace the source declared ignorable (its root's `mc:Ignorable`), or one of the
//   extension namespaces of ns.ts IGNORABLE_NS, is made ignorable on the target root too when
//   imported content uses it: Word refuses a part that uses a namespace it does not understand
//   unless it is ignorable.
// The serializer would add missing root declarations on its own (serialize.ts); doing it here
// explicitly also covers the mc lists and keeps the DOM itself consistent.

import type { Part } from '../opc/part.js';
import type { XElement } from '../xml/dom.js';
import { ignorablePrefixes, setIgnorable } from '../xml/mc.js';
import {
  canonicalNs,
  conventionalPrefix,
  IGNORABLE_NS,
  isStrictNs,
  localOf,
  NS,
} from '../xml/ns.js';
import type { ImportContext } from './types.js';

/** mc attributes holding a list of prefixes. */
const PREFIX_LISTS: ReadonlySet<string> = new Set(['Ignorable', 'MustUnderstand']);
/** mc attributes holding a list of qualified names (`p:local` or `p:*`). */
const QNAME_LISTS: ReadonlySet<string> = new Set([
  'ProcessContent',
  'PreserveElements',
  'PreserveAttributes',
]);

/** Strict OOXML twin of a transitional URI (inverse of ns.ts canonicalNs), or undefined. */
function strictTwin(uri: string): string | undefined {
  const m = /^http:\/\/schemas\.openxmlformats\.org\/(.+)\/2006\/(.+)$/.exec(uri);
  if (m === null) return undefined;
  const local = m[2]!.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());
  const cand = `http://purl.oclc.org/ooxml/${m[1]}/${local}`;
  return isStrictNs(cand) && canonicalNs(cand) === uri ? cand : undefined;
}

function splitList(v: string): string[] {
  return v.split(/[ \t\r\n]+/).filter((t) => t !== '');
}

/** prefix → canonical URI for the xmlns declarations among `el`'s attributes, over `outer`. */
function withDeclarations(outer: Map<string, string>, el: XElement): Map<string, string> {
  const a = el._attrs;
  let scope = outer;
  for (let i = 0; i < a.length; i += 3) {
    if (a[i + 1] !== NS.xmlns) continue;
    if (scope === outer) scope = new Map(outer);
    const name = a[i]!;
    const p = name === 'xmlns' ? '' : name.slice(6);
    const v = a[i + 2]!;
    if (v === '') scope.delete(p);
    else scope.set(p, canonicalNs(v));
  }
  return scope;
}

/** Prefix bookkeeping of the target root. */
class TargetScope {
  readonly root: XElement;
  /** prefix → canonical URI ('' = default namespace). */
  readonly prefixes = new Map<string, string>();
  /** canonical URI → named prefix chosen for it. */
  readonly #byUri = new Map<string, string>();
  readonly #strict: boolean;
  /** Prefixes that must be listed in the root's mc:Ignorable. */
  readonly ignorable = new Set<string>();
  /** URIs the source considers ignorable. */
  readonly #sourceIgnorable: ReadonlySet<string>;
  changed = false;

  constructor(root: XElement, sourceIgnorable: ReadonlySet<string>) {
    this.root = root;
    this.#sourceIgnorable = sourceIgnorable;
    let strict = false;
    const a = root._attrs;
    for (let i = 0; i < a.length; i += 3) {
      if (a[i + 1] !== NS.xmlns || a[i + 2] === '') continue;
      const name = a[i]!;
      const uri = canonicalNs(a[i + 2]!);
      if (isStrictNs(a[i + 2]!)) strict = true;
      const p = name === 'xmlns' ? '' : name.slice(6);
      this.prefixes.set(p, uri);
      if (p !== '' && !this.#byUri.has(uri)) this.#byUri.set(uri, p);
    }
    this.#strict = strict;
  }

  /**
   * The prefix to use for `uri` ('' only for elements in the root's default namespace),
   * declaring one on the root when needed; `preferred` is the source's prefix.
   */
  prefixFor(uri: string, preferred: string, element: boolean): string {
    if (uri === NS.xml) return 'xml';
    let p = this.#byUri.get(uri);
    if (p === undefined && element && this.prefixes.get('') === uri) return '';
    if (p === undefined) p = this.#declare(uri, preferred);
    if (this.#sourceIgnorable.has(uri) || IGNORABLE_NS.has(uri)) this.ignorable.add(p);
    return p;
  }

  #declare(uri: string, preferred: string): string {
    const free = (p: string | undefined): p is string =>
      p !== undefined && p !== '' && p !== 'xml' && p !== 'xmlns' && !this.prefixes.has(p);
    // The conventional prefix (`w14`, `wps`, ...) first, then the source's, then `<base><n>`.
    const conv = conventionalPrefix(uri);
    let p: string | undefined = free(conv) ? conv : free(preferred) ? preferred : undefined;
    const base = /^[A-Za-z_][\w.-]*$/.test(preferred) ? preferred : 'ns';
    for (let n = 1; p === undefined; n++) if (free(`${base}${n}`)) p = `${base}${n}`;
    const declared = this.#strict ? (strictTwin(uri) ?? uri) : uri;
    this.root.setAttr(`xmlns:${p}`, declared);
    this.prefixes.set(p, uri);
    this.#byUri.set(uri, p);
    this.changed = true;
    return p;
  }

  /** Lists every needed prefix in the root's mc:Ignorable. */
  commitIgnorable(): void {
    if (this.ignorable.size === 0) return;
    const have = ignorablePrefixes(this.root);
    const add = [...this.ignorable].filter((p) => !have.includes(p));
    if (add.length === 0) return;
    // An mc prefix must exist for the attribute: prefer the root's own.
    if (this.#byUri.get(NS.mc) === undefined) this.prefixFor(NS.mc, 'mc', false);
    setIgnorable(this.root, [...have, ...add]);
    this.changed = true;
  }
}

/**
 * Reconciles the namespace prefixes of imported `blocks` (already inserted into `targetPart`)
 * with the target part's root; see the file header. `sourcePart` (default: the source main
 * part) is the part the blocks were copied from: its root gives the source prefix scope.
 */
export function reconcileNamespaces(
  ctx: ImportContext,
  blocks: XElement[],
  targetPart: Part,
  sourcePart: Part = ctx.source.mainPart,
): void {
  if (blocks.length === 0) return;
  const troot = targetPart.xml().root;
  let sroot: XElement | undefined;
  try {
    sroot = sourcePart.xml().root;
  } catch {
    sroot = undefined;
  }
  const rootScope =
    sroot === undefined ? new Map<string, string>() : withDeclarations(new Map(), sroot);
  const sourceIgnorable = new Set<string>();
  if (sroot !== undefined) {
    for (const p of ignorablePrefixes(sroot)) {
      const uri = rootScope.get(p);
      if (uri !== undefined) sourceIgnorable.add(uri);
    }
  }
  const target = new TargetScope(troot, sourceIgnorable);
  let touched = false;

  const stack: [XElement, Map<string, string>][] = [];
  for (let i = blocks.length - 1; i >= 0; i--) stack.push([blocks[i]!, rootScope]);
  while (stack.length > 0) {
    const [el, outer] = stack.pop()!;
    const scope = withDeclarations(outer, el);
    const kids = el._children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push([k, scope]);
    }
    if (el.uri !== '') {
      const p = target.prefixFor(el.uri, prefixOfName(el.name), true);
      const name = p === '' ? el.local : `${p}:${el.local}`;
      if (name !== el.name) {
        el.name = name;
        touched = true;
      }
    }
    const a = el._attrs;
    if (a.length === 0) continue;
    let dropDecls = false;
    for (let i = 0; i < a.length; i += 3) {
      const uri = a[i + 1]!;
      if (uri === NS.xmlns) {
        dropDecls = true;
        continue;
      }
      const name = a[i]!;
      if (uri === '') {
        if (name === 'Requires' && el.uri === NS.mc && el.local === 'Choice') {
          const v = rewritePrefixes(a[i + 2]!, scope, target);
          if (v !== a[i + 2]) {
            a[i + 2] = v;
            touched = true;
          }
        }
        continue;
      }
      if (uri === NS.xml) continue;
      const local = localOf(name);
      const p = target.prefixFor(uri, prefixOfName(name), false);
      const qname = `${p}:${local}`;
      if (qname !== name) {
        a[i] = qname;
        touched = true;
      }
      if (uri === NS.mc) {
        const v = a[i + 2]!;
        const fixed = PREFIX_LISTS.has(local)
          ? rewritePrefixes(v, scope, target)
          : QNAME_LISTS.has(local)
            ? rewriteQNames(v, scope, target)
            : v;
        if (fixed !== v) {
          a[i + 2] = fixed;
          touched = true;
        }
      }
    }
    if (dropDecls) {
      for (let i = a.length - 3; i >= 0; i -= 3) if (a[i + 1] === NS.xmlns) a.splice(i, 3);
      touched = true;
    }
  }
  target.commitIgnorable();
  if (touched || target.changed) {
    targetPart.xml().markDirty();
    const owner = blocks[0]!.ownerDocument();
    if (owner !== undefined && owner !== targetPart.xml()) owner.markDirty();
  }
}

function prefixOfName(qname: string): string {
  const i = qname.indexOf(':');
  return i < 0 ? '' : qname.slice(0, i);
}

/** A prefix list with each source prefix replaced by the target's (unresolvable ones dropped). */
function rewritePrefixes(v: string, scope: Map<string, string>, target: TargetScope): string {
  const out: string[] = [];
  for (const p of splitList(v)) {
    const uri = scope.get(p);
    if (uri === undefined) continue;
    const tp = target.prefixFor(uri, p, false);
    if (!out.includes(tp)) out.push(tp);
  }
  const joined = out.join(' ');
  return joined === splitList(v).join(' ') ? v : joined;
}

/** A `p:local` / `p:*` list with prefixes mapped like {@link rewritePrefixes}. */
function rewriteQNames(v: string, scope: Map<string, string>, target: TargetScope): string {
  const out: string[] = [];
  for (const q of splitList(v)) {
    const c = q.indexOf(':');
    if (c < 0) continue;
    const uri = scope.get(q.slice(0, c));
    if (uri === undefined) continue;
    const t = `${target.prefixFor(uri, q.slice(0, c), false)}${q.slice(c)}`;
    if (!out.includes(t)) out.push(t);
  }
  const joined = out.join(' ');
  return joined === splitList(v).join(' ') ? v : joined;
}
