// Header and footer parts.
//
// A header/footer is its own part (`word/header1.xml`), related from the main document part and
// referenced by `w:headerReference`/`w:footerReference` in section properties. Several sections
// may reference the same part, so instances are cached per Part: whichever way a header is
// reached (the document's list or a section), callers get the same object.
//
// Editing: `createHeaderFooter` adds a new part the way Word writes one (its namespace
// declarations, one empty paragraph in the Header/Footer style), `copyHeaderFooter` duplicates a
// part with its relationships (images stay shared) and fresh document-unique ids, and
// `releaseReferences` removes a relationship and its part once no section references them. The
// content itself is edited through `HeaderFooter.story` like any other story.

import { UsageError } from '../core/errors.js';
import type { Part } from '../opc/part.js';
import { CT, REL, relTypeIs } from '../opc/rel-types.js';
import { partDir, partKey, relativeTarget } from '../opc/uri.js';
import { XML_DECLARATION } from '../opc/xml-util.js';
import { ELEMENT_NODE, XDocument, XElement, type XNode } from '../xml/dom.js';
import { isStrictNs, NS } from '../xml/ns.js';
import type { WordDocument } from './document.js';
import { freshId, relTypeFor } from './media.js';
import { type ImageOptions, Paragraph, type Run } from './paragraph.js';
import { Story } from './story.js';

/** Which kind of part a {@link HeaderFooter} is. */
export type HeaderFooterKind = 'header' | 'footer';

/** A header or footer part and its content story. */
export class HeaderFooter {
  readonly doc: WordDocument;
  readonly kind: HeaderFooterKind;
  readonly part: Part;
  #story: Story | undefined;

  constructor(doc: WordDocument, kind: HeaderFooterKind, part: Part) {
    this.doc = doc;
    this.kind = kind;
    this.part = part;
  }

  /**
   * The part's content (`w:hdr`/`w:ftr`) as a story. Parsed on first access. Editable like any
   * story (`appendParagraph`, paragraph and run setters, ...).
   */
  get story(): Story {
    const root = this.part.xml().root;
    // Re-created only if the part's DOM was replaced (Part.setXml/setBytes).
    if (this.#story === undefined || this.#story.element !== root) {
      this.#story = new Story(this.doc, this.kind, this.part, root);
    }
    return this.#story;
  }

  /** Plain text of the header/footer. */
  text(): string {
    return this.story.text();
  }

  /**
   * Removes all content, leaving one empty paragraph (in the Header/Footer style when the
   * document defines it), as in a newly created header/footer.
   */
  clear(): void {
    this.story.element.clear();
    appendStoryParagraph(this);
  }

  /**
   * Appends an inline picture (see {@link Paragraph.appendImage}; sizes in points) and returns
   * its run. The run goes into the story's last paragraph when that paragraph has no runs yet
   * (as in a new or cleared header), else into a new paragraph in the Header/Footer style.
   */
  appendImage(bytes: Uint8Array, o?: ImageOptions): Run {
    const story = this.story;
    const last = story.element.lastElement();
    const para =
      last?.is(NS.w, 'p') === true && last.descendants(NS.w, 'r').length === 0
        ? new Paragraph(last, story)
        : appendStoryParagraph(this);
    return para.appendImage(bytes, o);
  }
}

const cache = new WeakMap<Part, HeaderFooter>();

/** @internal The cached HeaderFooter for `part` (created on first use). */
export function headerFooterFor(
  doc: WordDocument,
  kind: HeaderFooterKind,
  part: Part,
): HeaderFooter {
  let hf = cache.get(part);
  if (hf === undefined || hf.doc !== doc) {
    hf = new HeaderFooter(doc, kind, part);
    cache.set(part, hf);
  }
  return hf;
}

/**
 * Every header and footer part related from the main document part, in relationship order,
 * each listed once (external and dangling relationships are skipped).
 */
export function headerFooterParts(doc: WordDocument): HeaderFooter[] {
  const main = doc.mainPart;
  const out: HeaderFooter[] = [];
  const seen = new Set<Part>();
  for (const rel of main.rels.all()) {
    if (rel.external) continue;
    const kind: HeaderFooterKind | undefined = relTypeIs(rel.type, 'header')
      ? 'header'
      : relTypeIs(rel.type, 'footer')
        ? 'footer'
        : undefined;
    if (kind === undefined) continue;
    const part = main.related(rel.id);
    if (part === undefined || seen.has(part)) continue;
    seen.add(part);
    out.push(headerFooterFor(doc, kind, part));
  }
  return out;
}

// ---- creating parts -------------------------------------------------------------------------

/** Namespace declarations of a new header/footer root, as Word writes them. */
const WORD_PREFIXES = [
  'wpc',
  'mc',
  'o',
  'r',
  'm',
  'v',
  'wp14',
  'wp',
  'w10',
  'w',
  'w14',
  'w15',
  'wpg',
  'wpi',
  'wne',
  'wps',
] as const;
/** Prefixes whose namespace has a Strict twin: not added to a Strict document's new parts. */
const STRICT_TWINNED: ReadonlySet<string> = new Set(['w', 'r', 'm', 'wp']);
/** Extension prefixes Word lists in `mc:Ignorable`. */
const IGNORABLE = ['w14', 'w15', 'wp14'];

/**
 * @internal A new XML document whose root is `w:<local>` with the main document root's
 * namespace declarations (so Strict stays Strict and prefixes match), completed with Word's
 * usual set, and an `mc:Ignorable` listing the extension prefixes.
 */
export function newWordRoot(doc: WordDocument, local: string): XDocument {
  const docRoot = doc.mainPart.xml().root;
  const attrs: string[] = [];
  const declared = new Set<string>();
  for (const a of docRoot.attributes()) {
    if (a.uri !== NS.xmlns) continue;
    attrs.push(a.name, NS.xmlns, a.value);
    declared.add(a.name === 'xmlns' ? '' : a.name.slice(6));
  }
  const wDeclared = docRoot.lookupNamespace(docRoot.lookupPrefix(NS.w) ?? 'w') ?? '';
  const strict = isStrictNs(wDeclared);
  for (const p of WORD_PREFIXES) {
    if (declared.has(p) || (strict && STRICT_TWINNED.has(p))) continue;
    attrs.push(`xmlns:${p}`, NS.xmlns, NS[p]);
    declared.add(p);
  }
  const root = new XElement(local, NS.w, attrs);
  const wp = root.lookupPrefix(NS.w);
  // A Strict root that does not declare WordprocessingML itself (declared deeper): add it.
  if (wp === undefined) root._attrs.push('xmlns:w', NS.xmlns, wDeclared || NS.w);
  const prefix = root.lookupPrefix(NS.w) ?? 'w';
  root.setName(prefix === '' ? local : `${prefix}:${local}`, NS.w);
  const mc = root.lookupPrefix(NS.mc);
  if (mc !== undefined && mc !== '') {
    const tokens: string[] = [];
    for (const t of [...(docRoot.attrNS(NS.mc, 'Ignorable') ?? '').split(/\s+/), ...IGNORABLE]) {
      if (t !== '' && declared.has(t) && !tokens.includes(t)) tokens.push(t);
    }
    if (tokens.length > 0) root._attrs.push(`${mc}:Ignorable`, NS.mc, tokens.join(' '));
  }
  return new XDocument(XML_DECLARATION, [root]);
}

/**
 * @internal Style id of the document's built-in Header/Footer paragraph style (matched by its
 * English `w:name`, which Word never localizes, then by the usual style id), or undefined.
 */
export function storyStyleId(doc: WordDocument, kind: HeaderFooterKind): string | undefined {
  const styles = doc.mainPart.relatedByType('styles')[0];
  if (styles === undefined) return undefined;
  const root = styles.xml().children.find((k): k is XElement => k.nodeType === ELEMENT_NODE);
  if (root === undefined) return undefined;
  const usualId = kind === 'header' ? 'Header' : 'Footer';
  let byId: string | undefined;
  for (const s of root.elements(NS.w, 'style')) {
    const type = s.attrNS(NS.w, 'type');
    if (type !== undefined && type !== 'paragraph') continue;
    const id = s.attrNS(NS.w, 'styleId');
    if (id === undefined) continue;
    if (s.element(NS.w, 'name')?.attrNS(NS.w, 'val')?.toLowerCase() === kind) return id;
    if (id === usualId) byId = id;
  }
  return byId;
}

/** Appends an empty paragraph in the Header/Footer style (when defined) to `hf`'s story. */
function appendStoryParagraph(hf: HeaderFooter): Paragraph {
  const style = storyStyleId(hf.doc, hf.kind);
  return hf.story.appendParagraph(undefined, style === undefined ? undefined : { style });
}

/** Adds the main part's relationship to a new header/footer part and returns the part. */
function addStoryPart(
  doc: WordDocument,
  kind: HeaderFooterKind,
  dir: string,
  xml: XDocument,
  contentType: string,
  relType: string,
): Part {
  const pkg = doc.package;
  const name = pkg.uniqueName(`${dir}${kind}{n}.xml`);
  const part = pkg.addPart(name, contentType, xml);
  doc.mainPart.rels.add(relType, relativeTarget(doc.mainPart.name, name));
  return part;
}

/**
 * Creates a new, empty header or footer part (`word/header{n}.xml`): content type, relationship
 * from the main part, Word's namespace declarations, and one empty paragraph in the
 * Header/Footer style when the document defines it. No section references it yet: pass it to
 * `Section.setHeader`/`setFooter`.
 */
export function createHeaderFooter(doc: WordDocument, kind: HeaderFooterKind): HeaderFooter {
  const xml = newWordRoot(doc, kind === 'header' ? 'hdr' : 'ftr');
  const part = addStoryPart(
    doc,
    kind,
    partDir(doc.mainPart.name),
    xml,
    CT[kind],
    relTypeFor(doc, REL[kind]),
  );
  const hf = headerFooterFor(doc, kind, part);
  appendStoryParagraph(hf);
  return hf;
}

/**
 * Duplicates a header/footer part of the same document: same content, same relationships (so
 * images and other targets are shared, not copied) under the same ids, and fresh
 * document-unique drawing, bookmark, content-control and paragraph ids. The copy is related from
 * the main part but referenced by no section.
 */
export function copyHeaderFooter(hf: HeaderFooter): HeaderFooter {
  const doc = hf.doc;
  const src = hf.part;
  if (doc.package.part(src.name) !== src) {
    throw new UsageError(`${hf.kind} part ${src.name} is no longer in the package`);
  }
  const xml = src.xml().clone();
  const mainRel = doc.mainPart.rels
    .all()
    .find(
      (r) => !r.external && r.partName !== undefined && partKey(r.partName) === partKey(src.name),
    );
  const part = addStoryPart(
    doc,
    hf.kind,
    partDir(src.name),
    xml,
    src.contentType ?? CT[hf.kind],
    mainRel?.type ?? relTypeFor(doc, REL[hf.kind]),
  );
  for (const rel of src.rels.all()) {
    const target =
      rel.external || rel.partName === undefined
        ? rel.target
        : relativeTarget(part.name, rel.partName);
    part.rels.add(rel.type, target, { id: rel.id, external: rel.external });
  }
  freshenIds(doc, xml.root);
  return headerFooterFor(doc, hf.kind, part);
}

/**
 * Gives the copied story fresh ids where Word requires document-wide uniqueness: drawing ids
 * (`wp:docPr/@id`), bookmark ids (start/end pairs kept matched) and content-control ids.
 * `w14:paraId` (optional; Word assigns missing ones on save) is dropped rather than renumbered.
 */
function freshenIds(doc: WordDocument, root: XElement): void {
  const bookmarks = new Map<string, string>();
  const stack: XNode[] = [root];
  while (stack.length > 0) {
    const n = stack.pop()!;
    if (n.nodeType !== ELEMENT_NODE) continue;
    const el = n;
    if (el.uri === NS.wp && el.local === 'docPr') {
      if (el.attr('id') !== undefined) el.setAttr('id', String(freshId(doc, 'docPr')));
    } else if (el.uri === NS.w) {
      if (el.local === 'bookmarkStart' || el.local === 'bookmarkEnd') {
        const old = el.attrNS(NS.w, 'id');
        if (old !== undefined) {
          let fresh = bookmarks.get(old);
          if (fresh === undefined) {
            fresh = String(freshId(doc, 'bookmark'));
            bookmarks.set(old, fresh);
          }
          el.setAttrNS(NS.w, 'id', fresh);
        }
      } else if (el.local === 'id' && el.parent?.nodeType === ELEMENT_NODE) {
        if (el.parent.is(NS.w, 'sdtPr') && el.attrNS(NS.w, 'val') !== undefined) {
          el.setAttrNS(NS.w, 'val', String(freshId(doc, 'sdt')));
        }
      }
      el.removeAttrNS(NS.w14, 'paraId');
    }
    const kids = el.children;
    for (let i = kids.length - 1; i >= 0; i--) stack.push(kids[i]!);
  }
}

// ---- references -----------------------------------------------------------------------------

/**
 * @internal Id of a relationship from the main part to `hf`'s part (an existing one when there
 * is one, else a new one). Throws for a header/footer of another document or one whose part was
 * removed.
 */
export function relIdFor(doc: WordDocument, hf: HeaderFooter): string {
  if (hf.doc !== doc) {
    throw new UsageError(
      `The ${hf.kind} belongs to another document (import it with the compose API)`,
    );
  }
  const part = hf.part;
  if (doc.package.part(part.name) !== part) {
    throw new UsageError(`${hf.kind} part ${part.name} is no longer in the package`);
  }
  const main = doc.mainPart;
  const key = partKey(part.name);
  for (const rel of main.rels.all()) {
    if (rel.external || rel.partName === undefined || partKey(rel.partName) !== key) continue;
    if (relTypeIs(rel.type, hf.kind)) return rel.id;
  }
  return main.rels.add(relTypeFor(doc, REL[hf.kind]), relativeTarget(main.name, part.name)).id;
}

/** r:ids used by header/footer references anywhere in the main part (runs are skipped). */
function referencedIds(root: XElement): Set<string> {
  const out = new Set<string>();
  const stack: XElement[] = [root];
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el.uri === NS.w) {
      if (el.local === 'headerReference' || el.local === 'footerReference') {
        const id = el.attrNS(NS.r, 'id');
        if (id !== undefined) out.add(id);
        continue;
      }
      if (el.local === 'r') continue;
    }
    for (const k of el.children) if (k.nodeType === ELEMENT_NODE) stack.push(k);
  }
  return out;
}

/**
 * @internal After references with the given r:ids were removed: drops each relationship that
 * no header/footer reference uses any more, then its header/footer part when no relationship
 * of the main part targets it.
 */
export function releaseReferences(doc: WordDocument, relIds: Iterable<string>): void {
  const main = doc.mainPart;
  let used: Set<string> | undefined;
  for (const id of relIds) {
    used ??= referencedIds(main.xml().root);
    if (used.has(id)) continue;
    const rel = main.rels.get(id);
    if (rel === undefined || rel.external || rel.partName === undefined) continue;
    if (!relTypeIs(rel.type, 'header') && !relTypeIs(rel.type, 'footer')) continue;
    main.rels.remove(id);
    const key = partKey(rel.partName);
    const stillTargeted = main.rels
      .all()
      .some((r) => !r.external && r.partName !== undefined && partKey(r.partName) === key);
    if (!stillTargeted) doc.package.removePart(rel.partName);
  }
}
