// Hyperlinks: views over `w:hyperlink` and the Hyperlink character style.
//
// A hyperlink is a `w:hyperlink` wrapper around runs. An external link points through `r:id`
// at a relationship with TargetMode External (Word moves a URL's `#fragment` into `w:anchor`);
// an internal link names a bookmark in `w:anchor`. Changing a URL rewrites the relationship's
// Target in place unless other markup shares that relationship, in which case the link gets its
// own: editing one link never retargets another.
//
// Field-based links (`HYPERLINK` complex or simple fields) are fields, not `w:hyperlink`, and
// are left to the fields module.

import type { Part } from '../opc/part.js';
import { CT, REL, relTypeIs } from '../opc/rel-types.js';
import { partDir, relsPathFor } from '../opc/uri.js';
import { XDocument, XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { WordDocument } from './document.js';
import { relTypeFor } from './media.js';
import { setWAttr, wElement, wPrefix } from './order.js';
import type { Story } from './story.js';
import { paragraphText } from './text.js';

/** A hyperlink (`w:hyperlink`) in a story. */
export class Hyperlink {
  /** The `w:hyperlink` element. */
  readonly element: XElement;
  /** The story containing it (its part holds the relationship). */
  readonly story: Story;

  constructor(el: XElement, story: Story) {
    this.element = el;
    this.story = story;
  }

  /** Relationship id (`r:id`) of an external link. */
  get relId(): string | undefined {
    return this.element.attrNS(NS.r, 'id');
  }

  /**
   * Target URL of an external link (relationship target, plus `#anchor` when Word split the
   * fragment off); undefined for a link to a bookmark.
   */
  get url(): string | undefined {
    const id = this.relId;
    if (id === undefined) return undefined;
    const rel = this.story.part.rels.get(id);
    if (rel === undefined) return undefined;
    const anchor = this.element.attrNS(NS.w, 'anchor');
    return rel.external && anchor !== undefined ? `${rel.target}#${anchor}` : rel.target;
  }

  /** Points the link at external `url` (and drops a bookmark anchor). */
  set url(v: string) {
    const el = this.element;
    const part = this.story.part;
    const id = this.relId;
    const rel = id === undefined ? undefined : part.rels.get(id);
    if (id !== undefined && rel?.external === true && references(part, id, el) === 0) {
      setRelTarget(part, id, v);
    } else {
      const added = part.rels.add(relTypeFor(this.story.doc, REL.hyperlink), v, { external: true });
      el.setAttrNS(NS.r, 'id', added.id);
      if (id !== undefined) dropIfUnused(part, id);
    }
    el.removeAttrNS(NS.w, 'anchor');
  }

  /** Bookmark name (`w:anchor`): the target of an internal link, or an external URL's fragment. */
  get anchor(): string | undefined {
    return this.element.attrNS(NS.w, 'anchor');
  }

  /**
   * Makes this a link to bookmark `v` (removing the external relationship reference, and the
   * relationship itself when nothing else uses it); undefined removes the anchor.
   */
  set anchor(v: string | undefined) {
    const el = this.element;
    if (v === undefined) {
      el.removeAttrNS(NS.w, 'anchor');
      return;
    }
    const id = this.relId;
    if (id !== undefined) {
      el.removeAttrNS(NS.r, 'id');
      dropIfUnused(this.story.part, id);
    }
    setWAttr(el, 'anchor', v);
  }

  /** Screen tip (`w:tooltip`). */
  get tooltip(): string | undefined {
    return this.element.attrNS(NS.w, 'tooltip');
  }

  set tooltip(v: string | undefined) {
    if (v === undefined) this.element.removeAttrNS(NS.w, 'tooltip');
    else setWAttr(this.element, 'tooltip', v);
  }

  /** Visible text of the link. */
  get text(): string {
    return paragraphText(this.element);
  }
}

/** Number of `r:*` attributes in `part` equal to `id`, not counting those on `except`. */
function references(part: Part, id: string, except?: XElement): number {
  let n = 0;
  for (const el of part.xml().root.descendants()) {
    if (el === except) continue;
    const a = el._attrs;
    for (let i = 0; i < a.length; i += 3) if (a[i + 1] === NS.r && a[i + 2] === id) n++;
  }
  return n;
}

/** Removes hyperlink relationship `id` of `part` when no markup references it any more. */
function dropIfUnused(part: Part, id: string): void {
  const rel = part.rels.get(id);
  if (rel !== undefined && relTypeIs(rel.type, 'hyperlink') && references(part, id) === 0) {
    part.rels.remove(id);
  }
}

/** Rewrites the Target of relationship `id` of `part` in place (as an external target). */
function setRelTarget(part: Part, id: string, target: string): void {
  const rels = part.package.part(relsPathFor(part.name));
  if (rels === undefined) return;
  for (const el of rels.xml().root.elements()) {
    if (el.attr('Id') !== id) continue;
    el.setAttr('Target', target);
    el.setAttr('TargetMode', 'External');
  }
}

/**
 * @internal Every `w:hyperlink` of `stories`, in document order per story (text boxes
 * included; a Fallback branch that duplicates its Choice is skipped).
 */
export function hyperlinksOf(stories: readonly Story[]): Hyperlink[] {
  const out: Hyperlink[] = [];
  for (const story of stories) {
    for (const el of story.element.descendants(NS.w, 'hyperlink')) {
      const fb = el.closest(NS.mc, 'Fallback');
      if (fb?.parent?.nodeType === 1 && fb.parent.element(NS.mc, 'Choice') !== undefined) continue;
      out.push(new Hyperlink(el, story));
    }
  }
  return out;
}

/** @internal The main part's styles part, created (empty) when the document has none. */
export function stylesPart(doc: WordDocument): Part {
  const found = doc.mainPart.relatedByType('styles')[0];
  if (found !== undefined) return found;
  const root = doc.mainPart.xml().root;
  const w = wPrefix(root) || 'w';
  const uri = root.lookupNamespace(wPrefix(root)) ?? NS.w;
  const styles = new XElement(`${w}:styles`, NS.w, [`xmlns:${w}`, NS.xmlns, uri]);
  const xml = new XDocument('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', [styles]);
  const name = doc.package.uniqueName(`${partDir(doc.mainPart.name)}styles{n}.xml`);
  const part = doc.package.addPart(name, CT.styles, xml);
  doc.mainPart.rels.add(relTypeFor(doc, REL.styles), name.slice(partDir(doc.mainPart.name).length));
  return part;
}

/**
 * @internal The style id of the built-in Hyperlink character style (matched by id or by its
 * English name, so localized ids are found), created when missing: blue (theme hyperlink
 * colour), single underline, based on the default paragraph font.
 */
export function ensureHyperlinkStyle(doc: WordDocument): string {
  const root = stylesPart(doc).xml().root;
  const ids = new Set<string>();
  let defaultFont: string | undefined;
  for (const s of root.elements(NS.w, 'style')) {
    const id = s.attrNS(NS.w, 'styleId');
    if (id !== undefined) ids.add(id);
    if (s.attrNS(NS.w, 'type') !== 'character' || id === undefined) continue;
    const name = s.element(NS.w, 'name')?.attrNS(NS.w, 'val');
    if (id === 'Hyperlink' || name?.toLowerCase() === 'hyperlink') return id;
    const isDefault = s.attrNS(NS.w, 'default');
    if (isDefault === '1' || isDefault === 'true' || isDefault === 'on') defaultFont = id;
  }
  let id = 'Hyperlink';
  for (let n = 1; ids.has(id); n++) id = `Hyperlink${n}`;
  const style = wElement(root, 'style', { type: 'character', styleId: id });
  root.append(style);
  style.append(wElement(style, 'name', { val: 'Hyperlink' }));
  if (defaultFont !== undefined) style.append(wElement(style, 'basedOn', { val: defaultFont }));
  style.append(wElement(style, 'uiPriority', { val: '99' }), wElement(style, 'unhideWhenUsed'));
  const rPr = wElement(style, 'rPr');
  style.append(rPr);
  rPr.append(
    wElement(rPr, 'color', { val: '0563C1', themeColor: 'hyperlink' }),
    wElement(rPr, 'u', { val: 'single' }),
  );
  return id;
}
