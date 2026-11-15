// Schema child order for the property containers we edit, plus helpers that insert children in
// that order.
//
// WordprocessingML property elements (`w:pPr`, `w:rPr`, ...) are XSD sequences: Word rejects a
// file whose children are out of order. Editing therefore never appends blindly: a new child is
// placed before the first existing sibling that the schema puts after it. Siblings we do not
// know (extensions, typos in the source) are left exactly where they are.
//
// Also here: building `w:` elements with whatever prefix the document binds to the
// WordprocessingML namespace (almost always `w`, but Strict files and odd generators exist).

import { XElement } from '../xml/dom.js';
import { conventionalUri, NS } from '../xml/ns.js';

/** Rank of each known child (`uri + ' ' + local` → position in the schema sequence). */
export type ChildOrder = ReadonlyMap<string, number>;

/**
 * Builds a {@link ChildOrder} from local names in schema order. Unprefixed entries are in the
 * WordprocessingML namespace; `w14:glow`-style entries use the conventional prefix's namespace.
 */
export function makeOrder(entries: readonly string[]): ChildOrder {
  const m = new Map<string, number>();
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i]!;
    const c = e.indexOf(':');
    const uri = c < 0 ? NS.w : (conventionalUri(e.slice(0, c)) ?? '');
    m.set(`${uri} ${c < 0 ? e : e.slice(c + 1)}`, i);
  }
  return m;
}

/** `w:pPr` children (CT_PPrBase + CT_PPr, ECMA-376 Part 1 §17.3.1.26). */
export const PPR_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'pStyle',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'framePr',
  'widowControl',
  'numPr',
  'suppressLineNumbers',
  'pBdr',
  'shd',
  'tabs',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'bidi',
  'adjustRightInd',
  'snapToGrid',
  'spacing',
  'ind',
  'contextualSpacing',
  'mirrorIndents',
  'suppressOverlap',
  'jc',
  'textDirection',
  'textAlignment',
  'textboxTightWrap',
  'outlineLvl',
  'divId',
  'cnfStyle',
  'rPr',
  'sectPr',
  'pPrChange',
]);

/**
 * `w:rPr` children (EG_RPrBase, §17.3.2.28), with the Word 2010 `w14:` extensions that follow
 * it. The paragraph-mark variant (CT_ParaRPr) starts with the revision markers `ins`, `del`,
 * `moveFrom`, `moveTo`; they never occur in a run's rPr, so one list serves both.
 */
export const RPR_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'oMath',
  'w14:glow',
  'w14:shadow',
  'w14:reflection',
  'w14:textOutline',
  'w14:textFill',
  'w14:scene3d',
  'w14:props3d',
  'w14:ligatures',
  'w14:numForm',
  'w14:numSpacing',
  'w14:stylisticSets',
  'w14:cntxtAlts',
  'rPrChange',
]);

function rankOf(el: XElement, order: ChildOrder): number | undefined {
  return order.get(`${el.uri} ${el.local}`);
}

/**
 * Inserts `child` into `parent` at its schema position: before the first child element ranked
 * after it. A child the order does not know goes before the final `*Change` revision element
 * (the last entry of every order), else at the end. Returns `child`.
 */
export function insertOrdered(parent: XElement, child: XElement, order: ChildOrder): XElement {
  const rank = rankOf(child, order) ?? order.size - 1.5;
  for (const k of parent.children) {
    if (k.nodeType !== 1) continue;
    const r = rankOf(k, order);
    if (r !== undefined && r > rank) {
      parent.insertBefore(child, k);
      return child;
    }
  }
  parent.append(child);
  return child;
}

/** Prefix bound to the WordprocessingML namespace at `scope` ('' = default namespace; `w` when unbound). */
export function wPrefix(scope: XElement): string {
  return scope.lookupPrefix(NS.w) ?? 'w';
}

/**
 * New detached `w:<local>` element named with the prefix `scope` binds to WordprocessingML.
 * `attrs` keys are local names in that namespace (`{ val: 'Heading1' }` → `w:val`).
 */
export function wElement(
  scope: XElement,
  local: string,
  attrs?: Readonly<Record<string, string>>,
): XElement {
  const p = wPrefix(scope);
  const el = new XElement(p === '' ? local : `${p}:${local}`, NS.w);
  if (attrs) {
    // Attributes need a prefix even when elements use the default namespace.
    const ap = p === '' ? 'w' : p;
    for (const k in attrs) el._attrs.push(`${ap}:${k}`, NS.w, attrs[k]!);
  }
  return el;
}

/** Sets attribute `w:<local>` on `el`, reusing whatever prefix the existing attribute or scope uses. */
export function setWAttr(el: XElement, local: string, value: string): void {
  if (el.attrNS(NS.w, local) !== undefined || (el.lookupPrefix(NS.w) ?? '') !== '') {
    el.setAttrNS(NS.w, local, value);
  } else {
    el.setAttr(`w:${local}`, value);
  }
}

/** The `w:<local>` child of `parent`, created at its schema position when missing. */
export function ensureChild(parent: XElement, local: string, order: ChildOrder): XElement {
  return parent.element(NS.w, local) ?? insertOrdered(parent, wElement(parent, local), order);
}

/**
 * The properties element `w:<local>` of `owner` (`pPr` of a paragraph, `rPr` of a run, ...),
 * which the schema always puts first; created as the first child when missing.
 */
export function ensureFirstChild(owner: XElement, local: string): XElement {
  const found = owner.element(NS.w, local);
  if (found) return found;
  const el = wElement(owner, local);
  const first = owner.firstElement();
  if (first) owner.insertBefore(el, first);
  else owner.append(el);
  return el;
}

/**
 * Removes `el` when it has no child elements, no non-whitespace text and no attributes, then
 * does the same for its parent if that is also a properties element (`*Pr`).
 */
export function pruneEmpty(el: XElement): void {
  let cur: XElement | undefined = el;
  while (cur && cur.parent !== null && cur.parent.nodeType === 1) {
    if (cur._attrs.length > 0 || cur.firstElement() !== undefined || cur.text.trim() !== '') return;
    const parent: XElement = cur.parent;
    cur.remove();
    cur = parent.local.endsWith('Pr') ? parent : undefined;
  }
}
