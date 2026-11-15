// Shared machinery of the style and numbering import steps (compose-styles agent): indexes over
// a document's styles and numbering parts, the basedOn flattening that style comparison and
// keepSourceFormatting expansion are built on, canonical fingerprints, property diffs, and the
// scan that finds which styles and lists a source document uses.
//
// Why flatten: two templates routinely define the same look through different basedOn chains (a
// German template's "Überschrift 1" based on "Standard", an English one's "Heading 1" based on
// "Normal"), and identical styles differ in rsids and UI metadata. Comparing raw XML would call
// them different and copy duplicates; comparing the *resolved* properties does not.
//
// Merge rules (style level, ECMA-376 §17.7.2): a property element of a derived style replaces the
// inherited one, except
//   - attribute-merged elements (spacing, ind, rFonts, lang): attributes override one by one, and
//     exclusive pairs (firstLine/hanging, ascii/asciiTheme, ...) clear each other;
//   - child-merged elements (borders, cell margins, tabs, numPr): children override per side or
//     tab position (`w:tab w:val="clear"` removes an inherited stop).
// Toggle properties (b, i, caps, ...) combine with XOR between the paragraph and the character
// level (§17.7.3); along one basedOn chain they simply override. Lists inside fingerprints are
// compared by their level definitions, never by numId (ids differ between documents).
//
// Everything here reads the DOM and never mutates it: property maps hold references to live
// elements, and every merged or synthesized value is a new detached element.

import type { Part } from '../opc/part.js';
import { relTypeIs } from '../opc/rel-types.js';
import type { WordDocument } from '../wml/document.js';
import { ELEMENT_NODE, XElement } from '../xml/dom.js';
import { localOf, NS } from '../xml/ns.js';

/** Property containers a style (or direct formatting) can carry. */
export type Container = 'pPr' | 'rPr' | 'tblPr' | 'trPr' | 'tcPr';

/** Containers in fingerprint order. */
export const CONTAINERS: readonly Container[] = ['pPr', 'rPr', 'tblPr', 'trPr', 'tcPr'];

/** One level of properties: merge key → property element (never mutated). */
export type PropMap = ReadonlyMap<string, XElement>;

/** Properties per container. */
export type PropSet = Readonly<Record<Container, PropMap>>;

/** A style with its basedOn chain folded in. */
export interface FlatStyle {
  readonly props: PropSet;
  /** Table-style conditional formatting (`w:tblStylePr`) by `w:type`, flattened likewise. */
  readonly cond: ReadonlyMap<string, PropSet>;
}

const EMPTY: PropMap = new Map();
/** A property set with nothing in it. */
export const EMPTY_SET: PropSet = {
  pPr: EMPTY,
  rPr: EMPTY,
  tblPr: EMPTY,
  trPr: EMPTY,
  tcPr: EMPTY,
};
const EMPTY_FLAT: FlatStyle = { props: EMPTY_SET, cond: new Map() };

/** Container children that are not formatting: references, revision records, metadata. */
const SKIP: ReadonlySet<string> = new Set([
  'pStyle',
  'rStyle',
  'tblStyle',
  'sectPr',
  'pPrChange',
  'rPrChange',
  'tblPrChange',
  'trPrChange',
  'tcPrChange',
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tblCaption',
  'tblDescription',
  'cnfStyle',
  'divId',
]);

/** Elements whose attributes merge one by one. */
const ATTR_MERGE: ReadonlySet<string> = new Set(['spacing', 'ind', 'rFonts', 'lang']);

/** Elements whose children merge per key (side, tab position, numPr part). */
const CHILD_MERGE: ReadonlySet<string> = new Set([
  'pBdr',
  'tabs',
  'tblBorders',
  'tblCellMar',
  'tcBorders',
  'tcMar',
  'numPr',
]);

/** Attribute pairs that override each other (setting one clears the other). */
const EXCLUSIVE: ReadonlyMap<string, string> = new Map([
  ['firstLine', 'hanging'],
  ['hanging', 'firstLine'],
  ['firstLineChars', 'hangingChars'],
  ['hangingChars', 'firstLineChars'],
  ['ascii', 'asciiTheme'],
  ['asciiTheme', 'ascii'],
  ['hAnsi', 'hAnsiTheme'],
  ['hAnsiTheme', 'hAnsi'],
  ['eastAsia', 'eastAsiaTheme'],
  ['eastAsiaTheme', 'eastAsia'],
  ['cs', 'cstheme'],
  ['cstheme', 'cs'],
  ['before', 'beforeLines'],
  ['beforeLines', 'before'],
  ['after', 'afterLines'],
  ['afterLines', 'after'],
]);

/** Toggle properties (§17.7.3): XOR between the paragraph and character style levels. */
const TOGGLES: ReadonlySet<string> = new Set([
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
  'vanish',
]);

/** On/off (ST_OnOff `w:val`) property elements of pPr and rPr. */
export const ONOFF: ReadonlySet<string> = new Set([
  ...TOGGLES,
  'noProof',
  'snapToGrid',
  'webHidden',
  'rtl',
  'cs',
  'specVanish',
  'oMath',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'widowControl',
  'suppressLineNumbers',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'bidi',
  'adjustRightInd',
  'contextualSpacing',
  'mirrorIndents',
  'suppressOverlap',
]);

// ---- small helpers ------------------------------------------------------------------------

/** `w:val` of the `w:<local>` child of `el`. */
export function childVal(el: XElement | undefined, local: string): string | undefined {
  return el?.element(NS.w, local)?.attrNS(NS.w, 'val');
}

/** ST_OnOff value of an element (absent `w:val` = on). */
export function onOff(el: XElement | undefined): boolean {
  if (el === undefined) return false;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined || !(v === '0' || v === 'false' || v === 'off');
}

/** ST_OnOff attribute value. */
export function onOffAttr(v: string | undefined): boolean {
  return v !== undefined && (v === '1' || v === 'true' || v === 'on');
}

/** Merge key of a property element (`b`, or `{uri}local` outside WordprocessingML). */
export function keyOf(el: XElement): string {
  return el.uri === NS.w ? el.local : `{${el.uri}}${el.local}`;
}

/** A new detached `w:<local>` element named with `prefix`. */
export function wEl(
  prefix: string,
  local: string,
  attrs?: Readonly<Record<string, string>>,
  kids?: XElement[],
): XElement {
  const el = new XElement(prefix === '' ? local : `${prefix}:${local}`, NS.w);
  if (attrs) {
    const ap = prefix === '' ? 'w' : prefix;
    for (const k in attrs) el._attrs.push(`${ap}:${k}`, NS.w, attrs[k]!);
  }
  if (kids) for (const k of kids) el.append(k);
  return el;
}

function prefixOfName(el: XElement): string {
  const c = el.name.indexOf(':');
  return c < 0 ? '' : el.name.slice(0, c);
}

/** 32-bit FNV-1a hash of a string (deterministic ids and nsids). */
export function fnv1a(s: string): number {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h >>> 0;
}

/** First element child of a part's DOM (undefined for an empty or non-XML part). */
export function rootOf(part: Part | undefined): XElement | undefined {
  if (part === undefined || !part.isXml) return undefined;
  for (const k of part.xml().children) if (k.nodeType === ELEMENT_NODE) return k;
  return undefined;
}

// ---- canonical serialization --------------------------------------------------------------

const HEX6 = /^[0-9a-f]{6}$/i;

function head(el: XElement): string {
  const a = el._attrs;
  const parts: string[] = [];
  let hasVal = false;
  for (let i = 0; i < a.length; i += 3) {
    const uri = a[i + 1]!;
    if (uri === NS.xmlns) continue;
    const local = localOf(a[i]!);
    if (local.startsWith('rsid') || local === 'tplc' || local === 'tentative') continue;
    if (uri === NS.w14 && (local === 'paraId' || local === 'textId')) continue;
    let v = a[i + 2]!;
    if (uri === NS.w && local === 'val') {
      hasVal = true;
      if (ONOFF.has(el.local) && el.uri === NS.w) v = onOff(el) ? '1' : '0';
    }
    if (HEX6.test(v)) v = v.toUpperCase();
    parts.push(`${uri === NS.w || uri === '' ? local : `{${uri}}${local}`}=${v}`);
  }
  if (!hasVal && el.uri === NS.w && ONOFF.has(el.local)) parts.push('val=1');
  parts.sort();
  return `${keyOf(el)}[${parts.join(' ')}]`;
}

/**
 * Order-insensitive canonical form of an element: attributes sorted (rsids dropped, on/off and
 * hex colors normalized), children sorted. `listKey` replaces each `w:numPr` (lists differ in id
 * between documents, so they are compared through their definitions).
 */
export function canon(el: XElement, listKey?: (numPr: XElement) => string): string {
  if (listKey && el.uri === NS.w && el.local === 'numPr') return listKey(el);
  interface Frame {
    el: XElement;
    kids: XElement[];
    i: number;
    out: string[];
  }
  const stack: Frame[] = [{ el, kids: el.elements(), i: 0, out: [] }];
  let result = '';
  while (stack.length > 0) {
    const f = stack[stack.length - 1]!;
    if (f.i < f.kids.length) {
      const k = f.kids[f.i++]!;
      if (listKey && k.uri === NS.w && k.local === 'numPr') f.out.push(listKey(k));
      else stack.push({ el: k, kids: k.elements(), i: 0, out: [] });
      continue;
    }
    stack.pop();
    let text = '';
    for (const c of f.el.children) if (c.nodeType === 3) text += c.text;
    const s = `${head(f.el)}${text.trim() === '' ? '' : `"${text}"`}${f.out.length > 0 ? `(${f.out.sort().join(',')})` : ''}`;
    if (stack.length > 0) stack[stack.length - 1]!.out.push(s);
    else result = s;
  }
  return result;
}

// ---- property maps ------------------------------------------------------------------------

/** One level of properties of a container element (`w:pPr`, `w:rPr`, ...). */
export function propsOf(container: XElement | undefined): PropMap {
  if (container === undefined) return EMPTY;
  const out = new Map<string, XElement>();
  for (const k of container.elements()) {
    const key = keyOf(k);
    if (SKIP.has(key)) continue;
    // `w:rPr` inside a pPr is the paragraph mark's run properties, not a paragraph property.
    if (key === 'rPr' && container.local === 'pPr') continue;
    const prev = out.get(key);
    out.set(key, prev === undefined ? k : mergeOne(prev, k, false));
  }
  return out;
}

function attrTriples(el: XElement): Map<string, [string, string, string]> {
  const m = new Map<string, [string, string, string]>();
  const a = el._attrs;
  for (let i = 0; i < a.length; i += 3) {
    if (a[i + 1] === NS.xmlns) continue;
    const local = localOf(a[i]!);
    const k = a[i + 1] === NS.w ? local : `{${a[i + 1]}}${local}`;
    m.set(k, [a[i]!, a[i + 1]!, a[i + 2]!]);
  }
  return m;
}

function mergeAttrs(base: XElement, top: XElement): XElement {
  const m = attrTriples(base);
  for (const [k, t] of attrTriples(top)) {
    const partner = EXCLUSIVE.get(k);
    if (partner !== undefined) m.delete(partner);
    m.set(k, t);
  }
  const out = new XElement(top.name, top.uri);
  for (const t of m.values()) out._attrs.push(t[0], t[1], t[2]);
  return out;
}

/** Key of a child of a child-merged element (border side, tab position, numPr part). */
function childKey(el: XElement): string {
  if (el.local === 'tab' && el.uri === NS.w) return `tab@${el.attrNS(NS.w, 'pos') ?? ''}`;
  return keyOf(el);
}

function mergeChildren(base: XElement, top: XElement): XElement {
  const m = new Map<string, XElement>();
  for (const k of base.elements()) m.set(childKey(k), k);
  for (const k of top.elements()) {
    const key = childKey(k);
    if (k.local === 'tab' && k.attrNS(NS.w, 'val') === 'clear') m.delete(key);
    else m.set(key, k);
  }
  const out = new XElement(top.name, top.uri, top._attrs.slice());
  for (const k of m.values()) out.append(k.clone());
  return out;
}

function mergeOne(prev: XElement, el: XElement, xor: boolean): XElement {
  const k = keyOf(el);
  if (xor && TOGGLES.has(k)) {
    return wEl(prefixOfName(el), el.local, { val: onOff(prev) !== onOff(el) ? '1' : '0' });
  }
  if (ATTR_MERGE.has(k)) return mergeAttrs(prev, el);
  if (CHILD_MERGE.has(k)) return mergeChildren(prev, el);
  return el;
}

/** `top` over `base`; `xor` combines toggle properties (paragraph → character level). */
export function mergeMaps(base: PropMap, top: PropMap, xor = false): PropMap {
  if (top.size === 0) return base;
  if (base.size === 0) return top;
  const out = new Map(base);
  for (const [k, el] of top) {
    const prev = out.get(k);
    out.set(k, prev === undefined ? el : mergeOne(prev, el, xor));
  }
  return out;
}

function mergeSets(base: PropSet, top: PropSet): PropSet {
  return {
    pPr: mergeMaps(base.pPr, top.pPr),
    rPr: mergeMaps(base.rPr, top.rPr),
    tblPr: mergeMaps(base.tblPr, top.tblPr),
    trPr: mergeMaps(base.trPr, top.trPr),
    tcPr: mergeMaps(base.tcPr, top.tcPr),
  };
}

function setOf(owner: XElement): PropSet {
  return {
    pPr: propsOf(owner.element(NS.w, 'pPr')),
    rPr: propsOf(owner.element(NS.w, 'rPr')),
    tblPr: propsOf(owner.element(NS.w, 'tblPr')),
    trPr: propsOf(owner.element(NS.w, 'trPr')),
    tcPr: propsOf(owner.element(NS.w, 'tcPr')),
  };
}

/** Fingerprint of a flattened style (`listKey` canonicalizes list references). */
export function fingerprint(flat: FlatStyle, listKey: (numPr: XElement) => string): string {
  const one = (set: PropSet): string =>
    CONTAINERS.map((c) => {
      const items: string[] = [];
      for (const el of set[c].values()) items.push(canon(el, listKey));
      return `${c}{${items.sort().join(',')}}`;
    }).join('');
  let s = one(flat.props);
  for (const t of [...flat.cond.keys()].sort()) s += `|${t}:${one(flat.cond.get(t)!)}`;
  return s;
}

// ---- style index --------------------------------------------------------------------------

/** `w:styleId` of a `w:style`. */
export function styleIdOf(s: XElement): string | undefined {
  return s.attrNS(NS.w, 'styleId');
}

/** `w:type` of a `w:style` (paragraph when absent). */
export function styleTypeOf(s: XElement): string {
  return s.attrNS(NS.w, 'type') ?? 'paragraph';
}

/** `w:name/@w:val` of a `w:style`. */
export function styleNameOf(s: XElement): string | undefined {
  return childVal(s, 'name');
}

/** Name key for matching: built-in names are English in every UI language, compared ignoring case. */
export function normName(name: string): string {
  return name.trim().replace(/\s+/g, ' ').toLowerCase();
}

/** The styles of one styles part: lookups by id, name and default, and flattened styles. */
export class StyleIndex {
  /** The styles part (undefined when the document has none). */
  readonly part: Part | undefined;
  /** Its `w:styles` root. */
  readonly root: XElement | undefined;
  readonly byId: Map<string, XElement> = new Map();
  /** Every `w:style`, in part order. */
  readonly order: XElement[] = [];
  /** docDefaults paragraph properties. */
  readonly docPPr: PropMap;
  /** docDefaults run properties. */
  readonly docRPr: PropMap;
  readonly #byName = new Map<string, XElement[]>();
  readonly #defaults = new Map<string, string>();
  readonly #flat = new Map<string, FlatStyle>();

  constructor(part: Part | undefined) {
    this.part = part;
    const root = rootOf(part);
    this.root = root?.is(NS.w, 'styles') === true ? root : undefined;
    const dd = this.root?.element(NS.w, 'docDefaults');
    this.docPPr = propsOf(dd?.element(NS.w, 'pPrDefault')?.element(NS.w, 'pPr'));
    this.docRPr = propsOf(dd?.element(NS.w, 'rPrDefault')?.element(NS.w, 'rPr'));
    if (this.root) for (const s of this.root.elements(NS.w, 'style')) this.add(s);
  }

  /** Registers a style element (after appending it to the part). */
  add(s: XElement): void {
    const id = styleIdOf(s);
    if (id === undefined) return;
    if (!this.byId.has(id)) this.byId.set(id, s);
    this.order.push(s);
    const name = styleNameOf(s);
    if (name !== undefined) {
      const k = normName(name);
      const list = this.#byName.get(k);
      if (list) list.push(s);
      else this.#byName.set(k, [s]);
    }
    const type = styleTypeOf(s);
    if (onOffAttr(s.attrNS(NS.w, 'default')) && !this.#defaults.has(type)) {
      this.#defaults.set(type, id);
    }
    this.#flat.clear();
  }

  /** Style by id. */
  get(id: string | undefined): XElement | undefined {
    return id === undefined ? undefined : this.byId.get(id);
  }

  /** The default style id of a type (`w:default="1"`). */
  defaultId(type: string): string | undefined {
    return this.#defaults.get(type);
  }

  /** First style named `name` (ignoring case), optionally of one type. */
  findByName(name: string, type?: string): XElement | undefined {
    const list = this.#byName.get(normName(name));
    if (!list) return undefined;
    return type === undefined ? list[0] : list.find((s) => styleTypeOf(s) === type);
  }

  /** True when a style of any type has this name (ignoring case). */
  hasName(name: string): boolean {
    return this.#byName.has(normName(name));
  }

  /** Forgets flattened styles (after editing styles in place). */
  invalidate(): void {
    this.#flat.clear();
  }

  /** The style with its basedOn chain folded in (empty for an unknown id). */
  flat(id: string | undefined): FlatStyle {
    if (id === undefined) return EMPTY_FLAT;
    const hit = this.#flat.get(id);
    if (hit) return hit;
    const chain: XElement[] = [];
    const seen = new Set<string>();
    let cur = this.byId.get(id);
    while (cur !== undefined && chain.length < 64) {
      const cid = styleIdOf(cur) ?? '';
      if (seen.has(cid)) break;
      seen.add(cid);
      chain.push(cur);
      cur = this.get(childVal(cur, 'basedOn'));
    }
    let props = EMPTY_SET;
    const cond = new Map<string, PropSet>();
    for (let i = chain.length - 1; i >= 0; i--) {
      const s = chain[i]!;
      props = mergeSets(props, setOf(s));
      for (const c of s.elements(NS.w, 'tblStylePr')) {
        const t = c.attrNS(NS.w, 'type') ?? '';
        cond.set(t, mergeSets(cond.get(t) ?? EMPTY_SET, setOf(c)));
      }
    }
    const flat: FlatStyle = { props, cond };
    this.#flat.set(id, flat);
    return flat;
  }
}

// ---- numbering index ----------------------------------------------------------------------

function intAttr(el: XElement, local: string): number | undefined {
  const v = el.attrNS(NS.w, local);
  if (v === undefined || !/^-?\d+$/.test(v.trim())) return undefined;
  return Number.parseInt(v, 10);
}

/** `w:numId/@w:val` of a `w:numPr` as a number. */
export function numIdOf(numPr: XElement): number | undefined {
  const el = numPr.element(NS.w, 'numId');
  return el ? intAttr(el, 'val') : undefined;
}

/** The lists of one numbering part. */
export class NumberingIndex {
  readonly part: Part | undefined;
  /** Its `w:numbering` root. */
  readonly root: XElement | undefined;
  readonly nums: Map<number, XElement> = new Map();
  readonly abstracts: Map<number, XElement> = new Map();
  readonly picBullets: Map<number, XElement> = new Map();
  readonly nsids: Set<string> = new Set();

  constructor(part: Part | undefined) {
    this.part = part;
    const root = rootOf(part);
    this.root = root?.is(NS.w, 'numbering') === true ? root : undefined;
    if (!this.root) return;
    for (const el of this.root.elements(NS.w)) {
      if (el.local === 'num') this.addNum(el);
      else if (el.local === 'abstractNum') this.addAbstract(el);
      else if (el.local === 'numPicBullet') this.addPicBullet(el);
    }
  }

  addNum(el: XElement): void {
    const id = intAttr(el, 'numId');
    if (id !== undefined && !this.nums.has(id)) this.nums.set(id, el);
  }

  addAbstract(el: XElement): void {
    const id = intAttr(el, 'abstractNumId');
    if (id !== undefined && !this.abstracts.has(id)) this.abstracts.set(id, el);
    const nsid = childVal(el, 'nsid');
    if (nsid !== undefined) this.nsids.add(nsid.toUpperCase());
  }

  addPicBullet(el: XElement): void {
    const id = intAttr(el, 'numPicBulletId');
    if (id !== undefined && !this.picBullets.has(id)) this.picBullets.set(id, el);
  }

  /** `w:abstractNumId` a `w:num` points to. */
  abstractIdOf(num: XElement): number | undefined {
    const el = num.element(NS.w, 'abstractNumId');
    return el ? intAttr(el, 'val') : undefined;
  }

  /** The abstractNum of numId `id`. */
  abstractOfNum(id: number | undefined): XElement | undefined {
    const num = id === undefined ? undefined : this.nums.get(id);
    const a = num ? this.abstractIdOf(num) : undefined;
    return a === undefined ? undefined : this.abstracts.get(a);
  }

  /** The abstractNum defining the levels of numId `id`, following `w:numStyleLink` to the list style. */
  definitionOfNum(id: number | undefined, styles: StyleIndex): XElement | undefined {
    let abs = this.abstractOfNum(id);
    for (let hops = 0; abs !== undefined && hops < 4; hops++) {
      const link = childVal(abs, 'numStyleLink');
      if (link === undefined) break;
      const numPr = styles.get(link)?.element(NS.w, 'pPr')?.element(NS.w, 'numPr');
      const next = this.abstractOfNum(numPr ? numIdOf(numPr) : undefined);
      if (next === undefined || next === abs) break;
      abs = next;
    }
    return abs;
  }

  /** Smallest unused numId (≥ 1, above every existing one). */
  nextNumId(): number {
    let max = 0;
    for (const k of this.nums.keys()) if (k > max) max = k;
    return max + 1;
  }

  /** Smallest abstractNumId above every existing one (≥ 0). */
  nextAbstractId(): number {
    let max = -1;
    for (const k of this.abstracts.keys()) if (k > max) max = k;
    return max + 1;
  }

  /** numPicBulletId above every existing one (≥ 0). */
  nextPicBulletId(): number {
    let max = -1;
    for (const k of this.picBullets.keys()) if (k > max) max = k;
    return max + 1;
  }
}

/** Canonical form of a `w:lvl` (its `w:pStyle` link mapped through `pStyle`, or dropped). */
export function canonLevel(lvl: XElement, pStyle?: (id: string) => string | undefined): string {
  const link = lvl.element(NS.w, 'pStyle');
  if (link === undefined) return canon(lvl);
  const copy = lvl.clone();
  const l = copy.element(NS.w, 'pStyle')!;
  const v = pStyle ? pStyle(l.attrNS(NS.w, 'val') ?? '') : undefined;
  if (v === undefined) l.remove();
  else l.setAttrNS(NS.w, 'val', v);
  return canon(copy);
}

/**
 * Fingerprint of an abstractNum's level definitions: multiLevelType and canonical levels; nsid,
 * tmpl, name and style links are identity, not formatting, and are left out.
 */
export function abstractFingerprint(
  abs: XElement | undefined,
  pStyle?: (id: string) => string | undefined,
): string {
  if (abs === undefined) return '?';
  const levels = abs.elements(NS.w, 'lvl').map((l) => canonLevel(l, pStyle));
  return `${childVal(abs, 'multiLevelType') ?? ''}|${levels.sort().join(';')}`;
}

/** Fingerprint of list numId `id` (definition levels plus the num's level overrides). */
export function listFingerprint(nums: NumberingIndex, styles: StyleIndex, id: number): string {
  const num = nums.nums.get(id);
  if (num === undefined) return `?${id}`;
  const def = abstractFingerprint(nums.definitionOfNum(id, styles));
  const ovr = num
    .elements(NS.w, 'lvlOverride')
    .map((o) => canon(o))
    .sort();
  return `${def}|${ovr.join(';')}`;
}

/** A `listKey` for {@link canon}: numPr compared by list definition and level. */
export function listKeyFor(nums: NumberingIndex, styles: StyleIndex): (numPr: XElement) => string {
  const cache = new Map<number, string>();
  return (numPr) => {
    const ilvl = childVal(numPr, 'ilvl') ?? '';
    const id = numIdOf(numPr);
    if (id === undefined) return `numPr[${ilvl}]`;
    if (id === 0) return 'numPr[none]';
    let fp = cache.get(id);
    if (fp === undefined) {
      fp = listFingerprint(nums, styles, id);
      cache.set(id, fp);
    }
    return `numPr[${ilvl}|${fp}]`;
  };
}

// ---- diffs (keepSourceFormatting) ---------------------------------------------------------

/** Value Word assumes for an attribute nobody sets (attribute-merged elements). */
function attrDefault(container: Container, key: string, attr: string): string | undefined {
  if (key === 'ind') return attr.endsWith('Chars') || !/^[a-z]+$/i.test(attr) ? undefined : '0';
  if (key === 'spacing') {
    if (container === 'rPr') return attr === 'val' ? '0' : undefined;
    if (attr === 'line') return '240';
    if (attr === 'lineRule') return 'auto';
    if (/^(before|after)(Lines|Autospacing)?$/.test(attr)) return '0';
    return undefined;
  }
  if (key === 'rFonts' && /^(ascii|hAnsi|eastAsia|cs)$/.test(attr)) return 'Times New Roman';
  return undefined;
}

/** Element that undoes an inherited property nobody on the source side sets. */
function resetOf(el: XElement): XElement | undefined {
  const p = prefixOfName(el);
  const k = keyOf(el);
  if (el.uri === NS.w && ONOFF.has(k)) return wEl(p, k, { val: '0' });
  switch (k) {
    case 'numPr':
      return wEl(p, 'numPr', undefined, [wEl(p, 'numId', { val: '0' })]);
    case 'shd':
      return wEl(p, 'shd', { val: 'clear', color: 'auto', fill: 'auto' });
    case 'jc':
      return wEl(p, 'jc', { val: 'left' });
    case 'outlineLvl':
      return wEl(p, 'outlineLvl', { val: '9' });
    case 'textAlignment':
      return wEl(p, 'textAlignment', { val: 'auto' });
    case 'textDirection':
      return wEl(p, 'textDirection', { val: 'lrTb' });
    case 'color':
      return wEl(p, 'color', { val: 'auto' });
    case 'sz':
    case 'szCs':
      return wEl(p, k, { val: '20' });
    case 'u':
    case 'highlight':
    case 'effect':
    case 'em':
      return wEl(p, k, { val: 'none' });
    case 'bdr':
      return wEl(p, k, { val: 'nil' });
    case 'vertAlign':
      return wEl(p, k, { val: 'baseline' });
    case 'position':
    case 'kern':
      return wEl(p, k, { val: '0' });
    case 'w':
      return wEl(p, k, { val: '100' });
    default:
      return undefined;
  }
}

function attrDiff(container: Container, s?: XElement, t?: XElement): XElement | undefined {
  const like = (s ?? t)!;
  const key = keyOf(like);
  const sa = s ? attrTriples(s) : new Map<string, [string, string, string]>();
  const ta = t ? attrTriples(t) : new Map<string, [string, string, string]>();
  const out = new XElement(like.name, like.uri);
  for (const [k, v] of sa) {
    if (k.startsWith('rsid')) continue;
    if (ta.get(k)?.[2] !== v[2]) out._attrs.push(v[0], v[1], v[2]);
  }
  const prefix = prefixOfName(like) === '' ? 'w' : prefixOfName(like);
  for (const [k, v] of ta) {
    if (sa.has(k) || k.startsWith('rsid')) continue;
    const partner = EXCLUSIVE.get(k);
    if (partner !== undefined && sa.has(partner)) continue;
    // An inherited theme font is overridden by the explicit slot it pairs with.
    const explicit = k.endsWith('Theme') ? k.slice(0, -5) : k === 'cstheme' ? 'cs' : undefined;
    if (explicit !== undefined) {
      if (sa.has(explicit) || out.attrNS(NS.w, explicit) !== undefined) continue;
      const d = attrDefault(container, key, explicit);
      if (d !== undefined) out._attrs.push(`${prefix}:${explicit}`, NS.w, d);
      continue;
    }
    const d = attrDefault(container, key, k);
    if (d !== undefined && out.attrNS(NS.w, k) === undefined) out._attrs.push(v[0], v[1], d);
  }
  return out._attrs.length > 0 ? out : undefined;
}

function childDiff(s?: XElement, t?: XElement): XElement | undefined {
  const like = (s ?? t)!;
  const sk = new Map<string, XElement>();
  const tk = new Map<string, XElement>();
  if (s) for (const k of s.elements()) sk.set(childKey(k), k);
  if (t) for (const k of t.elements()) tk.set(childKey(k), k);
  const kids: XElement[] = [];
  for (const [k, el] of sk) {
    const other = tk.get(k);
    if (other === undefined || canon(other) !== canon(el)) kids.push(el.clone());
  }
  const p = prefixOfName(like);
  for (const [k, el] of tk) {
    if (sk.has(k)) continue;
    if (el.local === 'tab') {
      kids.push(wEl(p, 'tab', { val: 'clear', pos: el.attrNS(NS.w, 'pos') ?? '0' }));
    } else if (like.local.endsWith('Borders') || like.local === 'pBdr') {
      kids.push(wEl(p, el.local, { val: 'nil' }));
    }
  }
  if (kids.length === 0) return undefined;
  const out = new XElement(like.name, like.uri, (s ?? like)._attrs.slice());
  for (const k of kids) out.append(k);
  return out;
}

/**
 * Property elements that, set as direct formatting over `tgt`, give `src`: what differs, plus
 * resets for what only the target side sets. Lists compare through their own side's key.
 */
export function diffMaps(
  container: Container,
  src: PropMap,
  tgt: PropMap,
  srcList: (numPr: XElement) => string,
  tgtList: (numPr: XElement) => string,
): XElement[] {
  const out: XElement[] = [];
  const keys = new Set<string>([...src.keys(), ...tgt.keys()]);
  for (const k of keys) {
    const s = src.get(k);
    const t = tgt.get(k);
    if (ONOFF.has(k)) {
      if (onOff(s) === onOff(t)) continue;
      out.push(s ? s : resetOf(t!)!);
      continue;
    }
    if (s && t && canon(s, srcList) === canon(t, tgtList)) continue;
    let d: XElement | undefined;
    if (k === 'numPr') d = s ?? resetOf(t!);
    else if (ATTR_MERGE.has(k)) d = attrDiff(container, s, t);
    else if (CHILD_MERGE.has(k)) d = childDiff(s, t);
    else d = s ?? resetOf(t!);
    if (d !== undefined) out.push(d);
  }
  return out;
}

/**
 * Merges what `diff` sets and `existing` (direct formatting, which wins) lacks into `existing`:
 * missing attributes of attribute-merged elements, missing children of child-merged ones.
 */
export function fillMissing(existing: XElement, diff: XElement): void {
  const k = keyOf(diff);
  if (ATTR_MERGE.has(k)) {
    const have = attrTriples(existing);
    for (const [ak, t] of attrTriples(diff)) {
      if (have.has(ak)) continue;
      const partner = EXCLUSIVE.get(ak);
      if (partner !== undefined && have.has(partner)) continue;
      existing._attrs.push(t[0], t[1], t[2]);
      existing.ownerDocument()?.markDirty();
    }
  } else if (CHILD_MERGE.has(k)) {
    const have = new Set(existing.elements().map(childKey));
    for (const c of diff.elements()) if (!have.has(childKey(c))) existing.append(c.clone());
  }
}

// ---- what a source document uses ----------------------------------------------------------

/** Styles and lists referenced by content. */
export interface Usage {
  /** pStyle/rStyle/tblStyle values, first-use order. */
  readonly styles: Set<string>;
  /** numIds (numPr/numId values), first-use order (0 included). */
  readonly numIds: Set<number>;
  /** A paragraph without pStyle exists (it uses the default paragraph style). */
  bareParagraph: boolean;
  /** A table without tblStyle exists (it uses the default table style). */
  bareTable: boolean;
}

/** An empty {@link Usage}. */
export function newUsage(): Usage {
  return { styles: new Set(), numIds: new Set(), bareParagraph: false, bareTable: false };
}

/** Adds what `roots` (and their descendants) reference to `u`. */
export function scanUsage(roots: Iterable<XElement>, u: Usage): void {
  for (const root of roots) {
    const all = root.descendants(NS.w);
    if (root.uri === NS.w) all.unshift(root);
    for (const el of all) {
      switch (el.local) {
        case 'pStyle':
        case 'rStyle':
        case 'tblStyle': {
          const v = el.attrNS(NS.w, 'val');
          if (v !== undefined) u.styles.add(v);
          break;
        }
        case 'p':
          if (el.element(NS.w, 'pPr')?.element(NS.w, 'pStyle') === undefined)
            u.bareParagraph = true;
          break;
        case 'tbl':
          if (el.element(NS.w, 'tblPr')?.element(NS.w, 'tblStyle') === undefined)
            u.bareTable = true;
          break;
        case 'numPr': {
          const id = numIdOf(el);
          if (id !== undefined) u.numIds.add(id);
          break;
        }
        default:
          break;
      }
    }
  }
}

const STORY_RELS = ['header', 'footer', 'footnotes', 'endnotes', 'comments'] as const;

/** Root elements of every story part of `doc`: main document, headers/footers, notes, comments. */
export function storyRoots(doc: WordDocument): XElement[] {
  const out: XElement[] = [];
  const seen = new Set<Part>();
  const main = doc.mainPart;
  const mainRoot = rootOf(main);
  if (mainRoot) out.push(mainRoot);
  seen.add(main);
  for (const rel of main.rels.all()) {
    if (rel.external || !STORY_RELS.some((k) => relTypeIs(rel.type, k))) continue;
    const part = main.related(rel.id);
    if (part === undefined || seen.has(part)) continue;
    seen.add(part);
    const root = rootOf(part);
    if (root) out.push(root);
  }
  return out;
}

/** The first part related from the main document part by relationship type `kind`. */
export function relatedPart(doc: WordDocument, kind: string): Part | undefined {
  return doc.mainPart.relatedByType(kind)[0];
}
