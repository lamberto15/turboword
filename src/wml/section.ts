// Sections, their page setup and their effective headers/footers.
//
// WordprocessingML has no section container: a `w:sectPr` inside a paragraph's `w:pPr` ends a
// section (that paragraph is its last one) and the `w:sectPr` that is the body's last child
// describes the final section. Paragraphs wrapped in body-level content controls (`w:sdt`) or
// custom XML still count as body-level.
//
// Headers and footers are inherited ("link to previous"): a section that does not reference a
// header of some type (default/first/even) uses the one of the nearest previous section that
// does. Each type inherits independently. Whether a `first` or `even` header is actually shown
// also depends on `titlePage` and the document-wide `w:evenAndOddHeaders` setting.
//
// Editing keeps that model: removing a section's reference makes it inherit again, unlinking
// gives it a private copy of what it showed, and inserting/removing a section break moves
// references so that every page keeps the header it had. New `w:sectPr` children go in schema
// order (SECTPR_ORDER). A Section view is tied to its `w:sectPr` (the final section to the body's
// last one), so it stays valid across edits; the body scan behind it is cached and refreshed
// whenever the main part's DOM version changes.

import { UsageError } from '../core/errors.js';
import { CT, REL } from '../opc/rel-types.js';
import { partDir, relativeTarget } from '../opc/uri.js';
import type { XDocument, XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { WordDocument } from './document.js';
import {
  copyHeaderFooter,
  createHeaderFooter,
  type HeaderFooter,
  type HeaderFooterKind,
  headerFooterFor,
  newWordRoot,
  releaseReferences,
  relIdFor,
} from './header-footer.js';
import { relTypeFor } from './media.js';
import {
  type ChildOrder,
  ensureChild,
  ensureFirstChild,
  insertOrdered,
  makeOrder,
  PPR_ORDER,
  setWAttr,
  wElement,
} from './order.js';
import { Paragraph } from './paragraph.js';
import { wrappedChildren } from './table.js';

/** How a section starts (`w:type`). */
export type SectionType = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';

/** Header/footer slot of a section (`w:type` of `w:headerReference`/`w:footerReference`). */
export type HeaderFooterType = 'default' | 'first' | 'even';

/** Page orientation (`w:pgSz/@w:orient`). */
export type PageOrientation = 'portrait' | 'landscape';

/** Page size in twips (`w:pgSz`). Width/height are undefined when the attribute is missing. */
export interface PageSize {
  readonly width: number | undefined;
  readonly height: number | undefined;
  readonly orientation: PageOrientation;
}

/** Page margins in twips (`w:pgMar`); missing attributes read as 0. */
export interface PageMargins {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
  readonly header: number;
  readonly footer: number;
  readonly gutter: number;
}

/** `w:sectPr` children in schema order (CT_SectPr, ECMA-376 Part 1 §17.6.17). */
export const SECTPR_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'headerReference',
  'footerReference',
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'w15:footnoteColumns',
  'sectPrChange',
]);

const SECTION_TYPES: ReadonlySet<string> = new Set([
  'nextPage',
  'continuous',
  'evenPage',
  'oddPage',
  'nextColumn',
]);

const HF_TYPES: readonly HeaderFooterType[] = ['default', 'first', 'even'];
const HF_KINDS: readonly HeaderFooterKind[] = ['header', 'footer'];

/** Page margins written when a section gets its first `w:pgMar` (Word's 1" / 0.5" defaults). */
const DEFAULT_MARGINS: PageMargins = {
  top: 1440,
  right: 1440,
  bottom: 1440,
  left: 1440,
  header: 720,
  footer: 720,
  gutter: 0,
};
const MARGIN_KEYS: readonly (keyof PageMargins)[] = [
  'top',
  'right',
  'bottom',
  'left',
  'header',
  'footer',
  'gutter',
];
/** US Letter in twips: the page size assumed when a section specifies none. */
const LETTER = { width: 12240, height: 15840 };

/** Twips per unit of an ST_UniversalMeasure suffix. */
const TWIPS_PER: Readonly<Record<string, number>> = {
  in: 1440,
  cm: 1440 / 2.54,
  mm: 144 / 2.54,
  pt: 20,
  pc: 240,
  pi: 240,
};

/** Parses an ST_(Signed)TwipsMeasure: a twips integer or a universal measure such as `2.5cm`. */
function twips(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const m = /^\s*(-?\d+(?:\.\d+)?)\s*(mm|cm|in|pt|pc|pi)?\s*$/.exec(value);
  if (m === null) return undefined;
  const n = Number(m[1]);
  const unit = m[2];
  return Math.round(unit === undefined ? n : n * TWIPS_PER[unit]!);
}

/** ST_OnOff value of a toggle element: present without `w:val` = on. */
function onOff(el: XElement | undefined): boolean {
  if (el === undefined) return false;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined || !(v === 'false' || v === '0' || v === 'off');
}

/** Rounds a length in twips, rejecting non-finite values (and negative ones unless `signed`). */
function twipsValue(what: string, v: number, signed = false): string {
  if (!Number.isFinite(v) || (!signed && v < 0)) {
    throw new UsageError(`${what} must be a ${signed ? '' : 'non-negative '}number of twips`);
  }
  return String(Math.round(v));
}

/** Removes `el` when it has no attributes and no child elements (never walks further up). */
function pruneIfEmpty(el: XElement | undefined): void {
  if (el !== undefined && el._attrs.length === 0 && el.firstElement() === undefined) el.remove();
}

function checkType(type: HeaderFooterType): void {
  if (!HF_TYPES.includes(type)) {
    throw new UsageError(`Unknown header/footer type ${JSON.stringify(type)}`);
  }
}

// ---- the body scan ----------------------------------------------------------------------------

/** One section as found in the body: its `w:sectPr` and its body-level paragraphs. */
interface SectionEntry {
  /** Undefined only for an implicit final section (no body-level sectPr). */
  readonly sectPr: XElement | undefined;
  readonly paragraphs: readonly XElement[];
}

interface Scan {
  readonly xml: XDocument;
  readonly version: number;
  readonly entries: readonly SectionEntry[];
}

const scans = new WeakMap<WordDocument, Scan>();
const BODY_LEVEL: ReadonlySet<string> = new Set(['p']);

/** The body element of the document (`w:body`), tolerating a story rooted at `w:document`. */
function bodyElement(doc: WordDocument): XElement | undefined {
  const el = doc.body.element;
  return el.is(NS.w, 'document') ? el.element(NS.w, 'body') : el;
}

/** The body's sections (cached until the main part's DOM changes). */
function scanSections(doc: WordDocument): readonly SectionEntry[] {
  const xml = doc.mainPart.xml();
  const hit = scans.get(doc);
  if (hit !== undefined && hit.xml === xml && hit.version === xml.version) return hit.entries;
  const entries: SectionEntry[] = [];
  const body = bodyElement(doc);
  if (body !== undefined) {
    let current: XElement[] = [];
    for (const p of wrappedChildren(body, BODY_LEVEL)) {
      current.push(p);
      const sectPr = p.element(NS.w, 'pPr')?.element(NS.w, 'sectPr');
      if (sectPr !== undefined) {
        entries.push({ sectPr, paragraphs: current });
        current = [];
      }
    }
    entries.push({ sectPr: body.element(NS.w, 'sectPr'), paragraphs: current });
  }
  scans.set(doc, { xml, version: xml.version, entries });
  return entries;
}

/** `w:<kind>Reference` elements of `type` in `sectPr`. */
function refsOf(
  sectPr: XElement | undefined,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): XElement[] {
  if (sectPr === undefined) return [];
  return sectPr
    .elements(NS.w, `${kind}Reference`)
    .filter((r) => (r.attrNS(NS.w, 'type') ?? 'default') === type);
}

/** The reference of `type` in `sectPr`, resolved (hf undefined when it dangles). */
function ownRef(
  doc: WordDocument,
  sectPr: XElement | undefined,
  kind: HeaderFooterKind,
  type: HeaderFooterType,
): { hf: HeaderFooter | undefined } | undefined {
  const ref = refsOf(sectPr, kind, type)[0];
  if (ref === undefined) return undefined;
  const id = ref.attrNS(NS.r, 'id');
  const part = id === undefined ? undefined : doc.mainPart.related(id);
  return { hf: part === undefined ? undefined : headerFooterFor(doc, kind, part) };
}

/** Sets `w:type` of a sectPr (`nextPage`, the default, removes the element). */
function setSectPrType(sectPr: XElement, type: SectionType): void {
  if (!SECTION_TYPES.has(type)) {
    throw new UsageError(`Unknown section type ${JSON.stringify(type)}`);
  }
  const el = sectPr.element(NS.w, 'type');
  if (type === 'nextPage') el?.remove();
  else setWAttr(el ?? ensureChild(sectPr, 'type', SECTPR_ORDER), 'val', type);
}

/** The body-level sectPr, created (empty, as the body's last child) when missing. */
function ensureBodySectPr(doc: WordDocument): XElement {
  const body = bodyElement(doc);
  if (body === undefined) throw new UsageError('The document has no body');
  const found = body.element(NS.w, 'sectPr');
  if (found !== undefined) return found;
  const sectPr = wElement(body, 'sectPr');
  body.append(sectPr);
  return sectPr;
}

// ---- Section ----------------------------------------------------------------------------------

/** A section of the document body. Obtain them from `doc.sections()`. */
export class Section {
  readonly doc: WordDocument;
  /** The `w:sectPr` of a non-final section; undefined for the final section. */
  readonly #sectPr: XElement | undefined;

  /** @internal Use {@link sectionsOf}. `sectPr` undefined = the final section. */
  constructor(doc: WordDocument, sectPr: XElement | undefined) {
    this.doc = doc;
    this.#sectPr = sectPr;
  }

  /** Position in the document (0-based), current as of the latest edit. */
  get index(): number {
    return this.#locate().index;
  }

  /**
   * The `w:sectPr`, or undefined for an implicit final section (body without a final sectPr;
   * the setters create it).
   */
  get element(): XElement | undefined {
    if (this.#sectPr !== undefined) return this.#sectPr;
    return bodyElement(this.doc)?.element(NS.w, 'sectPr');
  }

  #locate(): { entries: readonly SectionEntry[]; index: number } {
    const entries = scanSections(this.doc);
    const own = this.#sectPr;
    const index =
      own === undefined ? entries.length - 1 : entries.findIndex((e) => e.sectPr === own);
    if (index < 0) {
      throw new UsageError('This section no longer exists (its section break was removed)');
    }
    return { entries, index };
  }

  #child(local: string): XElement | undefined {
    return this.element?.element(NS.w, local);
  }

  #ensureSectPr(): XElement {
    return this.element ?? ensureBodySectPr(this.doc);
  }

  // ---- page setup ----

  /** How the section starts (`w:type`); `nextPage` when unspecified. */
  get type(): SectionType {
    const v = this.#child('type')?.attrNS(NS.w, 'val');
    return v !== undefined && SECTION_TYPES.has(v) ? (v as SectionType) : 'nextPage';
  }

  /** Sets how the section starts; `nextPage` (the default) removes `w:type`. */
  set type(v: SectionType) {
    if (this.type === v && SECTION_TYPES.has(v)) return;
    setSectPrType(this.#ensureSectPr(), v);
  }

  /** Page size (`w:pgSz`), or undefined when the section does not specify one. */
  get pageSize(): PageSize | undefined {
    const el = this.#child('pgSz');
    if (el === undefined) return undefined;
    return {
      width: twips(el.attrNS(NS.w, 'w')),
      height: twips(el.attrNS(NS.w, 'h')),
      orientation: el.attrNS(NS.w, 'orient') === 'landscape' ? 'landscape' : 'portrait',
    };
  }

  /**
   * Sets the page size in twips (US Letter is 12240 x 15840, A4 11906 x 16838). With an
   * `orientation`, width and height are swapped when needed to match it; without one, it
   * follows from the dimensions. A stale paper code (`w:code`) is dropped when the size changes.
   */
  setPageSize(width: number, height: number, orientation?: PageOrientation): this {
    let w = Number(twipsValue('Page width', width));
    let h = Number(twipsValue('Page height', height));
    if (w === 0 || h === 0) throw new UsageError('Page width and height must be positive');
    const orient = orientation ?? (w > h ? 'landscape' : 'portrait');
    if (orient !== 'portrait' && orient !== 'landscape') {
      throw new UsageError(`Unknown orientation ${JSON.stringify(orient)}`);
    }
    if ((orient === 'landscape' && w < h) || (orient === 'portrait' && w > h)) [w, h] = [h, w];
    const el = ensureChild(this.#ensureSectPr(), 'pgSz', SECTPR_ORDER);
    const oldW = twips(el.attrNS(NS.w, 'w'));
    const oldH = twips(el.attrNS(NS.w, 'h'));
    const same = (oldW === w && oldH === h) || (oldW === h && oldH === w);
    if (!same) el.removeAttrNS(NS.w, 'code');
    setWAttr(el, 'w', String(w));
    setWAttr(el, 'h', String(h));
    if (orient === 'landscape') setWAttr(el, 'orient', 'landscape');
    else el.removeAttrNS(NS.w, 'orient');
    return this;
  }

  /** Page orientation; portrait when the section specifies no page size. */
  get orientation(): PageOrientation {
    return this.pageSize?.orientation ?? 'portrait';
  }

  /**
   * Sets the orientation, swapping width and height as Word does (a section without a page
   * size gets US Letter).
   */
  set orientation(o: PageOrientation) {
    const size = this.pageSize;
    const w = size?.width ?? LETTER.width;
    const h = size?.height ?? LETTER.height;
    const long = Math.max(w, h);
    const short = Math.min(w, h);
    this.setPageSize(o === 'landscape' ? long : short, o === 'landscape' ? short : long, o);
  }

  /** Page margins (`w:pgMar`), or undefined when the section does not specify them. */
  get margins(): PageMargins | undefined {
    const el = this.#child('pgMar');
    if (el === undefined) return undefined;
    const get = (local: string): number => twips(el.attrNS(NS.w, local)) ?? 0;
    return {
      top: get('top'),
      right: get('right'),
      bottom: get('bottom'),
      left: get('left'),
      header: get('header'),
      footer: get('footer'),
      gutter: get('gutter'),
    };
  }

  /**
   * Sets the given page margins in twips, keeping the others. A section without `w:pgMar` gets
   * one with every (required) attribute, defaulting to 1" margins and 0.5" header/footer.
   */
  setMargins(m: Partial<PageMargins>): this {
    const values = new Map<keyof PageMargins, string>();
    for (const k of MARGIN_KEYS) {
      const v = m[k];
      if (v !== undefined)
        values.set(k, twipsValue(`Margin ${k}`, v, k === 'top' || k === 'bottom'));
    }
    const sectPr = this.#ensureSectPr();
    const existing = sectPr.element(NS.w, 'pgMar');
    const el = existing ?? ensureChild(sectPr, 'pgMar', SECTPR_ORDER);
    for (const k of MARGIN_KEYS) {
      const v = values.get(k) ?? (existing === undefined ? String(DEFAULT_MARGINS[k]) : undefined);
      if (v !== undefined) setWAttr(el, k, v);
    }
    return this;
  }

  /** Number of text columns (`w:cols`); 1 when unspecified. */
  get columns(): number {
    const cols = this.#child('cols');
    if (cols === undefined) return 1;
    const n = Number.parseInt(cols.attrNS(NS.w, 'num') ?? '', 10);
    if (Number.isFinite(n) && n > 0) return n;
    return Math.max(1, cols.elements(NS.w, 'col').length);
  }

  /** Space between columns in twips (`w:cols/@w:space`), or undefined. */
  get columnSpace(): number | undefined {
    return twips(this.#child('cols')?.attrNS(NS.w, 'space'));
  }

  /**
   * Sets `count` equal-width columns, `space` twips apart (kept when omitted; 720 = 0.5" for a
   * section that had none). Individual column widths (`w:col`) are removed.
   */
  setColumns(count: number, space?: number): this {
    if (!Number.isInteger(count) || count < 1) {
      throw new UsageError('Column count must be a positive integer');
    }
    const gap = space === undefined ? undefined : twipsValue('Column space', space);
    const cols = ensureChild(this.#ensureSectPr(), 'cols', SECTPR_ORDER);
    for (const c of cols.elements(NS.w, 'col')) c.remove();
    cols.removeAttrNS(NS.w, 'equalWidth');
    if (count === 1) cols.removeAttrNS(NS.w, 'num');
    else setWAttr(cols, 'num', String(count));
    if (gap !== undefined) setWAttr(cols, 'space', gap);
    else if (cols.attrNS(NS.w, 'space') === undefined) setWAttr(cols, 'space', '720');
    return this;
  }

  /** Different first page header/footer (`w:titlePg`). */
  get titlePage(): boolean {
    return onOff(this.#child('titlePg'));
  }

  /**
   * Turns the different first page on or off. The first page shows the `first` header/footer
   * (blank when the section has none, own or inherited).
   */
  set titlePage(v: boolean) {
    if (v === this.titlePage) return;
    if (v) ensureChild(this.#ensureSectPr(), 'titlePg', SECTPR_ORDER).removeAttrNS(NS.w, 'val');
    else this.#child('titlePg')?.remove();
  }

  /** Page number the section restarts at (`w:pgNumType/@w:start`), or undefined. */
  get pageNumberStart(): number | undefined {
    const v = this.#child('pgNumType')?.attrNS(NS.w, 'start');
    if (v === undefined) return undefined;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n : undefined;
  }

  /** Restarts page numbering at `v`; undefined continues from the previous section. */
  set pageNumberStart(v: number | undefined) {
    if (v === undefined) {
      const el = this.#child('pgNumType');
      el?.removeAttrNS(NS.w, 'start');
      pruneIfEmpty(el);
      return;
    }
    if (!Number.isInteger(v) || v < 0) {
      throw new UsageError('Page number start must be a non-negative integer');
    }
    setWAttr(ensureChild(this.#ensureSectPr(), 'pgNumType', SECTPR_ORDER), 'start', String(v));
  }

  /**
   * Page number format (`w:pgNumType/@w:fmt`: `decimal`, `lowerRoman`, `upperRoman`,
   * `lowerLetter`, `upperLetter`, ...), or undefined (decimal).
   */
  get pageNumberFormat(): string | undefined {
    return this.#child('pgNumType')?.attrNS(NS.w, 'fmt');
  }

  /** Sets the page number format (an ST_NumberFormat value); undefined removes it. */
  set pageNumberFormat(v: string | undefined) {
    if (v === undefined) {
      const el = this.#child('pgNumType');
      el?.removeAttrNS(NS.w, 'fmt');
      pruneIfEmpty(el);
      return;
    }
    if (!/^[A-Za-z]+$/.test(v)) throw new UsageError(`Invalid page number format ${v}`);
    setWAttr(ensureChild(this.#ensureSectPr(), 'pgNumType', SECTPR_ORDER), 'fmt', v);
  }

  // ---- headers and footers ----

  /** The header this section itself references for `type` (no inheritance), or undefined. */
  ownHeader(type: HeaderFooterType = 'default'): HeaderFooter | undefined {
    return ownRef(this.doc, this.element, 'header', type)?.hf;
  }

  /** The footer this section itself references for `type` (no inheritance), or undefined. */
  ownFooter(type: HeaderFooterType = 'default'): HeaderFooter | undefined {
    return ownRef(this.doc, this.element, 'footer', type)?.hf;
  }

  /**
   * The effective header for `type`: this section's own, else the one inherited from the
   * nearest previous section that references a header of that type; undefined when none does.
   */
  header(type: HeaderFooterType = 'default'): HeaderFooter | undefined {
    return this.#effective('header', type);
  }

  /** The effective footer for `type` (see {@link header}). */
  footer(type: HeaderFooterType = 'default'): HeaderFooter | undefined {
    return this.#effective('footer', type);
  }

  /**
   * Makes this section reference `hf` as its `type` header: an existing header of this
   * document, `'new'` for a new empty one, or `null` to remove the section's own reference so
   * it inherits from the previous section again. A part no section references any more is
   * removed from the package. `first` headers show only with {@link titlePage}, `even` ones
   * only with `setEvenAndOddHeaders(doc, true)`. Returns the section's own header afterwards.
   */
  setHeader(type: HeaderFooterType, hf: HeaderFooter | 'new'): HeaderFooter;
  setHeader(type: HeaderFooterType, hf: null): undefined;
  setHeader(type: HeaderFooterType, hf: HeaderFooter | 'new' | null): HeaderFooter | undefined;
  setHeader(type: HeaderFooterType, hf: HeaderFooter | 'new' | null): HeaderFooter | undefined {
    return this.#setRef('header', type, hf);
  }

  /** Footer counterpart of {@link setHeader}. */
  setFooter(type: HeaderFooterType, hf: HeaderFooter | 'new'): HeaderFooter;
  setFooter(type: HeaderFooterType, hf: null): undefined;
  setFooter(type: HeaderFooterType, hf: HeaderFooter | 'new' | null): HeaderFooter | undefined;
  setFooter(type: HeaderFooterType, hf: HeaderFooter | 'new' | null): HeaderFooter | undefined {
    return this.#setRef('footer', type, hf);
  }

  /**
   * Gives this section its own copy of the header (or footer) it currently shows for `type`,
   * so editing it leaves the previous sections alone: the effective one is copied with its
   * relationships (images stay shared), or a new empty one is created when it shows none. A
   * section that already has its own, unshared one keeps it. As in Word, later sections that
   * inherit from this one follow the copy. Without `type`, every type the section shows is
   * unlinked and the `default` one is returned.
   */
  unlinkFromPrevious(kind: HeaderFooterKind, type: HeaderFooterType): HeaderFooter;
  unlinkFromPrevious(kind: HeaderFooterKind, type?: HeaderFooterType): HeaderFooter | undefined;
  unlinkFromPrevious(kind: HeaderFooterKind, type?: HeaderFooterType): HeaderFooter | undefined {
    checkKind(kind);
    if (type !== undefined) return this.#unlink(kind, type);
    for (const t of HF_TYPES) if (this.#effective(kind, t) !== undefined) this.#unlink(kind, t);
    return ownRef(this.doc, this.element, kind, 'default')?.hf;
  }

  /**
   * Removes this section's own header (or footer) references of `type` (all types when
   * omitted), so it shows the previous section's again. On the first section this leaves no
   * header for that type.
   */
  linkToPrevious(kind: HeaderFooterKind, type?: HeaderFooterType): void {
    checkKind(kind);
    for (const t of type === undefined ? HF_TYPES : [type]) this.#setRef(kind, t, null);
  }

  /** Body-level paragraphs belonging to this section (not those inside tables). */
  paragraphs(): Paragraph[] {
    const { entries, index } = this.#locate();
    const body = this.doc.body;
    return entries[index]!.paragraphs.map((p) => new Paragraph(p, body));
  }

  #effective(kind: HeaderFooterKind, type: HeaderFooterType): HeaderFooter | undefined {
    const { entries, index } = this.#locate();
    for (let i = index; i >= 0; i--) {
      const own = ownRef(this.doc, entries[i]!.sectPr, kind, type);
      // A reference defines the slot even when its target is missing: no further inheritance.
      if (own !== undefined) return own.hf;
    }
    return undefined;
  }

  #setRef(
    kind: HeaderFooterKind,
    type: HeaderFooterType,
    value: HeaderFooter | 'new' | null,
  ): HeaderFooter | undefined {
    checkType(type);
    const doc = this.doc;
    if (value === null) {
      const refs = refsOf(this.element, kind, type);
      const ids = refs.map((r) => r.attrNS(NS.r, 'id') ?? '');
      for (const r of refs) r.remove();
      releaseReferences(doc, ids);
      return undefined;
    }
    if (value !== 'new' && value.kind !== kind) {
      throw new UsageError(`A ${value.kind} cannot be used as a ${kind}`);
    }
    const hf = value === 'new' ? createHeaderFooter(doc, kind) : value;
    // Throws for a header of another document (or a removed one) before anything changes.
    const id = relIdFor(doc, hf);
    const sectPr = this.#ensureSectPr();
    const refs = refsOf(sectPr, kind, type);
    const released: string[] = [];
    const first = refs[0];
    if (first === undefined) {
      const ref = insertOrdered(
        sectPr,
        wElement(sectPr, `${kind}Reference`, { type }),
        SECTPR_ORDER,
      );
      ref.setAttrNS(NS.r, 'id', id);
    } else {
      released.push(first.attrNS(NS.r, 'id') ?? '');
      first.setAttrNS(NS.r, 'id', id);
      for (const r of refs.slice(1)) {
        released.push(r.attrNS(NS.r, 'id') ?? '');
        r.remove();
      }
    }
    releaseReferences(
      doc,
      released.filter((x) => x !== id),
    );
    return hf;
  }

  #unlink(kind: HeaderFooterKind, type: HeaderFooterType): HeaderFooter {
    checkType(type);
    const own = ownRef(this.doc, this.element, kind, type);
    if (own?.hf !== undefined && !this.#sharedElsewhere(own.hf)) return own.hf;
    const source = own !== undefined ? own.hf : this.#effective(kind, type);
    const hf = source === undefined ? createHeaderFooter(this.doc, kind) : copyHeaderFooter(source);
    this.#setRef(kind, type, hf);
    return hf;
  }

  /** True when another section references `hf`'s part. */
  #sharedElsewhere(hf: HeaderFooter): boolean {
    const self = this.element;
    const main = this.doc.mainPart;
    for (const e of this.#locate().entries) {
      if (e.sectPr === undefined || e.sectPr === self) continue;
      for (const ref of e.sectPr.elements(NS.w, `${hf.kind}Reference`)) {
        const id = ref.attrNS(NS.r, 'id');
        if (id !== undefined && main.related(id) === hf.part) return true;
      }
    }
    return false;
  }
}

function checkKind(kind: HeaderFooterKind): void {
  if (!HF_KINDS.includes(kind)) throw new UsageError(`Unknown kind ${JSON.stringify(kind)}`);
}

/**
 * The document's sections in order. Each `w:p/w:pPr/w:sectPr` among the body-level paragraphs
 * ends a section; the body's final `w:sectPr` (or an implicit one) describes the last section.
 */
export function sectionsOf(doc: WordDocument): Section[] {
  const entries = scanSections(doc);
  const last = entries.length - 1;
  return entries.map((e, i) => new Section(doc, i === last ? undefined : e.sectPr));
}

// ---- section breaks -----------------------------------------------------------------------------

/** The child of `body` that is or contains `el`. */
function topLevel(body: XElement, el: XElement): XElement | undefined {
  let cur: XElement = el;
  while (cur.parent !== body) {
    const p = cur.parent;
    if (p === null || p.nodeType !== 1) return undefined;
    cur = p;
  }
  return cur;
}

/**
 * Ends a section at `paragraph` (a body-level paragraph): a copy of its section's properties
 * goes into the paragraph's `w:pPr`, so the content after it starts a new section, which
 * starts as `type` says (`nextPage`: on a new page). Every page keeps its header and footer:
 * the new section after the break inherits them from the one ending at `paragraph` (link to
 * previous), and a page-number restart stays with the first of the two. When nothing follows
 * the paragraph in the document, an empty paragraph is added for the new section.
 * Returns the section that now ends at `paragraph`.
 */
export function insertSectionBreak(paragraph: Paragraph, type: SectionType = 'nextPage'): Section {
  const doc = paragraph.story.doc;
  if (paragraph.story.part !== doc.mainPart || paragraph.story.kind !== 'body') {
    throw new UsageError('Section breaks can only be inserted in the document body');
  }
  if (!SECTION_TYPES.has(type)) {
    throw new UsageError(`Unknown section type ${JSON.stringify(type)}`);
  }
  const p = paragraph.element;
  if (p.element(NS.w, 'pPr')?.element(NS.w, 'sectPr') !== undefined) {
    throw new UsageError('The paragraph already ends a section');
  }
  const entries = scanSections(doc);
  const k = entries.findIndex((e) => e.paragraphs.includes(p));
  if (k < 0) {
    throw new UsageError('Section breaks go after a body-level paragraph (not in a table)');
  }
  const original = entries[k]!.sectPr ?? ensureBodySectPr(doc);
  const copy = original.clone();
  for (const c of copy.elements(NS.w, 'sectPrChange')) c.remove();
  // The section after the break is linked to the previous one: references move to the copy.
  for (const kind of HF_KINDS) {
    for (const ref of original.elements(NS.w, `${kind}Reference`)) ref.remove();
  }
  const pgNum = original.element(NS.w, 'pgNumType');
  if (pgNum !== undefined) {
    pgNum.removeAttrNS(NS.w, 'start');
    pruneIfEmpty(pgNum);
  }
  setSectPrType(original, type);
  insertOrdered(ensureFirstChild(p, 'pPr'), copy, PPR_ORDER);

  if (k === entries.length - 1) {
    const body = bodyElement(doc)!;
    const top = topLevel(body, p);
    let next = top?.nextElement();
    while (next?.is(NS.w, 'sectPr')) next = next.nextElement();
    if (next === undefined) doc.body.appendParagraph();
  }
  return new Section(doc, copy);
}

/**
 * Removes the section break that ends `section`, merging it with the next section, whose
 * properties then apply to both (as in Word). Header/footer references the next section was
 * inheriting from this one move to it, so its pages keep their headers; parts no longer
 * referenced are removed. The paragraph that held the break stays. Throws for the last section.
 */
export function removeSectionBreak(section: Section): void {
  const doc = section.doc;
  const entries = scanSections(doc);
  const index = section.index;
  if (index >= entries.length - 1) {
    throw new UsageError('The last section has no section break to remove');
  }
  const sectPr = entries[index]!.sectPr!;
  const nextIsFinal = index + 1 === entries.length - 1;
  let next = entries[index + 1]!.sectPr;
  const released: string[] = [];
  for (const kind of HF_KINDS) {
    for (const type of HF_TYPES) {
      const own = refsOf(sectPr, kind, type);
      for (const r of own) released.push(r.attrNS(NS.r, 'id') ?? '');
      if (own.length === 0 || refsOf(next, kind, type).length > 0) continue;
      next ??= nextIsFinal ? ensureBodySectPr(doc) : undefined;
      if (next !== undefined) insertOrdered(next, own[0]!.clone(), SECTPR_ORDER);
    }
  }
  const pPr = sectPr.parent;
  sectPr.remove();
  if (pPr !== null && pPr.nodeType === 1) pruneIfEmpty(pPr);
  releaseReferences(doc, released);
}

// ---- settings: even and odd headers ---------------------------------------------------------------

/** `w:settings` children in schema order (CT_Settings, §17.15.1.78), enough to place new ones. */
const SETTINGS_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'writeProtection',
  'view',
  'zoom',
  'removePersonalInformation',
  'removeDateAndTime',
  'doNotDisplayPageBoundaries',
  'displayBackgroundShape',
  'printPostScriptOverText',
  'printFractionalCharacterWidth',
  'printFormsData',
  'embedTrueTypeFonts',
  'embedSystemFonts',
  'saveSubsetFonts',
  'saveFormsData',
  'mirrorMargins',
  'alignBordersAndEdges',
  'bordersDoNotSurroundHeader',
  'bordersDoNotSurroundFooter',
  'gutterAtTop',
  'hideSpellingErrors',
  'hideGrammaticalErrors',
  'activeWritingStyle',
  'proofState',
  'formsDesign',
  'attachedTemplate',
  'linkStyles',
  'stylePaneFormatFilter',
  'stylePaneSortMethod',
  'documentType',
  'mailMerge',
  'revisionView',
  'trackRevisions',
  'doNotTrackMoves',
  'doNotTrackFormatting',
  'documentProtection',
  'autoFormatOverride',
  'styleLockTheme',
  'styleLockQFSet',
  'defaultTabStop',
  'autoHyphenation',
  'consecutiveHyphenLimit',
  'hyphenationZone',
  'doNotHyphenateCaps',
  'showEnvelope',
  'summaryLength',
  'clickAndTypeStyle',
  'defaultTableStyle',
  'evenAndOddHeaders',
  'bookFoldRevPrinting',
  'bookFoldPrinting',
  'bookFoldPrintingSheets',
  'drawingGridHorizontalSpacing',
  'drawingGridVerticalSpacing',
  'displayHorizontalDrawingGridEvery',
  'displayVerticalDrawingGridEvery',
  'doNotUseMarginsForDrawingGridOrigin',
  'drawingGridHorizontalOrigin',
  'drawingGridVerticalOrigin',
  'doNotShadeFormData',
  'noPunctuationKerning',
  'characterSpacingControl',
  'printTwoOnOne',
  'strictFirstAndLastChars',
  'noLineBreaksAfter',
  'noLineBreaksBefore',
  'savePreviewPicture',
  'doNotValidateAgainstSchema',
  'saveInvalidXml',
  'ignoreMixedContent',
  'alwaysShowPlaceholderText',
  'doNotDemarcateInvalidXml',
  'saveXmlDataOnly',
  'useXSLTWhenSaving',
  'saveThroughXslt',
  'showXMLTags',
  'alwaysMergeEmptyNamespace',
  'updateFields',
  'hdrShapeDefaults',
  'footnotePr',
  'endnotePr',
  'compat',
  'docVars',
  'rsids',
  'm:mathPr',
  'attachedSchema',
  'themeFontLang',
  'clrSchemeMapping',
  'doNotIncludeSubdocsInStats',
  'doNotAutoCompressPictures',
  'forceUpgrade',
  'captions',
  'readModeInkLockDown',
  'smartTagType',
  'schemaLibrary',
  'shapeDefaults',
  'doNotEmbedSmartTags',
  'decimalSymbol',
  'listSeparator',
]);

/** The schema library namespace (`sl:schemaLibrary` sits among the w: settings). */
const SL = 'http://schemas.openxmlformats.org/schemaLibrary/2006/main';

/**
 * Inserts a new `w:settings` child at its schema position. Unlike {@link insertOrdered}, a
 * foreign element (Word's w14:/w15: extensions, which follow every w: setting) counts as
 * ranked last, so a new setting never lands after them.
 */
function insertSetting(root: XElement, child: XElement): XElement {
  const rank = SETTINGS_ORDER.get(`${NS.w} ${child.local}`) ?? SETTINGS_ORDER.size;
  for (const k of root.children) {
    if (k.nodeType !== 1) continue;
    let r = SETTINGS_ORDER.get(`${k.uri === SL ? NS.w : k.uri} ${k.local}`);
    if (r === undefined && k.uri !== NS.w) r = SETTINGS_ORDER.size;
    if (r !== undefined && r > rank) {
      root.insertBefore(child, k);
      return child;
    }
  }
  root.append(child);
  return child;
}

/** The settings part's root element; with `create`, a missing settings part is added. */
function settingsRoot(doc: WordDocument, create: boolean): XElement | undefined {
  const main = doc.mainPart;
  const part = main.relatedByType('settings')[0];
  if (part !== undefined) {
    return part.xml().children.find((k): k is XElement => k.nodeType === 1);
  }
  if (!create) return undefined;
  const pkg = doc.package;
  const dir = partDir(main.name);
  const name =
    pkg.part(`${dir}settings.xml`) === undefined
      ? `${dir}settings.xml`
      : pkg.uniqueName(`${dir}settings{n}.xml`);
  const xml = newWordRoot(doc, 'settings');
  pkg.addPart(name, CT.settings, xml);
  main.rels.add(relTypeFor(doc, REL.settings), relativeTarget(main.name, name));
  return xml.root;
}

/**
 * Whether even pages show the `even` header/footer (`w:evenAndOddHeaders` in the settings
 * part). Document-wide: it applies to every section.
 */
export function evenAndOddHeaders(doc: WordDocument): boolean {
  return onOff(settingsRoot(doc, false)?.element(NS.w, 'evenAndOddHeaders'));
}

/**
 * Turns different even/odd page headers and footers on or off for the whole document (the
 * settings part is created when missing). When on, even pages show each section's `even`
 * header/footer, blank where a section has none, own or inherited.
 */
export function setEvenAndOddHeaders(doc: WordDocument, on: boolean): void {
  if (evenAndOddHeaders(doc) === on) return;
  if (!on) {
    const root = settingsRoot(doc, false);
    for (const el of root?.elements(NS.w, 'evenAndOddHeaders') ?? []) el.remove();
    return;
  }
  const root = settingsRoot(doc, true)!;
  const el = root.element(NS.w, 'evenAndOddHeaders');
  if (el !== undefined) el.removeAttrNS(NS.w, 'val');
  else insertSetting(root, wElement(root, 'evenAndOddHeaders'));
}
