// Document properties: core (docProps/core.xml), extended "app" (docProps/app.xml) and custom
// (docProps/custom.xml).
//
// Why a view and not a model: the property parts are small, but producers put all sorts of extra
// elements in them (thumbnails, HeadingPairs, vendor fields). Reading and writing straight through
// the part DOM keeps everything else byte-for-byte, and every change goes through DOM methods so
// the part is marked dirty and re-serialized on save. Setting a value equal to the current one
// changes nothing, so a no-op edit keeps the package byte-identical.
//
// Parts are found through the package relationships (as Word does). The first write creates the
// part, the package relationship and the content-type override, each only when missing. Reads
// never create anything.

import { UsageError } from '../core/errors.js';
import type { OpcPackage } from '../opc/package.js';
import type { Part } from '../opc/part.js';
import { CT, REL, relTypeIs } from '../opc/rel-types.js';
import { TEXT_NODE, XDocument, XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';

/** Extended ("app") properties, from docProps/app.xml. Numbers are parsed; absent = omitted. */
export interface AppProperties {
  application?: string;
  appVersion?: string;
  company?: string;
  template?: string;
  /** Page count as last laid out by the producing application (Word keeps it up to date). */
  pages?: number;
  words?: number;
  characters?: number;
  charactersWithSpaces?: number;
  lines?: number;
  paragraphs?: number;
  /** Total editing time in minutes. */
  totalTime?: number;
  docSecurity?: number;
}

/** A custom property value: `vt:lpwstr`, `vt:i4`/`vt:r8`, `vt:bool` or `vt:filetime`. */
export type CustomPropertyValue = string | number | boolean | Date;

/** Names of the text-valued core properties. */
export type CoreTextProperty =
  | 'title'
  | 'subject'
  | 'creator'
  | 'keywords'
  | 'description'
  | 'lastModifiedBy'
  | 'revision'
  | 'category'
  | 'contentStatus'
  | 'language'
  | 'identifier'
  | 'version';

/** Names of the date-valued core properties. */
export type CoreDateProperty = 'created' | 'modified' | 'lastPrinted';

type CoreProperty = CoreTextProperty | CoreDateProperty;

/** Core property → element, in the order Word writes them (new elements are inserted in it). */
const CORE: Readonly<Record<CoreProperty, readonly [string, string]>> = {
  title: [NS.dc, 'title'],
  subject: [NS.dc, 'subject'],
  creator: [NS.dc, 'creator'],
  keywords: [NS.cp, 'keywords'],
  description: [NS.dc, 'description'],
  lastModifiedBy: [NS.cp, 'lastModifiedBy'],
  revision: [NS.cp, 'revision'],
  lastPrinted: [NS.cp, 'lastPrinted'],
  created: [NS.dcterms, 'created'],
  modified: [NS.dcterms, 'modified'],
  category: [NS.cp, 'category'],
  contentStatus: [NS.cp, 'contentStatus'],
  language: [NS.dc, 'language'],
  identifier: [NS.dc, 'identifier'],
  version: [NS.cp, 'version'],
};
const CORE_ORDER: readonly (readonly [string, string])[] = Object.values(CORE);

/** Preferred prefixes when a namespace must be declared on a property part's root. */
const PREFIX: ReadonlyMap<string, string> = new Map([
  [NS.cp, 'cp'],
  [NS.dc, 'dc'],
  [NS.dcterms, 'dcterms'],
  [NS.xsi, 'xsi'],
  [NS.vt, 'vt'],
]);

/** app.xml element → AppProperties key (numbers marked). */
const APP: ReadonlyMap<string, readonly [keyof AppProperties, boolean]> = new Map([
  ['Application', ['application', false]],
  ['AppVersion', ['appVersion', false]],
  ['Company', ['company', false]],
  ['Template', ['template', false]],
  ['Pages', ['pages', true]],
  ['Words', ['words', true]],
  ['Characters', ['characters', true]],
  ['CharactersWithSpaces', ['charactersWithSpaces', true]],
  ['Lines', ['lines', true]],
  ['Paragraphs', ['paragraphs', true]],
  ['TotalTime', ['totalTime', true]],
  ['DocSecurity', ['docSecurity', true]],
]);
/** The order Word writes app.xml elements in. */
const APP_ORDER: readonly (readonly [string, string])[] = [
  'Template',
  'TotalTime',
  'Pages',
  'Words',
  'Characters',
  'Application',
  'DocSecurity',
  'Lines',
  'Paragraphs',
  'ScaleCrop',
  'HeadingPairs',
  'TitlesOfParts',
  'Company',
  'LinksUpToDate',
  'CharactersWithSpaces',
  'SharedDoc',
  'HyperlinksChanged',
  'AppVersion',
].map((l) => [NS.ep, l] as const);

/** Format id every custom document property uses (FMTID_UserDefinedProperties). */
const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const I4_MIN = -2147483648;
const I4_MAX = 2147483647;

interface PartSpec {
  /** Relationship type last segment, for relTypeIs. */
  readonly kind: string;
  readonly name: string;
  readonly relType: string;
  readonly contentType: string;
  readonly root: () => XElement;
}

function decls(...pairs: [string, string][]): string[] {
  const out: string[] = [];
  for (const [name, uri] of pairs) out.push(name, NS.xmlns, uri);
  return out;
}

const CORE_PART: PartSpec = {
  kind: 'core-properties',
  name: 'docProps/core.xml',
  relType: REL.coreProperties,
  contentType: CT.coreProperties,
  root: () =>
    new XElement(
      'cp:coreProperties',
      NS.cp,
      decls(
        ['xmlns:cp', NS.cp],
        ['xmlns:dc', NS.dc],
        ['xmlns:dcterms', NS.dcterms],
        ['xmlns:dcmitype', NS.dcmitype],
        ['xmlns:xsi', NS.xsi],
      ),
    ),
};

const APP_PART: PartSpec = {
  kind: 'extended-properties',
  name: 'docProps/app.xml',
  relType: REL.extendedProperties,
  contentType: CT.extendedProperties,
  root: () => new XElement('Properties', NS.ep, decls(['xmlns', NS.ep], ['xmlns:vt', NS.vt])),
};

const CUSTOM_PART: PartSpec = {
  kind: 'custom-properties',
  name: 'docProps/custom.xml',
  relType: REL.customProperties,
  contentType: CT.customProperties,
  root: () =>
    new XElement('Properties', NS.custom, decls(['xmlns', NS.custom], ['xmlns:vt', NS.vt])),
};

/**
 * Core, extended and custom properties of a package. Obtain it from `WordDocument.properties`
 * (or construct one over any OpcPackage).
 *
 * Text getters return `undefined` when the element is absent; setting `undefined` removes it.
 */
export class DocumentProperties {
  readonly #pkg: OpcPackage;

  constructor(pkg: OpcPackage) {
    this.#pkg = pkg;
  }

  // ---- core properties --------------------------------------------------------------------

  /** `dc:title` */
  get title(): string | undefined {
    return this.#coreText('title');
  }
  set title(value: string | undefined) {
    this.#setCoreText('title', value);
  }

  /** `dc:subject` */
  get subject(): string | undefined {
    return this.#coreText('subject');
  }
  set subject(value: string | undefined) {
    this.#setCoreText('subject', value);
  }

  /** `dc:creator` (the author). */
  get creator(): string | undefined {
    return this.#coreText('creator');
  }
  set creator(value: string | undefined) {
    this.#setCoreText('creator', value);
  }

  /** `cp:keywords` */
  get keywords(): string | undefined {
    return this.#coreText('keywords');
  }
  set keywords(value: string | undefined) {
    this.#setCoreText('keywords', value);
  }

  /** `dc:description` (Word's "Comments"). */
  get description(): string | undefined {
    return this.#coreText('description');
  }
  set description(value: string | undefined) {
    this.#setCoreText('description', value);
  }

  /** `cp:lastModifiedBy` */
  get lastModifiedBy(): string | undefined {
    return this.#coreText('lastModifiedBy');
  }
  set lastModifiedBy(value: string | undefined) {
    this.#setCoreText('lastModifiedBy', value);
  }

  /** `cp:revision` (a string in the schema; Word writes a number). */
  get revision(): string | undefined {
    return this.#coreText('revision');
  }
  set revision(value: string | undefined) {
    this.#setCoreText('revision', value);
  }

  /** `cp:category` */
  get category(): string | undefined {
    return this.#coreText('category');
  }
  set category(value: string | undefined) {
    this.#setCoreText('category', value);
  }

  /** `cp:contentStatus` */
  get contentStatus(): string | undefined {
    return this.#coreText('contentStatus');
  }
  set contentStatus(value: string | undefined) {
    this.#setCoreText('contentStatus', value);
  }

  /** `dc:language` */
  get language(): string | undefined {
    return this.#coreText('language');
  }
  set language(value: string | undefined) {
    this.#setCoreText('language', value);
  }

  /** `dc:identifier` */
  get identifier(): string | undefined {
    return this.#coreText('identifier');
  }
  set identifier(value: string | undefined) {
    this.#setCoreText('identifier', value);
  }

  /** `cp:version` */
  get version(): string | undefined {
    return this.#coreText('version');
  }
  set version(value: string | undefined) {
    this.#setCoreText('version', value);
  }

  /** `dcterms:created` (written as W3CDTF, whole seconds, UTC). */
  get created(): Date | undefined {
    return parseW3cdtf(this.#coreText('created'));
  }
  set created(value: Date | undefined) {
    this.#setCoreDate('created', value);
  }

  /** `dcterms:modified` (written as W3CDTF, whole seconds, UTC). */
  get modified(): Date | undefined {
    return parseW3cdtf(this.#coreText('modified'));
  }
  set modified(value: Date | undefined) {
    this.#setCoreDate('modified', value);
  }

  /** `cp:lastPrinted` */
  get lastPrinted(): Date | undefined {
    return parseW3cdtf(this.#coreText('lastPrinted'));
  }
  set lastPrinted(value: Date | undefined) {
    this.#setCoreDate('lastPrinted', value);
  }

  // ---- extended (app) properties ----------------------------------------------------------

  /** Extended properties (docProps/app.xml); `{}` when the part is absent. */
  get app(): AppProperties {
    const out: Record<string, string | number> = {};
    const root = this.#root(APP_PART);
    if (root === undefined) return out;
    for (const el of root.elements(NS.ep)) {
      const entry = APP.get(el.local);
      if (entry === undefined || entry[0] in out) continue;
      const [key, numeric] = entry;
      const text = el.text;
      if (!numeric) out[key] = text;
      else if (/^\s*-?\d+\s*$/.test(text)) out[key] = Number.parseInt(text, 10);
    }
    return out as AppProperties;
  }

  /** `Company` in docProps/app.xml. */
  get company(): string | undefined {
    return this.#root(APP_PART)?.element(NS.ep, 'Company')?.text;
  }
  set company(value: string | undefined) {
    this.#setText(APP_PART, APP_ORDER, NS.ep, 'Company', value);
  }

  /** `Application` in docProps/app.xml (the producing application's name). */
  get application(): string | undefined {
    return this.#root(APP_PART)?.element(NS.ep, 'Application')?.text;
  }
  set application(value: string | undefined) {
    this.#setText(APP_PART, APP_ORDER, NS.ep, 'Application', value);
  }

  // ---- custom properties ------------------------------------------------------------------

  /** All custom properties by name, typed from their `vt:*` element. */
  custom(): Record<string, CustomPropertyValue> {
    const out: Record<string, CustomPropertyValue> = {};
    const root = this.#root(CUSTOM_PART);
    if (root === undefined) return out;
    for (const prop of root.elements(NS.custom, 'property')) {
      const name = prop.attr('name');
      if (name === undefined || Object.hasOwn(out, name)) continue;
      const value = readVariant(prop.elements(NS.vt)[0]);
      if (value === undefined) continue;
      // defineProperty: a property named "__proto__" must not reach the prototype setter.
      Object.defineProperty(out, name, {
        value,
        enumerable: true,
        writable: true,
        configurable: true,
      });
    }
    return out;
  }

  /**
   * Sets a custom property (names compare case-insensitively, as in Word). Strings are written
   * as `vt:lpwstr`, integers in the 32-bit range as `vt:i4`, other numbers as `vt:r8`, booleans
   * as `vt:bool` and dates as `vt:filetime`. An existing property keeps its `pid`.
   */
  setCustom(name: string, value: CustomPropertyValue): void {
    if (name === '') throw new UsageError('Custom property name must not be empty');
    const [type, text] = writeVariant(value);
    const existing = this.#customProp(name);
    if (existing !== undefined) {
      const cur = existing.elements(NS.vt);
      if (cur.length === 1 && cur[0]!.local === type && cur[0]!.text === text) return;
    }
    const root = existing === undefined ? this.#ensure(CUSTOM_PART) : this.#root(CUSTOM_PART)!;
    ensureNs(root, NS.vt);
    const doc = root.ownerDocument()!;
    const v = doc.createElement(NS.vt, type);
    v.text = text;
    if (existing !== undefined) {
      existing.clear();
      existing.append(v);
      return;
    }
    let maxPid = 1;
    for (const p of root.elements(NS.custom, 'property')) {
      const pid = Number.parseInt(p.attr('pid') ?? '', 10);
      if (pid > maxPid) maxPid = pid;
    }
    const prop = doc.createElement(NS.custom, 'property', {
      fmtid: FMTID,
      pid: String(maxPid + 1),
      name,
    });
    prop.append(v);
    root.append(prop);
  }

  /** Removes a custom property (case-insensitive name); returns whether it existed. */
  removeCustom(name: string): boolean {
    const prop = this.#customProp(name);
    if (prop === undefined) return false;
    prop.remove();
    return true;
  }

  // ---- internals --------------------------------------------------------------------------

  #customProp(name: string): XElement | undefined {
    const root = this.#root(CUSTOM_PART);
    if (root === undefined) return undefined;
    const props = root.elements(NS.custom, 'property');
    const exact = props.find((p) => p.attr('name') === name);
    if (exact !== undefined) return exact;
    const lower = name.toLowerCase();
    return props.find((p) => p.attr('name')?.toLowerCase() === lower);
  }

  /** The part targeted by the first internal package relationship of `spec.kind`. */
  #find(spec: PartSpec): Part | undefined {
    for (const r of this.#pkg.rels.all()) {
      if (r.external || r.partName === undefined || !relTypeIs(r.type, spec.kind)) continue;
      const part = this.#pkg.part(r.partName);
      if (part !== undefined) return part;
    }
    return undefined;
  }

  #root(spec: PartSpec): XElement | undefined {
    const part = this.#find(spec);
    return part === undefined ? undefined : part.xml().root;
  }

  /** Root element of the part, creating the part, relationship and override when missing. */
  #ensure(spec: PartSpec): XElement {
    const pkg = this.#pkg;
    let part = this.#find(spec);
    if (part === undefined) {
      part = pkg.part(spec.name);
      if (part === undefined) {
        const doc = new XDocument(XML_DECL);
        doc.append(spec.root());
        part = pkg.addPart(spec.name, spec.contentType, doc);
      }
      pkg.rels.add(spec.relType, part.name);
    }
    const ct = pkg.contentTypes;
    if (ct.get(part.name) !== spec.contentType) ct.setOverride(part.name, spec.contentType);
    return part.xml().root;
  }

  #coreText(prop: CoreProperty): string | undefined {
    const [uri, local] = CORE[prop];
    return this.#root(CORE_PART)?.element(uri, local)?.text;
  }

  #setCoreText(prop: CoreTextProperty, value: string | undefined): void {
    const [uri, local] = CORE[prop];
    this.#setText(CORE_PART, CORE_ORDER, uri, local, value);
  }

  #setCoreDate(prop: CoreDateProperty, value: Date | undefined): void {
    const [uri, local] = CORE[prop];
    const el = this.#setText(
      CORE_PART,
      CORE_ORDER,
      uri,
      local,
      value === undefined ? undefined : formatW3cdtf(value),
    );
    // Word types dcterms dates explicitly (and requires it); cp:lastPrinted has no xsi:type.
    if (el !== undefined && uri === NS.dcterms && el.attrNS(NS.xsi, 'type') === undefined) {
      const root = el.ownerDocument()!.root;
      ensureNs(root, NS.xsi);
      el.setAttrNS(NS.xsi, 'type', el.ownerDocument()!.qualify(NS.dcterms, 'W3CDTF'));
    }
  }

  /**
   * Sets (or, for undefined, removes) the text of element `{uri}local` under the part's root,
   * inserting a new element at its place in `order`. Returns the element when one remains.
   */
  #setText(
    spec: PartSpec,
    order: readonly (readonly [string, string])[],
    uri: string,
    local: string,
    value: string | undefined,
  ): XElement | undefined {
    if (value === undefined) {
      const root = this.#root(spec);
      if (root !== undefined) for (const el of root.elements(uri, local)) el.remove();
      return undefined;
    }
    const root = this.#ensure(spec);
    let el = root.element(uri, local);
    if (el === undefined) {
      ensureNs(root, uri);
      el = root.ownerDocument()!.createElement(uri, local);
      insertOrdered(root, el, order);
    }
    setTextIfChanged(el, value);
    return el;
  }
}

/** Declares `uri` on `root` when no prefix is bound to it (preferred prefix, else a free one). */
function ensureNs(root: XElement, uri: string): void {
  if (root.lookupPrefix(uri) !== undefined) return;
  const base = PREFIX.get(uri) ?? 'ns';
  let prefix = base;
  for (let n = 1; root.lookupNamespace(prefix) !== undefined; n++) prefix = `${base}${n}`;
  root.setAttr(`xmlns:${prefix}`, uri);
}

/** Inserts `el` before the first child element that comes after it in `order`. */
function insertOrdered(
  root: XElement,
  el: XElement,
  order: readonly (readonly [string, string])[],
): void {
  const rank = (e: XElement): number => order.findIndex(([u, l]) => e.uri === u && e.local === l);
  const mine = rank(el);
  for (const child of root.elements()) {
    if (rank(child) > mine) {
      root.insertBefore(el, child);
      return;
    }
  }
  root.append(el);
}

/** Sets an element's text only when it differs (so a no-op edit leaves the part clean). */
function setTextIfChanged(el: XElement, value: string): void {
  const kids = el.children;
  const same =
    value === ''
      ? kids.length === 0
      : kids.length === 1 && kids[0]!.nodeType === TEXT_NODE && kids[0]!.text === value;
  if (!same) el.text = value;
}

const W3CDTF =
  /^(\d{4})(?:-(\d{2})(?:-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?)?)?$/;

/**
 * Parses a W3CDTF date (ISO 8601 profile: `YYYY`, `YYYY-MM`, `YYYY-MM-DD`, `YYYY-MM-DDThh:mm`,
 * with optional seconds, fraction and zone). A missing zone is read as UTC, so results never
 * depend on the machine's time zone.
 */
function parseW3cdtf(text: string | undefined): Date | undefined {
  if (text === undefined) return undefined;
  const m = W3CDTF.exec(text.trim());
  if (m === null) return undefined;
  const d = new Date(0);
  d.setUTCFullYear(Number(m[1]), Number(m[2] ?? 1) - 1, Number(m[3] ?? 1));
  const ms = m[7] === undefined ? 0 : Math.round(Number(m[7]) * 1000);
  d.setUTCHours(Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0), ms);
  const tz = m[8];
  if (tz !== undefined && tz !== 'Z') {
    const sign = tz[0] === '-' ? -1 : 1;
    const minutes = Number(tz.slice(1, 3)) * 60 + Number(tz.slice(4, 6));
    d.setTime(d.getTime() - sign * minutes * 60_000);
  }
  return Number.isNaN(d.getTime()) ? undefined : d;
}

/** ISO 8601 in UTC without milliseconds, as Word writes dates (`2024-03-05T10:20:30Z`). */
function formatW3cdtf(date: Date): string {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    throw new UsageError('Invalid Date');
  }
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** `vt:*` element type and text for a custom property value. */
function writeVariant(value: CustomPropertyValue): [string, string] {
  switch (typeof value) {
    case 'string':
      return ['lpwstr', value];
    case 'boolean':
      return ['bool', value ? 'true' : 'false'];
    case 'number':
      if (!Number.isFinite(value))
        throw new UsageError(`Custom property value ${value} is not finite`);
      if (Number.isInteger(value) && value >= I4_MIN && value <= I4_MAX)
        return ['i4', String(value)];
      return ['r8', String(value)];
    default:
      if (value instanceof Date) return ['filetime', formatW3cdtf(value)];
      throw new UsageError('Custom property values must be string, number, boolean or Date');
  }
}

/** Value of a `vt:*` element; unknown or unparsable types fall back to their text. */
function readVariant(el: XElement | undefined): CustomPropertyValue | undefined {
  if (el === undefined) return undefined;
  const text = el.text;
  switch (el.local) {
    case 'i1':
    case 'i2':
    case 'i4':
    case 'i8':
    case 'int':
    case 'ui1':
    case 'ui2':
    case 'ui4':
    case 'ui8':
    case 'uint':
    case 'r4':
    case 'r8':
    case 'decimal': {
      const t = text.trim();
      const n = Number(t);
      return t === '' || Number.isNaN(n) ? text : n;
    }
    case 'bool': {
      const t = text.trim().toLowerCase();
      return t === 'true' || t === '1';
    }
    case 'filetime':
    case 'date':
      return parseW3cdtf(text) ?? text;
    default:
      return text;
  }
}
