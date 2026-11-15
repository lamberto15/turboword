// Per-document resolution context and its caches.
//
// Everything the resolver derives (style chains, numbering levels, resolved paragraphs, table
// grids, list counters) is cached on one context object per package. The context remembers the
// parts it was built from (styles, numbering, theme, settings, plus the main part's rels that
// locate them) together with their `XDocument.version`; any change to one of them, a replaced
// part or a re-pointed relationship rebuilds the context and so drops every cache at once.
// Caches keyed by story content (paragraphs, tables, counters) additionally record the version
// of the story part they were computed from.

import type { OpcPackage } from '../opc/package.js';
import type { Part } from '../opc/part.js';
import { relsPathFor } from '../opc/uri.js';
import type { WordDocument } from '../wml/document.js';
import type { XDocument, XElement } from '../xml/dom.js';
import type { StoryCounters } from './labels.js';
import type { EffectiveLevel, NumberingIndex } from './numbering.js';
import { indexNumbering, levelOwners } from './numbering.js';
import type { ParaResult } from './paragraph.js';
import type { ParaLayer } from './ppr.js';
import { readPPr } from './ppr.js';
import type { RunLayer, ThemeCtx } from './rpr.js';
import { readRPr } from './rpr.js';
import type { CharStyleInfo, ParaStyleInfo, StyleIndex, TableStyleInfo } from './styles.js';
import { charStyleInfo, indexStyles, paraStyleInfo, tableStyleInfo } from './styles.js';
import type { TableInfo } from './table.js';
import type { SettingsData, ThemeData } from './theme.js';
import { parseSettings, parseTheme } from './theme.js';

interface PartSnapshot {
  readonly part: Part | undefined;
  readonly doc: XDocument | undefined;
  readonly version: number;
}

/** A cache entry tied to the versions of several parts (a counter flow over many stories). */
export interface FlowEntry<T> {
  readonly docs: readonly XDocument[];
  readonly versions: readonly number[];
  readonly value: T;
}

/** A cache entry tied to the version of the story part it was computed from. */
export interface Versioned<T> {
  readonly doc: XDocument | undefined;
  readonly version: number;
  readonly value: T;
}

function snapshot(part: Part | undefined): PartSnapshot {
  let doc: XDocument | undefined;
  try {
    doc = part?.xml();
  } catch {
    doc = undefined;
  }
  return { part, doc, version: doc?.version ?? 0 };
}

function same(s: PartSnapshot): boolean {
  if (s.part === undefined) return true;
  let doc: XDocument | undefined;
  try {
    doc = s.part.xml();
  } catch {
    doc = undefined;
  }
  return doc === s.doc && (doc?.version ?? 0) === s.version;
}

/** Resolution context of one package (see the file comment). */
export class DocContext implements ThemeCtx {
  readonly theme: ThemeData;
  readonly settings: SettingsData;
  readonly styles: StyleIndex;
  readonly numbering: NumberingIndex;
  readonly docPPr: ParaLayer;
  readonly docRPr: RunLayer;

  /** @internal */ readonly paraStyles: Map<string, ParaStyleInfo> = new Map();
  /** @internal */ readonly charStyles: Map<string, CharStyleInfo> = new Map();
  /** @internal */ readonly tableStyles: Map<string, TableStyleInfo> = new Map();
  /** @internal */ readonly levels: Map<string, EffectiveLevel | undefined> = new Map();
  /** @internal */ readonly paragraphs: WeakMap<XElement, Versioned<ParaResult>> = new WeakMap();
  /** @internal */ readonly tables: WeakMap<XElement, Versioned<TableInfo>> = new WeakMap();
  /** @internal */ readonly flows: WeakMap<object, FlowEntry<StoryCounters>> = new WeakMap();
  /** @internal Key of the counter flow shared by all headers and footers. */
  readonly hfKey: object = {};
  #owners: Map<string, string> | undefined;

  readonly #mainName: string;
  readonly #rels: PartSnapshot;
  readonly #parts: readonly PartSnapshot[];

  /** @internal Use {@link contextOf}. */
  constructor(doc: WordDocument) {
    const pkg = doc.package;
    const main = doc.mainPart;
    this.#mainName = main.name;
    this.#rels = snapshot(pkg.part(relsPathFor(main.name)));
    const first = (kind: string): Part | undefined => {
      try {
        return main.relatedByType(kind)[0];
      } catch {
        return undefined;
      }
    };
    const styles = snapshot(first('styles'));
    const numbering = snapshot(first('numbering'));
    const theme = snapshot(first('theme'));
    const settings = snapshot(first('settings'));
    this.#parts = [styles, numbering, theme, settings];
    this.theme = parseTheme(theme.doc);
    this.settings = parseSettings(settings.doc);
    this.styles = indexStyles(styles.doc);
    this.numbering = indexNumbering(numbering.doc);
    this.docPPr = readPPr(this, this.styles.docPPr);
    this.docRPr = readRPr(this, this.styles.docRPr);
  }

  /** @internal True while the parts this context was built from are unchanged. */
  isCurrent(doc: WordDocument): boolean {
    if (doc.mainPart.name !== this.#mainName) return false;
    const relsPart = doc.package.part(relsPathFor(this.#mainName));
    if (relsPart !== this.#rels.part || !same(this.#rels)) return false;
    return this.#parts.every(same);
  }

  /** Owner paragraph style of a numbering level (see numbering.ts), if any style claims it. */
  levelOwner(abstractId: number, ilvl: number): string | undefined {
    this.#owners ??= levelOwners(this.numbering, this.styles);
    return this.#owners.get(`${abstractId}:${ilvl}`);
  }

  /** Paragraph style chain layers (cached). */
  paraStyle(id: string | undefined): ParaStyleInfo {
    const key = id ?? '';
    let v = this.paraStyles.get(key);
    if (v === undefined) {
      v = paraStyleInfo(this, this.styles, id);
      this.paraStyles.set(key, v);
    }
    return v;
  }

  /** Character style chain layer (cached). */
  charStyle(id: string | undefined): CharStyleInfo {
    const key = id ?? '';
    let v = this.charStyles.get(key);
    if (v === undefined) {
      v = charStyleInfo(this, this.styles, id);
      this.charStyles.set(key, v);
    }
    return v;
  }

  /** Table style chain layers (cached). */
  tableStyle(id: string | undefined): TableStyleInfo {
    const key = id ?? '';
    let v = this.tableStyles.get(key);
    if (v === undefined) {
      v = tableStyleInfo(this, this.styles, id);
      this.tableStyles.set(key, v);
    }
    return v;
  }
}

const contexts = new WeakMap<OpcPackage, DocContext>();

/** The (cached, revalidated) resolution context of a document. */
export function contextOf(doc: WordDocument): DocContext {
  const cur = contexts.get(doc.package);
  if (cur?.isCurrent(doc)) return cur;
  const fresh = new DocContext(doc);
  contexts.set(doc.package, fresh);
  return fresh;
}

/** A cached value for `el` when its part is still at the recorded version. */
export function cached<T>(map: WeakMap<XElement, Versioned<T>>, el: XElement): T | undefined {
  const e = map.get(el);
  if (e === undefined) return undefined;
  const doc = el.ownerDocument();
  if (doc === undefined || doc !== e.doc || doc.version !== e.version) return undefined;
  return e.value;
}

/** Stores `value` for `el`, stamped with its part's version (detached elements are not cached). */
export function store<T>(map: WeakMap<XElement, Versioned<T>>, el: XElement, value: T): T {
  const doc = el.ownerDocument();
  if (doc !== undefined) map.set(el, { doc, version: doc.version, value });
  return value;
}
