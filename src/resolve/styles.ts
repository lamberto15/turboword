// The styles part: style definitions, basedOn chains and their merged property layers.
//
// - Paragraph styles: a paragraph without `w:pStyle` (or naming a style that does not exist)
//   uses the default paragraph style (`w:default="1"`). Its chain is split at that default style
//   because, inside tables, Word lets the table style override what the default paragraph style
//   (and the document defaults) say, while any other paragraph style still overrides the table
//   style.
// - Character styles: only applied when a run names one; the default character style ("Default
//   Paragraph Font") is never applied implicitly.
// - Table styles: a table without (or with an unknown) `w:tblStyle` uses the default table style.
//   Each style contributes whole-table properties and conditional regions (`w:tblStylePr`).
// basedOn chains stop at a cycle, at a style of another type, or after MAX_CHAIN levels.

import type { XDocument, XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { ParaLayer } from './ppr.js';
import { mergePara, readPPr } from './ppr.js';
import type { RunLayer, ThemeCtx } from './rpr.js';
import { mergeRun, readRPr } from './rpr.js';
import type { TblLayer, TcLayer } from './tblpr.js';
import { mergeTbl, mergeTc, readTblPr, readTcPr } from './tblpr.js';
import { onOffValue, wAttr } from './values.js';

/** Longest basedOn chain followed (Word's own UI limit is lower). */
export const MAX_CHAIN = 64;

export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

/** One `w:style`. */
export interface StyleDef {
  readonly id: string;
  readonly type: StyleType;
  readonly isDefault: boolean;
  readonly basedOn: string | undefined;
  readonly el: XElement;
}

/** Index of the styles part. */
export interface StyleIndex {
  readonly byId: ReadonlyMap<string, StyleDef>;
  /** Default style id per type. */
  readonly defaults: Readonly<Partial<Record<StyleType, string>>>;
  readonly docRPr: XElement | undefined;
  readonly docPPr: XElement | undefined;
}

/** Builds the index of a styles part (an empty index when there is none). */
export function indexStyles(doc: XDocument | undefined): StyleIndex {
  const byId = new Map<string, StyleDef>();
  const defaults: Partial<Record<StyleType, string>> = {};
  let root: XElement | undefined;
  try {
    root = doc?.root;
  } catch {
    root = undefined;
  }
  if (root === undefined || !root.is(NS.w, 'styles')) {
    return { byId, defaults, docRPr: undefined, docPPr: undefined };
  }
  for (const el of root.elements(NS.w, 'style')) {
    const id = wAttr(el, 'styleId');
    if (id === undefined || byId.has(id)) continue;
    const t = wAttr(el, 'type');
    const type: StyleType =
      t === 'character' || t === 'table' || t === 'numbering' ? t : 'paragraph';
    const isDefault = onOffValue(wAttr(el, 'default')) === true;
    const basedOn = el.element(NS.w, 'basedOn')?.attrNS(NS.w, 'val');
    byId.set(id, { id, type, isDefault, basedOn, el });
    if (isDefault && defaults[type] === undefined) defaults[type] = id;
  }
  const dd = root.element(NS.w, 'docDefaults');
  return {
    byId,
    defaults,
    docRPr: dd?.element(NS.w, 'rPrDefault')?.element(NS.w, 'rPr'),
    docPPr: dd?.element(NS.w, 'pPrDefault')?.element(NS.w, 'pPr'),
  };
}

/** The basedOn chain of `id`, leaf first (empty when `id` is not a style of `type`). */
export function styleChain(index: StyleIndex, id: string | undefined, type: StyleType): StyleDef[] {
  const out: StyleDef[] = [];
  const seen = new Set<string>();
  let cur = id;
  while (cur !== undefined && out.length < MAX_CHAIN && !seen.has(cur)) {
    const def = index.byId.get(cur);
    if (def === undefined || def.type !== type) break;
    seen.add(cur);
    out.push(def);
    cur = def.basedOn;
  }
  return out;
}

/** Merged layers of a paragraph style chain. */
export interface ParaStyleInfo {
  /** Effective style id (the default paragraph style when none or an unknown one is named). */
  readonly id: string | undefined;
  readonly chain: readonly StyleDef[];
  /** Whole chain. */
  readonly pFull: ParaLayer;
  readonly rFull: RunLayer;
  /** The default paragraph style and its bases (applied below the table style). */
  readonly pLower: ParaLayer;
  readonly rLower: RunLayer;
  /** The styles above the default paragraph style (applied above the table style). */
  readonly pUpper: ParaLayer;
  readonly rUpper: RunLayer;
}

/** Resolves a paragraph style chain into layers. */
export function paraStyleInfo(
  t: ThemeCtx,
  index: StyleIndex,
  requested: string | undefined,
): ParaStyleInfo {
  let id = requested;
  if (id === undefined || index.byId.get(id)?.type !== 'paragraph') id = index.defaults.paragraph;
  const chain = styleChain(index, id, 'paragraph');
  const def = index.defaults.paragraph;
  const split = def === undefined ? -1 : chain.findIndex((s) => s.id === def);
  let pFull: ParaLayer = {};
  let rFull: RunLayer = {};
  let pLower: ParaLayer = {};
  let rLower: RunLayer = {};
  let pUpper: ParaLayer = {};
  let rUpper: RunLayer = {};
  for (let k = chain.length - 1; k >= 0; k--) {
    const s = chain[k]!;
    const p = readPPr(t, s.el.element(NS.w, 'pPr'));
    const r = readRPr(t, s.el.element(NS.w, 'rPr'));
    pFull = mergePara(pFull, p);
    rFull = mergeRun(rFull, r);
    if (split >= 0 && k >= split) {
      pLower = mergePara(pLower, p);
      rLower = mergeRun(rLower, r);
    } else {
      pUpper = mergePara(pUpper, p);
      rUpper = mergeRun(rUpper, r);
    }
  }
  return { id, chain, pFull, rFull, pLower, rLower, pUpper, rUpper };
}

/** Merged run layer of a character style chain. */
export interface CharStyleInfo {
  readonly id: string | undefined;
  readonly r: RunLayer;
}

/** Resolves a character style chain (unknown or missing id → nothing applied). */
export function charStyleInfo(
  t: ThemeCtx,
  index: StyleIndex,
  id: string | undefined,
): CharStyleInfo {
  const chain = styleChain(index, id, 'character');
  if (chain.length === 0) return { id: undefined, r: {} };
  let r: RunLayer = {};
  for (let k = chain.length - 1; k >= 0; k--) {
    r = mergeRun(r, readRPr(t, chain[k]!.el.element(NS.w, 'rPr')));
  }
  return { id: chain[0]!.id, r };
}

/** One set of table-style properties (whole table or one conditional region). */
export interface TableStyleLayers {
  pPr: ParaLayer;
  rPr: RunLayer;
  tbl: TblLayer;
  tc: TcLayer;
}

/** Merged layers of a table style chain. */
export interface TableStyleInfo {
  readonly id: string | undefined;
  readonly whole: TableStyleLayers;
  /** Conditional regions (`w:tblStylePr/@w:type`) defined anywhere in the chain. */
  readonly regions: ReadonlyMap<string, TableStyleLayers>;
}

function readLayers(t: ThemeCtx, el: XElement): TableStyleLayers {
  return {
    pPr: readPPr(t, el.element(NS.w, 'pPr')),
    rPr: readRPr(t, el.element(NS.w, 'rPr')),
    tbl: readTblPr(t, el.element(NS.w, 'tblPr')),
    tc: readTcPr(t, el.element(NS.w, 'tcPr')),
  };
}

function mergeLayers(a: TableStyleLayers, b: TableStyleLayers): TableStyleLayers {
  return {
    pPr: mergePara(a.pPr, b.pPr),
    rPr: mergeRun(a.rPr, b.rPr),
    tbl: mergeTbl(a.tbl, b.tbl),
    tc: mergeTc(a.tc, b.tc),
  };
}

const EMPTY_LAYERS = (): TableStyleLayers => ({ pPr: {}, rPr: {}, tbl: {}, tc: {} });

/** Resolves a table style chain (missing/unknown id → the default table style). */
export function tableStyleInfo(
  t: ThemeCtx,
  index: StyleIndex,
  requested: string | undefined,
): TableStyleInfo {
  let id = requested;
  if (id === undefined || index.byId.get(id)?.type !== 'table') id = index.defaults.table;
  const chain = styleChain(index, id, 'table');
  let whole = EMPTY_LAYERS();
  const regions = new Map<string, TableStyleLayers>();
  for (let k = chain.length - 1; k >= 0; k--) {
    const s = chain[k]!;
    whole = mergeLayers(whole, readLayers(t, s.el));
    for (const cp of s.el.elements(NS.w, 'tblStylePr')) {
      const type = wAttr(cp, 'type');
      if (type === undefined) continue;
      regions.set(type, mergeLayers(regions.get(type) ?? EMPTY_LAYERS(), readLayers(t, cp)));
    }
  }
  return { id: chain.length > 0 ? id : undefined, whole, regions };
}
