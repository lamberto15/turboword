// importStyles: compose pipeline step 3 — brings the source's styles into the target.
//
// Scope: every style the source uses anywhere (body, headers, footers, notes, comments, the
// default paragraph/table style of unstyled content, styles linked from lists), plus, for each
// style that gets copied, its dependencies (basedOn, next, link, list styles). latentStyles are
// never touched.
//
// Matching. A source style matches a target style of the same type with the same styleId, unless
// their names disagree and another target style carries the source's name: names win then,
// because Word identifies styles by name (styleIds of built-ins are localized — a German
// template's "Überschrift1" is named "heading 1", built-in names are English in every UI
// language — and two styles can never share a name). Without an id match, styles match by name
// (compared ignoring case). Matched styles are compared by a fingerprint of their flattened
// properties (styles-props.ts), so identical styles defined differently never conflict.
//
// Modes (ImportOptions.styles), for a matched style whose fingerprint differs:
//   - useDestinationStyles: the target definition wins.
//   - keepSourceFormatting: the content keeps the target style, and the difference between the
//     source and target resolved formatting (docDefaults → table style → paragraph style →
//     character style) is written as direct formatting on the imported paragraphs, runs and
//     paragraph marks, as Aspose does; differing docDefaults are expanded the same way. Table
//     and numbering styles cannot be expanded and are copied renamed instead. Copied styles that
//     inherit from an expanded one get the inherited difference materialized. With
//     smartStyleBehavior (default) styles the target lacks are copied; without it they are
//     expanded too and the target gains no new paragraph/character styles.
//   - keepDifferentStyles: the source style is copied renamed (`Heading1_1`, name `heading 1_1`),
//     keeping its outline level so TOCs still pick it up; an identical earlier copy is reused.
// Unmatched styles are copied (renamed only when their id or name is taken by another type).
//
// Bookkeeping: ctx.styleMap, ctx.importedStyles (created target styles), pStyle/rStyle/tblStyle
// in the blocks. Later content (headers, notes, comments) goes through remapStyleRefs, which
// applies the same mapping and expansion.

import type { Part } from '../opc/part.js';
import { CT, REL } from '../opc/rel-types.js';
import { partDir, relativeTarget } from '../opc/uri.js';
import {
  type ChildOrder,
  ensureChild,
  ensureFirstChild,
  insertOrdered,
  makeOrder,
  PPR_ORDER,
  RPR_ORDER,
} from '../wml/order.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { parseXml } from '../xml/parse.js';
import {
  CONTAINERS,
  type Container,
  childVal,
  diffMaps,
  type FlatStyle,
  fillMissing,
  fingerprint,
  listKeyFor,
  mergeMaps,
  NumberingIndex,
  newUsage,
  normName,
  numIdOf,
  onOffAttr,
  type PropMap,
  relatedPart,
  rootOf,
  StyleIndex,
  scanUsage,
  storyRoots,
  styleIdOf,
  styleNameOf,
  styleTypeOf,
  wEl,
} from './styles-props.js';
import type { ImportContext } from './types.js';

/** `w:style` children (CT_Style, §17.7.4.17). */
const STYLE_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'name',
  'aliases',
  'basedOn',
  'next',
  'link',
  'autoRedefine',
  'hidden',
  'uiPriority',
  'semiHidden',
  'unhideWhenUsed',
  'qFormat',
  'locked',
  'personal',
  'personalCompose',
  'personalReply',
  'rsid',
  'pPr',
  'rPr',
  'tblPr',
  'trPr',
  'tcPr',
  'tblStylePr',
]);

/** `w:tblPr` children (§17.4.60). */
const TBLPR_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'tblStyle',
  'tblpPr',
  'tblOverlap',
  'bidiVisual',
  'tblStyleRowBandSize',
  'tblStyleColBandSize',
  'tblW',
  'jc',
  'tblCellSpacing',
  'tblInd',
  'tblBorders',
  'shd',
  'tblLayout',
  'tblCellMar',
  'tblLook',
  'tblCaption',
  'tblDescription',
  'tblPrChange',
]);

/** `w:trPr` children (§17.4.82). */
const TRPR_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'cnfStyle',
  'divId',
  'gridBefore',
  'gridAfter',
  'wBefore',
  'wAfter',
  'cantSplit',
  'trHeight',
  'tblHeader',
  'tblCellSpacing',
  'jc',
  'hidden',
  'ins',
  'del',
  'trPrChange',
]);

/** `w:tcPr` children (§17.4.70). */
const TCPR_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
  'cnfStyle',
  'tcW',
  'gridSpan',
  'hMerge',
  'vMerge',
  'tcBorders',
  'shd',
  'noWrap',
  'tcMar',
  'textDirection',
  'tcFitText',
  'vAlign',
  'hideMark',
  'headers',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tcPrChange',
]);

const ORDERS: Readonly<Record<Container, ChildOrder>> = {
  pPr: PPR_ORDER,
  rPr: RPR_ORDER,
  tblPr: TBLPR_ORDER,
  trPr: TRPR_ORDER,
  tcPr: TCPR_ORDER,
};

/**
 * What happens to one source style:
 * - same: matched an identical target style; - dest: matched, target definition wins;
 * - expand: formatting expanded as direct formatting over `target` (undefined: no style);
 * - copy: copied as `target`; - rename: conflicting style copied renamed as `target`.
 */
type Action = 'same' | 'dest' | 'expand' | 'copy' | 'rename';

interface Decision {
  readonly action: Action;
  readonly target: string | undefined;
  /** New `w:name` for a renamed copy. */
  readonly name?: string;
}

interface StyleState {
  readonly src: StyleIndex;
  readonly tgt: StyleIndex;
  readonly srcNums: NumberingIndex;
  readonly ksf: boolean;
  readonly decisions: Map<string, Decision>;
  /** Target styleIds created by this import. */
  readonly created: Set<string>;
  /** Source numIds that expansion writes as direct formatting (importNumbering imports them). */
  readonly expandNumIds: Set<number>;
  /**
   * [source numId, target numId] of paragraph styles mapped onto a pre-existing target style
   * (same/dest) when both define a list: content numbered through such a style now follows the
   * target's list, so the source list joins it (importNumbering).
   */
  readonly listPairs: [number, number][];
}

const STATE = new WeakMap<ImportContext, StyleState>();

const COPIES: ReadonlySet<Action> = new Set<Action>(['copy', 'rename']);

/**
 * Brings every style the source uses into the target according to `ctx.options.styles`, fills
 * `ctx.styleMap` / `ctx.importedStyles`, and rewrites the style references (and expands
 * formatting, in keepSourceFormatting) in `blocks`.
 */
export function importStyles(ctx: ImportContext, blocks: XElement[]): void {
  const src = new StyleIndex(relatedPart(ctx.source, 'styles'));
  if (src.root === undefined) return; // nothing to import; references stay as they are
  const tgtPart = relatedPart(ctx.target, 'styles') ?? createStylesPart(ctx);
  const tgt = new StyleIndex(tgtPart);
  const tgtRoot = tgt.root;
  if (tgtRoot === undefined) {
    ctx.warn('STYLES_UNREADABLE', 'The target styles part has no w:styles root', tgtPart.name);
    return;
  }
  const srcNums = new NumberingIndex(relatedPart(ctx.source, 'numbering'));
  const tgtNums = new NumberingIndex(relatedPart(ctx.target, 'numbering'));
  const mode = ctx.options.styles;
  const ksf = mode === 'keepSourceFormatting';
  const st: StyleState = {
    src,
    tgt,
    srcNums,
    ksf,
    decisions: new Map(),
    created: new Set(),
    expandNumIds: new Set(),
    listPairs: [],
  };
  STATE.set(ctx, st);

  const srcList = listKeyFor(srcNums, src);
  const tgtList = listKeyFor(tgtNums, tgt);
  const srcFp = new Map<string, string>();
  const tgtFp = new Map<string, string>();
  const fpOf = (side: 'src' | 'tgt', id: string): string => {
    const cache = side === 'src' ? srcFp : tgtFp;
    let fp = cache.get(id);
    if (fp === undefined) {
      const idx = side === 'src' ? src : tgt;
      fp = fingerprint(withImplicitOutline(idx, id), side === 'src' ? srcList : tgtList);
      cache.set(id, fp);
    }
    return fp;
  };

  // ---- 1. what the source uses ----
  const usage = newUsage();
  scanUsage(storyRoots(ctx.source), usage);
  scanUsage(blocks, usage);
  const queue: string[] = [];
  const queued = new Set<string>();
  /** Styles reached through lists: never expanded away (the list needs the style to exist). */
  const listLinked = new Set<string>();
  const enqueue = (id: string | undefined): void => {
    if (id === undefined || queued.has(id)) return;
    queued.add(id);
    queue.push(id);
  };
  const enqueueList = (numId: number | undefined): void => {
    for (const id of listStyleRefs(srcNums, numId)) {
      listLinked.add(id);
      enqueue(id);
    }
  };
  for (const id of usage.styles) enqueue(id);
  if (usage.bareParagraph) enqueue(src.defaultId('paragraph'));
  if (usage.bareTable) enqueue(src.defaultId('table'));
  for (const n of usage.numIds) enqueueList(n);

  // ---- 2. decide, following the dependencies of whatever gets copied ----
  const reservedIds = new Set<string>();
  const reservedNames = new Set<string>();
  const unique = (baseId: string, baseName: string | undefined, force: boolean): Decision => {
    const idFree = (id: string): boolean => !tgt.byId.has(id) && !reservedIds.has(id);
    const nameFree = (name: string): boolean =>
      !tgt.hasName(name) && !reservedNames.has(normName(name));
    let n = 0;
    if (force || !idFree(baseId)) for (n = 1; !idFree(`${baseId}_${n}`); n++);
    const id = n === 0 ? baseId : `${baseId}_${n}`;
    let name = baseName;
    if (baseName !== undefined && (force || !nameFree(baseName))) {
      let m = Math.max(n, 1);
      while (!nameFree(`${baseName}_${m}`)) m++;
      name = `${baseName}_${m}`;
    }
    reservedIds.add(id);
    if (name !== undefined) reservedNames.add(normName(name));
    return name === undefined
      ? { action: 'copy', target: id }
      : { action: 'copy', target: id, name };
  };

  while (queue.length > 0) {
    const id = queue.shift()!;
    const s = src.get(id);
    if (s === undefined) continue; // dangling reference in the source: left as is
    const type = styleTypeOf(s);
    const expandable = type === 'paragraph' || type === 'character';
    const m = matchOf(tgt, s);
    let d: Decision;
    if (m !== undefined) {
      const tid = styleIdOf(m)!;
      if (fpOf('src', id) === fpOf('tgt', tid)) d = { action: 'same', target: tid };
      else if (mode === 'useDestinationStyles') d = { action: 'dest', target: tid };
      else if (ksf && expandable) d = { action: 'expand', target: tid };
      else {
        const variant = identicalVariant(tgt, tid, type, fpOf('src', id), (v) => fpOf('tgt', v));
        if (variant !== undefined) d = { action: 'same', target: variant };
        else d = { ...unique(tid, styleNameOf(s) ?? tid, true), action: 'rename' };
      }
    } else if (ksf && !ctx.options.smartStyleBehavior && expandable && !listLinked.has(id)) {
      d = { action: 'expand', target: type === 'paragraph' ? tgt.defaultId(type) : undefined };
    } else {
      d = unique(id, styleNameOf(s), false);
    }
    st.decisions.set(id, d);
    if (COPIES.has(d.action)) {
      enqueue(childVal(s, 'basedOn'));
      enqueue(childVal(s, 'next'));
      enqueue(childVal(s, 'link'));
      const numPr = s.element(NS.w, 'pPr')?.element(NS.w, 'numPr');
      if (numPr) enqueueList(numIdOf(numPr));
    } else if (d.action === 'expand') {
      const numPr = src.flat(id).props.pPr.get('numPr');
      const n = numPr ? numIdOf(numPr) : undefined;
      if (n !== undefined && n !== 0) {
        st.expandNumIds.add(n);
        enqueueList(n);
      }
    }
  }

  // Lists of styles that keep the target definition (computed before any copy exists, so the
  // target chains hold target numIds only).
  for (const [id, d] of st.decisions) {
    if ((d.action !== 'same' && d.action !== 'dest') || d.target === undefined) continue;
    if (styleTypeOf(src.get(id)!) !== 'paragraph') continue;
    const sn = src.flat(id).props.pPr.get('numPr');
    const tn = tgt.flat(d.target).props.pPr.get('numPr');
    const a = sn ? numIdOf(sn) : undefined;
    const b = tn ? numIdOf(tn) : undefined;
    if (a !== undefined && b !== undefined && a !== 0 && b !== 0) st.listPairs.push([a, b]);
  }

  // ---- 3. create the copies, in source order ----
  const made: { srcId: string; el: XElement; d: Decision }[] = [];
  for (const s of src.order) {
    const srcId = styleIdOf(s);
    const d = srcId === undefined ? undefined : st.decisions.get(srcId);
    if (srcId === undefined || d === undefined || !COPIES.has(d.action) || d.target === undefined)
      continue;
    if (st.created.has(d.target)) continue; // duplicate styleId in the source
    const el = s.clone();
    el.setAttrNS(NS.w, 'styleId', d.target);
    el.removeAttrNS(NS.w, 'default'); // the target keeps its own defaults
    if (d.name !== undefined && d.name !== styleNameOf(s)) {
      const nameEl = el.element(NS.w, 'name');
      if (nameEl) nameEl.setAttrNS(NS.w, 'val', d.name);
      else insertOrdered(el, wEl(prefixOf(el), 'name', { val: d.name }), STYLE_ORDER);
      // A renamed built-in is no longer the built-in.
      if (!onOffAttr(el.attrNS(NS.w, 'customStyle'))) el.setAttrNS(NS.w, 'customStyle', '1');
      el.element(NS.w, 'aliases')?.remove();
    }
    tgtRoot.append(el);
    tgt.add(el);
    ctx.importedStyles.push(el);
    st.created.add(d.target);
    made.push({ srcId, el, d });
  }

  // ---- 4. the style map ----
  for (const [id, d] of st.decisions) if (d.target !== undefined) ctx.styleMap.set(id, d.target);

  // ---- 5. references inside the copies ----
  for (const { el } of made) {
    for (const local of ['basedOn', 'next', 'link']) {
      const ref = el.element(NS.w, local);
      const v = ref?.attrNS(NS.w, 'val');
      if (ref === undefined || v === undefined) continue;
      const mapped = ctx.styleMap.get(v);
      if (mapped === undefined) {
        ref.remove();
        continue;
      }
      ref.setAttrNS(NS.w, 'val', mapped);
      // A link must be mutual: drop it when it points at a target style linked elsewhere.
      if (local === 'link' && !st.created.has(mapped)) {
        if (childVal(tgt.get(mapped), 'link') !== styleIdOf(el)) ref.remove();
      }
    }
  }
  tgt.invalidate();

  // ---- 6. renamed paragraph styles keep their outline level (TOC) ----
  for (const { srcId, el, d } of made) {
    if (d.action !== 'rename' || styleTypeOf(el) !== 'paragraph') continue;
    const lvl =
      src.flat(srcId).props.pPr.get('outlineLvl')?.attrNS(NS.w, 'val') ??
      headingLevel(styleNameOf(src.get(srcId)!));
    if (lvl === undefined) continue;
    const pPr = el.element(NS.w, 'pPr') ?? insertOrdered(el, wEl(prefixOf(el), 'pPr'), STYLE_ORDER);
    if (pPr.element(NS.w, 'outlineLvl') === undefined) {
      insertOrdered(pPr, wEl(prefixOf(el), 'outlineLvl', { val: lvl }), PPR_ORDER);
    }
  }
  tgt.invalidate();

  // ---- 7. keepSourceFormatting: copies inheriting from expanded styles keep their look ----
  if (ksf) {
    const depth = (id: string): number => {
      let n = 0;
      for (let s = src.get(id); s !== undefined && n < 64; n++) s = src.get(childVal(s, 'basedOn'));
      return n;
    };
    const ordered = made.slice().sort((a, b) => depth(a.srcId) - depth(b.srcId));
    for (const { srcId, el, d } of ordered) {
      const fs = src.flat(srcId).props;
      const ft = tgt.flat(d.target).props;
      for (const c of CONTAINERS) {
        // numPr is left out: copies still carry source numIds until importNumbering runs.
        const diffs = diffMaps(c, withoutNumPr(fs[c]), withoutNumPr(ft[c]), srcList, tgtList);
        if (diffs.length > 0) applyInto(el, c, diffs, STYLE_ORDER);
      }
      tgt.invalidate();
    }
  }

  // ---- 8. the imported content ----
  remapStyleRefs(ctx, blocks);
}

/**
 * Applies the style import to content copied from the source after {@link importStyles} ran
 * (header/footer, note and comment content): keepSourceFormatting expansion, explicit styles for
 * unstyled content whose default style maps elsewhere, and pStyle/rStyle/tblStyle rewritten
 * through ctx.styleMap. Call it once per element, before numIds are remapped (expansion can add
 * source numIds). Without importStyles state it only rewrites through ctx.styleMap.
 */
export function remapStyleRefs(ctx: ImportContext, elements: readonly XElement[]): void {
  const st = STATE.get(ctx);
  const refs: XElement[] = [];
  const paragraphs: XElement[] = [];
  const tables: XElement[] = [];
  for (const root of elements) {
    const all = root.descendants(NS.w);
    if (root.uri === NS.w) all.unshift(root);
    for (const el of all) {
      switch (el.local) {
        case 'pStyle':
        case 'rStyle':
        case 'tblStyle':
          refs.push(el);
          break;
        case 'p':
          paragraphs.push(el);
          break;
        case 'tbl':
          tables.push(el);
          break;
        default:
          break;
      }
    }
  }
  if (st === undefined) {
    for (const r of refs) {
      const m = ctx.styleMap.get(r.attrNS(NS.w, 'val') ?? '');
      if (m !== undefined) r.setAttrNS(NS.w, 'val', m);
    }
    return;
  }

  if (st.ksf) expandAll(ctx, st, paragraphs);

  // Unstyled content uses the default style; say so explicitly when that maps elsewhere.
  const implicit = (type: 'paragraph' | 'table'): string | undefined => {
    const id = st.src.defaultId(type);
    const d = id === undefined ? undefined : st.decisions.get(id);
    const t = d?.target;
    return t !== undefined && t !== st.tgt.defaultId(type) ? t : undefined;
  };
  const p0 = implicit('paragraph');
  if (p0 !== undefined) {
    for (const p of paragraphs) {
      if (p.element(NS.w, 'pPr')?.element(NS.w, 'pStyle') !== undefined) continue;
      const pPr = ensureFirstChild(p, 'pPr');
      insertOrdered(pPr, wEl(prefixOf(pPr), 'pStyle', { val: p0 }), PPR_ORDER);
    }
  }
  const t0 = implicit('table');
  if (t0 !== undefined) {
    for (const t of tables) {
      if (t.element(NS.w, 'tblPr')?.element(NS.w, 'tblStyle') !== undefined) continue;
      const tblPr = ensureFirstChild(t, 'tblPr');
      insertOrdered(tblPr, wEl(prefixOf(tblPr), 'tblStyle', { val: t0 }), TBLPR_ORDER);
    }
  }

  for (const r of refs) {
    const v = r.attrNS(NS.w, 'val');
    const d = v === undefined ? undefined : st.decisions.get(v);
    if (d === undefined) continue;
    if (d.target !== undefined) {
      r.setAttrNS(NS.w, 'val', d.target);
      continue;
    }
    // Expanded with no target style: the formatting is direct now, the reference goes.
    const parent = r.parent;
    r.remove();
    if (parent !== null && parent.nodeType === 1 && parent.firstElement() === undefined) {
      if (parent._attrs.length === 0) parent.remove();
    }
  }
}

/**
 * [source numId, target numId] pairs of paragraph styles that kept the target's definition while
 * both define a list (for importNumbering: the source list joins the target's).
 */
export function styleListPairs(ctx: ImportContext): readonly (readonly [number, number])[] {
  return STATE.get(ctx)?.listPairs ?? [];
}

/** Source numIds written as direct formatting by keepSourceFormatting expansion (for importNumbering). */
export function expandedNumIds(ctx: ImportContext): ReadonlySet<number> {
  return STATE.get(ctx)?.expandNumIds ?? new Set();
}

// ---- matching ---------------------------------------------------------------------------

/** The target style a source style matches (see the header), or undefined. */
function matchOf(tgt: StyleIndex, s: XElement): XElement | undefined {
  const id = styleIdOf(s)!;
  const type = styleTypeOf(s);
  const name = styleNameOf(s);
  const byIdEl = tgt.get(id);
  const byId = byIdEl !== undefined && styleTypeOf(byIdEl) === type ? byIdEl : undefined;
  const byIdName = byId === undefined ? undefined : styleNameOf(byId);
  if (
    byId !== undefined &&
    (name === undefined || byIdName === undefined || normName(name) === normName(byIdName))
  ) {
    return byId;
  }
  const byName = name === undefined ? undefined : tgt.findByName(name, type);
  return byName ?? byId;
}

/** An existing `<id>_<n>` copy of the same type with fingerprint `fp` (from an earlier import). */
function identicalVariant(
  tgt: StyleIndex,
  id: string,
  type: string,
  fp: string,
  fpOf: (id: string) => string,
): string | undefined {
  const prefix = `${id}_`;
  for (const s of tgt.order) {
    const sid = styleIdOf(s);
    if (sid === undefined || !sid.startsWith(prefix) || !/^\d+$/.test(sid.slice(prefix.length))) {
      continue;
    }
    if (styleTypeOf(s) === type && fpOf(sid) === fp) return sid;
  }
  return undefined;
}

/** Styles a list refers to: numStyleLink/styleLink of its abstractNum, pStyle of its levels. */
function listStyleRefs(nums: NumberingIndex, numId: number | undefined): string[] {
  const out: string[] = [];
  if (numId === undefined || numId === 0) return out;
  const num = nums.nums.get(numId);
  if (num === undefined) return out;
  const abs = nums.abstractOfNum(numId);
  for (const el of [abs, num]) {
    if (el === undefined) continue;
    for (const local of ['numStyleLink', 'styleLink']) {
      const v = childVal(el, local);
      if (v !== undefined) out.push(v);
    }
    for (const lvl of el.descendants(NS.w, 'lvl')) {
      const v = childVal(lvl, 'pStyle');
      if (v !== undefined) out.push(v);
    }
  }
  return out;
}

/**
 * The flattened style, with the outline level a built-in heading implies ("heading 2" → 1) made
 * explicit: renamed copies carry it explicitly, and must still compare equal to the original.
 */
function withImplicitOutline(idx: StyleIndex, id: string): FlatStyle {
  const flat = idx.flat(id);
  if (flat.props.pPr.has('outlineLvl')) return flat;
  const lvl = headingLevel(styleNameOf(idx.get(id)!));
  if (lvl === undefined) return flat;
  const pPr = new Map(flat.props.pPr);
  pPr.set('outlineLvl', wEl('w', 'outlineLvl', { val: lvl }));
  return { props: { ...flat.props, pPr }, cond: flat.cond };
}

function headingLevel(name: string | undefined): string | undefined {
  const m = name === undefined ? null : /^heading ([1-9])$/i.exec(name.trim());
  return m ? String(Number(m[1]) - 1) : undefined;
}

// ---- keepSourceFormatting expansion -----------------------------------------------------

interface Diff {
  readonly pPr: XElement[];
  readonly rPr: XElement[];
}

function withoutNumPr(m: PropMap): PropMap {
  if (!m.has('numPr')) return m;
  const out = new Map(m);
  out.delete('numPr');
  return out;
}

/** Writes the source-vs-target difference of every paragraph, paragraph mark and run as direct formatting. */
function expandAll(ctx: ImportContext, st: StyleState, paragraphs: readonly XElement[]): void {
  const { src, tgt } = st;
  const srcList = listKeyFor(st.srcNums, src);
  const tgtList = listKeyFor(new NumberingIndex(relatedPart(ctx.target, 'numbering')), tgt);
  const cache = new Map<string, Diff>();

  /** Target style standing in for source style `id` of `type`. */
  const targetOf = (id: string | undefined, type: string): string | undefined => {
    if (id === undefined) return type === 'character' ? undefined : tgt.defaultId(type);
    const d = st.decisions.get(id);
    if (d !== undefined) {
      return d.target ?? (type === 'character' ? undefined : tgt.defaultId(type));
    }
    return tgt.get(id) !== undefined ? id : undefined;
  };

  const diffFor = (
    srcT: string | undefined,
    srcP: string | undefined,
    srcC: string | undefined,
  ): Diff => {
    const key = `${srcT ?? ''}\u0000${srcP ?? ''}\u0000${srcC ?? ''}`;
    const hit = cache.get(key);
    if (hit) return hit;
    const tgtT = srcT === undefined ? undefined : targetOf(srcT, 'table');
    const tgtP = targetOf(srcP, 'paragraph');
    const tgtC = srcC === undefined ? undefined : targetOf(srcC, 'character');
    const side = (
      idx: StyleIndex,
      t: string | undefined,
      p: string | undefined,
      c: string | undefined,
    ): { pPr: PropMap; rPr: PropMap } => {
      const ft = idx.flat(t).props;
      const fp = idx.flat(p).props;
      const pPr = mergeMaps(mergeMaps(idx.docPPr, ft.pPr), fp.pPr);
      const rPr = mergeMaps(
        mergeMaps(mergeMaps(idx.docRPr, ft.rPr), fp.rPr),
        idx.flat(c).props.rPr,
        true,
      );
      return { pPr, rPr };
    };
    const s = side(src, srcT, srcP, srcC);
    const t = side(tgt, tgtT, tgtP, tgtC);
    // A copied style carries the source's lists itself (renumbered by importNumbering).
    const copied = tgtP !== undefined && st.created.has(tgtP);
    const d: Diff = {
      pPr: diffMaps(
        'pPr',
        copied ? withoutNumPr(s.pPr) : s.pPr,
        copied ? withoutNumPr(t.pPr) : t.pPr,
        srcList,
        tgtList,
      ),
      rPr: diffMaps('rPr', s.rPr, t.rPr, srcList, tgtList),
    };
    cache.set(key, d);
    return d;
  };

  for (const p of paragraphs) {
    const pPr = p.element(NS.w, 'pPr');
    const srcP = childVal(pPr, 'pStyle') ?? src.defaultId('paragraph');
    const tbl = p.closest(NS.w, 'tbl');
    const srcT =
      tbl === undefined
        ? undefined
        : (childVal(tbl.element(NS.w, 'tblPr'), 'tblStyle') ?? src.defaultId('table'));
    const base = diffFor(srcT, srcP, undefined);
    for (const el of base.pPr) {
      const n = el.local === 'numPr' ? numIdOf(el) : undefined;
      if (n !== undefined && n !== 0) st.expandNumIds.add(n);
    }
    applyDirect(p, 'pPr', base.pPr);
    // The paragraph mark (list label, empty-paragraph height).
    const markStyle = childVal(pPr?.element(NS.w, 'rPr'), 'rStyle');
    const mark = markStyle === undefined ? base : diffFor(srcT, srcP, markStyle);
    if (mark.rPr.length > 0) applyDirect(ensureFirstChild(p, 'pPr'), 'rPr', mark.rPr);
    for (const r of p.descendants(NS.w, 'r')) {
      if (r.closest(NS.w, 'p') !== p) continue;
      const rs = childVal(r.element(NS.w, 'rPr'), 'rStyle');
      const d = rs === undefined ? base : diffFor(srcT, srcP, rs);
      applyDirect(r, 'rPr', d.rPr);
    }
  }
}

/**
 * Adds `diffs` to the direct properties of `owner` (`w:p` → pPr, `w:r` → rPr, `w:pPr` → the
 * mark's rPr). Direct formatting already there wins; containers are created only when needed.
 */
function applyDirect(owner: XElement, local: 'pPr' | 'rPr', diffs: readonly XElement[]): void {
  if (diffs.length === 0) return;
  let container = owner.element(NS.w, local);
  const todo: XElement[] = [];
  for (const d of diffs) {
    const existing = container?.element(d.uri, d.local);
    if (existing === undefined) todo.push(d);
    else fillMissing(existing, d);
  }
  if (todo.length === 0) return;
  container ??=
    owner.local === 'pPr' ? ensureChild(owner, 'rPr', PPR_ORDER) : ensureFirstChild(owner, local);
  const order = local === 'pPr' ? PPR_ORDER : RPR_ORDER;
  for (const d of todo) insertOrdered(container, d.clone(), order);
}

/** Adds `diffs` to a style's own `c` container (created at its schema position when missing). */
function applyInto(
  style: XElement,
  c: Container,
  diffs: readonly XElement[],
  order: ChildOrder,
): void {
  let container = style.element(NS.w, c);
  const todo: XElement[] = [];
  for (const d of diffs) {
    const existing = container?.element(d.uri, d.local);
    if (existing === undefined) todo.push(d);
    else fillMissing(existing, d);
  }
  if (todo.length === 0) return;
  container ??= insertOrdered(style, wEl(prefixOf(style), c), order);
  for (const d of todo) insertOrdered(container, d.clone(), ORDERS[c]);
}

// ---- parts ------------------------------------------------------------------------------

function prefixOf(el: XElement): string {
  return el.lookupPrefix(NS.w) ?? 'w';
}

/** Relationship type `rel` in the flavour (transitional or Strict) the target already uses. */
export function relTypeFor(ctx: ImportContext, rel: string): string {
  const strict = ctx.target.package.rels
    .all()
    .some((r) => r.type.startsWith('http://purl.oclc.org/ooxml/officeDocument/relationships/'));
  return strict
    ? `http://purl.oclc.org/ooxml/officeDocument/relationships/${rel.slice(rel.lastIndexOf('/') + 1)}`
    : rel;
}

/** WordprocessingML namespace URI as the target's main part declares it (Strict or transitional). */
export function wUriFor(ctx: ImportContext): string {
  const root = rootOf(ctx.target.mainPart);
  return root?.lookupNamespace(root.lookupPrefix(NS.w) ?? 'w') ?? NS.w;
}

/** Adds an empty part named `<main dir>/<file>` (unique) related from the main part. */
export function addMainPart(
  ctx: ImportContext,
  file: string,
  rootLocal: string,
  contentType: string,
  rel: string,
): Part {
  const pkg = ctx.target.package;
  const main = ctx.target.mainPart;
  const dir = partDir(main.name);
  let name = `${dir}${file}`;
  if (pkg.part(name) !== undefined)
    name = pkg.uniqueName(`${dir}${file.replace(/\.xml$/, '{n}.xml')}`);
  const doc = parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:${rootLocal} xmlns:w="${wUriFor(ctx)}"/>`,
  );
  const part = pkg.addPart(name, contentType, doc);
  main.rels.add(relTypeFor(ctx, rel), relativeTarget(main.name, name));
  ctx.addedParts.push(name);
  return part;
}

function createStylesPart(ctx: ImportContext): Part {
  return addMainPart(ctx, 'styles.xml', 'styles', CT.styles, REL.styles);
}
