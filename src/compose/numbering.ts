// importNumbering: compose pipeline step 4 — brings the source's lists into the target.
//
// Scope: every numId the source uses anywhere (all stories, the blocks, the styles importStyles
// copied, the lists keepSourceFormatting expansion writes as direct formatting). numId 0 means
// "no list" and is kept; a numId the source does not define maps to 0.
//
// Modes (Syncfusion's list import options):
//   - lists 'keepSource' (default): each source abstractNum is copied under a new abstractNumId
//     with a fresh, unique `w:nsid`. Word identifies a list across documents by its nsid, so a
//     copy that kept the source's nsid would merge with (and continue) any list Word already knows
//     under that nsid. The nsid is deterministic: an FNV-1a hash of the source nsid and id plus a
//     counter, re-hashed until unused in the target.
//   - lists 'mergeWithDestination': a source list reuses a target abstractNum linked to the same
//     list style (`w:numStyleLink`), or with identical level definitions (fingerprint, styles
//     mapped through ctx.styleMap); it then continues the target's list. Otherwise it is copied.
//   - restartListsAtImport: the first num of each imported list gets `w:lvlOverride` +
//     `w:startOverride` for every numbered level (bullets need none), so numbering starts over.
// Source nums without level overrides that share an abstractNum share one target num, which keeps
// "continues the previous list" exactly as it was; nums with overrides get their own num.
//
// Lists tied to styles that keep the target's definition are joined, not copied:
//   - a list style's definition (`w:styleLink`) is reused when the style mapped to a target style
//     that existed before the import (two definitions for one style would conflict);
//   - the list of a paragraph style mapped onto a target style (same or destination definition)
//     joins the target style's list, so paragraphs numbered directly and paragraphs numbered
//     through the style stay one sequence (as docxcompose does).
// With restartListsAtImport, the first imported paragraph numbered only through such a style gets
// a direct numPr with a restarting num (the content would otherwise continue the target list).
//
// Picture bullets (`w:numPicBullet`, VML imagedata or DrawingML) are copied with the numbering
// part's own relationships through ctx.copier; their drawing ids are renumbered.
//
// Bookkeeping: ctx.numIdMap; numIds rewritten in the blocks and in ctx.importedStyles. The target
// numbering part (content type, relationship) is created when missing. Later content goes through
// remapNumIds.

import { CT, REL } from '../opc/rel-types.js';
import { ensureFirstChild, insertOrdered, PPR_ORDER } from '../wml/order.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { addMainPart, expandedNumIds, styleListPairs } from './styles.js';
import {
  abstractFingerprint,
  childVal,
  fnv1a,
  NumberingIndex,
  newUsage,
  numIdOf,
  relatedPart,
  StyleIndex,
  scanUsage,
  storyRoots,
  styleIdOf,
  wEl,
} from './styles-props.js';
import type { ImportContext } from './types.js';

/** Number formats that need no restart (nothing is counted). */
const UNCOUNTED: ReadonlySet<string> = new Set(['bullet', 'none']);

/**
 * Brings every list the source uses into the target according to `ctx.options.lists` and
 * `restartListsAtImport`, fills `ctx.numIdMap`, and rewrites numIds in `blocks` and
 * `ctx.importedStyles`.
 */
export function importNumbering(ctx: ImportContext, blocks: XElement[]): void {
  const state = importLists(ctx, blocks);
  if (ctx.options.restartListsAtImport) restartStyleLists(ctx, blocks, state);
}

/** What importLists did to the target numbering (for restartStyleLists). */
interface ListState {
  readonly tgt: NumberingIndex;
  /** Target abstractNumIds that already have a restarting num. */
  readonly restarted: Set<number>;
  /** Target abstractNumIds created by this import (fresh lists start over anyway). */
  readonly createdAbs: Set<number>;
}

function importLists(ctx: ImportContext, blocks: XElement[]): ListState | undefined {
  const usage = newUsage();
  scanUsage(storyRoots(ctx.source), usage);
  scanUsage(blocks, usage);
  scanUsage(ctx.importedStyles, usage);
  for (const n of expandedNumIds(ctx)) usage.numIds.add(n);
  const wanted = [...usage.numIds].filter((n) => n !== 0 && !ctx.numIdMap.has(n));
  if (wanted.length === 0) {
    remapNumIds(ctx, [...blocks, ...ctx.importedStyles]);
    return undefined;
  }

  const srcPart = relatedPart(ctx.source, 'numbering');
  const src = new NumberingIndex(srcPart);
  const srcStyles = new StyleIndex(relatedPart(ctx.source, 'styles'));
  const defined: number[] = [];
  for (const n of wanted) {
    if (src.abstractOfNum(n) !== undefined) defined.push(n);
    else {
      ctx.numIdMap.set(n, 0);
      ctx.warn('NUMID_UNDEFINED', `numId ${n} is not defined by the source; imported as no list`);
    }
  }
  if (defined.length === 0 || srcPart === undefined) {
    remapNumIds(ctx, [...blocks, ...ctx.importedStyles]);
    return undefined;
  }

  const tgtPart =
    relatedPart(ctx.target, 'numbering') ??
    addMainPart(ctx, 'numbering.xml', 'numbering', CT.numbering, REL.numbering);
  const tgt = new NumberingIndex(tgtPart);
  const root = tgt.root;
  if (root === undefined) {
    ctx.warn('NUMBERING_UNREADABLE', 'The target numbering part has no w:numbering root');
    return undefined;
  }
  const prefix = root.lookupPrefix(NS.w) ?? 'w';
  const merge = ctx.options.lists === 'mergeWithDestination';
  const restart = ctx.options.restartListsAtImport;

  /** Target styles that existed before this import. */
  const imported = new Set(ctx.importedStyles.map((s) => styleIdOf(s)));
  const tgtStyles = new StyleIndex(relatedPart(ctx.target, 'styles'));
  const preexisting = (id: string): boolean => tgtStyles.get(id) !== undefined && !imported.has(id);
  const mapStyle = (id: string): string => ctx.styleMap.get(id) ?? id;

  const absMap = new Map<number, number>(); // source abstractNumId → target abstractNumId
  const createdAbs = new Set<number>();
  const sharedNum = new Map<number, number>(); // target abstractNumId → num for override-less nums
  const restarted = new Set<number>(); // target abstractNumIds whose restart num exists
  const picMap = new Map<number, number>();
  const forbiddenNsids = new Set<string>(src.nsids);

  // ---- picture bullets ----
  const copyPicBullet = (srcId: number): number | undefined => {
    const known = picMap.get(srcId);
    if (known !== undefined) return known;
    const pb = src.picBullets.get(srcId);
    if (pb === undefined) return undefined;
    const el = pb.clone();
    const id = tgt.nextPicBulletId();
    el.setAttrNS(NS.w, 'numPicBulletId', String(id));
    for (const node of [el, ...el.descendants()]) {
      const a = node._attrs;
      for (let i = 0; i < a.length; i += 3) {
        const uri = a[i + 1];
        const isRel = uri === NS.r || (uri === NS.o && a[i]!.endsWith(':relid'));
        if (!isRel) continue;
        try {
          node.setAttr(a[i]!, ctx.copier.copyRelationship(srcPart, a[i + 2]!, tgtPart));
        } catch (e) {
          ctx.warn('PICTURE_BULLET_REL', `Picture bullet ${srcId}: ${(e as Error).message}`);
        }
      }
    }
    for (const docPr of el.descendants(NS.wp, 'docPr')) {
      docPr.setAttr('id', String(ctx.ids.next('docPr')));
    }
    const last = lastChild(root, 'numPicBullet');
    if (last) root.insertAfter(el, last);
    else root.insertBefore(el, firstElement(root));
    tgt.addPicBullet(el);
    picMap.set(srcId, id);
    return id;
  };

  /** Rewrites the style link and picture bullet of a copied `w:lvl`. */
  const mapLevel = (lvl: XElement): void => {
    const ps = lvl.element(NS.w, 'pStyle');
    const v = ps?.attrNS(NS.w, 'val');
    if (ps && v !== undefined) ps.setAttrNS(NS.w, 'val', mapStyle(v));
    const pic = lvl.element(NS.w, 'lvlPicBulletId');
    const pv = pic?.attrNS(NS.w, 'val');
    if (pic && pv !== undefined) {
      const nid = copyPicBullet(Number.parseInt(pv, 10));
      if (nid !== undefined) pic.setAttrNS(NS.w, 'val', String(nid));
      else pic.remove();
    }
  };

  // ---- abstractNums ----
  const findTarget = (pred: (a: XElement) => boolean): number | undefined => {
    const ids = [...tgt.abstracts.keys()].filter((k) => !createdAbs.has(k)).sort((a, b) => a - b);
    for (const k of ids) if (pred(tgt.abstracts.get(k)!)) return k;
    return undefined;
  };

  const freshNsid = (abs: XElement, absId: number): string => {
    const seed = `${childVal(abs, 'nsid') ?? ''}|${absId}|${srcPart.name}`;
    for (let n = 0; ; n++) {
      const v = fnv1a(`${seed}|${n}`).toString(16).toUpperCase().padStart(8, '0');
      if (!tgt.nsids.has(v) && !forbiddenNsids.has(v)) return v;
    }
  };

  const copyAbstract = (abs: XElement, absId: number): number => {
    const el = abs.clone();
    const id = tgt.nextAbstractId();
    el.setAttrNS(NS.w, 'abstractNumId', String(id));
    let nsid = el.element(NS.w, 'nsid');
    if (nsid === undefined) {
      nsid = wEl(prefix, 'nsid');
      el.prepend(nsid);
    }
    nsid.setAttrNS(NS.w, 'val', freshNsid(abs, absId));
    for (const local of ['styleLink', 'numStyleLink']) {
      const link = el.element(NS.w, local);
      const v = link?.attrNS(NS.w, 'val');
      if (link === undefined || v === undefined) continue;
      const mapped = mapStyle(v);
      // The target style already has its own definition: a second one would conflict.
      if (local === 'styleLink' && preexisting(mapped)) link.remove();
      else link.setAttrNS(NS.w, 'val', mapped);
    }
    for (const lvl of el.elements(NS.w, 'lvl')) mapLevel(lvl);
    const last = lastChild(root, 'abstractNum');
    if (last) root.insertAfter(el, last);
    else root.insertBefore(el, firstChildOf(root, ['num', 'numIdMacAtCleanup']) ?? null);
    tgt.addAbstract(el);
    createdAbs.add(id);
    return id;
  };

  const resolveAbstract = (abs: XElement, absId: number): number => {
    const styleLink = childVal(abs, 'styleLink');
    if (styleLink !== undefined && preexisting(mapStyle(styleLink))) {
      const want = mapStyle(styleLink);
      const t = findTarget((a) => childVal(a, 'styleLink') === want);
      if (t !== undefined) return t;
    }
    if (merge) {
      const link = childVal(abs, 'numStyleLink');
      if (link !== undefined) {
        const want = mapStyle(link);
        const t = findTarget((a) => childVal(a, 'numStyleLink') === want);
        if (t !== undefined) return t;
      } else if (styleLink === undefined && abs.descendants(NS.w, 'lvlPicBulletId').length === 0) {
        const fp = abstractFingerprint(abs, mapStyle);
        const plain = (a: XElement): boolean =>
          childVal(a, 'numStyleLink') === undefined && childVal(a, 'styleLink') === undefined;
        const t = findTarget((a) => plain(a) && abstractFingerprint(a, (s) => s) === fp);
        if (t !== undefined) return t;
      }
    }
    return copyAbstract(abs, absId);
  };

  // ---- nums ----
  const plainNumOf = (absId: number): number | undefined => {
    const ids = [...tgt.nums.keys()].sort((a, b) => a - b);
    for (const k of ids) {
      const n = tgt.nums.get(k)!;
      if (tgt.abstractIdOf(n) === absId && n.element(NS.w, 'lvlOverride') === undefined) return k;
    }
    return undefined;
  };

  const createNum = (
    absId: number,
    overrides: readonly XElement[],
    startFrom?: XElement,
  ): number => {
    const kids = overrides.map((o) => {
      const c = o.clone();
      for (const lvl of c.elements(NS.w, 'lvl')) mapLevel(lvl);
      return c;
    });
    return appendNum(tgt, absId, kids, startFrom);
  };

  // A source list numbering a style that kept the target's definition joins the target list
  // that style uses: paragraphs numbered directly and through the style stay one sequence.
  for (const [sn, tn] of styleListPairs(ctx)) {
    const sNum = src.nums.get(sn);
    const tNum = tgt.nums.get(tn);
    const a = sNum ? src.abstractIdOf(sNum) : undefined;
    const b = tNum ? tgt.abstractIdOf(tNum) : undefined;
    if (a === undefined || b === undefined || absMap.has(a)) continue;
    absMap.set(a, b);
    if (tNum!.element(NS.w, 'lvlOverride') === undefined && !sharedNum.has(b)) {
      sharedNum.set(b, tn);
    }
  }

  for (const n of defined) {
    const num = src.nums.get(n)!;
    const absId = src.abstractIdOf(num)!;
    let tAbs = absMap.get(absId);
    if (tAbs === undefined) {
      tAbs = resolveAbstract(src.abstracts.get(absId)!, absId);
      absMap.set(absId, tAbs);
    }
    const overrides = num.elements(NS.w, 'lvlOverride');
    const startFrom =
      restart && !restarted.has(tAbs) ? src.definitionOfNum(n, srcStyles) : undefined;
    let tNum: number | undefined;
    if (overrides.length === 0 && startFrom === undefined) {
      tNum = sharedNum.get(tAbs);
      // Merging continues the target's own list, unless the import restarts lists.
      if (tNum === undefined && !createdAbs.has(tAbs) && startFrom === undefined) {
        tNum = plainNumOf(tAbs);
      }
    }
    tNum ??= createNum(tAbs, overrides, startFrom);
    if (startFrom !== undefined) restarted.add(tAbs);
    if (overrides.length === 0) sharedNum.set(tAbs, tNum);
    ctx.numIdMap.set(n, tNum);
  }

  remapNumIds(ctx, [...blocks, ...ctx.importedStyles]);

  return { tgt, restarted, createdAbs };
}

/**
 * restartListsAtImport for lists the imported content follows only through its (target)
 * paragraph styles: the first such paragraph of each list gets a direct numPr with a restarting
 * num; the following ones continue it through the style.
 */
function restartStyleLists(
  ctx: ImportContext,
  blocks: readonly XElement[],
  state: ListState | undefined,
): void {
  const tgt = state?.tgt ?? new NumberingIndex(relatedPart(ctx.target, 'numbering'));
  const root = tgt.root;
  if (root === undefined) return;
  const restarted = state?.restarted ?? new Set<number>();
  const createdAbs = state?.createdAbs ?? new Set<number>();
  const prefix = root.lookupPrefix(NS.w) ?? 'w';
  const styles = new StyleIndex(relatedPart(ctx.target, 'styles'));
  for (const block of blocks) {
    const ps = block.descendants(NS.w, 'p');
    if (block.is(NS.w, 'p')) ps.unshift(block);
    for (const p of ps) {
      const pPr = p.element(NS.w, 'pPr');
      if (pPr?.element(NS.w, 'numPr') !== undefined) continue; // numbered directly (or "none")
      const sid = childVal(pPr, 'pStyle') ?? styles.defaultId('paragraph');
      const numPr = styles.flat(sid).props.pPr.get('numPr');
      const n = numPr ? numIdOf(numPr) : undefined;
      const num = n === undefined || n === 0 ? undefined : tgt.nums.get(n);
      const a = num ? tgt.abstractIdOf(num) : undefined;
      if (a === undefined || restarted.has(a) || createdAbs.has(a)) continue;
      const id = appendNum(tgt, a, [], tgt.definitionOfNum(n, styles));
      restarted.add(a);
      const direct = wEl(prefix, 'numPr', undefined, [
        wEl(prefix, 'ilvl', { val: childVal(numPr, 'ilvl') ?? '0' }),
        wEl(prefix, 'numId', { val: String(id) }),
      ]);
      insertOrdered(ensureFirstChild(p, 'pPr'), direct, PPR_ORDER);
    }
  }
}

/**
 * Appends a new `w:num` for abstractNum `absId` (after the last num) with `overrides` as its
 * children and, given `startFrom` (the list's level definitions), restart overrides.
 */
function appendNum(
  tgt: NumberingIndex,
  absId: number,
  overrides: readonly XElement[],
  startFrom: XElement | undefined,
): number {
  const root = tgt.root!;
  const prefix = root.lookupPrefix(NS.w) ?? 'w';
  const id = tgt.nextNumId();
  const el = wEl(prefix, 'num', { numId: String(id) }, [
    wEl(prefix, 'abstractNumId', { val: String(absId) }),
  ]);
  for (const o of overrides) el.append(o);
  if (startFrom) addStartOverrides(el, startFrom, prefix);
  const last = lastChild(root, 'num');
  if (last) root.insertAfter(el, last);
  else root.insertBefore(el, firstChildOf(root, ['numIdMacAtCleanup']) ?? null);
  tgt.addNum(el);
  return id;
}

/**
 * Rewrites `w:numPr/w:numId` in `elements` (and descendants) through ctx.numIdMap; 0 and ids
 * not in the map stay. Call it once per element: ids are not idempotent.
 */
export function remapNumIds(ctx: ImportContext, elements: readonly XElement[]): void {
  if (ctx.numIdMap.size === 0) return;
  for (const root of elements) {
    const list = root.descendants(NS.w, 'numPr');
    if (root.is(NS.w, 'numPr')) list.unshift(root);
    for (const numPr of list) {
      const id = numIdOf(numPr);
      if (id === undefined || id === 0) continue;
      const m = ctx.numIdMap.get(id);
      if (m !== undefined && m !== id) {
        numPr.element(NS.w, 'numId')!.setAttrNS(NS.w, 'val', String(m));
      }
    }
  }
}

/**
 * Adds `w:lvlOverride/w:startOverride` to `num` for every counted level of `def` that has no
 * start override yet (the level's own `w:start`, 0 when absent as the schema says).
 */
function addStartOverrides(num: XElement, def: XElement, prefix: string): void {
  for (const lvl of def.elements(NS.w, 'lvl')) {
    const ilvl = lvl.attrNS(NS.w, 'ilvl');
    if (ilvl === undefined || UNCOUNTED.has(childVal(lvl, 'numFmt') ?? 'decimal')) continue;
    let o = num.elements(NS.w, 'lvlOverride').find((x) => x.attrNS(NS.w, 'ilvl') === ilvl);
    if (o?.element(NS.w, 'startOverride') !== undefined) continue;
    if (o === undefined) {
      o = wEl(prefix, 'lvlOverride', { ilvl });
      const after = num
        .elements(NS.w, 'lvlOverride')
        .find((x) => Number(x.attrNS(NS.w, 'ilvl')) > Number(ilvl));
      num.insertBefore(o, after ?? null);
    }
    o.prepend(wEl(prefix, 'startOverride', { val: childVal(lvl, 'start') ?? '0' }));
  }
}

function lastChild(root: XElement, local: string): XElement | undefined {
  let last: XElement | undefined;
  for (const k of root.elements(NS.w, local)) last = k;
  return last;
}

function firstChildOf(root: XElement, locals: readonly string[]): XElement | undefined {
  for (const k of root.elements(NS.w)) if (locals.includes(k.local)) return k;
  return undefined;
}

function firstElement(root: XElement): XElement | null {
  return root.firstElement() ?? null;
}
