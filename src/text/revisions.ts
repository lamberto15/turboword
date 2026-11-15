// acceptAllRevisions() / rejectAllRevisions(): resolve every tracked change (ECMA-376 Part 1
// §17.13) in a story, or in every story part of a document (body, headers, footers, footnotes,
// endnotes, comments) plus styles and numbering (which can carry formatting changes).
//
// What each revision becomes:
// | revision                                   | accept                    | reject                     |
// |--------------------------------------------|---------------------------|----------------------------|
// | `w:ins` / `w:moveTo` (content)             | unwrapped                 | removed                    |
// | `w:del` / `w:moveFrom` (content)           | removed                   | unwrapped, delText → t     |
// | paragraph mark ins/moveTo (`pPr/rPr/ins`)  | marker dropped            | paragraph merged with next |
// | paragraph mark del/moveFrom                | paragraph merged with next| marker dropped             |
// | `*PrChange`, `tblGridChange`               | record dropped            | old properties restored    |
// | `numberingChange`                          | dropped                   | dropped (no machine form)  |
// | `numPr/ins`                                | marker dropped            | numbering removed          |
// | row `trPr/ins` / `trPr/del`                | marker / row removed      | row removed / marker       |
// | `cellIns` / `cellDel`                      | marker / cell removed     | cell removed / marker      |
// | `cellMerge`                                | vMerge := @vMerge         | vMerge := @vMergeOrig      |
// | move ranges, custom XML revision ranges    | removed                   | removed                    |
// | stray `w:delText`/`w:delInstrText`         | removed                   | renamed `w:t`/`w:instrText`|
//
// A merged paragraph keeps the properties of the paragraph whose mark survives (the next one),
// as in Word. Removed content gives up its bookmarks, comment ranges and permission ranges,
// which are moved out first so they stay balanced. A table left without rows is removed (and a
// cell left without a final paragraph gets an empty one). Contents of the revision records
// themselves (the old properties inside `*Change`) are not traversed.

import type { Part } from '../opc/part.js';
import type { WordDocument } from '../wml/document.js';
import { insertOrdered, makeOrder, pruneEmpty, setWAttr, wElement } from '../wml/order.js';
import { Story } from '../wml/story.js';
import type { XElement, XNode } from '../xml/dom.js';
import { NS } from '../xml/ns.js';

/** Property-change records: their content is the old state, never traversed. */
const PROP_CHANGES: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'rPrChange',
  'pPrChange',
  'sectPrChange',
  'tblPrChange',
  'tblPrExChange',
  'trPrChange',
  'tcPrChange',
  'tblGridChange',
  'numberingChange',
]);

/** Range markers that only annotate revisions. */
const RANGES: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
  'customXmlInsRangeStart',
  'customXmlInsRangeEnd',
  'customXmlDelRangeStart',
  'customXmlDelRangeEnd',
  'customXmlMoveFromRangeStart',
  'customXmlMoveFromRangeEnd',
  'customXmlMoveToRangeStart',
  'customXmlMoveToRangeEnd',
]);

/** Range markers rescued from removed content (they pair with markers elsewhere). */
const KEEP_MARKERS: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'permStart',
  'permEnd',
]);

const MARKS: ReadonlySet<string> = /* @__PURE__ */ new Set(['ins', 'del', 'moveFrom', 'moveTo']);

/** `w:tcPr` children (CT_TcPr, §17.4.70), for placing a restored `w:vMerge`. */
const TCPR_ORDER = /* @__PURE__ */ makeOrder([
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

/**
 * How a property change is rejected: the old properties replace the holder's children except
 * `before` (kept first) and `after` (kept last, the old ones go in front of them).
 */
const RESTORE: Readonly<Record<string, { old: string; before?: string[]; after?: string[] }>> = {
  rPrChange: { old: 'rPr', before: ['ins', 'del', 'moveFrom', 'moveTo'] },
  pPrChange: { old: 'pPr', after: ['rPr', 'sectPr'] },
  sectPrChange: { old: 'sectPr', before: ['headerReference', 'footerReference'] },
  tblPrChange: { old: 'tblPr' },
  tblPrExChange: { old: 'tblPrEx' },
  trPrChange: { old: 'trPr', after: ['ins', 'del'] },
  tcPrChange: { old: 'tcPr', after: ['cellIns', 'cellDel', 'cellMerge'] },
  tblGridChange: { old: 'tblGrid' },
};

function attached(el: XElement, root: XElement): boolean {
  let n: XElement | null = el;
  while (n !== null && n !== root) {
    const p: XElement['parent'] = n.parent;
    n = p !== null && p.nodeType === 1 ? p : null;
  }
  return n === root;
}

function parentEl(el: XElement): XElement | undefined {
  const p = el.parent;
  return p !== null && p.nodeType === 1 ? p : undefined;
}

/** Rejects a property change: puts the old properties back. */
function restore(change: XElement): void {
  const holder = parentEl(change);
  const spec = RESTORE[change.local];
  change.remove();
  if (holder === undefined || spec === undefined) return;
  const keep = new Set([...(spec.before ?? []), ...(spec.after ?? [])]);
  const after = new Set(spec.after ?? []);
  for (const k of holder.elements()) if (!(k.uri === NS.w && keep.has(k.local))) k.remove();
  const old = change.element(NS.w, spec.old);
  if (old === undefined) return;
  const ref = holder.elements().find((k) => k.uri === NS.w && after.has(k.local)) ?? null;
  for (const k of old.elements()) {
    if (k.uri === NS.w && keep.has(k.local)) continue;
    holder.insertBefore(k, ref);
  }
}

/** Removes a marker element and the property containers it leaves empty. */
function dropMarker(el: XElement): void {
  const holder = parentEl(el);
  el.remove();
  if (holder !== undefined) pruneEmpty(holder);
}

function removeTable(tbl: XElement): void {
  const parent = parentEl(tbl);
  tbl.remove();
  if (parent?.is(NS.w, 'tc') && parent.lastElement()?.is(NS.w, 'p') !== true) {
    parent.append(wElement(parent, 'p'));
  }
}

function removeRow(tr: XElement): void {
  const tbl = tr.closest(NS.w, 'tbl');
  tr.remove();
  if (tbl !== undefined && tbl.descendants(NS.w, 'tr').length === 0) removeTable(tbl);
}

function removeCell(tc: XElement): void {
  const tr = tc.closest(NS.w, 'tr');
  tc.remove();
  if (tr !== undefined && tr.descendants(NS.w, 'tc').length === 0) removeRow(tr);
}

/** Moves the bookmarks/comment ranges/permissions out of `el`, then removes it. */
function removeContent(el: XElement): void {
  for (const m of el.descendants(NS.w)) if (KEEP_MARKERS.has(m.local)) el.before(m);
  el.remove();
}

/** Replaces `el` with its children. */
function unwrap(el: XElement): void {
  el.replaceWith(...el.children.slice());
}

/** Turns deleted text and field codes back into live ones. */
function revive(el: XElement): void {
  for (const k of el.descendants(NS.w)) {
    if (k.local === 'delText') k.setName(k.name.replace(/delText$/, 't'), NS.w);
    else if (k.local === 'delInstrText') {
      k.setName(k.name.replace(/delInstrText$/, 'instrText'), NS.w);
    }
  }
}

/** Sets `w:vMerge` of a cell from a `cellMerge` value (`rest`, `cont` or absent). */
function setVMerge(tcPr: XElement, v: string | undefined): void {
  const cur = tcPr.element(NS.w, 'vMerge');
  if (v === undefined) {
    cur?.remove();
    return;
  }
  const el = cur ?? insertOrdered(tcPr, wElement(tcPr, 'vMerge'), TCPR_ORDER);
  if (v === 'rest' || v === 'restart') setWAttr(el, 'val', 'restart');
  else el.removeAttrNS(NS.w, 'val');
}

function hasContent(p: XElement): boolean {
  return p.elements().some((k) => !k.is(NS.w, 'pPr'));
}

/** Merges paragraph `p` (whose mark goes away) into the next paragraph. */
function mergeWithNext(p: XElement): void {
  const next = p.nextElement();
  if (next?.is(NS.w, 'p')) {
    const pPr = p.element(NS.w, 'pPr');
    let ref: XNode | null = next.element(NS.w, 'pPr') ?? null;
    for (const k of p.children.slice()) {
      if (k === pPr) continue;
      next.insertAfter(k, ref);
      ref = k;
    }
    p.remove();
    return;
  }
  // No following paragraph (a table, the end of a cell or story): drop an empty paragraph when
  // a later paragraph still closes its container, otherwise keep it.
  if (hasContent(p)) return;
  let later = next;
  while (later !== undefined && !later.is(NS.w, 'p')) later = later.nextElement();
  if (later !== undefined) p.remove();
}

/** Resolves every revision under `root`; returns how many there were. */
function resolveIn(root: XElement, accept: boolean): number {
  const props: XElement[] = [];
  const numIns: XElement[] = [];
  const cells: XElement[] = [];
  const rows: XElement[] = [];
  const containers: XElement[] = [];
  const marks: XElement[] = [];
  const ranges: XElement[] = [];
  const leaves: XElement[] = [];
  const stack: XElement[] = [root];
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el.uri === NS.w && el !== root) {
      const l = el.local;
      if (PROP_CHANGES.has(l)) {
        props.push(el);
        continue;
      }
      if (MARKS.has(l)) {
        const parent = parentEl(el);
        if (parent?.is(NS.w, 'trPr')) rows.push(el);
        else if (parent?.is(NS.w, 'numPr')) numIns.push(el);
        else if (parent?.is(NS.w, 'rPr') && parentEl(parent)?.is(NS.w, 'pPr')) marks.push(el);
        else containers.push(el);
      } else if (l === 'cellIns' || l === 'cellDel' || l === 'cellMerge') cells.push(el);
      else if (RANGES.has(l)) ranges.push(el);
      else if (l === 'delText' || l === 'delInstrText') leaves.push(el);
    }
    const kids = el.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
  }
  const count =
    props.length + numIns.length + cells.length + rows.length + containers.length + marks.length;

  for (const el of props) {
    if (!attached(el, root)) continue;
    if (accept || el.local === 'numberingChange') el.remove();
    else restore(el);
  }
  for (const el of numIns) {
    if (!attached(el, root)) continue;
    if (accept) dropMarker(el);
    else {
      const numPr = parentEl(el)!;
      const pPr = parentEl(numPr);
      numPr.remove();
      if (pPr !== undefined) pruneEmpty(pPr);
    }
  }
  for (const el of cells) {
    if (!attached(el, root)) continue;
    const tc = el.closest(NS.w, 'tc');
    if (el.local === 'cellMerge') {
      const tcPr = parentEl(el)!;
      const v = el.attrNS(NS.w, accept ? 'vMerge' : 'vMergeOrig');
      el.remove();
      setVMerge(tcPr, v);
      pruneEmpty(tcPr);
    } else if ((el.local === 'cellDel') === accept && tc !== undefined) removeCell(tc);
    else dropMarker(el);
  }
  for (const el of rows) {
    if (!attached(el, root)) continue;
    const tr = el.closest(NS.w, 'tr');
    if ((el.local === 'del' || el.local === 'moveFrom') === accept && tr !== undefined) {
      removeRow(tr);
    } else dropMarker(el);
  }
  for (const el of containers) {
    if (!attached(el, root)) continue;
    const removed = el.local === 'del' || el.local === 'moveFrom';
    if (removed === accept) removeContent(el);
    else {
      if (removed) revive(el);
      unwrap(el);
    }
  }
  for (const el of marks) {
    if (!attached(el, root)) continue;
    const p = parentEl(parentEl(parentEl(el)!)!);
    const goes = (el.local === 'del' || el.local === 'moveFrom') === accept;
    dropMarker(el);
    if (goes && p?.is(NS.w, 'p')) mergeWithNext(p);
  }
  for (const el of ranges) if (attached(el, root)) el.remove();
  for (const el of leaves) {
    if (!attached(el, root) || (el.local !== 'delText' && el.local !== 'delInstrText')) continue;
    if (accept) el.remove();
    else revive(parentEl(el)!);
  }
  return count;
}

/** Root elements of the parts that can hold revisions. */
function partRoots(doc: WordDocument): XElement[] {
  const parts = new Set<Part>([doc.mainPart]);
  for (const kind of [
    'header',
    'footer',
    'footnotes',
    'endnotes',
    'comments',
    'styles',
    'numbering',
  ]) {
    for (const p of doc.mainPart.relatedByType(kind)) parts.add(p);
  }
  const out: XElement[] = [];
  for (const part of parts) {
    if (!part.isXml) continue;
    for (const k of part.xml().children) {
      if (k.nodeType === 1) {
        out.push(k);
        break;
      }
    }
  }
  return out;
}

function resolve(target: WordDocument | Story, accept: boolean): number {
  const roots = target instanceof Story ? [target.element] : partRoots(target);
  let n = 0;
  for (const r of roots) n += resolveIn(r, accept);
  return n;
}

/**
 * Accepts every tracked change in `target` (a story, or the whole document: every story part
 * plus styles and numbering): insertions stay, deletions go, moves land, formatting changes
 * become final, deleted paragraph marks merge paragraphs, deleted rows/cells go. Returns the
 * number of revisions resolved.
 */
export function acceptAllRevisions(target: WordDocument | Story): number {
  return resolve(target, true);
}

/**
 * Rejects every tracked change in `target`: insertions go, deletions come back (deleted text
 * and field codes revive), moves return, old formatting is restored, inserted paragraph marks
 * merge paragraphs, inserted rows/cells go. Returns the number of revisions resolved.
 */
export function rejectAllRevisions(target: WordDocument | Story): number {
  return resolve(target, false);
}
