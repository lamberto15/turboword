// Rewriting the content of a content control (and unwrapping one).
//
// The rules Word follows when a control receives a value, and that we reproduce:
// - the new text takes the formatting of the content's first run (or, while the placeholder is
//   shown, the control's own `w:sdtPr/w:rPr`), minus the placeholder character style;
// - block-level content keeps its first paragraph (and so its paragraph properties and ids);
//   extra paragraphs are shallow copies of it without `w14:paraId`/`w14:textId`;
// - zero-width markers inside the old content (bookmarks, comment ranges, permission ranges,
//   comment references) are kept, starts at the beginning and ends at the end, so ranges that
//   cross the control's boundary stay balanced; a section break carried by a dropped paragraph
//   moves to the last new one;
// - containers that must hold a paragraph (cells, the body, headers, notes...) keep one.

import { ensureFirstChild, insertOrdered, PPR_ORDER, wElement } from '../wml/order.js';
import type { XElement, XNode } from '../xml/dom.js';
import { XText } from '../xml/dom.js';
import { NS } from '../xml/ns.js';

/** Where a content control sits. */
export type Level = 'block' | 'inline' | 'row' | 'cell';

/** Elements a level walk passes through to reach the deciding ancestor. */
const PASS: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'sdtContent',
  'sdt',
  'customXml',
  'smartTag',
  'hyperlink',
  'ins',
  'del',
  'moveTo',
  'moveFrom',
  'fldSimple',
  'dir',
  'bdo',
]);

/** Level of a `w:sdt`: decided by its nearest paragraph, table, row or other container. */
export function levelOf(sdt: XElement): Level {
  for (let p = sdt.parent; p !== null && p.nodeType === 1; p = p.parent) {
    if (p.uri === NS.w) {
      if (p.local === 'p') return 'inline';
      if (p.local === 'tbl') return 'row';
      if (p.local === 'tr') return 'cell';
      if (!PASS.has(p.local)) return 'block';
    } else if (p.uri === NS.m) {
      return 'inline';
    } else if (p.uri !== NS.mc) {
      return 'block';
    }
  }
  return 'block';
}

/** The `w:sdtContent` of `sdt`, created when missing. */
export function ensureContent(sdt: XElement): XElement {
  let c = sdt.element(NS.w, 'sdtContent');
  if (c === undefined) {
    c = wElement(sdt, 'sdtContent');
    sdt.append(c);
  }
  return c;
}

/** True when `el` lies (strictly below `top`) inside deleted content or an `mc:Fallback`. */
function hidden(el: XElement, top: XElement): boolean {
  for (let p = el.parent; p !== null && p !== top && p.nodeType === 1; p = p.parent) {
    if (p.uri === NS.mc && p.local === 'Fallback') return true;
    if (p.uri === NS.w && (p.local === 'del' || p.local === 'moveFrom')) return true;
  }
  return false;
}

/** First visible `w:r` under `root`. */
export function firstRun(root: XElement | undefined): XElement | undefined {
  if (root === undefined) return undefined;
  return root.descendants(NS.w, 'r').find((r) => !hidden(r, root));
}

/** Run content for `text`: `w:t` pieces, TAB → `w:tab`, LF/CR/CRLF → `w:br`. */
function runNodes(scope: XElement, text: string): XElement[] {
  const out: XElement[] = [];
  for (const piece of text.split(/(\t|\r\n|\r|\n)/)) {
    if (piece === '') continue;
    if (piece === '\t' || piece === '\n' || piece === '\r' || piece === '\r\n') {
      out.push(wElement(scope, piece === '\t' ? 'tab' : 'br'));
      continue;
    }
    const t = wElement(scope, 't');
    if (/^\s|\s$/.test(piece)) t.setAttr('xml:space', 'preserve');
    t.append(new XText(piece));
    out.push(t);
  }
  return out;
}

/** A new `w:r` with a copy of `rPr` and `text`, named in `scope`'s prefixes. */
export function makeRun(scope: XElement, rPr: XElement | undefined, text: string): XElement {
  const r = wElement(scope, 'r');
  if (rPr !== undefined) r.append(rPr.clone());
  for (const n of runNodes(scope, text)) r.append(n);
  return r;
}

/** Detaches the markers to keep from `root` (see the header). */
function takeMarkers(root: XElement): { starts: XElement[]; ends: XElement[] } {
  const starts: XElement[] = [];
  const ends: XElement[] = [];
  for (const el of root.descendants()) {
    if (el.uri !== NS.w) continue;
    const l = el.local;
    if (l === 'bookmarkStart' || l === 'commentRangeStart' || l === 'permStart') starts.push(el);
    else if (l === 'bookmarkEnd' || l === 'commentRangeEnd' || l === 'permEnd') ends.push(el);
    else if (l === 'commentReference') {
      const r = el.closest(NS.w, 'r');
      if (r !== undefined && !ends.includes(r)) ends.push(r);
    }
  }
  for (const m of [...starts, ...ends]) m.remove();
  return { starts, ends };
}

/**
 * Replaces the content of `sdt` with `paras`: one list of inline nodes (runs) per paragraph.
 * Inline controls get the nodes of every entry in one flow (callers pass one entry). Row and
 * cell controls write into their first cell.
 */
export function replaceContent(sdt: XElement, level: Level, paras: XNode[][]): void {
  const content = ensureContent(sdt);
  if (paras.length === 0) paras = [[]];
  if (level === 'inline') {
    const kept = takeMarkers(content);
    content.clear();
    content.append(...kept.starts, ...paras.flat(), ...kept.ends);
    return;
  }
  const container = level === 'block' ? content : (content.descendants(NS.w, 'tc')[0] ?? content);
  const template = container.element(NS.w, 'p') ?? container.descendants(NS.w, 'p')[0];
  const reuse = template?.parent === container ? template : undefined;
  let sectPr: XElement | undefined;
  for (const p of container.descendants(NS.w, 'p')) {
    sectPr = p.element(NS.w, 'pPr')?.element(NS.w, 'sectPr') ?? sectPr;
  }
  const kept = takeMarkers(container);
  for (const k of container.children.slice()) {
    if (k === reuse || (k.nodeType === 1 && k.is(NS.w, 'tcPr'))) continue;
    k.remove();
  }
  const out: XElement[] = [];
  for (let i = 0; i < paras.length; i++) {
    let p: XElement;
    if (i === 0 && reuse !== undefined) {
      p = reuse;
      for (const k of p.children.slice()) if (!(k.nodeType === 1 && k.is(NS.w, 'pPr'))) k.remove();
    } else {
      p = template !== undefined ? template.clone(false) : wElement(sdt, 'p');
      p.removeAttrNS(NS.w14, 'paraId');
      p.removeAttrNS(NS.w14, 'textId');
      const pPr = template?.element(NS.w, 'pPr');
      if (pPr !== undefined) {
        const copy = pPr.clone();
        copy.element(NS.w, 'sectPr')?.remove();
        p.append(copy);
      }
      container.append(p);
    }
    p.append(...paras[i]!);
    out.push(p);
  }
  const first = out[0]!;
  const last = out[out.length - 1]!;
  const fpPr = first.element(NS.w, 'pPr');
  for (const s of kept.starts.reverse()) {
    if (fpPr !== undefined) first.insertAfter(s, fpPr);
    else first.prepend(s);
  }
  last.append(...kept.ends);
  if (sectPr !== undefined && sectPr.parent?.parent !== last) {
    insertOrdered(ensureFirstChild(last, 'pPr'), sectPr, PPR_ORDER);
  }
}

/** Containers that must hold at least one paragraph. */
const NEEDS_P: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'tc',
  'body',
  'hdr',
  'ftr',
  'footnote',
  'endnote',
  'comment',
  'txbxContent',
  'docPartBody',
]);

/** Block-level children that can end a cell. */
const BLOCKS: ReadonlySet<string> = /* @__PURE__ */ new Set(['p', 'tbl', 'sdt', 'customXml']);

/** True when the last block of `tc` (looking into content controls and custom XML) is a `w:p`. */
function endsWithP(tc: XElement): boolean {
  let c: XElement | undefined = tc;
  while (c !== undefined) {
    const kids: XElement[] = c.elements(NS.w).filter((k) => BLOCKS.has(k.local));
    const last: XElement | undefined = kids[kids.length - 1];
    if (last === undefined) return false;
    if (last.local === 'p') return true;
    if (last.local === 'tbl') return false;
    c = last.local === 'sdt' ? last.element(NS.w, 'sdtContent') : last;
  }
  return false;
}

/**
 * After content was removed under `start`: drops tables left without rows and gives the nearest
 * paragraph container an empty paragraph when it has none left (a cell: when it no longer ends
 * with one).
 */
export function ensureValid(start: XElement): void {
  let el: XElement | null = start;
  while (el !== null) {
    const up: XElement | null = el.parent?.nodeType === 1 ? el.parent : null;
    if (el.uri === NS.w) {
      if (el.local === 'tbl' && el.descendants(NS.w, 'tr').length === 0) {
        el.remove();
      } else if (NEEDS_P.has(el.local)) {
        const ok = el.local === 'tc' ? endsWithP(el) : el.descendants(NS.w, 'p').length > 0;
        if (!ok) {
          const p = wElement(el, 'p');
          const last = el.lastElement();
          if (last?.is(NS.w, 'sectPr')) el.insertBefore(p, last);
          else el.append(p);
        }
        return;
      }
    }
    el = up;
  }
}

/** Replaces `sdt` with its content (no-op when detached). */
export function unwrap(sdt: XElement): void {
  const parent = sdt.parent;
  if (parent === null || parent.nodeType !== 1) return;
  const content = sdt.element(NS.w, 'sdtContent');
  const kids = content === undefined ? [] : content.children.slice();
  sdt.replaceWith(...kids);
  if (kids.length === 0) ensureValid(parent);
}
