// flatten(): a paragraph's visible text with a map from every character back to the DOM.
//
// The characters come from `walkVisible` (wml/text.ts), the one definition of what a paragraph
// shows, so search can never disagree with `text()`. Each contributing leaf (`w:t`, `w:tab`,
// `w:br`, `w:sym`, `m:t`, ...) becomes a segment [start, end) of the text; offsets are UTF-16
// code units, so a supplementary character counts as 2.
//
// A second pass, with exactly the same skip rules, records the zero-width things that sit
// between characters: range markers (bookmarks, comment ranges, permissions, move ranges),
// `w:proofErr`, and field boundaries (every `w:fldChar` and both edges of a `w:fldSimple`).
// Each gets the position of the character that follows it. Replace uses them to move range
// starts/ends to the match edges, to drop proofing marks and to refuse matches that would cut
// through a field.

import type { Paragraph } from '../wml/paragraph.js';
import { chooseAlternate, walkVisible } from '../wml/text.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { FlatText } from './types.js';

/** One leaf element contributing `text` at [start, end) of the paragraph text. */
export interface Segment {
  readonly node: XElement;
  readonly text: string;
  readonly start: number;
  readonly end: number;
}

/**
 * Zero-width item kinds: the start or end of a range (bookmark, comment, permission, move,
 * custom XML revision range), a proofing mark, or a field boundary.
 */
export type MarkerKind = 'start' | 'end' | 'proof' | 'field';

/** A zero-width item found before the character at `pos`. */
export interface Marker {
  readonly node: XElement;
  readonly pos: number;
  readonly kind: MarkerKind;
}

const START: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'bookmarkStart',
  'commentRangeStart',
  'permStart',
  'moveFromRangeStart',
  'moveToRangeStart',
  'customXmlInsRangeStart',
  'customXmlDelRangeStart',
  'customXmlMoveFromRangeStart',
  'customXmlMoveToRangeStart',
]);

const END: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'bookmarkEnd',
  'commentRangeEnd',
  'permEnd',
  'moveFromRangeEnd',
  'moveToRangeEnd',
  'customXmlInsRangeEnd',
  'customXmlDelRangeEnd',
  'customXmlMoveFromRangeEnd',
  'customXmlMoveToRangeEnd',
]);

/** The `w:` subtrees walkVisible skips (keep in sync with wml/text.ts). */
const SKIP: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'delText',
  'instrText',
  'delInstrText',
  'del',
  'moveFrom',
  'drawing',
  'pict',
  'object',
]);

/** The `w:` text leaves walkVisible reports (when they contribute characters). */
const LEAVES: ReadonlySet<string> = /* @__PURE__ */ new Set([
  't',
  'tab',
  'ptab',
  'br',
  'cr',
  'noBreakHyphen',
  'softHyphen',
  'sym',
]);

/** Marker kind of a `w:` element, if it is one. */
export function markerKind(local: string): MarkerKind | undefined {
  if (START.has(local)) return 'start';
  if (END.has(local)) return 'end';
  if (local === 'proofErr') return 'proof';
  if (local === 'fldChar') return 'field';
  return undefined;
}

/** Exit event of a `w:fldSimple` (its closing edge is a field boundary too). */
class Exit {
  readonly el: XElement;
  constructor(el: XElement) {
    this.el = el;
  }
}

/** A flattened paragraph: text, segments and zero-width markers. */
export class Flat implements FlatText {
  readonly text: string;
  readonly segments: readonly Segment[];
  readonly markers: readonly Marker[];

  constructor(text: string, segments: Segment[], markers: Marker[]) {
    this.text = text;
    this.segments = segments;
    this.markers = markers;
  }

  /** Index of the segment containing character `index`, or -1. */
  segmentIndex(index: number): number {
    const segs = this.segments;
    let lo = 0;
    let hi = segs.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const s = segs[mid]!;
      if (index < s.start) hi = mid - 1;
      else if (index >= s.end) lo = mid + 1;
      else return mid;
    }
    return -1;
  }

  at(index: number): { node: XElement; offset: number } | undefined {
    const i = this.segmentIndex(index);
    if (i < 0) return undefined;
    const s = this.segments[i]!;
    return { node: s.node, offset: index - s.start };
  }
}

/** Flattens any element (normally a `w:p`); see the file comment. */
export function flattenElement(root: XElement): Flat {
  const segments: Segment[] = [];
  let text = '';
  walkVisible(root, (node, t) => {
    segments.push({ node, text: t, start: text.length, end: text.length + t.length });
    text += t;
  });
  const index = new Map<XElement, number>();
  for (let i = 0; i < segments.length; i++) index.set(segments[i]!.node, i);

  const markers: Marker[] = [];
  let pos = 0;
  let seen = 0;
  const stack: (XElement | Exit)[] = [root];
  while (stack.length > 0) {
    const item = stack.pop()!;
    if (item instanceof Exit) {
      markers.push({ node: item.el, pos, kind: 'field' });
      continue;
    }
    const el = item;
    const si = index.get(el);
    if (si !== undefined) {
      pos = segments[si]!.end;
      seen++;
      continue;
    }
    if (el.uri === NS.w) {
      if (LEAVES.has(el.local) || SKIP.has(el.local)) continue;
      const kind = markerKind(el.local);
      if (kind !== undefined) {
        markers.push({ node: el, pos, kind });
        continue;
      }
      if (el.local === 'fldSimple') {
        markers.push({ node: el, pos, kind: 'field' });
        stack.push(new Exit(el));
      }
    } else if (el.uri === NS.mc) {
      if (el.local === 'AlternateContent') {
        const branch = chooseAlternate(el);
        if (branch) stack.push(branch);
        continue;
      }
    } else if (el.uri === NS.m && el.local === 't') {
      continue;
    }
    if (el.local.endsWith('Pr')) continue;
    const kids = el.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
  }
  if (seen !== segments.length) {
    throw new Error('turboword internal error: flatten() disagrees with walkVisible()');
  }
  return new Flat(text, segments, markers);
}

/** Flattens a paragraph: its visible text and the DOM position of every character. */
export function flatten(p: Paragraph): FlatText {
  return flattenElement(p.element);
}
