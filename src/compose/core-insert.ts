// Where insertDocument / importBlocks put the imported blocks, and the paragraph surgery that
// makes room for them.
//
// Every target resolves to an InsertionPoint: a block container (`w:body`, `w:tc`,
// `w:sdtContent`, a header root, ...) and the child before which the blocks go. Inline targets
// (a placeholder, a bookmark or a content control inside a paragraph) split the paragraph:
// - the split is a classic DOM range split: every ancestor between the split point and the
//   paragraph (runs, hyperlinks, insertions, inline content controls, ...) is cut in two, the
//   right halves being shallow copies that keep the property children (`w:pPr`, `w:rPr`,
//   `w:sdtPr`, ...), so both halves keep their formatting;
// - ids copied into the right halves (paraId/textId, revision and content-control ids) are
//   replaced with fresh ones, and the section properties stay with the paragraph mark, i.e. on
//   the right half;
// - a half left without content (only properties) is removed after the insertion, unless it
//   carries the section properties.
// The placeholder text itself is removed, but markers inside it (bookmark and comment ranges,
// references, field characters) are kept on the left half so pairs stay balanced.

import { UsageError } from '../core/errors.js';
import type { Part } from '../opc/part.js';
import { findText } from '../text/index.js';
import type { WordDocument } from '../wml/document.js';
import type { DocumentIds } from '../wml/ids.js';
import { REVISION_ELEMENTS } from '../wml/ids.js';
import type { Story } from '../wml/story.js';
import { walkVisible } from '../wml/text.js';
import { ELEMENT_NODE, XElement, type XNode, XText } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { InsertTarget } from './types.js';

/** Where imported blocks go. */
export interface InsertionPoint {
  /** The block container receiving the blocks. */
  readonly container: XElement;
  /** The child the blocks are inserted before (null: at the end, before a body's final sectPr). */
  readonly before: XNode | null;
  /** The story part holding the container. */
  readonly part: Part;
  /** Paragraphs created by a split, removed after the insertion when left empty. */
  readonly halves: readonly XElement[];
  /** The container is body-level (sections can be imported there). */
  readonly bodyLevel: boolean;
}

/** Markers that must survive the removal of the placeholder text (balanced pairs, references). */
const MARKERS: ReadonlySet<string> = new Set([
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
  'permStart',
  'permEnd',
  'customXmlInsRangeStart',
  'customXmlInsRangeEnd',
  'customXmlDelRangeStart',
  'customXmlDelRangeEnd',
  'customXmlMoveFromRangeStart',
  'customXmlMoveFromRangeEnd',
  'customXmlMoveToRangeStart',
  'customXmlMoveToRangeEnd',
]);
/** Run content that must survive with its run (references and field structure). */
const KEEP_RUN_CONTENT: ReadonlySet<string> = new Set([
  'commentReference',
  'footnoteReference',
  'endnoteReference',
  'fldChar',
]);
const BLOCKS: ReadonlySet<string> = new Set(['p', 'tbl', 'sdt', 'customXml', 'altChunk']);

const isEl = (n: XNode | null | undefined): n is XElement => n?.nodeType === ELEMENT_NODE;

/** A shallow copy of `el` with deep copies of its leading property children (`*Pr`). */
function shellOf(el: XElement): XElement {
  const shell = el.clone(false);
  for (const k of el._children) {
    if (!isEl(k)) continue;
    if (!k.local.endsWith('Pr')) break;
    shell.append(k.clone());
  }
  return shell;
}

/** Fresh paraId/textId, revision and content-control ids for the copied parts of a shell. */
function freshen(shell: XElement, ids: DocumentIds): void {
  for (const a of shell.attributes()) {
    if (a.uri !== NS.w14) continue;
    if (a.name.endsWith(':paraId')) shell.setAttr(a.name, ids.nextHex('paraId'));
    else if (a.name.endsWith(':textId') && a.value !== '77777777') {
      shell.setAttr(a.name, ids.nextHex('textId'));
    }
  }
  const copied: XElement[] = [shell];
  for (const k of shell._children) {
    if (isEl(k) && k.local.endsWith('Pr')) copied.push(k, ...k.descendants());
  }
  for (const el of copied) {
    if (el.uri !== NS.w) continue;
    if (REVISION_ELEMENTS.has(el.local) && el.attrNS(NS.w, 'id') !== undefined) {
      el.setAttrNS(NS.w, 'id', String(ids.next('revision')));
    } else if (el.local === 'id' && (el.parent as XElement | null)?.is(NS.w, 'sdtPr')) {
      el.setAttrNS(NS.w, 'val', String(ids.next('sdt')));
    }
  }
}

/**
 * Splits `root` (a `w:p`) so that `first` and everything after it (in document order) moves to a
 * new paragraph inserted right after `root`; `first` null = nothing moves. Returns the new
 * paragraph. Ancestors between `first` and `root` are cut in two (see the file header).
 */
export function splitBefore(root: XElement, first: XNode | null, ids: DocumentIds): XElement {
  const shell = shellOf(root);
  const shells: XElement[] = [shell];
  if (first !== null) {
    let cut: XNode = first;
    let inclusive = true;
    let built: XElement | null = null;
    let level = first.parent as XElement;
    for (;;) {
      const right = level === root ? shell : shellOf(level);
      if (right !== shell) shells.push(right);
      if (built !== null) right.append(built);
      const kids = level._children.slice();
      let i = kids.indexOf(cut);
      if (!inclusive) i++;
      for (let j = i; j < kids.length; j++) right.append(kids[j]!);
      if (level === root) break;
      built = right;
      cut = level;
      inclusive = false;
      level = level.parent as XElement;
    }
  }
  root.after(shell);
  for (const s of shells) freshen(s, ids);
  // The paragraph mark (and its section properties / mark revision) now belongs to the right half.
  const pPr = root.element(NS.w, 'pPr');
  pPr?.element(NS.w, 'sectPr')?.remove();
  const markRPr = pPr?.element(NS.w, 'rPr');
  if (markRPr !== undefined) {
    for (const r of markRPr.elements(NS.w)) {
      if (r.local === 'ins' || r.local === 'del' || r.local === 'moveFrom' || r.local === 'moveTo')
        r.remove();
    }
  }
  return shell;
}

/** The node right after `node` in document order within `root` (null when it is the last). */
function nextInRoot(root: XElement, node: XNode): XNode | null {
  let n: XNode = node;
  for (;;) {
    const p = n.parent as XElement | null;
    if (p === null) return null;
    const kids = p._children;
    const i = kids.indexOf(n);
    if (i + 1 < kids.length) return kids[i + 1]!;
    if (p === root) return null;
    n = p;
  }
}

/**
 * Splits paragraph `p` at visible-text offset `offset` (see wml/text.ts) and returns the new
 * right half (text from `offset` on). Splitting inside a `w:t` cuts its text in two runs.
 */
export function splitAtOffset(p: XElement, offset: number, ids: DocumentIds): XElement {
  let first: XNode | null | undefined;
  let pos = 0;
  let last: XElement | undefined;
  walkVisible(p, (node, text) => {
    last = node;
    if (first !== undefined) return;
    if (offset < pos + text.length) {
      const inner = offset - pos;
      if (inner > 0 && node.is(NS.w, 't')) {
        const t2 = node.clone(false);
        t2.append(new XText(text.slice(inner)));
        node.text = text.slice(0, inner);
        preserveSpace(node);
        preserveSpace(t2);
        node.after(t2);
        first = t2;
      } else first = node;
    }
    pos += text.length;
  });
  if (first === undefined) first = last === undefined ? firstContent(p) : nextInRoot(p, last);
  return splitBefore(p, first, ids);
}

/** The first child of `p` that is not its `w:pPr` (null when none). */
function firstContent(p: XElement): XNode | null {
  for (const k of p._children) if (!(isEl(k) && k.is(NS.w, 'pPr'))) return k;
  return null;
}

function preserveSpace(t: XElement): void {
  const text = t.text;
  if (text !== text.trim() && t.attrNS(NS.xml, 'space') === undefined) {
    t.setAttr('xml:space', 'preserve');
  }
}

/** True when `p` has nothing but its properties (and proofing marks). */
function isEmptyParagraph(p: XElement): boolean {
  for (const k of p._children) {
    if (!isEl(k)) {
      if (k.nodeType === 3 && k.text.trim() !== '') return false;
      continue;
    }
    if (k.uri === NS.w && (k.local === 'pPr' || k.local === 'proofErr')) continue;
    return false;
  }
  return true;
}

/** Removes the halves of a split left empty (a half holding section properties stays). */
export function dropEmptyHalves(halves: readonly XElement[]): void {
  for (const h of halves) {
    if (h.parent === null || !isEmptyParagraph(h)) continue;
    if (h.element(NS.w, 'pPr')?.element(NS.w, 'sectPr') !== undefined) continue;
    h.remove();
  }
}

/**
 * Removes the paragraph `mid` (the placeholder text), first moving its markers and the runs
 * holding references or field characters to the end of `left`.
 */
function removeKeepingMarkers(mid: XElement, left: XElement): void {
  const keep: XElement[] = [];
  const stack: XElement[] = [];
  for (let i = mid._children.length - 1; i >= 0; i--) {
    const k = mid._children[i]!;
    if (isEl(k)) stack.push(k);
  }
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el.uri === NS.w) {
      if (el.local === 'pPr') continue;
      if (MARKERS.has(el.local)) {
        keep.push(el);
        continue;
      }
      if (el.local === 'r') {
        if (el.elements(NS.w).some((c) => KEEP_RUN_CONTENT.has(c.local))) {
          const run = el.clone(false);
          for (const c of el.elements(NS.w)) {
            if (c.local === 'rPr' || KEEP_RUN_CONTENT.has(c.local)) run.append(c.clone());
          }
          keep.push(run);
        }
        continue;
      }
    }
    for (let i = el._children.length - 1; i >= 0; i--) {
      const k = el._children[i]!;
      if (isEl(k)) stack.push(k);
    }
  }
  for (const k of keep) left.append(k);
  mid.remove();
}

/** True when `el` is the body or nested in it only through block-level sdt/customXml wrappers. */
function isBodyLevel(el: XElement): boolean {
  let cur: XElement | null = el;
  while (cur !== null) {
    if (cur.is(NS.w, 'body')) return true;
    if (
      cur.uri !== NS.w ||
      !(cur.local === 'sdtContent' || cur.local === 'sdt' || cur.local === 'customXml')
    ) {
      return false;
    }
    cur = cur.parent?.nodeType === ELEMENT_NODE ? (cur.parent as XElement) : null;
  }
  return false;
}

function point(
  doc: WordDocument,
  container: XElement,
  before: XNode | null,
  part: Part,
  halves: XElement[] = [],
): InsertionPoint {
  return {
    container,
    before,
    part,
    halves,
    bodyLevel: part === doc.mainPart && isBodyLevel(container),
  };
}

/** The paragraph containing `node` (closest `w:p`), or undefined for block-level nodes. */
function paragraphOf(node: XElement): XElement | undefined {
  return node.closest(NS.w, 'p');
}

/** Stories searched for inline targets, in order: the body, then headers and footers. */
function searchStories(doc: WordDocument): Story[] {
  return doc.stories(['body', 'header', 'footer']);
}

/** The first element of `local` in the searched stories that `match` accepts. */
function findInStories(
  doc: WordDocument,
  local: string,
  match: (el: XElement) => boolean,
): { el: XElement; part: Part } | undefined {
  for (const story of searchStories(doc)) {
    const el = story.element.descendants(NS.w, local).find(match);
    if (el !== undefined) return { el, part: story.part };
  }
  return undefined;
}

/** Insertion point right before/after an inline `node` of a paragraph (the paragraph is split). */
function splitPointAt(
  doc: WordDocument,
  node: XElement,
  part: Part,
  ids: DocumentIds,
  after: boolean,
): InsertionPoint {
  const p = paragraphOf(node)!;
  const first = after ? nextInRoot(p, node) : node;
  const right = splitBefore(p, first, ids);
  return point(doc, p.parent as XElement, right, part, [p, right]);
}

/** Paragraph and visible-text offsets of the first occurrence of `q` (body, then headers/footers). */
function findPlaceholder(
  doc: WordDocument,
  q: string,
): { p: XElement; part: Part; start: number; end: number } | undefined {
  const m = findText(doc, q, { in: ['body', 'header', 'footer'] })[0];
  if (m === undefined) return undefined;
  return { p: m.paragraph.element, part: m.story.part, start: m.start, end: m.end };
}

/**
 * Resolves an {@link InsertTarget} of `doc` (splitting a paragraph when the target is inline).
 * Placeholders, bookmarks and content controls are looked up in the body first, then in the
 * headers and footers.
 */
export function resolveInsertionPoint(
  doc: WordDocument,
  at: InsertTarget,
  ids: DocumentIds,
): InsertionPoint {
  if ('placeholder' in at) {
    if (at.placeholder === '') throw new UsageError('insertDocument(): empty placeholder');
    const hit = findPlaceholder(doc, at.placeholder);
    if (hit === undefined) {
      throw new UsageError(
        `insertDocument(): placeholder ${JSON.stringify(at.placeholder)} not found`,
      );
    }
    const { p, part, start, end } = hit;
    const right = splitAtOffset(p, end, ids);
    const mid = splitAtOffset(p, start, ids);
    removeKeepingMarkers(mid, p);
    return point(doc, p.parent as XElement, right, part, [p, right]);
  }
  if ('bookmark' in at) {
    const hit = findInStories(doc, 'bookmarkStart', (b) => b.attrNS(NS.w, 'name') === at.bookmark);
    if (hit === undefined) {
      throw new UsageError(`insertDocument(): bookmark ${JSON.stringify(at.bookmark)} not found`);
    }
    const { el: bm, part } = hit;
    if (paragraphOf(bm) === undefined) {
      return point(doc, bm.parent as XElement, nextSiblingNode(bm), part);
    }
    return splitPointAt(doc, bm, part, ids, true);
  }
  if ('contentControl' in at) {
    const { tag, alias } = at.contentControl;
    if (tag === undefined && alias === undefined) {
      throw new UsageError('insertDocument(): contentControl needs a tag or an alias');
    }
    const hit = findInStories(doc, 'sdt', (s) => {
      const pr = s.element(NS.w, 'sdtPr');
      if (tag !== undefined && pr?.element(NS.w, 'tag')?.attrNS(NS.w, 'val') !== tag) return false;
      if (alias !== undefined && pr?.element(NS.w, 'alias')?.attrNS(NS.w, 'val') !== alias) {
        return false;
      }
      return true;
    });
    if (hit === undefined) throw new UsageError('insertDocument(): content control not found');
    const { el: sdt, part } = hit;
    const content = sdt.element(NS.w, 'sdtContent');
    const host = paragraphOf(sdt);
    if (host !== undefined) {
      // Inline control: cut the paragraph around it and replace the control with the blocks.
      const right = splitBefore(host, nextInRoot(host, sdt), ids);
      const mid = splitBefore(host, sdt, ids);
      mid.remove();
      return point(doc, host.parent as XElement, right, part, [host, right]);
    }
    if (content === undefined) {
      throw new UsageError('insertDocument(): content control has no content');
    }
    const kinds = content.elements(NS.w).map((e) => e.local);
    if (kinds.some((k) => k === 'tr' || k === 'tc')) {
      throw new UsageError(
        'insertDocument(): row/cell-level content controls cannot hold documents',
      );
    }
    content.clear();
    sdt.element(NS.w, 'sdtPr')?.element(NS.w, 'showingPlcHdr')?.remove();
    return point(doc, content, null, part);
  }
  if ('paragraph' in at) {
    const el = at.paragraph.element;
    const container = el.parent;
    if (container === null || container.nodeType !== ELEMENT_NODE) {
      throw new UsageError('insertDocument(): the paragraph is not in a document');
    }
    const part = at.paragraph.story.part;
    switch (at.mode) {
      case 'before':
        return point(doc, container, el, part);
      case 'after':
        return point(doc, container, nextSiblingNode(el), part);
      case 'replace': {
        // Keep the paragraph mark when it ends a section; otherwise the paragraph goes.
        const sectPr = el.element(NS.w, 'pPr')?.element(NS.w, 'sectPr');
        if (sectPr === undefined) {
          const next = nextSiblingNode(el);
          el.remove();
          return point(doc, container, next, part);
        }
        for (const k of el._children.slice()) if (!(isEl(k) && k.is(NS.w, 'pPr'))) k.remove();
        return point(doc, container, el, part);
      }
    }
  }
  if ('story' in at) {
    if (!Number.isInteger(at.index) || at.index < 0) {
      throw new UsageError(`insertDocument(): invalid story index ${at.index}`);
    }
    const blocks = at.story.blocks();
    const b = blocks[at.index];
    const before = b === undefined ? null : b instanceof XElement ? b : b.element;
    return point(doc, at.story.element, before, at.story.part);
  }
  throw new UsageError('insertDocument(): unknown insert target');
}

function nextSiblingNode(n: XNode): XNode | null {
  const p = n.parent;
  if (p === null) return null;
  const kids = p._children;
  return kids[kids.indexOf(n) + 1] ?? null;
}

/** Inserts `blocks` at `at` (in a body: before the final sectPr when appending). */
export function insertBlocks(at: InsertionPoint, blocks: readonly XElement[]): void {
  let before = at.before;
  if (before === null) {
    const last = at.container.lastElement();
    if (last?.is(NS.w, 'sectPr') && at.container.is(NS.w, 'body')) before = last;
  }
  for (const b of blocks) at.container.insertBefore(b, before);
}

/** Last block-level element of a cell, looking through sdt/customXml (like validatePackage). */
function lastBlock(container: XElement): XElement | undefined {
  let c: XElement = container;
  for (;;) {
    let found: XElement | undefined;
    for (let i = c._children.length - 1; i >= 0; i--) {
      const k = c._children[i]!;
      if (isEl(k) && k.uri === NS.w && BLOCKS.has(k.local)) {
        found = k;
        break;
      }
    }
    if (found === undefined) return undefined;
    if (found.local === 'sdt') {
      const content = found.element(NS.w, 'sdtContent');
      if (content === undefined) return found;
      c = content;
    } else if (found.local === 'customXml') c = found;
    else return found;
  }
}

/** Appends an empty paragraph to every cell around `container` that does not end with one. */
export function ensureCellsEndWithParagraph(container: XElement): void {
  let tc: XElement | undefined = container.is(NS.w, 'tc')
    ? container
    : container.closest(NS.w, 'tc');
  while (tc !== undefined) {
    if (!lastBlock(tc)?.is(NS.w, 'p')) {
      const p = tc.ownerDocument()?.createElement(NS.w, 'p');
      if (p !== undefined) tc.append(p);
    }
    tc = tc.closest(NS.w, 'tc');
  }
}

/** Removes paragraph-level section properties (`w:p/w:pPr/w:sectPr`) inside `blocks`. */
export function stripSectionProperties(blocks: readonly XElement[]): number {
  let n = 0;
  for (const b of blocks) {
    const all = b.is(NS.w, 'sectPr') ? [b] : b.descendants(NS.w, 'sectPr');
    for (const s of all) {
      if (s.parent !== null && (s.parent as XElement).is(NS.w, 'pPr')) {
        const pPr = s.parent as XElement;
        s.remove();
        if (pPr._children.length === 0 && pPr._attrs.length === 0) pPr.remove();
        n++;
      }
    }
  }
  return n;
}
