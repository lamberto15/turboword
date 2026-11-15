// Run surgery for replace: splitting text leaves and runs at match boundaries, isolating the
// matched leaves into runs of their own, and merging the pieces back afterwards.
//
// Every split keeps the *left* part in the original node and moves the right part into a new
// node inserted after it. Replace processes a paragraph's matches from last to first, so the
// segments of the earlier matches (still unprocessed) keep pointing at valid nodes and offsets.
//
// New runs copy the original run's attributes and a clone of its `w:rPr`, so formatting is
// unchanged. All pieces of one original run are remembered; when the edit is done, pieces that
// ended up adjacent with identical properties are merged again (adjacent `w:t` texts too), so
// replacing a word inside a single run leaves a single run, as if it had been edited in place.

import { domEqual, XElement, XText } from '../xml/dom.js';
import { NS } from '../xml/ns.js';

/** Text that needs `xml:space="preserve"` to keep its whitespace. */
export function needsPreserve(s: string): boolean {
  return /^\s|\s$|\s\s/.test(s);
}

/** Adds `xml:space="preserve"` to a `w:t`/`w:delText` whose text needs it. */
export function fixSpace(t: XElement): void {
  if (needsPreserve(t.text) && t.attrNS(NS.xml, 'space') !== 'preserve') {
    t.setAttr('xml:space', 'preserve');
  }
}

/**
 * Splits text leaf `node` at `offset` (0 < offset < length): `node` keeps the left part and the
 * right part becomes a new sibling (same name and attributes), which is returned.
 */
export function splitText(node: XElement, offset: number): XElement {
  const s = node.text;
  const right = new XElement(node.name, node.uri, node._attrs.slice());
  right.append(new XText(s.slice(offset)));
  node.text = s.slice(0, offset);
  node.after(right);
  fixSpace(node);
  fixSpace(right);
  return right;
}

/** Remembers which original run every piece came from. */
export class Pieces {
  readonly #origin = new Map<XElement, XElement>();

  originOf(run: XElement): XElement {
    return this.#origin.get(run) ?? run;
  }

  /** A new empty run like `run` (attributes and a clone of `w:rPr`), recorded as its piece. */
  piece(run: XElement): XElement {
    const nr = new XElement(run.name, run.uri, run._attrs.slice());
    const rPr = run.element(NS.w, 'rPr');
    if (rPr) nr.append(rPr.clone());
    const origin = this.originOf(run);
    this.#origin.set(run, origin);
    this.#origin.set(nr, origin);
    return nr;
  }

  /** Every recorded run (original runs and pieces). */
  runs(): XElement[] {
    return [...this.#origin.keys()];
  }
}

function isContent(k: { nodeType: number }, rPr: XElement | undefined): boolean {
  return k.nodeType === 1 && k !== rPr;
}

/**
 * Makes `child` the first content of a run: moves it and its following siblings into a new
 * piece after `run`. Returns the run that now starts with `child` (`run` itself when nothing but
 * `w:rPr` precedes it).
 */
export function splitRunBefore(run: XElement, child: XElement, pieces: Pieces): XElement {
  const rPr = run.element(NS.w, 'rPr');
  const kids = run.children;
  const at = kids.indexOf(child);
  let before = false;
  for (let i = 0; i < at; i++) {
    if (isContent(kids[i]!, rPr)) {
      before = true;
      break;
    }
  }
  if (!before) return run;
  const nr = pieces.piece(run);
  for (const k of kids.slice(at)) nr.append(k);
  run.after(nr);
  return nr;
}

/** Moves the siblings after `child` into a new piece after `run` (when there are any). */
export function splitRunAfter(run: XElement, child: XElement, pieces: Pieces): void {
  const kids = run.children;
  const tail = kids.slice(kids.indexOf(child) + 1);
  if (!tail.some((k) => k.nodeType === 1)) return;
  const nr = pieces.piece(run);
  for (const k of tail) nr.append(k);
  run.after(nr);
}

/**
 * Splits `run` so that each maximal sequence of covered (in `covered`) or uncovered content
 * children forms a run of its own.
 */
export function isolate(run: XElement, covered: ReadonlySet<XElement>, pieces: Pieces): void {
  const rPr = run.element(NS.w, 'rPr');
  const content = run.children.filter((k): k is XElement => isContent(k, rPr));
  const heads: XElement[] = [];
  for (let i = 1; i < content.length; i++) {
    if (covered.has(content[i]!) !== covered.has(content[i - 1]!)) heads.push(content[i]!);
  }
  for (let i = heads.length - 1; i >= 0; i--) splitRunBefore(run, heads[i]!, pieces);
}

function sameRun(a: XElement, b: XElement): boolean {
  if (a.name !== b.name || a._attrs.length !== b._attrs.length) return false;
  for (let i = 0; i < a._attrs.length; i++) if (a._attrs[i] !== b._attrs[i]) return false;
  const ra = a.element(NS.w, 'rPr');
  const rb = b.element(NS.w, 'rPr');
  if (ra === undefined || rb === undefined) return ra === rb;
  return domEqual(ra, rb);
}

/** Appends `from`'s content to `into`, joining a `w:t` pair at the seam. */
function absorb(into: XElement, from: XElement): void {
  const fromPr = from.element(NS.w, 'rPr');
  const moving = from.children.filter((k) => k !== fromPr);
  const last = into.lastElement();
  const first = moving.find((k): k is XElement => k.nodeType === 1);
  if (
    last !== undefined &&
    first !== undefined &&
    last.is(NS.w, 't') &&
    first.is(NS.w, 't') &&
    into.children[into.children.length - 1] === last &&
    moving[0] === first
  ) {
    last.text = last.text + first.text;
    fixSpace(last);
    moving.shift();
  }
  for (const k of moving) into.append(k);
  from.remove();
}

/** Merges pieces of the same original run that are adjacent and have identical properties. */
export function mergePieces(pieces: Pieces): void {
  for (const run of pieces.runs()) {
    if (run.parent === null) continue;
    const origin = pieces.originOf(run);
    for (;;) {
      const next = run.nextElement();
      if (next === undefined || pieces.originOf(next) !== origin || next === run) break;
      if (!sameRun(run, next)) break;
      absorb(run, next);
    }
  }
}
