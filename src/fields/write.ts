// Writing a field's new result into the DOM, leaving its code alone.
//
// Complex field: the runs between the `separate` and `end` markers are replaced; when the
// field has no result section a `separate` run is added before `end`. The removal may span
// runs, paragraphs and tables: paragraphs fully inside the old result are dropped, the
// paragraph holding `separate` and the one holding `end` are merged, and bookmark/comment range
// markers whose partner lies outside the removed range are kept (re-inserted after
// `separate`), so no range is left dangling. A range that would cut through table rows or
// cells, or remove a section break, is not touched (the update is skipped).
// Simple field: the children of `w:fldSimple` are replaced (`w:fldData` is kept).
//
// Formatting of the new runs: by default the first run of the old result (what Word does
// with `\* MERGEFORMAT`), else the field code's first run; `\* CHARFORMAT` takes the run that
// holds the first letter of the field type. A field nested in another field's instruction
// writes its result as `w:instrText` (as Word does), any other as `w:t`.
// Result text: '\r' starts a new paragraph (the paragraph holding the field is split, the new
// paragraphs copy its properties), '\n' is a line break, '\t' a tab.

import { insertOrdered, PPR_ORDER, RPR_ORDER, wElement } from '../wml/order.js';
import { type XElement, type XNode, XText } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { Scan } from './scan.js';

/** Extra formatting for a result (SYMBOL's font and size, NOTEREF's style). */
export interface RunFormat {
  font?: string;
  /** Half-points. */
  size?: number;
  /** Character style id. */
  style?: string;
}

/** Splits `child`'s run so the run holds only its properties and `child`; returns the run. */
export function isolate(child: XElement): XElement | undefined {
  const run = child.parent;
  if (run === null || run.nodeType !== 1 || !run.is(NS.w, 'r')) return undefined;
  const rPr = run.element(NS.w, 'rPr');
  const kids = run.children.slice();
  const at = kids.indexOf(child);
  const before = kids.slice(0, at).filter((k) => k !== rPr);
  const after = kids.slice(at + 1);
  const hasContent = (nodes: XNode[]): boolean => nodes.some((k) => k.nodeType === 1);
  if (hasContent(before)) {
    const r = run.clone(false);
    if (rPr) r.append(rPr.clone());
    for (const k of before) r.append(k);
    run.before(r);
  }
  if (hasContent(after)) {
    const r = run.clone(false);
    if (rPr) r.append(rPr.clone());
    for (const k of after) r.append(k);
    run.after(r);
  }
  return run;
}

/** Elements whose removal would break table structure. */
const STRUCTURAL: ReadonlySet<string> = new Set(['tr', 'tbl']);

/** Range markers kept when their partner survives. */
const MARKERS: Readonly<Record<string, string>> = {
  bookmarkStart: 'bookmarkEnd',
  bookmarkEnd: 'bookmarkStart',
  commentRangeStart: 'commentRangeEnd',
  commentRangeEnd: 'commentRangeStart',
  permStart: 'permEnd',
  permEnd: 'permStart',
};

const isProps = (n: XNode): boolean =>
  n.nodeType === 1 && (n.local.endsWith('Pr') || n.local === 'sdtEndPr');

function ancestors(el: XElement): XElement[] {
  const out: XElement[] = [];
  let n = el.parent;
  while (n !== null && n.nodeType === 1) {
    out.push(n);
    n = n.parent;
  }
  return out;
}

/**
 * Removes every node strictly between `a` and `b` (document order, `a` first, neither
 * containing the other) and merges their paragraphs. Returns false (and changes nothing) when
 * the range crosses table rows/cells or holds a section break.
 */
export function removeBetween(a: XElement, b: XElement): boolean {
  const ancA = new Set(ancestors(a));
  let common: XElement | undefined;
  for (const n of ancestors(b)) {
    if (ancA.has(n)) {
      common = n;
      break;
    }
  }
  if (common === undefined || STRUCTURAL.has(common.local)) return false;
  const doomed: XNode[] = [];
  let x = a;
  while (x.parent !== common) {
    const p = x.parent as XElement;
    if (STRUCTURAL.has(p.local)) return false;
    const kids = p.children;
    for (let i = kids.indexOf(x) + 1; i < kids.length; i++) {
      if (!isProps(kids[i]!)) doomed.push(kids[i]!);
    }
    x = p;
  }
  let y = b;
  const tail: XNode[] = [];
  while (y.parent !== common) {
    const p = y.parent as XElement;
    if (STRUCTURAL.has(p.local)) return false;
    const kids = p.children;
    for (let i = 0; i < kids.indexOf(y); i++) if (!isProps(kids[i]!)) tail.push(kids[i]!);
    y = p;
  }
  const kids = common.children;
  for (let i = kids.indexOf(x) + 1; i < kids.indexOf(y); i++) doomed.push(kids[i]!);
  doomed.push(...tail);

  // Section breaks and range markers inside the doomed nodes.
  const markers: XElement[] = [];
  const seen = new Map<string, number>();
  for (const n of doomed) {
    if (n.nodeType !== 1) continue;
    for (const e of [n, ...n.descendants(NS.w)]) {
      if (e.uri !== NS.w) continue;
      if (e.local === 'sectPr') return false;
      if (MARKERS[e.local] !== undefined) {
        markers.push(e);
        const key = `${e.local}|${e.attrNS(NS.w, 'id') ?? ''}`;
        seen.set(key, (seen.get(key) ?? 0) + 1);
      }
    }
  }
  for (const n of doomed) n.remove();
  let anchor: XElement = a;
  for (const m of markers) {
    const partner = `${MARKERS[m.local]}|${m.attrNS(NS.w, 'id') ?? ''}`;
    if (seen.has(partner)) continue;
    anchor.after(m);
    anchor = m;
  }
  if (x !== a && y !== b && x.is(NS.w, 'p') && y.is(NS.w, 'p') && x !== y) {
    for (const k of y.children.slice()) {
      if (k.nodeType === 1 && k.is(NS.w, 'pPr')) continue;
      x.append(k);
    }
    y.remove();
  }
  return true;
}

/** Run nodes (`w:t`/`w:instrText`, `w:tab`, `w:br`) for one line of text. */
function runContent(scope: XElement, text: string, instr: boolean): XNode[] {
  const out: XNode[] = [];
  const tag = instr ? 'instrText' : 't';
  let start = 0;
  const flush = (end: number): void => {
    if (end <= start) return;
    const s = text.slice(start, end);
    const t = wElement(scope, tag);
    if (/^\s|\s$|\s\s/.test(s)) t.setAttr('xml:space', 'preserve');
    t.append(new XText(s));
    out.push(t);
  };
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c !== '\t' && c !== '\n') continue;
    flush(i);
    out.push(wElement(scope, c === '\t' ? 'tab' : 'br'));
    start = i + 1;
  }
  flush(text.length);
  return out;
}

/** A new run with a copy of `rPr`, `format` applied, holding `text`. */
export function makeRun(
  scope: XElement,
  rPr: XElement | undefined,
  text: string,
  instr: boolean,
  format?: RunFormat,
): XElement {
  const r = wElement(scope, 'r');
  let props = rPr?.clone();
  if (format && (format.font !== undefined || format.size !== undefined || format.style)) {
    props ??= wElement(scope, 'rPr');
    if (format.style !== undefined) {
      props.element(NS.w, 'rStyle')?.remove();
      insertOrdered(props, wElement(scope, 'rStyle', { val: format.style }), RPR_ORDER);
    }
    if (format.font !== undefined) {
      props.element(NS.w, 'rFonts')?.remove();
      const f = format.font;
      insertOrdered(
        props,
        wElement(scope, 'rFonts', { ascii: f, hAnsi: f, eastAsia: f, cs: f }),
        RPR_ORDER,
      );
    }
    if (format.size !== undefined) {
      for (const n of ['sz', 'szCs']) {
        props.element(NS.w, n)?.remove();
        insertOrdered(props, wElement(scope, n, { val: String(format.size) }), RPR_ORDER);
      }
    }
  }
  if (props) r.append(props);
  for (const n of runContent(scope, text, instr)) r.append(n);
  return r;
}

/** Runs of the old result region of a complex field, in document order. */
function resultRuns(sepRun: XElement, endRun: XElement): XElement[] {
  const out: XElement[] = [];
  // Walk forward in document order from sepRun to endRun.
  let n: XElement | undefined = nextInOrder(sepRun, true);
  let guard = 0;
  while (n !== undefined && n !== endRun && guard++ < 100000) {
    if (n.is(NS.w, 'r')) out.push(n);
    n = nextInOrder(n, !n.is(NS.w, 'r'));
  }
  return out;
}

/** Next element in document order (descending into `el` when `descend`). */
function nextInOrder(el: XElement, descend: boolean): XElement | undefined {
  if (descend) {
    const first = el.firstElement();
    if (first) return first;
  }
  let n: XElement = el;
  for (;;) {
    const sib = n.nextElement();
    if (sib) return sib;
    const p = n.parent;
    if (p === null || p.nodeType !== 1) return undefined;
    n = p;
  }
}

/** The run properties to copy for a field's new result (see the file header). */
export function templateProps(s: Scan, charFormat: boolean): XElement | undefined {
  const codeRun = (): XElement | undefined => {
    if (s.kind === 'simple') return undefined;
    // The run holding the first non-space instruction character.
    let n: XElement | undefined = nextInOrder(s.begin.parent as XElement, false);
    let guard = 0;
    while (n !== undefined && guard++ < 10000) {
      if (n.is(NS.w, 'fldChar')) break;
      if (n.is(NS.w, 'instrText') && n.text.trim() !== '') return n.parent as XElement;
      n = nextInOrder(n, true);
    }
    return s.begin.parent as XElement;
  };
  if (charFormat) return codeRun()?.element(NS.w, 'rPr');
  const textTag = s.inInstr ? 'instrText' : 't';
  let runs: XElement[];
  if (s.kind === 'simple') runs = s.begin.descendants(NS.w, 'r');
  else if (s.separate) runs = resultRuns(s.separate.parent as XElement, s.end.parent as XElement);
  else runs = [];
  const withText = runs.find((r) => r.element(NS.w, textTag) !== undefined) ?? runs[0];
  if (withText) return withText.element(NS.w, 'rPr');
  return codeRun()?.element(NS.w, 'rPr');
}

/**
 * The isolated run of a complex field's `separate` marker; when the field has none, a
 * `separate` run is created right before `endRun` (and recorded in `s`).
 */
export function ensureSeparate(s: Scan, endRun: XElement): XElement | undefined {
  if (s.separate) return isolate(s.separate);
  const sepRun = endRun.clone(false);
  const endProps = endRun.element(NS.w, 'rPr');
  if (endProps) sepRun.append(endProps.clone());
  const sep = wElement(endRun, 'fldChar', { fldCharType: 'separate' });
  sepRun.append(sep);
  endRun.before(sepRun);
  s.separate = sep;
  return sepRun;
}

/** Removes `w:dirty` from a field's begin marker. */
export function clearDirty(s: Scan): void {
  s.begin.removeAttrNS(NS.w, 'dirty');
}

/**
 * Replaces a field's result with `text`. Returns false when the field's structure does not
 * allow it (markers outside runs, a result crossing table cells): nothing is changed then.
 */
export function writeResult(
  s: Scan,
  text: string,
  rPr: XElement | undefined,
  format?: RunFormat,
): boolean {
  if (s.kind === 'simple') {
    const fld = s.begin;
    for (const k of fld.children.slice()) {
      if (k.nodeType === 1 && k.is(NS.w, 'fldData')) continue;
      k.remove();
    }
    const flat = text.replace(/\r\n?/g, '\n');
    if (flat !== '') fld.append(makeRun(fld, rPr, flat, s.inInstr, format));
    return true;
  }
  const endRun = isolate(s.end);
  if (endRun === undefined) return false;
  const hadResult = s.separate !== undefined;
  const sepRun = ensureSeparate(s, endRun);
  if (sepRun === undefined) return false;
  if (hadResult && !removeBetween(sepRun, endRun)) return false;
  const paras = text.split(/\r\n?/);
  const P = sepRun.parent as XElement;
  if (paras.length > 1 && !P.is(NS.w, 'p')) {
    paras.splice(0, paras.length, paras.join('\n'));
  }
  const first = paras[0]!;
  let last: XElement = sepRun;
  if (first !== '') {
    last = makeRun(P, rPr, first, s.inInstr, format);
    sepRun.after(last);
  }
  if (paras.length > 1) splitInto(P, last, paras.slice(1), rPr, s.inInstr, format);
  return true;
}

/**
 * Splits paragraph `P` after its child `after`: one new paragraph per entry of `lines`
 * (copying P's properties; a section break moves to the last one), and whatever followed
 * `after` in P moves to the end of the last new paragraph.
 */
function splitInto(
  P: XElement,
  after: XElement,
  lines: string[],
  rPr: XElement | undefined,
  instr: boolean,
  format?: RunFormat,
): void {
  const kids = P.children;
  const tail = kids.slice(kids.indexOf(after) + 1);
  const pPr = P.element(NS.w, 'pPr');
  const sectPr = pPr?.element(NS.w, 'sectPr');
  let prev = P;
  for (let i = 0; i < lines.length; i++) {
    const np = wElement(P, 'p');
    if (pPr) {
      const copy = pPr.clone();
      copy.element(NS.w, 'sectPr')?.remove();
      np.append(copy);
    }
    const line = lines[i]!;
    if (line !== '') np.append(makeRun(P, rPr, line, instr, format));
    if (i === lines.length - 1) {
      for (const k of tail) np.append(k);
      const npPr = np.element(NS.w, 'pPr');
      if (sectPr && npPr) insertOrdered(npPr, sectPr, PPR_ORDER);
    }
    prev.after(np);
    prev = np;
  }
}
