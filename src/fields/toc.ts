// Table of contents (TOC field, ECMA-376 Part 1 §17.16.5.68): rebuilds the field's result from
// the document's headings, keeping the field code.
//
// Switches: `\o "1-3"` heading levels (outline levels of the paragraph style chain; built-in
// "heading N" styles count even without `w:outlineLvl`), `\u` also honours outline levels set
// on the paragraph itself, `\t "Style,1,Other,2"` extra styles and their levels, `\b name`
// only headings inside a bookmark, `\f [id]` + `\l "1-3"` TC fields, `\c "Figure"` captions
// (paragraphs holding a SEQ of that identifier), `\h` hyperlinks, `\n ["1-1"]` no page
// numbers (for those levels), `\p "sep"` separator instead of a tab, `\z` (`w:webHidden` on the
// page-number runs). Without `\o`, `\t`, `\u`, `\f` or `\c`, levels 1-9 are used.
//
// Each entry is a paragraph in the TOC1..TOC9 style (found by name "toc N", created when
// missing) with a right tab stop at the text width and a dot leader, the heading text (list
// label first), and a PAGEREF field to a `_Toc` bookmark on the heading (an existing one is
// reused; new ones get deterministic names and ids from DocumentIds, with a local fallback).
// Page numbers come from the PageInfoProvider; without one, the previous results are kept for
// headings that already had an entry, and the others are left blank (Word fills them on F9).

import type { WordDocument } from '../wml/document.js';
import { DocumentIds } from '../wml/ids.js';
import { insertOrdered, PPR_ORDER, wElement } from '../wml/order.js';
import { paragraphText } from '../wml/text.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { UpdateState } from './evaluate.js';
import type { Scan } from './scan.js';
import { ensurePart } from './settings.js';
import { findSwitch, hasSwitch, type ParsedInstruction } from './tokenize.js';
import type { PageInfoProvider } from './types.js';
import { clearDirty, ensureSeparate, isolate, makeRun, removeBetween } from './write.js';

/** Text Word writes when a TOC finds nothing. */
const NO_ENTRIES = 'No table of contents entries found.';

interface Entry {
  level: number;
  text: string;
  para: XElement;
  /** TC field entries: the field (bookmarked instead of the paragraph). */
  tc?: Scan;
  noPage: boolean;
  bookmark?: string;
}

function range(arg: string | undefined, fallback: [number, number]): [number, number] {
  if (arg === undefined) return fallback;
  const m = /^\s*(\d+)\s*(?:-\s*(\d+))?\s*$/.exec(arg);
  if (!m) return fallback;
  const a = Number(m[1]);
  const b = m[2] === undefined ? a : Number(m[2]);
  return [Math.max(1, Math.min(a, b)), Math.min(9, Math.max(a, b))];
}

const within = (n: number, [a, b]: [number, number]): boolean => n >= a && n <= b;

/** FNV-1a, for deterministic bookmark names. */
function hash(s: string): number {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h;
}

/** Hands out fresh bookmark ids: DocumentIds when available, else max + 1 over every story. */
export function bookmarkIds(doc: WordDocument): () => string {
  let fallback: number | undefined;
  let useRegistry = true;
  return () => {
    if (useRegistry) {
      try {
        return String(DocumentIds.of(doc).next('bookmark'));
      } catch {
        useRegistry = false;
      }
    }
    if (fallback === undefined) {
      let max = -1;
      for (const story of doc.stories()) {
        for (const el of story.element.descendants(NS.w)) {
          if (el.local !== 'bookmarkStart' && el.local !== 'bookmarkEnd') continue;
          const n = Number(el.attrNS(NS.w, 'id'));
          if (Number.isInteger(n) && n > max) max = n;
        }
      }
      fallback = max + 1;
    }
    return String(fallback++);
  };
}

/** Style id of "toc N", creating the style when missing. */
function tocStyle(st: UpdateState, level: number): string {
  const info = st.info;
  const found = info.styleId(`toc ${level}`, 'paragraph');
  if (found !== undefined) return found;
  const part = info.stylesPart() ?? ensurePart(info.doc, 'styles', 'styles');
  const root = part.xml().root;
  let id = `TOC${level}`;
  for (let n = 1; info.styles().has(id); n++) id = `TOC${level}_${n}`;
  const base = info.defaultParagraphStyle();
  const style = wElement(root, 'style', { type: 'paragraph', styleId: id });
  style.append(wElement(root, 'name', { val: `toc ${level}` }));
  if (base !== undefined) {
    style.append(wElement(root, 'basedOn', { val: base }));
    style.append(wElement(root, 'next', { val: base }));
  }
  style.append(wElement(root, 'autoRedefine'));
  style.append(wElement(root, 'uiPriority', { val: '39' }));
  style.append(wElement(root, 'unhideWhenUsed'));
  const pPr = wElement(root, 'pPr');
  pPr.append(wElement(root, 'spacing', { after: '100' }));
  if (level > 1) pPr.append(wElement(root, 'ind', { left: String(220 * (level - 1)) }));
  style.append(pPr);
  root.append(style);
  info.invalidateStyles();
  return id;
}

/** Heading text as a TOC shows it: visible text, line breaks as spaces, trimmed. */
function headingText(p: XElement): string {
  return paragraphText(p)
    .replace(/[\n\r]+/g, ' ')
    .trim();
}

/** Paragraphs inside the result of any TOC field of the body. */
function tocResultParagraphs(st: UpdateState, scans: readonly Scan[]): Set<XElement> {
  const out = new Set<XElement>();
  const order = st.info.paragraphOrder(st.info.doc.body);
  const paras = [...order.keys()];
  for (const s of scans) {
    if (s.parsed.type !== 'TOC' || s.separate === undefined) continue;
    const a = order.get(s.separate.closest(NS.w, 'p') as XElement);
    const b = order.get(s.end.closest(NS.w, 'p') as XElement);
    if (a === undefined || b === undefined) continue;
    for (let i = a; i <= b; i++) out.add(paras[i]!);
  }
  return out;
}

/**
 * The paragraph holding `el`, or for a block-level marker the nearest paragraph after
 * (`dir` 1) or before (`dir` -1) it.
 */
function nearParagraph(el: XElement, dir: 1 | -1): XElement | undefined {
  const own = el.closest(NS.w, 'p');
  if (own) return own;
  for (let n = dir > 0 ? el.nextElement() : el.previousElement(); n; ) {
    if (n.is(NS.w, 'p')) return n;
    const inner = n.descendants(NS.w, 'p');
    const pick = dir > 0 ? inner[0] : inner[inner.length - 1];
    if (pick) return pick;
    n = dir > 0 ? n.nextElement() : n.previousElement();
  }
  return undefined;
}

function collect(
  st: UpdateState,
  p: ParsedInstruction,
  scans: readonly Scan[],
  levels?: [number, number],
): Entry[] {
  const info = st.info;
  const sw = p.switches;
  const has = (n: string): boolean => hasSwitch(sw, n);
  const argOf = (n: string): string | undefined => findSwitch(sw, n)?.arg;
  const applied = has('\\u');
  const custom = new Map<string, number>();
  const t = argOf('\\t');
  if (t !== undefined) {
    const parts = t.split(/[,;]/).map((x) => x.trim());
    for (let i = 0; i < parts.length; i += 2) {
      const id = parts[i] ? info.styleId(parts[i]!) : undefined;
      const lvl = Number(parts[i + 1] ?? 1);
      if (id !== undefined) custom.set(id, Number.isInteger(lvl) && lvl >= 1 && lvl <= 9 ? lvl : 1);
    }
  }
  const tcId = has('\\f') ? (argOf('\\f') ?? 'C').toUpperCase() : undefined;
  const caption = argOf('\\c') ?? argOf('\\a');
  let outline = levels;
  if (outline === undefined && has('\\o')) outline = range(argOf('\\o'), [1, 9]);
  if (
    outline === undefined &&
    !applied &&
    t === undefined &&
    tcId === undefined &&
    caption === undefined
  ) {
    outline = [1, 9];
  }
  const noPage = has('\\n') ? range(argOf('\\n'), [1, 9]) : undefined;
  const excluded = tocResultParagraphs(st, scans);
  const order = info.paragraphOrder(info.doc.body);
  let span: [number, number] | undefined;
  const bm = argOf('\\b');
  if (bm !== undefined) {
    const b = info.bookmark(bm);
    const endEl = b
      ? info.doc.body.element
          .descendants(NS.w, 'bookmarkEnd')
          .find((e) => e.attrNS(NS.w, 'id') === b.id)
      : undefined;
    const start = b ? nearParagraph(b.start, 1) : undefined;
    const end = endEl ? nearParagraph(endEl, -1) : undefined;
    span = [
      start === undefined ? 0 : (order.get(start) ?? 0),
      end === undefined ? Number.MAX_SAFE_INTEGER : (order.get(end) ?? Number.MAX_SAFE_INTEGER),
    ];
  }
  const inSpan = (para: XElement): boolean => {
    if (span === undefined) return true;
    const i = order.get(para);
    return i !== undefined && i >= span[0] && i <= span[1];
  };

  const entries: Entry[] = [];
  const captions = new Set<XElement>();
  if (caption !== undefined) {
    for (const c of scans) {
      if (
        c.parsed.type === 'SEQ' &&
        !c.secondary &&
        c.parsed.args[0]?.text.toLowerCase() === caption.toLowerCase()
      ) {
        const para = c.begin.closest(NS.w, 'p');
        if (para) captions.add(para);
      }
    }
  }
  const tcs = new Map<XElement, Scan[]>();
  if (tcId !== undefined) {
    for (const c of scans) {
      if (c.parsed.type !== 'TC' || c.secondary) continue;
      const id = (findSwitch(c.parsed.switches, '\\f')?.arg ?? 'C').toUpperCase();
      if (id !== tcId) continue;
      const para = c.begin.closest(NS.w, 'p');
      if (para) tcs.set(para, [...(tcs.get(para) ?? []), c]);
    }
  }
  const tcLevels = range(argOf('\\l'), [1, 9]);
  const pageless = (lvl: number): boolean => noPage !== undefined && within(lvl, noPage);
  for (const para of order.keys()) {
    if (excluded.has(para) || !inSpan(para)) continue;
    let level: number | undefined;
    const styleId = info.paragraphStyle(para);
    if (styleId !== undefined && custom.has(styleId)) level = custom.get(styleId);
    if (level === undefined && (outline !== undefined || applied)) {
      const lvl = info.headingLevel(para, applied);
      if (lvl !== undefined && within(lvl, outline ?? [1, 9])) level = lvl;
    }
    if (level === undefined && captions.has(para)) level = 1;
    if (level !== undefined) {
      const text = headingText(para);
      if (text !== '') {
        const label = info.listLabel(para, info.doc.body);
        entries.push({
          level,
          text: label ? `${label}\t${text}` : text,
          para,
          noPage: pageless(level),
        });
      }
    }
    for (const c of tcs.get(para) ?? []) {
      const lvl = Number(findSwitch(c.parsed.switches, '\\l')?.arg ?? 1);
      const level2 = Number.isInteger(lvl) && lvl >= 1 && lvl <= 9 ? lvl : 1;
      if (!within(level2, tcLevels)) continue;
      const text = (c.parsed.args[0]?.text ?? '').trim();
      if (text === '') continue;
      entries.push({
        level: level2,
        text,
        para,
        tc: c,
        noPage: pageless(level2) || hasSwitch(c.parsed.switches, '\\n'),
      });
    }
  }
  return entries;
}

/** The `_Toc` bookmark on an entry's target, created when missing. */
function ensureBookmark(
  entry: Entry,
  taken: Set<string>,
  nextId: () => string,
  index: number,
): string {
  const p = entry.para;
  if (entry.tc === undefined) {
    for (const b of p.elements(NS.w, 'bookmarkStart')) {
      const name = b.attrNS(NS.w, 'name');
      if (name !== undefined && /^_Toc/i.test(name)) return name;
    }
  }
  let n = hash(`${index}|${entry.text}`) % 1e9;
  let name = `_Toc${String(n).padStart(9, '0')}`;
  while (taken.has(name.toLowerCase())) {
    n = (n + 1) % 1e9;
    name = `_Toc${String(n).padStart(9, '0')}`;
  }
  taken.add(name.toLowerCase());
  const id = nextId();
  const start = wElement(p, 'bookmarkStart', { id, name });
  const end = wElement(p, 'bookmarkEnd', { id });
  if (entry.tc) {
    const beginRun = entry.tc.begin.parent as XElement;
    const endRun = entry.tc.end.parent as XElement;
    beginRun.before(start);
    endRun.after(end);
  } else {
    const pPr = p.element(NS.w, 'pPr');
    if (pPr) pPr.after(start);
    else p.prepend(start);
    p.append(end);
  }
  return name;
}

function fldRun(scope: XElement, rPr: XElement, kind: string): XElement {
  const r = wElement(scope, 'r');
  r.append(rPr.clone());
  r.append(wElement(scope, 'fldChar', { fldCharType: kind }));
  return r;
}

/** Page-number text of an entry. */
function entryPage(
  entry: Entry,
  pages: PageInfoProvider | undefined,
  old: Map<XElement, string>,
): string {
  if (pages) {
    const node = entry.tc ? entry.tc.begin : entry.para;
    const text = pages.pageNumberText?.(node);
    if (text !== undefined) return text;
    const n = pages.pageOf(node);
    if (n !== undefined) return String(n);
  }
  return entry.tc ? '' : (old.get(entry.para) ?? '');
}

/**
 * Rebuilds the result of TOC field `s`. Returns false when its structure does not allow it
 * (markers outside runs, a result crossing table cells, not in the body).
 */
export function rebuildToc(
  st: UpdateState,
  s: Scan,
  p: ParsedInstruction,
  scans: readonly Scan[],
  o: { pages?: PageInfoProvider | undefined; levels?: [number, number] | undefined },
): boolean {
  if (s.kind !== 'complex' || s.story.kind !== 'body') return false;
  const info = st.info;
  const entries = collect(st, p, scans, o.levels);

  // Old page numbers, by heading paragraph (through the old PAGEREF bookmarks).
  const old = new Map<XElement, string>();
  for (const c of s.children) {
    if (c.inInstr || c.parsed.type !== 'PAGEREF') continue;
    const name = c.parsed.args[0]?.text;
    const para = name === undefined ? undefined : info.bookmark(name)?.start.closest(NS.w, 'p');
    if (para) old.set(para, c.resultText.trim());
  }

  // Structure checks and the DOM surgery.
  const endRun = isolate(s.end);
  if (endRun === undefined) return false;
  const sepRun = ensureSeparate(s, endRun);
  if (sepRun === undefined) return false;
  const P = sepRun.parent as XElement;
  if (!P.is(NS.w, 'p')) return false;
  const endPara = endRun.closest(NS.w, 'p');
  const contentRun = (r: XElement): boolean => r.elements().some((k) => !k.is(NS.w, 'rPr'));
  let dedicated = false;
  if (endPara !== undefined && endPara !== P) {
    const runs = endPara.descendants(NS.w, 'r');
    dedicated = runs.slice(0, runs.indexOf(endRun)).every((r) => !contentRun(r));
  }
  const endPPr = (endPara ?? P).element(NS.w, 'pPr')?.clone();
  const lead = P.children
    .slice(0, P.children.indexOf(sepRun))
    .some((k) => k.nodeType === 1 && (k.is(NS.w, 't') || k.descendants(NS.w, 't').length > 0));

  // Page numbers and bookmarks first, while the document is as it was laid out (both are
  // outside the TOC result, so independent of the surgery).
  const pageTexts = entries.map((e) => (e.noPage ? '' : entryPage(e, o.pages, old)));
  const taken = new Set(info.bookmarks().keys());
  const nextId = bookmarkIds(info.doc);
  entries.forEach((e, i) => {
    e.bookmark = ensureBookmark(e, taken, nextId, i);
  });
  info.invalidateBookmarks();

  if (!removeBetween(sepRun, endRun)) return false;
  const tail = P.children.slice(P.children.indexOf(sepRun) + 1);
  for (const k of tail) k.remove();
  const tailHasContent = tail.some(
    (k) =>
      k !== endRun &&
      k.nodeType === 1 &&
      (k.is(NS.w, 'r') ? contentRun(k) : !/^bookmark(Start|End)$/.test(k.local)),
  );

  const width = String(info.textWidth(P));
  const hyper = hasSwitch(p.switches, '\\h');
  const webHidden = hasSwitch(p.switches, '\\z');
  const sep = findSwitch(p.switches, '\\p')?.arg;
  const paras = entries.map((e, i) => {
    const np = wElement(P, 'p');
    const pPr = wElement(P, 'pPr');
    pPr.append(wElement(P, 'pStyle', { val: tocStyle(st, e.level) }));
    const tabs = wElement(P, 'tabs');
    tabs.append(wElement(P, 'tab', { val: 'right', leader: 'dot', pos: width }));
    pPr.append(tabs);
    const markRPr = wElement(P, 'rPr');
    markRPr.append(wElement(P, 'noProof'));
    pPr.append(markRPr);
    np.append(pPr);
    const box = hyper ? wElement(P, 'hyperlink', { anchor: e.bookmark!, history: '1' }) : np;
    if (box !== np) np.append(box);
    const textRPr = wElement(P, 'rPr');
    textRPr.append(wElement(P, 'noProof'));
    box.append(makeRun(P, textRPr, e.text, false));
    if (!e.noPage) {
      const pageRPr = textRPr.clone();
      if (webHidden) pageRPr.append(wElement(P, 'webHidden'));
      box.append(makeRun(P, pageRPr, sep ?? '\t', false));
      box.append(fldRun(P, pageRPr, 'begin'));
      box.append(makeRun(P, pageRPr, ` PAGEREF ${e.bookmark} \\h `, true));
      box.append(fldRun(P, pageRPr, 'separate'));
      const page = pageTexts[i]!;
      if (page !== '') box.append(makeRun(P, pageRPr, page, false));
      box.append(fldRun(P, pageRPr, 'end'));
    }
    return np;
  });

  const ownPPr = P.element(NS.w, 'pPr');
  const sectPr = ownPPr?.element(NS.w, 'sectPr');
  sectPr?.remove();
  let last = P;
  if (!lead && paras.length > 0) {
    const first = paras.shift()!;
    const firstPPr = first.element(NS.w, 'pPr')!;
    if (ownPPr) ownPPr.replaceWith(firstPPr);
    else P.prepend(firstPPr);
    for (const k of first.children.slice()) P.append(k);
  } else if (entries.length === 0) {
    P.append(makeRun(P, undefined, NO_ENTRIES, false));
  }
  for (const np of paras) {
    last.after(np);
    last = np;
  }
  let final = last;
  if (dedicated || tailHasContent) {
    const tp = wElement(P, 'p');
    if (endPPr) tp.append(endPPr);
    for (const k of tail) tp.append(k);
    last.after(tp);
    final = tp;
  } else {
    for (const k of tail) last.append(k);
  }
  if (sectPr) {
    let pPr = final.element(NS.w, 'pPr');
    if (!pPr) {
      pPr = wElement(P, 'pPr');
      final.prepend(pPr);
    }
    pPr.element(NS.w, 'sectPr')?.remove();
    insertOrdered(pPr, sectPr, PPR_ORDER);
  }
  clearDirty(s);
  return true;
}
