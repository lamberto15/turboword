// The visible-text flattener: the single definition of "what text does this content show".
//
// `text()`, search/replace (v0.3), list labels and the layout input all derive from
// `walkVisible`, so they can never disagree about what a paragraph says. The walker reports
// every contributing element together with the characters it stands for; mapping a match back to
// DOM positions is then a matter of recording those (node, text) pairs.
//
// Rules (ECMA-376 Part 1 §17.3.3 run content, §17.13 revisions, §17.16 fields):
// - characters: `w:t`; `w:tab`/`w:ptab` → TAB; `w:br`/`w:cr` → LF; `w:noBreakHyphen` → U+2011;
//   `w:softHyphen` → U+00AD; `w:sym` → the character of its hex `w:char`; `m:t` (math) as is;
// - skipped: deleted and moved-away content (`w:del`, `w:moveFrom` subtrees, `w:delText`), field
//   instructions (`w:instrText`, `w:delInstrText`), drawings/VML/objects (text boxes are separate
//   stories) and every property element (`*Pr`);
// - traversed: everything else (hyperlinks, smart tags, content controls, insertions, moves-to,
//   simple fields, custom XML, `w:dir`/`w:bdo`, ...). A complex field therefore shows its result
//   text, because only its instructions are skipped;
// - `mc:AlternateContent`: the Fallback when present, else the first Choice.
//
// Explicit stacks only: documents can be nested arbitrarily deep.

import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';

/** Receives each element that contributes visible text, with the text it contributes. */
export type VisibleVisitor = (node: XElement, text: string) => void;

/** `w:` elements whose whole subtree shows nothing. */
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

/** The branch of an `mc:AlternateContent` that text and structure are read from. */
export function chooseAlternate(ac: XElement): XElement | undefined {
  return ac.element(NS.mc, 'Fallback') ?? ac.element(NS.mc, 'Choice');
}

/** Character for `w:sym` (`w:char` is hex, e.g. `F0E0`); '' when missing or invalid. */
function symChar(el: XElement): string {
  const hex = el.attrNS(NS.w, 'char');
  if (hex === undefined || !/^[0-9A-Fa-f]{1,6}$/.test(hex)) return '';
  const cp = Number.parseInt(hex, 16);
  return cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : '';
}

/** Text of a `w:` leaf element, or undefined when `el` is not a text-bearing leaf. */
function leafText(el: XElement): string | undefined {
  switch (el.local) {
    case 't':
      return el.text;
    case 'tab':
    case 'ptab':
      return '\t';
    case 'br':
    case 'cr':
      return '\n';
    case 'noBreakHyphen':
      return '‑';
    case 'softHyphen':
      return '­';
    case 'sym':
      return symChar(el);
    default:
      return undefined;
  }
}

/**
 * Walks `root` (itself included) in document order and calls `visit(node, text)` for every
 * element that contributes visible text (`w:t`, `w:tab`, `w:br`, `w:sym`, ...). Elements
 * contributing nothing (an empty `w:t`) are not reported. No paragraph separators are emitted:
 * walk paragraphs one by one for that (see {@link blockParagraphs}).
 */
export function walkVisible(root: XElement, visit: VisibleVisitor): void {
  const stack: XElement[] = [root];
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el.uri === NS.w) {
      const t = leafText(el);
      if (t !== undefined) {
        if (t !== '') visit(el, t);
        continue;
      }
      if (SKIP.has(el.local)) continue;
    } else if (el.uri === NS.mc) {
      if (el.local === 'AlternateContent') {
        const branch = chooseAlternate(el);
        if (branch) stack.push(branch);
        continue;
      }
    } else if (el.uri === NS.m && el.local === 't') {
      const t = el.text;
      if (t !== '') visit(el, t);
      continue;
    }
    if (el.local.endsWith('Pr')) continue;
    const kids = el.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
  }
}

/** Visible text of a paragraph (or of any element: a run, a hyperlink, ...). */
export function paragraphText(p: XElement): string {
  let out = '';
  walkVisible(p, (_node, text) => {
    out += text;
  });
  return out;
}

/**
 * The paragraphs (`w:p`) of a story container, deep and in document order: paragraphs inside
 * tables, content controls and custom XML are included; paragraphs inside text boxes
 * (drawings/VML) and deleted content are not. `mc:AlternateContent` follows
 * {@link chooseAlternate}.
 */
export function blockParagraphs(root: XElement): XElement[] {
  const out: XElement[] = [];
  const stack: XElement[] = [root];
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el.uri === NS.w) {
      if (el.local === 'p') {
        out.push(el);
        continue;
      }
      if (SKIP.has(el.local)) continue;
    } else if (el.uri === NS.mc && el.local === 'AlternateContent') {
      const branch = chooseAlternate(el);
      if (branch) stack.push(branch);
      continue;
    }
    if (el.local.endsWith('Pr')) continue;
    const kids = el.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
  }
  return out;
}
