// The field index: finds every field of a story without changing anything.
//
// Ported in spirit from turbomerge `src/parse.ts` (MIT, same author), which *replaced* fields
// with placeholders. Here the walk is read-only: it records where each field's begin, separate
// and end markers are and what text its instruction and result hold, so updates can later
// rewrite the result alone and leave every field code in place.
//
// Complex fields (ECMA-376 Part 1 §17.16.18) are a stream of markers in document order:
// `fldChar begin` … instruction … `fldChar separate` … result … `fldChar end`. They span runs,
// paragraphs and whole tables, and nest: a field that begins inside another's instruction is
// part of that instruction (its *current result* is what the outer field reads), and one that
// begins inside the result is part of the result. `w:fldSimple` is the one-element form.
// The walk keeps a stack of open fields per *flow*: text boxes (`w:txbxContent`) and each
// branch of `mc:AlternateContent` are separate flows that fields cannot cross. Deleted content
// (`w:del`, `w:moveFrom`) is ignored. Unbalanced markers are tolerated: a stray `separate`/`end`
// is ignored and a field still open at the end of its flow is dropped (its nested fields are
// kept and move up to its parent).
//
// Text conventions inside the index: paragraph marks are '\r', line breaks '\n', tabs '\t'.
// Explicit stacks only: documents can be arbitrarily deep.

import type { Story } from '../wml/story.js';
import { chooseAlternate } from '../wml/text.js';
import type { XDocument, XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { type ParsedInstruction, parseInstruction } from './tokenize.js';

/** A piece of an instruction or result: text, or a nested field. */
export type Piece = string | Scan;

/** One field found by {@link scanStory} (internal; the public view is `Field`). */
export interface Scan {
  kind: 'complex' | 'simple';
  story: Story;
  /** `w:fldChar` begin element (complex) or the `w:fldSimple` element. */
  begin: XElement;
  /** `w:fldChar` separate element (complex fields with a result section). */
  separate: XElement | undefined;
  /** `w:fldChar` end element (complex) or the `w:fldSimple` element. */
  end: XElement;
  instr: Piece[];
  result: Piece[];
  parent: Scan | undefined;
  /** Nested in the parent's instruction (true) or result (false). */
  inInstr: boolean;
  /** Nested fields in order (instruction and result). */
  children: Scan[];
  /** Inside an `mc:AlternateContent` branch that readers do not show. */
  secondary: boolean;
  /** Document order within the story. */
  index: number;
  /** Instruction text, nested fields read as their current results. */
  instrText: string;
  /** Result text. */
  resultText: string;
  /** Instruction with nested fields shown as `{code}` (Word's Alt+F9 view). */
  code: string;
  parsed: ParsedInstruction;
  /** @internal Scanning state. */
  phase: 'instr' | 'result';
}

/** ST_OnOff (§17.17.4). */
export function onOff(v: string | undefined): boolean {
  return v !== undefined && (v === 'true' || v === '1' || v === 'on');
}

/** Text of a run-content leaf, or undefined when `el` is not one. */
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
      return '\u2011';
    case 'softHyphen':
      return '\u00ad';
    case 'sym': {
      const hex = el.attrNS(NS.w, 'char');
      if (hex === undefined || !/^[0-9A-Fa-f]{1,6}$/.test(hex)) return '';
      const cp = Number.parseInt(hex, 16);
      return cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : '';
    }
    default:
      return undefined;
  }
}

/** `w:` elements whose subtree is never part of the visible content or of a field. */
const SKIP: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'del',
  'moveFrom',
  'delText',
  'delInstrText',
  'fldData',
]);

interface Frame {
  el: XElement;
  i: number;
  /** Exit action: 0 none, 1 paragraph mark, 2 close a simple field, 3 end of a flow. */
  exit: 0 | 1 | 2 | 3;
  simple?: Scan;
  savedOpen?: Scan[];
  savedSecondary?: boolean;
}

/**
 * Indexes the fields of `story`. `allBranches` also walks the `mc:AlternateContent` branches
 * readers do not show (their fields are marked `secondary`), which updates need so both
 * representations of a text box stay in step.
 */
export function scanStory(story: Story, allBranches = false): Scan[] {
  const out: Scan[] = [];
  let open: Scan[] = [];
  let secondary = false;

  const top = (): Scan | undefined => open[open.length - 1];
  const push = (text: string, instrOk: boolean): void => {
    const t = top();
    if (t === undefined) return;
    if (t.phase === 'instr') {
      if (instrOk) t.instr.push(text);
    } else {
      t.result.push(text);
    }
  };
  const openField = (kind: 'complex' | 'simple', el: XElement): Scan => {
    const parent = top();
    const s: Scan = {
      kind,
      story,
      begin: el,
      separate: undefined,
      end: el,
      instr: [],
      result: [],
      parent,
      inInstr: parent?.phase === 'instr',
      children: [],
      secondary,
      index: out.length,
      instrText: '',
      resultText: '',
      code: '',
      parsed: { type: '', args: [], switches: [] },
      phase: kind === 'simple' ? 'result' : 'instr',
    };
    if (parent) {
      parent.children.push(s);
      (parent.phase === 'instr' ? parent.instr : parent.result).push(s);
    }
    out.push(s);
    open.push(s);
    return s;
  };
  /** Drops fields left open in the current flow (keeping their nested fields). */
  const closeFlow = (keep: number): void => {
    while (open.length > keep) {
      const s = open.pop()!;
      if (s.kind === 'complex') drop(s);
    }
  };
  const drop = (s: Scan): void => {
    const i = out.indexOf(s);
    if (i >= 0) out.splice(i, 1);
    const parent = s.parent;
    for (const c of s.children) {
      c.parent = parent;
      c.inInstr = parent === undefined ? false : s.inInstr;
    }
    if (parent) {
      const at = parent.children.indexOf(s);
      parent.children.splice(at, 1, ...s.children);
      for (const list of [parent.instr, parent.result]) {
        const k = list.indexOf(s);
        if (k >= 0) list.splice(k, 1, ...s.children);
      }
    }
  };

  const stack: Frame[] = [
    { el: story.element, i: 0, exit: 3, savedOpen: [], savedSecondary: false },
  ];
  const enter = (el: XElement): void => {
    if (el.uri === NS.w) {
      const local = el.local;
      if (SKIP.has(local) || local.endsWith('Pr')) return;
      const text = leafText(el);
      if (text !== undefined) {
        push(text, local === 't');
        return;
      }
      switch (local) {
        case 'instrText':
          push(el.text, true);
          return;
        case 'fldChar':
          fldChar(el);
          return;
        case 'fldSimple': {
          const s = openField('simple', el);
          stack.push({ el, i: 0, exit: 2, simple: s });
          return;
        }
        case 'txbxContent':
          stack.push({ el, i: 0, exit: 3, savedOpen: open, savedSecondary: secondary });
          open = [];
          return;
        case 'p':
          stack.push({ el, i: 0, exit: 1 });
          return;
        default:
          break;
      }
    } else if (el.uri === NS.mc && el.local === 'AlternateContent') {
      const shown = chooseAlternate(el);
      const branches = el.elements(NS.mc);
      for (let i = branches.length - 1; i >= 0; i--) {
        const b = branches[i]!;
        if (b !== shown && !allBranches) continue;
        stack.push({ el: b, i: 0, exit: 3, savedOpen: open, savedSecondary: secondary });
      }
      return;
    }
    stack.push({ el, i: 0, exit: 0 });
  };
  const fldChar = (el: XElement): void => {
    const type = el.attrNS(NS.w, 'fldCharType');
    const t = top();
    if (type === 'begin') openField('complex', el);
    else if (type === 'separate') {
      if (t?.kind === 'complex' && t.phase === 'instr') {
        t.separate = el;
        t.phase = 'result';
      }
    } else if (type === 'end' && t?.kind === 'complex') {
      t.end = el;
      open.pop();
    }
  };

  while (stack.length > 0) {
    const f = stack[stack.length - 1]!;
    const kids = f.el.children;
    // Flow frames start fresh: set on first visit.
    if (f.i === 0 && f.exit === 3 && f.el !== story.element && f.el.uri === NS.mc) {
      secondary = f.savedSecondary === true || f.el !== chooseAlternate(f.el.parent as XElement);
      open = [];
    }
    if (f.i < kids.length) {
      const k = kids[f.i++]!;
      if (k.nodeType === 1) enter(k);
      continue;
    }
    stack.pop();
    if (f.exit === 1) push('\r', true);
    else if (f.exit === 2) {
      const at = open.indexOf(f.simple!);
      if (at >= 0) closeFlow(at + 1);
      if (open[open.length - 1] === f.simple) open.pop();
    } else if (f.exit === 3) {
      closeFlow(0);
      open = f.savedOpen ?? [];
      secondary = f.savedSecondary ?? false;
    }
  }

  for (let i = 0; i < out.length; i++) out[i]!.index = i;
  finish(out);
  return out;
}

/** Computes the texts and parsed instruction of each field (children first). */
function finish(fields: readonly Scan[]): void {
  for (let i = fields.length - 1; i >= 0; i--) {
    const s = fields[i]!;
    s.resultText = joinPieces(s.result, (c) => c.resultText);
    if (s.kind === 'simple') {
      s.instrText = s.begin.attrNS(NS.w, 'instr') ?? '';
      s.code = s.instrText;
    } else {
      s.instrText = joinPieces(s.instr, (c) => c.resultText);
      s.code = joinPieces(s.instr, (c) => `{${c.code}}`);
    }
    s.parsed = parseInstruction(s.instrText);
  }
}

/** Joins pieces, mapping nested fields through `nested`. */
export function joinPieces(pieces: readonly Piece[], nested: (s: Scan) => string): string {
  let out = '';
  for (const p of pieces) out += typeof p === 'string' ? p : nested(p);
  return out;
}

/** True when `el` is still inside `root` (not detached by an earlier update). */
export function attached(el: XElement, root: XElement): boolean {
  let n: XElement | XDocument | null = el;
  while (n !== null && n.nodeType === 1) {
    if (n === root) return true;
    n = n.parent;
  }
  return false;
}
