// Fields as the IR builder sees them.
//
// A complex field is `fldChar begin` → instruction (`w:instrText`, possibly split over runs and
// containing nested fields) → `fldChar separate` → result runs → `fldChar end`, and may span
// paragraphs (a TOC). The builder only needs to know, for every piece of content it meets, where
// that content goes:
// - into an instruction (some enclosing field is still before its `separate`): not shown;
// - into the fallback of a page-dependent field (PAGE, NUMPAGES, SECTIONPAGES, PAGEREF): not shown,
//   the paginator evaluates the field per page and uses the current result only as a fallback;
// - otherwise it is shown, with the link of an enclosing HYPERLINK field if any.
// The state persists across the paragraphs of one story.

import type { ResolvedRunProps } from '../../resolve/types.js';
import type { IRInline, IRLink } from '../types.js';

export type FieldInline = Extract<IRInline, { kind: 'field' }>;

/** Field types evaluated per page by the paginator. */
export const PAGE_FIELDS: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'PAGE',
  'NUMPAGES',
  'SECTIONPAGES',
  'PAGEREF',
]);

/** A parsed field instruction. */
export interface Instruction {
  /** Upper-case field type ('' when the instruction is empty). */
  type: string;
  /** Positional arguments (unquoted). */
  args: string[];
  /** Switches in order: name (e.g. `\*`, `\h`) and argument when it has one. */
  switches: { name: string; arg?: string }[];
}

/** Switches that take an argument (for the field types the layout cares about). */
const ARG_SWITCHES: ReadonlySet<string> = /* @__PURE__ */ new Set([
  '\\*',
  '\\#',
  '\\@',
  '\\l',
  '\\o',
  '\\t',
  '\\m',
  '\\n',
  '\\s',
  '\\c',
  '\\f',
  '\\b',
  '\\a',
]);

/** Splits an instruction into tokens, honouring double quotes (`\"` inside quotes is a quote). */
function tokens(instr: string): { text: string; quoted: boolean }[] {
  const out: { text: string; quoted: boolean }[] = [];
  let i = 0;
  const n = instr.length;
  while (i < n) {
    const c = instr[i]!;
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === '\u00a0') {
      i++;
      continue;
    }
    if (c === '"' || c === '\u201c' || c === '\u201d') {
      let j = i + 1;
      let text = '';
      while (j < n && instr[j] !== '"' && instr[j] !== '\u201d' && instr[j] !== '\u201c') {
        if (instr[j] === '\\' && instr[j + 1] === '"') {
          text += '"';
          j += 2;
          continue;
        }
        text += instr[j];
        j++;
      }
      out.push({ text, quoted: true });
      i = j + 1;
      continue;
    }
    let j = i;
    while (j < n && !/[\s"\u00a0]/.test(instr[j]!)) j++;
    out.push({ text: instr.slice(i, j), quoted: false });
    i = j;
  }
  return out;
}

/** Parses a field instruction (`PAGEREF _Toc1 \h`, `HYPERLINK "url" \l "x"`, ...). */
export function parseInstruction(instr: string): Instruction {
  const toks = tokens(instr);
  const out: Instruction = { type: '', args: [], switches: [] };
  if (toks.length === 0) return out;
  out.type = toks[0]!.quoted ? toks[0]!.text : toks[0]!.text.toUpperCase();
  for (let i = 1; i < toks.length; i++) {
    const t = toks[i]!;
    if (!t.quoted && t.text.startsWith('\\') && t.text.length >= 2) {
      const name = t.text.slice(0, 2);
      const rest = t.text.slice(2);
      if (rest !== '') {
        out.switches.push({ name, arg: rest });
      } else if (ARG_SWITCHES.has(name) && i + 1 < toks.length) {
        const next = toks[i + 1]!;
        if (next.quoted || !next.text.startsWith('\\')) {
          out.switches.push({ name, arg: next.text });
          i++;
          continue;
        }
        out.switches.push({ name });
      } else {
        out.switches.push({ name });
      }
      continue;
    }
    out.args.push(t.text);
  }
  return out;
}

/** The number format of a page field: the first `\*` argument that is not MERGEFORMAT/CHARFORMAT. */
export function fieldFormat(ins: Instruction): string | undefined {
  for (const s of ins.switches) {
    if (s.name !== '\\*' || s.arg === undefined) continue;
    const up = s.arg.toUpperCase();
    if (up === 'MERGEFORMAT' || up === 'CHARFORMAT') continue;
    return s.arg;
  }
  return undefined;
}

/** A HYPERLINK field's target. */
export function hyperlinkTarget(ins: Instruction): IRLink | undefined {
  const link: IRLink = {};
  const url = ins.args[0];
  if (url !== undefined && url !== '') link.url = url;
  for (const s of ins.switches) {
    if (s.name === '\\l' && s.arg !== undefined) link.anchor = s.arg;
  }
  return link.url === undefined && link.anchor === undefined ? undefined : link;
}

/** A page field inline for an instruction (undefined when it is not a page field). */
export function pageFieldInline(
  ins: Instruction,
  props: ResolvedRunProps,
  fallback: string,
): FieldInline | undefined {
  if (!PAGE_FIELDS.has(ins.type)) return undefined;
  const f: FieldInline = { kind: 'field', type: ins.type, fallback, props };
  if (ins.type === 'PAGEREF' && ins.args[0] !== undefined) f.arg = ins.args[0];
  const fmt = fieldFormat(ins);
  if (fmt !== undefined) f.format = fmt;
  return f;
}

interface Frame {
  instr: string;
  phase: 'instr' | 'result';
  /** The page field inline this field became (result text feeds its fallback). */
  page?: FieldInline;
  /** Whether `page.props` was taken from a result run yet. */
  propsTaken: boolean;
  link?: IRLink;
}

/** Where a piece of content goes, see the file header. */
export type Route =
  | { to: 'show'; link: IRLink | undefined }
  | { to: 'instr' }
  | { to: 'fallback'; field: FieldInline; first: boolean };

/** Complex-field state of one story. */
export class FieldState {
  readonly #stack: Frame[] = [];

  /** True while some field is open. */
  get active(): boolean {
    return this.#stack.length > 0;
  }

  /** Where content met now goes. `first`: no result run gave the page field its formatting yet. */
  route(): Route {
    const s = this.#stack;
    for (let i = s.length - 1; i >= 0; i--) if (s[i]!.phase === 'instr') return { to: 'instr' };
    for (let i = s.length - 1; i >= 0; i--) {
      const f = s[i]!;
      if (f.page !== undefined) return { to: 'fallback', field: f.page, first: !f.propsTaken };
    }
    for (let i = s.length - 1; i >= 0; i--) {
      const l = s[i]!.link;
      if (l !== undefined) return { to: 'show', link: l };
    }
    return { to: 'show', link: undefined };
  }

  /** Records that the innermost page field took its formatting from a result run. */
  tookProps(): void {
    const s = this.#stack;
    for (let i = s.length - 1; i >= 0; i--) {
      if (s[i]!.page !== undefined) {
        s[i]!.propsTaken = true;
        return;
      }
    }
  }

  /** Instruction text (appended to the innermost field still in its instruction). */
  instrText(text: string): void {
    const s = this.#stack;
    for (let i = s.length - 1; i >= 0; i--) {
      if (s[i]!.phase === 'instr') {
        s[i]!.instr += text;
        return;
      }
    }
  }

  /** `fldChar begin`. Nested result text of a field inside an instruction joins that instruction. */
  begin(): void {
    this.#stack.push({ instr: '', phase: 'instr', propsTaken: false });
  }

  /**
   * `fldChar separate`: the innermost field's instruction is complete. Returns the page field
   * inline to insert at this point (when the field is page-dependent and visible).
   */
  separate(props: ResolvedRunProps): FieldInline | undefined {
    const f = this.#stack[this.#stack.length - 1];
    if (f === undefined || f.phase !== 'instr') return undefined;
    f.phase = 'result';
    return this.#classify(f, props);
  }

  /**
   * `fldChar end`. Returns a page field inline to insert when the field had no result part
   * (so it was never inserted at `separate`).
   */
  end(props: ResolvedRunProps): FieldInline | undefined {
    const f = this.#stack.pop();
    if (f === undefined) return undefined;
    if (f.phase === 'instr') {
      // No separate: the field shows nothing in Word until updated; a page field still evaluates.
      const shownHere = this.#stack.every((g) => g.phase === 'result' && g.page === undefined);
      if (!shownHere) return undefined;
      const ins = parseInstruction(f.instr);
      const inline = pageFieldInline(ins, props, '');
      if (inline !== undefined) return inline;
      // Content of a non-page field without a result: nested result text it collected is lost.
      return undefined;
    }
    return undefined;
  }

  #classify(f: Frame, props: ResolvedRunProps): FieldInline | undefined {
    const ins = parseInstruction(f.instr);
    if (ins.type === 'HYPERLINK') {
      const link = hyperlinkTarget(ins);
      if (link !== undefined) f.link = link;
      return undefined;
    }
    // Only a field whose result would be shown becomes a page field inline.
    const outer = this.#stack.slice(0, -1);
    const shown = outer.every((g) => g.phase === 'result' && g.page === undefined);
    if (!shown) return undefined;
    const inline = pageFieldInline(ins, props, '');
    if (inline !== undefined) f.page = inline;
    return inline;
  }
}
