// Legacy form fields: FORMTEXT, FORMCHECKBOX and FORMDROPDOWN (Word 97 "protected forms").
//
// Each is a complex field whose `begin` fldChar carries `w:ffData` (name, enabled, text-input /
// checkbox / dropdown settings). A checkbox or dropdown shows what ffData says (`w:checked`,
// `w:ddList/w:result`); a text field shows its result runs, between `separate` and `end`, and
// shows five EN SPACEs (U+2002) when empty so it can still be clicked. The scanner pairs field
// characters with a stack, so fields nested in a result are handled, and fields may span runs
// and paragraphs.

import { UsageError } from '../core/errors.js';
import type { WordDocument } from '../wml/document.js';
import { setWAttr, wElement } from '../wml/order.js';
import type { Story } from '../wml/story.js';
import { paragraphText } from '../wml/text.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { makeRun } from './content.js';
import { formatWordDate, parseIsoDate } from './date-format.js';
import type { FillResult, LegacyFormField } from './types.js';

const EMPTY_TEXT = ' '.repeat(5);
const KIND = /^\s*FORM(TEXT|CHECKBOX|DROPDOWN)\b/i;

/** ST_OnOff value of `w:<local>` in `parent` (absent element = `dflt`, absent val = on). */
function flag(parent: XElement | undefined, local: string, dflt: boolean): boolean {
  const el = parent?.element(NS.w, local);
  if (el === undefined) return dflt;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined || !(v === '0' || v === 'false' || v === 'off');
}

function val(parent: XElement | undefined, local: string): string | undefined {
  return parent?.element(NS.w, local)?.attrNS(NS.w, 'val');
}

/** Applies a text form field `w:format` (case conversions and date pictures). */
function formatText(text: string, input: XElement | undefined): string {
  const format = val(input, 'format');
  if (format === undefined || format === '') return text;
  const type = val(input, 'type') ?? 'regular';
  if (type === 'date') {
    const d = parseIsoDate(text);
    return d === undefined ? text : formatWordDate(d, format);
  }
  if (type !== 'regular') return text;
  switch (format.toUpperCase()) {
    case 'UPPERCASE':
      return text.toUpperCase();
    case 'LOWERCASE':
      return text.toLowerCase();
    case 'FIRST CAPITAL':
      return text.charAt(0).toUpperCase() + text.slice(1);
    case 'TITLE CASE':
      return text.replace(/(^|\s)(\S)/g, (_m, s: string, c: string) => s + c.toUpperCase());
    default:
      return text;
  }
}

class FormField implements LegacyFormField {
  readonly type: 'text' | 'checkbox' | 'dropdown';
  /** The `begin` `w:fldChar` (holding `w:ffData`). */
  readonly element: XElement;
  readonly #story: Story;
  #sep: XElement | undefined;
  readonly #end: XElement;

  constructor(
    type: 'text' | 'checkbox' | 'dropdown',
    story: Story,
    begin: XElement,
    sep: XElement | undefined,
    end: XElement,
  ) {
    this.type = type;
    this.#story = story;
    this.element = begin;
    this.#sep = sep;
    this.#end = end;
  }

  get #ff(): XElement {
    return this.element.element(NS.w, 'ffData')!;
  }

  get name(): string | undefined {
    return val(this.#ff, 'name');
  }

  get value(): string | boolean {
    const ff = this.#ff;
    if (this.type === 'checkbox') {
      const cb = ff.element(NS.w, 'checkBox');
      return cb?.element(NS.w, 'checked') !== undefined
        ? flag(cb, 'checked', false)
        : flag(cb, 'default', false);
    }
    if (this.type === 'dropdown') {
      const dd = ff.element(NS.w, 'ddList');
      const entries = (dd?.elements(NS.w, 'listEntry') ?? []).map((e) => e.attrNS(NS.w, 'val'));
      const i = Number(val(dd, 'result') ?? val(dd, 'default') ?? 0);
      return entries[i] ?? '';
    }
    const text = this.#resultRuns().map(paragraphText).join('');
    return text === EMPTY_TEXT ? '' : text;
  }

  setValue(v: string | boolean): void {
    const ff = this.#ff;
    const label = JSON.stringify(this.name ?? '?');
    if (this.type === 'checkbox') {
      let on: boolean;
      if (typeof v === 'boolean') on = v;
      else if (/^\s*(1|true|on|yes|x)\s*$/i.test(v)) on = true;
      else if (/^\s*(0|false|off|no|)\s*$/i.test(v)) on = false;
      else throw new UsageError(`Checkbox form field ${label} takes a boolean`);
      const cb = ff.element(NS.w, 'checkBox')!;
      let checked = cb.element(NS.w, 'checked');
      if (checked === undefined) {
        checked = wElement(cb, 'checked');
        const after =
          cb.element(NS.w, 'default') ?? cb.element(NS.w, 'size') ?? cb.element(NS.w, 'sizeAuto');
        if (after !== undefined) after.after(checked);
        else cb.prepend(checked);
      }
      setWAttr(checked, 'val', on ? '1' : '0');
      return;
    }
    if (typeof v !== 'string') throw new UsageError(`Form field ${label} takes text`);
    if (this.type === 'dropdown') {
      const dd = ff.element(NS.w, 'ddList')!;
      const entries = dd.elements(NS.w, 'listEntry').map((e) => e.attrNS(NS.w, 'val'));
      const i = entries.indexOf(v);
      if (i < 0) {
        throw new UsageError(
          `${JSON.stringify(v)} is not an entry of dropdown form field ${label} (entries: ${entries.map((e) => JSON.stringify(e)).join(', ')})`,
        );
      }
      let result = dd.element(NS.w, 'result');
      if (result === undefined) {
        result = wElement(dd, 'result');
        dd.prepend(result);
      }
      setWAttr(result, 'val', String(i));
      if (this.#sep !== undefined) this.#writeResult(v);
      return;
    }
    const input = ff.element(NS.w, 'textInput');
    const max = Number(val(input, 'maxLength') ?? 0);
    if (max > 0 && v.length > max) {
      throw new UsageError(`Text form field ${label} takes at most ${max} characters`);
    }
    this.#writeResult(formatText(v, input));
  }

  /** The runs strictly between the `separate` and `end` field characters. */
  #resultRuns(): XElement[] {
    const sepRun = this.#sep?.parent;
    const endRun = this.#end.parent;
    if (sepRun === undefined || sepRun === null || endRun === null) return [];
    const all = this.#story.element.descendants(NS.w, 'r');
    const i = all.indexOf(sepRun as XElement);
    const j = all.indexOf(endRun as XElement);
    return i < 0 || j < 0 ? [] : all.slice(i + 1, j);
  }

  #writeResult(text: string): void {
    const runs = this.#resultRuns();
    const beginRun = this.element.parent as XElement;
    const rPr = (runs.find((r) => r.element(NS.w, 'rPr') !== undefined) ?? beginRun).element(
      NS.w,
      'rPr',
    );
    const run = makeRun(beginRun, rPr, text === '' ? EMPTY_TEXT : text);
    if (this.#sep === undefined) {
      const sepRun = wElement(beginRun, 'r');
      const sep = wElement(beginRun, 'fldChar', { fldCharType: 'separate' });
      sepRun.append(sep);
      (this.#end.parent as XElement).before(sepRun, run);
      this.#sep = sep;
      return;
    }
    for (const r of runs) r.remove();
    (this.#sep.parent as XElement).after(run);
  }
}

interface Open {
  begin: XElement;
  instr: string;
  sep: XElement | undefined;
}

/** Every legacy form field of the document (all stories, document order). */
export function legacyFields(doc: WordDocument): LegacyFormField[] {
  const out: LegacyFormField[] = [];
  for (const story of doc.stories()) {
    const open: Open[] = [];
    const stack: XElement[] = [story.element];
    while (stack.length > 0) {
      const el = stack.pop()!;
      if (el.uri === NS.w) {
        if (el.local === 'fldChar') {
          const t = el.attrNS(NS.w, 'fldCharType');
          if (t === 'begin') open.push({ begin: el, instr: '', sep: undefined });
          else if (t === 'separate') {
            const top = open[open.length - 1];
            if (top !== undefined) top.sep = el;
          } else if (t === 'end') {
            const f = open.pop();
            const m = f === undefined ? null : KIND.exec(f.instr);
            if (f !== undefined && m !== null && f.begin.element(NS.w, 'ffData') !== undefined) {
              const kind = m[1]!.toUpperCase();
              const type = kind === 'TEXT' ? 'text' : kind === 'CHECKBOX' ? 'checkbox' : 'dropdown';
              out.push(new FormField(type, story, f.begin, f.sep, el));
            }
          }
          continue;
        }
        if (el.local === 'instrText') {
          const top = open[open.length - 1];
          if (top !== undefined && top.sep === undefined) top.instr += el.text;
          continue;
        }
        if (
          el.local === 'del' ||
          el.local === 'moveFrom' ||
          el.local === 'rPr' ||
          el.local === 'pPr'
        ) {
          continue;
        }
      } else if (el.uri === NS.mc && el.local === 'Fallback') {
        continue;
      }
      const kids = el.children;
      for (let i = kids.length - 1; i >= 0; i--) {
        const k = kids[i]!;
        if (k.nodeType === 1) stack.push(k);
      }
    }
  }
  return out;
}

/** See index.ts. */
export function fillLegacy(
  doc: WordDocument,
  values: Readonly<Record<string, string | boolean>>,
): FillResult {
  const fields = legacyFields(doc);
  const res: FillResult = { filled: 0, missing: [] };
  for (const [k, v] of Object.entries(values)) {
    const hits = fields.filter((f) => f.name === k);
    if (hits.length === 0) res.missing.push(k);
    for (const f of hits) {
      f.setValue(v);
      res.filled++;
    }
  }
  return res;
}
