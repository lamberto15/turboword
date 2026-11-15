// Built-in field evaluators (ECMA-376 Part 1 §17.16.5): each turns a parsed instruction into
// the new result text, or undefined to leave the field as it is (unknown bookmark, no data, no
// page information, no clock for DATE/TIME...).
//
// The comparison logic of IF/COMPARE is ported from turbomerge `src/fields.ts` (MIT, same
// author; docx-mailmerge lineage, see NOTICE), made case-sensitive as Word is: `=` and `<>`
// match `*`/`?` wildcards when the right operand is quoted; both operands numeric compare as
// numbers, otherwise as strings. Operands may be bookmark names, numbers, quoted strings or
// formulas (`IF {Qty} * 2 > 10 ...`).

import { Table } from '../wml/table.js';
import { paragraphText } from '../wml/text.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { DocInfo } from './docinfo.js';
import {
  dateText,
  type FormatContext,
  formatValue,
  hasTextFormat,
  separators,
  toNumber,
} from './format.js';
import { type CellValue, evaluateFormula, type FormulaEnv, roundHalfAway } from './formula.js';
import type { Scan } from './scan.js';
import { findSwitch, hasSwitch, type ParsedInstruction, type Token } from './tokenize.js';
import type { PageInfoProvider } from './types.js';
import { alphabetic, roman } from './words.js';
import type { RunFormat } from './write.js';

/** What an evaluator produces. */
export interface Outcome {
  text: string;
  format?: RunFormat;
  /** Evaluated, but the result is not written (SET, ASK). */
  noWrite?: boolean;
}

/** Shared state of one update run. */
export interface UpdateState {
  info: DocInfo;
  values: Readonly<Record<string, unknown>>;
  fmt: FormatContext;
  now: Date | undefined;
  pages: PageInfoProvider | undefined;
  /** SET/ASK bookmark values (lower-case names). */
  overlay: Map<string, string>;
  /** SEQ values, computed per story before updating. */
  seq: Map<Scan, number>;
  /** Set by an evaluator that lacked data (so a parent IF is skipped too). */
  noData: boolean;
  counts?: { words: number; chars: number };
}

type BuiltIn = (s: Scan, p: ParsedInstruction, st: UpdateState) => Outcome | string | undefined;

const MISSING: unique symbol = Symbol('missing');

/** `values[name]`: exact key, then a dot path through nested objects, then case-insensitive. */
export function lookup(values: Readonly<Record<string, unknown>>, name: string): unknown {
  if (Object.hasOwn(values, name)) return values[name];
  if (name.includes('.')) {
    let cur: unknown = values;
    let ok = true;
    for (const seg of name.split('.')) {
      if (cur !== null && typeof cur === 'object' && Object.hasOwn(cur, seg)) {
        cur = (cur as Record<string, unknown>)[seg];
      } else {
        ok = false;
        break;
      }
    }
    if (ok) return cur;
  }
  const lower = name.toLowerCase();
  for (const k of Object.keys(values)) if (k.toLowerCase() === lower) return values[k];
  return MISSING;
}

const arg = (p: ParsedInstruction, i: number): string | undefined => p.args[i]?.text;
const sw = (p: ParsedInstruction, name: string): string | undefined =>
  findSwitch(p.switches, name)?.arg;

// ---- comparisons (IF, COMPARE) ---------------------------------------------------------------

const OPS: ReadonlySet<string> = new Set(['=', '<>', '<', '<=', '>', '>=']);

interface Operand {
  text: string;
  num: number | undefined;
  quoted: boolean;
}

function wildcard(text: string, pattern: string): boolean {
  const re = pattern
    .replace(/[.+^${}()|[\]\\]/g, '\\$&')
    .replace(/\*/g, '.*')
    .replace(/\?/g, '.');
  return new RegExp(`^(?:${re})$`, 's').test(text);
}

function bookmarkText(st: UpdateState, name: string): string | undefined {
  const o = st.overlay.get(name.toLowerCase());
  if (o !== undefined) return o;
  const t = st.info.bookmark(name)?.text;
  return t?.replace(/[\r\n]+$/, '');
}

function operand(st: UpdateState, toks: readonly Token[], s: Scan): Operand {
  const text = toks.map((t) => t.text).join(' ');
  const loc = st.fmt.locale;
  if (toks.length === 1 && toks[0]!.quoted) return { text, num: toNumber(text, loc), quoted: true };
  if (toks.length === 1) {
    const n = toNumber(text, loc);
    if (n !== undefined) return { text, num: n, quoted: false };
    const b = /^[\p{L}_][\p{L}\p{N}_]*$/u.test(text) ? bookmarkText(st, text) : undefined;
    if (b !== undefined) return { text: b, num: toNumber(b, loc), quoted: false };
    return { text, num: undefined, quoted: false };
  }
  try {
    const n = evaluateFormula(text, formulaEnv(st, s));
    return { text: String(n), num: n, quoted: false };
  } catch {
    return { text, num: undefined, quoted: false };
  }
}

function compare(a: Operand, op: string, b: Operand): boolean {
  const numeric = a.num !== undefined && b.num !== undefined;
  if (op === '=' || op === '<>') {
    let eq: boolean;
    if (b.quoted && /[*?]/.test(b.text)) eq = wildcard(a.text, b.text);
    else eq = numeric ? a.num === b.num : a.text === b.text;
    return op === '=' ? eq : !eq;
  }
  const x = numeric ? a.num! : a.text;
  const y = numeric ? b.num! : b.text;
  switch (op) {
    case '<':
      return x < y;
    case '<=':
      return x <= y;
    case '>':
      return x > y;
    default:
      return x >= y;
  }
}

/** Splits IF/COMPARE arguments into [left, op, right, rest]; undefined without an operator. */
function condition(
  st: UpdateState,
  p: ParsedInstruction,
  s: Scan,
): { value: boolean; rest: Token[] } | undefined {
  const args = p.args;
  const at = args.findIndex((t, i) => i > 0 && !t.quoted && OPS.has(t.text));
  if (at < 0) {
    const n = args[0] === undefined ? undefined : toNumber(args[0].text, st.fmt.locale);
    if (n === undefined) return undefined;
    return { value: n !== 0, rest: args.slice(1) };
  }
  const right = args[at + 1];
  if (right === undefined) return undefined;
  const value = compare(operand(st, args.slice(0, at), s), args[at]!.text, operand(st, [right], s));
  return { value, rest: args.slice(at + 2) };
}

// ---- formula environment ---------------------------------------------------------------------

function firstNumber(text: string, locale: string): number | undefined {
  const n = toNumber(text, locale);
  if (n !== undefined) return n;
  const m = /[-+]?[\d.,]*\d/.exec(text);
  return m ? toNumber(m[0], locale) : undefined;
}

function tableEnv(s: Scan, st: UpdateState): FormulaEnv['table'] {
  const tc = s.begin.closest(NS.w, 'tc');
  const tbl = tc?.closest(NS.w, 'tbl');
  if (tc === undefined || tbl === undefined) return undefined;
  const grid = new Table(tbl, s.story).grid();
  let row = -1;
  let col = -1;
  for (let r = 0; r < grid.rows && row < 0; r++) {
    for (let c = 0; c < grid.cols; c++) {
      const g = grid.at(r, c);
      if (g?.origin && g.cell.element === tc) {
        row = r;
        col = c;
        break;
      }
    }
  }
  if (row < 0) return undefined;
  const cache = new Map<XElement, CellValue>();
  return {
    row,
    col,
    rows: grid.rows,
    cols: grid.cols,
    cell(r, c) {
      const g = grid.at(r, c);
      if (g === undefined) return undefined;
      const el = g.cell.element;
      let v = cache.get(el);
      if (v === undefined) {
        const text = g.cell.text().trim();
        v = { id: el, value: text === '' ? null : (toNumber(text, st.fmt.locale) ?? null) };
        cache.set(el, v);
      }
      return v;
    },
  };
}

export function formulaEnv(st: UpdateState, s: Scan): FormulaEnv {
  const fromSettings = st.info.separatorsFromSettings();
  const loc = separators(st.fmt.locale);
  const decimal = fromSettings.decimal ?? loc.decimal;
  return {
    decimal,
    group: decimal === loc.decimal ? loc.group : decimal === ',' ? '.' : ',',
    listSep: fromSettings.list ?? (decimal === ',' ? ';' : ','),
    bookmark: (name) => {
      const t = bookmarkText(st, name);
      return t === undefined ? undefined : (firstNumber(t, st.fmt.locale) ?? 0);
    },
    table: tableEnv(s, st),
  };
}

// ---- helpers ---------------------------------------------------------------------------------

function paragraphOf(el: XElement): XElement | undefined {
  return el.is(NS.w, 'p') ? el : el.closest(NS.w, 'p');
}

/** -1 when `a` comes before `b` in document order, 1 after, 0 when unrelated. */
export function docOrder(a: XElement, b: XElement): number {
  if (a === b) return 0;
  const path = (n: XElement): XElement[] => {
    const out: XElement[] = [n];
    let p = n.parent;
    while (p !== null && p.nodeType === 1) {
      out.unshift(p);
      p = p.parent;
    }
    return out;
  };
  const pa = path(a);
  const pb = path(b);
  let i = 0;
  while (i < pa.length && i < pb.length && pa[i] === pb[i]) i++;
  if (i === pa.length) return -1;
  if (i === pb.length) return 1;
  if (i === 0) return 0;
  const parent = pa[i - 1]!;
  return parent.children.indexOf(pa[i]!) < parent.children.indexOf(pb[i]!) ? -1 : 1;
}

/** `\n`/`\r`/`\w`/`\t` paragraph-number forms of a list label. */
function numberForm(label: string, p: ParsedInstruction): string {
  let out = label;
  if (!hasSwitch(p.switches, '\\w')) out = out.replace(/[.\s]+$/, '');
  if (hasSwitch(p.switches, '\\t')) {
    const m = /[\dA-Za-z]+(?:[.\-:][\dA-Za-z]+)*[.)]?$/.exec(out.replace(/[.\s]+$/, ''));
    if (m) out = m[0];
  }
  return out;
}

function wantsNumber(p: ParsedInstruction): boolean {
  return ['\\n', '\\r', '\\w'].some((n) => hasSwitch(p.switches, n));
}

const fmtText = (st: UpdateState, v: unknown, p: ParsedInstruction): string =>
  formatValue(v, p.switches, st.fmt);

/** Number of a note reference (footnote or endnote), formatted per the note settings. */
function noteNumber(st: UpdateState, ref: XElement): string | undefined {
  const kind = ref.local === 'endnoteReference' ? 'endnote' : 'footnote';
  const refs = st.info.doc.body.element
    .descendants(NS.w, ref.local)
    .filter((r) => r.attrNS(NS.w, 'customMarkFollows') === undefined);
  const at = refs.indexOf(ref);
  if (at < 0) return undefined;
  const pr = st.info.settings()?.element(NS.w, `${kind}Pr`);
  const start = Number(pr?.element(NS.w, 'numStart')?.attrNS(NS.w, 'val') ?? 1) || 1;
  const fmt =
    pr?.element(NS.w, 'numFmt')?.attrNS(NS.w, 'val') ??
    (kind === 'endnote' ? 'lowerRoman' : 'decimal');
  const n = start + at;
  switch (fmt) {
    case 'lowerRoman':
      return roman(n).toLowerCase();
    case 'upperRoman':
      return roman(n);
    case 'lowerLetter':
      return alphabetic(n).toLowerCase();
    case 'upperLetter':
      return alphabetic(n);
    default:
      return String(n);
  }
}

function pageText(pages: PageInfoProvider, node: XElement, p: ParsedInstruction): unknown {
  const n = pages.pageOf(node);
  if (n === undefined) return undefined;
  if (hasTextFormat(p.switches)) return n;
  return pages.pageNumberText?.(node) ?? n;
}

function counts(st: UpdateState): { words: number; chars: number } {
  if (st.counts) return st.counts;
  const text = st.info.doc.body.text();
  const words = text.split(/\s+/).filter((w) => w !== '').length;
  const chars = text.replace(/\s/g, '').length;
  st.counts = { words, chars };
  return st.counts;
}

// ---- evaluators ------------------------------------------------------------------------------

const CORE_TEXT: Readonly<
  Record<
    string,
    | 'title'
    | 'subject'
    | 'creator'
    | 'keywords'
    | 'description'
    | 'lastModifiedBy'
    | 'category'
    | 'revision'
    | 'contentStatus'
  >
> = {
  TITLE: 'title',
  SUBJECT: 'subject',
  AUTHOR: 'creator',
  KEYWORDS: 'keywords',
  COMMENTS: 'description',
  LASTSAVEDBY: 'lastModifiedBy',
  CATEGORY: 'category',
  REVISIONNUMBER: 'revision',
  CONTENTSTATUS: 'contentStatus',
};

const APP_NUMBERS: Readonly<
  Record<
    string,
    | 'pages'
    | 'words'
    | 'characters'
    | 'charactersWithSpaces'
    | 'lines'
    | 'paragraphs'
    | 'totalTime'
    | 'docSecurity'
  >
> = {
  PAGES: 'pages',
  WORDS: 'words',
  CHARACTERS: 'characters',
  CHARACTERSWITHSPACES: 'charactersWithSpaces',
  LINES: 'lines',
  PARAGRAPHS: 'paragraphs',
  TOTALEDITINGTIME: 'totalTime',
  SECURITY: 'docSecurity',
};

/** A document property by Word's DOCPROPERTY name (built-in first, then custom). */
function docProperty(st: UpdateState, name: string): unknown {
  const props = st.info.doc.properties;
  const key = name.replace(/\s+/g, '').toUpperCase();
  const core = CORE_TEXT[key];
  if (core !== undefined) return props[core];
  switch (key) {
    case 'CREATETIME':
      return props.created;
    case 'LASTSAVEDTIME':
      return props.modified;
    case 'LASTPRINTED':
      return props.lastPrinted;
    case 'COMPANY':
      return props.app.company;
    case 'TEMPLATE':
      return props.app.template;
    case 'NAMEOFAPPLICATION':
      return props.app.application;
    case 'MANAGER':
      return st.info.appValue('Manager');
    case 'HYPERLINKBASE':
      return st.info.appValue('HyperlinkBase');
    default:
      break;
  }
  const app = APP_NUMBERS[key];
  if (app !== undefined) return props.app[app];
  const custom = props.custom();
  if (Object.hasOwn(custom, name)) return custom[name];
  const lower = name.toLowerCase();
  for (const k of Object.keys(custom)) if (k.toLowerCase() === lower) return custom[k];
  return undefined;
}

function dateField(
  st: UpdateState,
  p: ParsedInstruction,
  d: Date | undefined,
  time?: 'time' | 'date',
): string | undefined {
  if (d === undefined) return undefined;
  if (findSwitch(p.switches, '\\@')?.arg !== undefined) return fmtText(st, d, p);
  let base: string;
  if (time === 'time') {
    base = new Intl.DateTimeFormat(st.fmt.locale, {
      hour: 'numeric',
      minute: '2-digit',
      ...(st.fmt.timeZone === undefined ? {} : { timeZone: st.fmt.timeZone }),
    }).format(d);
  } else {
    base = dateText(d, st.fmt, time !== 'date');
  }
  return fmtText(st, base, p);
}

function refField(
  st: UpdateState,
  s: Scan,
  p: ParsedInstruction,
  name: string,
): string | undefined {
  const b = st.info.bookmark(name);
  const overlay = st.overlay.get(name.toLowerCase());
  let text = overlay ?? b?.text?.replace(/[\r\n]+$/, '');
  if (text === undefined) return undefined;
  if (wantsNumber(p)) {
    if (b === undefined) return undefined;
    const para = paragraphOf(b.start);
    const label = para ? st.info.listLabel(para, b.story) : undefined;
    if (label === undefined) return undefined;
    text = numberForm(label, p);
  }
  if (hasSwitch(p.switches, '\\p') && b !== undefined) {
    const where = docOrder(b.start, s.begin) < 0 ? 'above' : 'below';
    text = wantsNumber(p) ? `${text} ${where}` : where;
  }
  return fmtText(st, text, p);
}

function styleRef(st: UpdateState, s: Scan, p: ParsedInstruction): string | undefined {
  if (s.story.kind !== 'body') return undefined;
  let name = arg(p, 0);
  if (name === undefined) return undefined;
  if (/^[1-9]$/.test(name)) name = `heading ${name}`;
  const info = st.info;
  const paraId = info.styleId(name, 'paragraph');
  const charId = paraId === undefined ? info.styleId(name, 'character') : undefined;
  if (paraId === undefined && charId === undefined) return undefined;
  const index = info.paragraphOrder(s.story);
  const paras = [...index.keys()];
  const here = paragraphOf(s.begin);
  const at = here === undefined ? -1 : (index.get(here) ?? -1);
  const sequence = [...paras.slice(0, at + 1).reverse(), ...paras.slice(at + 1)];
  for (const para of sequence) {
    let text: string | undefined;
    if (paraId !== undefined) {
      if (info.paragraphStyle(para) === paraId) text = paragraphText(para);
    } else {
      const runs = para
        .descendants(NS.w, 'r')
        .filter(
          (r) => r.element(NS.w, 'rPr')?.element(NS.w, 'rStyle')?.attrNS(NS.w, 'val') === charId,
        );
      if (runs.length > 0) text = runs.map((r) => paragraphText(r)).join('');
    }
    if (text === undefined) continue;
    let out = text.replace(/\s+$/, '');
    if (wantsNumber(p)) {
      const label = info.listLabel(para, s.story);
      if (label === undefined) return undefined;
      out = numberForm(label, p);
    }
    if (hasSwitch(p.switches, '\\p')) {
      const where = index.get(para)! <= at ? 'above' : 'below';
      out = wantsNumber(p) ? `${out} ${where}` : where;
    }
    return fmtText(st, out, p);
  }
  return undefined;
}

const SYMBOL_FONTS = /^(symbol|wingdings( [23])?|webdings|marlett)$/i;

export const BUILTINS: Readonly<Record<string, BuiltIn>> = {
  MERGEFIELD(_s, p, st) {
    const name = arg(p, 0);
    if (name === undefined) return undefined;
    const v = lookup(st.values, name);
    if (v === MISSING) {
      st.noData = true;
      return undefined;
    }
    const text = fmtText(st, v, p);
    if (text === '') return '';
    // Paragraph marks in \b/\f text become line breaks, as Word writes them.
    const lines = (t: string | undefined): string => (t ?? '').replace(/\r\n?/g, '\n');
    return lines(sw(p, '\\b')) + text + lines(sw(p, '\\f'));
  },

  IF(s, p, st) {
    const c = condition(st, p, s);
    if (c === undefined) return undefined;
    const branch = c.value ? c.rest[0] : c.rest[1];
    return fmtText(st, branch?.text ?? '', p);
  },

  COMPARE(s, p, st) {
    const c = condition(st, p, s);
    if (c === undefined) return undefined;
    return c.value ? '1' : '0';
  },

  '='(s, p, st) {
    const expr = arg(p, 0);
    if (expr === undefined) return undefined;
    let v: number;
    try {
      v = evaluateFormula(expr, formulaEnv(st, s));
    } catch {
      return undefined;
    }
    if (findSwitch(p.switches, '\\#') === undefined) v = roundHalfAway(v, 2);
    return fmtText(st, v, p);
  },

  DATE: (_s, p, st) => dateField(st, p, st.now, 'date'),
  TIME: (_s, p, st) => dateField(st, p, st.now, 'time'),
  CREATEDATE: (_s, p, st) => dateField(st, p, st.info.doc.properties.created),
  SAVEDATE: (_s, p, st) => dateField(st, p, st.info.doc.properties.modified),
  PRINTDATE: (_s, p, st) => dateField(st, p, st.info.doc.properties.lastPrinted),

  DOCPROPERTY(_s, p, st) {
    const name = arg(p, 0);
    if (name === undefined) return undefined;
    const v = docProperty(st, name);
    if (v === undefined) return undefined;
    if (typeof v === 'boolean') return fmtText(st, v ? 'Y' : 'N', p);
    // Custom date properties show the date alone (CreateTime and the like also the time).
    if (v instanceof Date && !hasSwitch(p.switches, '\\@') && !/time|printed/i.test(name)) {
      return fmtText(st, dateText(v, st.fmt, false), p);
    }
    return fmtText(st, v, p);
  },

  DOCVARIABLE(_s, p, st) {
    const name = arg(p, 0);
    const v = name === undefined ? undefined : st.info.docVar(name);
    return v === undefined ? undefined : fmtText(st, v, p);
  },

  NUMWORDS: (_s, p, st) => fmtText(st, counts(st).words, p),
  NUMCHARS: (_s, p, st) => fmtText(st, counts(st).chars, p),

  REVNUM(_s, p, st) {
    const v = st.info.doc.properties.revision;
    return v === undefined ? undefined : fmtText(st, v, p);
  },

  TEMPLATE(_s, p, st) {
    const v = st.info.doc.properties.app.template;
    return v === undefined ? undefined : fmtText(st, v, p);
  },

  EDITTIME(_s, p, st) {
    const v = st.info.doc.properties.app.totalTime;
    return v === undefined ? undefined : fmtText(st, v, p);
  },

  SECTION(s, p, st) {
    const para = paragraphOf(s.begin);
    if (s.story.kind !== 'body' || para === undefined) return undefined;
    return fmtText(st, st.info.sectionOf(para).index + 1, p);
  },

  REF(s, p, st) {
    const name = arg(p, 0);
    return name === undefined ? undefined : refField(st, s, p, name);
  },

  NOTEREF(s, p, st) {
    const name = arg(p, 0);
    const b = name === undefined ? undefined : st.info.bookmark(name);
    if (b?.noteRef === undefined) return undefined;
    let text = noteNumber(st, b.noteRef);
    if (text === undefined) return undefined;
    if (hasSwitch(p.switches, '\\p')) text += docOrder(b.start, s.begin) < 0 ? ' above' : ' below';
    const out: Outcome = { text: fmtText(st, text, p) };
    if (hasSwitch(p.switches, '\\f')) {
      const style = st.info.styleId(
        `${b.noteRef.local === 'endnoteReference' ? 'endnote' : 'footnote'} reference`,
        'character',
      );
      if (style !== undefined) out.format = { style };
    }
    return out;
  },

  PAGEREF(s, p, st) {
    const name = arg(p, 0);
    const b = name === undefined ? undefined : st.info.bookmark(name);
    if (st.pages === undefined || b === undefined) return undefined;
    const v = pageText(st.pages, b.start, p);
    if (v === undefined) return undefined;
    const text = fmtText(st, v, p);
    if (!hasSwitch(p.switches, '\\p')) return text;
    const here = st.pages.pageOf(s.begin);
    if (here !== undefined && here === st.pages.pageOf(b.start)) {
      return docOrder(b.start, s.begin) < 0 ? 'above' : 'below';
    }
    return `on page ${text}`;
  },

  PAGE(s, p, st) {
    if (st.pages === undefined) return undefined;
    const v = pageText(st.pages, s.begin, p);
    return v === undefined ? undefined : fmtText(st, v, p);
  },

  NUMPAGES: (_s, p, st) =>
    st.pages === undefined ? undefined : fmtText(st, st.pages.pageCount(), p),

  SECTIONPAGES(s, p, st) {
    const n = st.pages?.sectionPageCount?.(s.begin);
    return n === undefined ? undefined : fmtText(st, n, p);
  },

  SEQ(s, p, st) {
    const n = st.seq.get(s);
    if (n === undefined) return undefined;
    if (hasSwitch(p.switches, '\\h') && !hasTextFormat(p.switches)) return '';
    return fmtText(st, n, p);
  },

  STYLEREF: (s, p, st) => styleRef(st, s, p),

  FILLIN(_s, p, st) {
    const prompt = arg(p, 0) ?? '';
    let v = lookup(st.values, prompt);
    if (v === MISSING) v = sw(p, '\\d') ?? MISSING;
    if (v === MISSING) {
      st.noData = true;
      return undefined;
    }
    return fmtText(st, v, p);
  },

  ASK(_s, p, st) {
    const name = arg(p, 0);
    if (name === undefined) return undefined;
    let v = lookup(st.values, name);
    const prompt = arg(p, 1);
    if (v === MISSING && prompt !== undefined) v = lookup(st.values, prompt);
    if (v === MISSING) v = sw(p, '\\d') ?? MISSING;
    if (v === MISSING) {
      st.noData = true;
      return undefined;
    }
    st.overlay.set(name.toLowerCase(), fmtText(st, v, { type: 'ASK', args: [], switches: [] }));
    return { text: '', noWrite: true };
  },

  SET(_s, p, st) {
    const name = arg(p, 0);
    if (name === undefined) return undefined;
    st.overlay.set(
      name.toLowerCase(),
      p.args
        .slice(1)
        .map((t) => t.text)
        .join(' '),
    );
    return { text: '', noWrite: true };
  },

  QUOTE: (_s, p, st) => fmtText(st, p.args.map((t) => t.text).join(' '), p),

  SYMBOL(_s, p, _st) {
    const raw = arg(p, 0);
    if (raw === undefined) return undefined;
    const code = /^0x[0-9a-f]+$/i.test(raw) ? Number.parseInt(raw.slice(2), 16) : Number(raw);
    if (!Number.isInteger(code) || code <= 0 || code > 0x10ffff) return undefined;
    const font = sw(p, '\\f');
    const size = sw(p, '\\s');
    const unicode = hasSwitch(p.switches, '\\u');
    const text =
      !unicode && font !== undefined && SYMBOL_FONTS.test(font) && code < 256
        ? String.fromCharCode(0xf000 + code)
        : String.fromCodePoint(code);
    const format: RunFormat = {};
    if (font !== undefined) format.font = font;
    const pts = size === undefined ? Number.NaN : Number(size);
    if (Number.isFinite(pts) && pts > 0) format.size = Math.round(pts * 2);
    return { text, format };
  },

  AUTHOR: (_s, p, st) => coreField(st, p, 'creator'),
  TITLE: (_s, p, st) => coreField(st, p, 'title'),
  SUBJECT: (_s, p, st) => coreField(st, p, 'subject'),
  KEYWORDS: (_s, p, st) => coreField(st, p, 'keywords'),
  COMMENTS: (_s, p, st) => coreField(st, p, 'description'),
  LASTSAVEDBY: (_s, p, st) => coreField(st, p, 'lastModifiedBy'),
};

/** AUTHOR, TITLE, ...: the property, or the field's own argument when it has one. */
function coreField(
  st: UpdateState,
  p: ParsedInstruction,
  key: 'creator' | 'title' | 'subject' | 'keywords' | 'description' | 'lastModifiedBy',
): string {
  const given = p.args.length > 0 ? p.args.map((t) => t.text).join(' ') : undefined;
  return fmtText(st, given ?? st.info.doc.properties[key] ?? '', p);
}

/** Field types with a built-in evaluator (TOC is handled by the TOC updater). */
export const KNOWN_TYPES: ReadonlySet<string> = new Set([...Object.keys(BUILTINS), 'TOC']);
