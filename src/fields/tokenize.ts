// Field instruction syntax: tokens, the field type, positional arguments and switches.
//
// Ported from turbomerge `src/tokenize.ts` (MIT, same author; itself a port of docx-mailmerge,
// MIT, see NOTICE) and extended to Word's own rules (ECMA-376 Part 1 §17.16.1):
// - whitespace separates tokens; only double quotes group (single quotes are literal characters
//   that date/number pictures use for literal text);
// - inside quotes `\"` is a literal quote and `\\` a literal backslash;
// - adjacent quoted and unquoted text concatenate; `""` is an empty token;
// - an unterminated quote closes at the end of the instruction;
// - outside quotes `\x` starts a switch. The general switches `\*`, `\#` and `\@` always take
//   an argument; a letter switch takes one when the field type says so (see ARG_SWITCHES), and
//   for unknown types when the next token is quoted;
// - a formula field (`= expr`) keeps its expression whole: it is not whitespace-separated.

import type { FieldSwitch } from './types.js';

/** One token of an instruction. */
export interface Token {
  text: string;
  /** Some or all of the token was in double quotes (so it is never a switch or an operator). */
  quoted: boolean;
}

/** An instruction split into its parts. */
export interface ParsedInstruction {
  /** Upper-case field type ('' when the instruction is empty). */
  type: string;
  /** Positional arguments after the type (formula: the expression). */
  args: Token[];
  switches: FieldSwitch[];
}

/** Splits an instruction into tokens (see the file header for the rules). */
export function tokenize(instr: string): Token[] {
  const out: Token[] = [];
  let text = '';
  let inToken = false;
  let quoted = false;
  let inQuote = false;
  /** The token began outside quotes with a backslash: a switch, even if quoted text follows. */
  let switchy = false;
  const flushToken = (): void => {
    out.push({ text, quoted: quoted && !switchy });
    text = '';
    inToken = false;
    quoted = false;
    switchy = false;
  };
  for (let i = 0; i < instr.length; i++) {
    const ch = instr[i]!;
    if (inQuote) {
      if (ch === '\\' && (instr[i + 1] === '"' || instr[i + 1] === '\\')) {
        text += instr[++i];
      } else if (ch === '"') {
        inQuote = false;
      } else {
        text += ch;
      }
      continue;
    }
    if (ch === ' ' || ch === '\t' || ch === '\r' || ch === '\n' || ch === '\u00a0') {
      if (inToken) flushToken();
    } else if (ch === '"') {
      inQuote = true;
      inToken = true;
      quoted = true;
    } else {
      if (!inToken && ch === '\\') switchy = true;
      text += ch;
      inToken = true;
    }
  }
  if (inToken) flushToken();
  return out;
}

/** Letter switches that take an argument, per field type (the others are flags). */
const ARG_SWITCHES: Readonly<Record<string, string>> = {
  MERGEFIELD: 'bf',
  REF: 'd',
  SEQ: 'rs',
  ASK: 'd',
  FILLIN: 'd',
  HYPERLINK: 'lot',
  INCLUDETEXT: 'cx',
  INCLUDEPICTURE: 'c',
  SYMBOL: 'fs',
  TC: 'fl',
  TOC: 'abcdfstnolp',
  XE: 'frty',
  INDEX: 'bcdefghklprsyz',
  DATE: '',
  TIME: '',
};

/** Switches whose argument is optional (TOC `\o`, `\n`, `\l`): only a range or quoted text. */
const OPTIONAL_ARG: Readonly<Record<string, string>> = { TOC: 'nol' };

const RANGE = /^\d+(?:-\d+)?$/;

function isSwitch(t: Token): boolean {
  return !t.quoted && t.text.length >= 2 && t.text[0] === '\\';
}

/**
 * Parses switches from `tokens` (starting at `from`) for a field of `type`; tokens that are not
 * switches or switch arguments are returned as positional arguments.
 */
function parseRest(type: string, tokens: readonly Token[], from: number): ParsedInstruction {
  const args: Token[] = [];
  const switches: FieldSwitch[] = [];
  const takes = ARG_SWITCHES[type];
  const optional = OPTIONAL_ARG[type] ?? '';
  for (let i = from; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (!isSwitch(t)) {
      args.push(t);
      continue;
    }
    const name = t.text.slice(0, 2);
    const inline = t.text.slice(2);
    if (inline !== '') {
      switches.push({ name, arg: inline });
      continue;
    }
    const letter = name[1]!.toLowerCase();
    const next = tokens[i + 1];
    let wants: boolean;
    if (next === undefined || isSwitch(next)) wants = false;
    else if (name === '\\*' || name === '\\#' || name === '\\@') wants = true;
    else if (optional.includes(letter)) wants = next.quoted || RANGE.test(next.text);
    else if (takes !== undefined) wants = takes.includes(letter);
    else wants = next.quoted;
    if (wants && next !== undefined) {
      switches.push({ name, arg: next.text });
      i++;
    } else {
      switches.push({ name });
    }
  }
  return { type, args, switches };
}

/** Index of the first `\` outside double quotes in `s`, or -1. */
function firstSwitch(s: string): number {
  let inQuote = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') inQuote = !inQuote;
    else if (ch === '\\' && !inQuote) return i;
  }
  return -1;
}

/** Splits an instruction into type, positional arguments and switches. */
export function parseInstruction(instr: string): ParsedInstruction {
  const trimmed = instr.replace(/^[\s\u00a0]+/, '');
  if (trimmed.startsWith('=')) {
    const body = trimmed.slice(1);
    const cut = firstSwitch(body);
    const expr = (cut < 0 ? body : body.slice(0, cut)).trim();
    const rest = parseRest('=', cut < 0 ? [] : tokenize(body.slice(cut)), 0);
    return { type: '=', args: [{ text: expr, quoted: false }], switches: rest.switches };
  }
  const tokens = tokenize(trimmed);
  const first = tokens[0];
  if (first === undefined || isSwitch(first)) return parseRest('', tokens, 0);
  return parseRest(first.text.toUpperCase(), tokens, 1);
}

/** First switch named `name` (e.g. '\\h'); case-insensitive for letter switches. */
export function findSwitch(
  switches: readonly FieldSwitch[],
  name: string,
): FieldSwitch | undefined {
  const lower = name.toLowerCase();
  return switches.find((s) => s.name.toLowerCase() === lower);
}

/** True when a switch named `name` is present. */
export function hasSwitch(switches: readonly FieldSwitch[], name: string): boolean {
  return findSwitch(switches, name) !== undefined;
}
