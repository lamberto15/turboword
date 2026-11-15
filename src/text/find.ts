// findText(): search visible text across run boundaries.
//
// Queries compile to one global RegExp. A string is matched literally (escaped) with the `u`
// flag, case-sensitively unless `caseSensitive: false`; a RegExp keeps its own flags (`g` is
// forced, `y` dropped) and `caseSensitive` only overrides `i` when given. Matching runs on each
// paragraph's flattened text, so a match never crosses a paragraph and offsets are UTF-16.
//
// Empty matches (e.g. /x*/) are skipped. `wholeWord` rejects a match whose neighbouring
// characters are letters, digits, marks or `_`, then retries one character further on (the
// same semantics as wrapping the pattern in look-arounds, without editing the user's pattern).

import { UsageError } from '../core/errors.js';
import type { WordDocument } from '../wml/document.js';
import type { Paragraph } from '../wml/paragraph.js';
import type { Story } from '../wml/story.js';
import { paragraphText } from '../wml/text.js';
import { paragraphOf, targetParagraphs } from './stories.js';
import type { FindOptions, TextMatch } from './types.js';

/** A match in one string. */
export interface RawMatch {
  readonly start: number;
  readonly end: number;
  readonly text: string;
  readonly groups: readonly (string | undefined)[];
}

const WORD_CHAR = /[\p{L}\p{N}\p{M}_]/u;

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** The global RegExp for a query (see the file comment). */
export function compileQuery(q: string | RegExp, o: FindOptions): RegExp {
  if (typeof q === 'string') {
    if (q === '') throw new UsageError('Cannot search for an empty string');
    return new RegExp(escapeRegExp(q), o.caseSensitive === false ? 'giu' : 'gu');
  }
  let flags = q.flags.replace(/[gy]/g, '');
  if (o.caseSensitive === false && !flags.includes('i')) flags += 'i';
  else if (o.caseSensitive === true) flags = flags.replace('i', '');
  return new RegExp(q.source, `${flags}g`);
}

/** Index after the code point at `i` (or the code unit, outside unicode mode). */
function advance(text: string, i: number, unicode: boolean): number {
  if (unicode) {
    const c = text.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff && i + 1 < text.length) {
      const d = text.charCodeAt(i + 1);
      if (d >= 0xdc00 && d <= 0xdfff) return i + 2;
    }
  }
  return i + 1;
}

function wordBefore(text: string, i: number): boolean {
  if (i <= 0) return false;
  let from = i - 1;
  const c = text.charCodeAt(from);
  if (c >= 0xdc00 && c <= 0xdfff && from > 0) {
    const h = text.charCodeAt(from - 1);
    if (h >= 0xd800 && h <= 0xdbff) from--;
  }
  return WORD_CHAR.test(text.slice(from, i));
}

function wordAfter(text: string, i: number): boolean {
  if (i >= text.length) return false;
  return WORD_CHAR.test(String.fromCodePoint(text.codePointAt(i)!));
}

/** Every non-empty match of global `re` in `text` (see the file comment for `wholeWord`). */
export function matchText(text: string, re: RegExp, wholeWord: boolean): RawMatch[] {
  const out: RawMatch[] = [];
  const unicode = re.flags.includes('u') || re.flags.includes('v');
  re.lastIndex = 0;
  for (;;) {
    const m = re.exec(text);
    if (m === null) break;
    const start = m.index;
    const end = start + m[0].length;
    if (end === start) {
      re.lastIndex = advance(text, start, unicode);
      continue;
    }
    if (wholeWord && (wordBefore(text, start) || wordAfter(text, end))) {
      re.lastIndex = advance(text, start, unicode);
      continue;
    }
    out.push({ start, end, text: m[0], groups: m.slice(1) });
  }
  return out;
}

/**
 * Finds `q` in the visible text of `target` (a document, story or paragraph), across run,
 * hyperlink, content-control and field-result boundaries, in document order. Matches never cross
 * paragraphs. Stories: `o.in` (default `['body']`; `'all'` adds headers, footers, footnotes,
 * endnotes, comments and text boxes).
 */
export function findText(
  target: WordDocument | Story | Paragraph,
  q: string | RegExp,
  o: FindOptions = {},
): TextMatch[] {
  const re = compileQuery(q, o);
  const out: TextMatch[] = [];
  for (const t of targetParagraphs(target, o.in)) {
    const text = paragraphText(t.el);
    if (text === '') continue;
    const found = matchText(text, re, o.wholeWord === true);
    if (found.length === 0) continue;
    const paragraph = paragraphOf(t);
    for (const m of found) {
      out.push({
        story: t.story,
        paragraph,
        start: m.start,
        end: m.end,
        text: m.text,
        groups: m.groups,
      });
    }
  }
  return out;
}
