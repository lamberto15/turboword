// importBookmarks: unique bookmark names and ids for the imported content, references rewritten.
//
// Bookmark names are document-wide (and compared case-insensitively by Word), ids must be unique
// too. Every imported bookmark gets a fresh id; one whose name the target already uses is
// renamed `name_1`, `name_2`, ... (Word's 40-character limit kept) or, with
// `bookmarks: 'drop'`, removed. Renames are then applied to everything that refers to a bookmark
// by name inside the imported content:
//   - field instructions REF / PAGEREF / NOTEREF / `HYPERLINK \l` (and the bare `{ name }` REF
//     shorthand), both in `w:fldSimple/@w:instr` and in complex fields whose instruction is
//     split over several `w:instrText` runs (possibly over several paragraphs: a TOC);
//   - `w:hyperlink/@w:anchor`.
// The source's `_GoBack` (Word's "last edit position") is always dropped: two of them would
// conflict and it means nothing in the result.

import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { allElements, importedExtras, removeElement } from './parts-util.js';
import type { ImportContext } from './types.js';

/** Word truncates bookmark names longer than this. */
const MAX_NAME = 40;

/**
 * Renames conflicting bookmarks of `blocks` (and of note/comment content imported before), gives
 * them fresh ids, drops `_GoBack`, rewrites references and fills `ctx.bookmarkRenames`.
 */
export function importBookmarks(ctx: ImportContext, blocks: XElement[]): void {
  const extras = importedExtras(ctx);
  renameBookmarks(ctx, [...blocks, ...extras], blocks);
  if (ctx.bookmarkRenames.size === 0) return;
  rewriteReferences(blocks, ctx.bookmarkRenames);
  for (const e of extras) rewriteReferences([e], ctx.bookmarkRenames);
}

/**
 * The bookmark pass over imported `roots` (which are already in the target): fresh ids,
 * conflicting names renamed (or dropped), `_GoBack` and unbalanced markers removed. Names used
 * anywhere else in the target count as taken. `blocks` is kept in sync when a removed marker is
 * one of them. Exported for the sections step (header/footer content).
 */
export function renameBookmarks(
  ctx: ImportContext,
  roots: readonly XElement[],
  blocks: XElement[] = [],
): void {
  const starts: XElement[] = [];
  const ends: XElement[] = [];
  for (const el of allElements(roots)) {
    if (el.uri !== NS.w) continue;
    if (el.local === 'bookmarkStart') starts.push(el);
    else if (el.local === 'bookmarkEnd') ends.push(el);
  }
  if (starts.length === 0 && ends.length === 0) return;

  const taken = targetBookmarkNames(ctx, new Set(roots));
  const idMap = new Map<string, string>();
  const dropped = new Set<string>();
  for (const start of starts) {
    const oldId = start.attrNS(NS.w, 'id') ?? '';
    const name = start.attrNS(NS.w, 'name');
    if (name !== undefined && name.toLowerCase() === '_goback') {
      dropped.add(oldId);
      removeElement(blocks, start);
      continue;
    }
    if (name !== undefined) {
      const key = name.toLowerCase();
      if (taken.has(key)) {
        if (ctx.options.bookmarks === 'drop') {
          dropped.add(oldId);
          removeElement(blocks, start);
          continue;
        }
        const renamed = freeName(name, taken);
        start.setAttrNS(NS.w, 'name', renamed);
        if (!ctx.bookmarkRenames.has(name)) ctx.bookmarkRenames.set(name, renamed);
        taken.add(renamed.toLowerCase());
      } else {
        taken.add(key);
      }
    }
    const fresh = String(ctx.ids.next('bookmark'));
    idMap.set(oldId, fresh);
    start.setAttrNS(NS.w, 'id', fresh);
  }

  const ended = new Set<string>();
  for (const end of ends) {
    const oldId = end.attrNS(NS.w, 'id') ?? '';
    const fresh = idMap.get(oldId);
    if (dropped.has(oldId) || fresh === undefined || ended.has(fresh)) {
      removeElement(blocks, end);
      continue;
    }
    ended.add(fresh);
    end.setAttrNS(NS.w, 'id', fresh);
  }
  // A start whose end was not imported (a partial selection) would be unbalanced.
  for (const start of starts) {
    if (start.parent === null) continue;
    const id = start.attrNS(NS.w, 'id');
    if (id !== undefined && !ended.has(id)) removeElement(blocks, start);
  }
}

/** First `name_N` (N = 1, 2, ...) not in `taken`, cut to Word's 40-character limit. */
function freeName(name: string, taken: ReadonlySet<string>): string {
  for (let n = 1; ; n++) {
    const suffix = `_${n}`;
    const candidate = name.slice(0, Math.max(1, MAX_NAME - suffix.length)) + suffix;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}

/** Lower-cased names of every bookmark in the target's stories, outside `exclude` subtrees. */
function targetBookmarkNames(ctx: ImportContext, exclude: ReadonlySet<XElement>): Set<string> {
  const out = new Set<string>();
  for (const story of ctx.target.stories()) {
    const stack: XElement[] = [story.element];
    while (stack.length > 0) {
      const el = stack.pop()!;
      if (exclude.has(el)) continue;
      if (el.is(NS.w, 'bookmarkStart')) {
        const name = el.attrNS(NS.w, 'name');
        if (name !== undefined) out.add(name.toLowerCase());
      }
      for (const k of el.children) if (k.nodeType === 1) stack.push(k);
    }
  }
  return out;
}

/**
 * Rewrites bookmark references in `roots` (one story's content, in document order) through
 * `renames` (source name → new name; matched case-insensitively): field instructions in
 * `w:fldSimple` and complex fields, and `w:hyperlink/@w:anchor`.
 */
export function rewriteReferences(
  roots: readonly XElement[],
  renames: ReadonlyMap<string, string>,
): void {
  if (renames.size === 0) return;
  const lookup = new Map<string, string>();
  for (const [from, to] of renames) lookup.set(from.toLowerCase(), to);

  // Complex fields: begin ... instrText* ... separate ... end, possibly nested.
  const fields: { instr: XElement[]; open: boolean }[] = [];
  const finish = (f: { instr: XElement[]; open: boolean }): void => {
    f.open = false;
    if (f.instr.length === 0) return;
    const text = f.instr.map((e) => e.text).join('');
    const rewritten = rewriteInstruction(text, lookup);
    if (rewritten === text) return;
    const [first, ...rest] = f.instr;
    first!.text = rewritten;
    first!.setAttr('xml:space', 'preserve');
    for (const e of rest) e.text = '';
  };
  for (const el of allElements(roots)) {
    if (el.uri !== NS.w) continue;
    switch (el.local) {
      case 'fldChar': {
        const type = el.attrNS(NS.w, 'fldCharType');
        if (type === 'begin') fields.push({ instr: [], open: true });
        else if (type === 'separate') {
          const top = fields.at(-1);
          if (top?.open === true) finish(top);
        } else if (type === 'end') {
          const top = fields.pop();
          if (top?.open === true) finish(top);
        }
        break;
      }
      case 'instrText': {
        const top = fields.at(-1);
        if (top?.open === true) top.instr.push(el);
        break;
      }
      case 'fldSimple': {
        const instr = el.attrNS(NS.w, 'instr');
        if (instr === undefined) break;
        const rewritten = rewriteInstruction(instr, lookup);
        if (rewritten !== instr) el.setAttrNS(NS.w, 'instr', rewritten);
        break;
      }
      case 'hyperlink': {
        const anchor = el.attrNS(NS.w, 'anchor');
        const renamed = anchor === undefined ? undefined : lookup.get(anchor.toLowerCase());
        if (renamed !== undefined) el.setAttrNS(NS.w, 'anchor', renamed);
        break;
      }
      default:
        break;
    }
  }
}

interface Token {
  readonly value: string;
  readonly start: number;
  readonly end: number;
  readonly quoted: boolean;
}

/** Field instruction tokens: whitespace-separated words and `"quoted strings"` (with offsets). */
function tokenize(s: string): Token[] {
  const out: Token[] = [];
  const space = (c: string | undefined): boolean => c !== undefined && /\s/.test(c);
  let i = 0;
  while (i < s.length) {
    while (i < s.length && space(s[i])) i++;
    if (i >= s.length) break;
    const start = i;
    if (s[i] === '"') {
      i++;
      let value = '';
      while (i < s.length && s[i] !== '"') {
        if (s[i] === '\\' && (s[i + 1] === '"' || s[i + 1] === '\\')) {
          value += s[i + 1];
          i += 2;
        } else value += s[i++];
      }
      i = Math.min(i + 1, s.length);
      out.push({ value, start, end: i, quoted: true });
    } else {
      while (i < s.length && !space(s[i])) i++;
      out.push({ value: s.slice(start, i), start, end: i, quoted: false });
    }
  }
  return out;
}

/** Field keywords that are never the bare `{ bookmark }` REF shorthand. */
const KEYWORDS: ReadonlySet<string> = new Set(
  (
    'ADDRESSBLOCK ADVANCE ASK AUTHOR AUTONUM AUTONUMLGL AUTONUMOUT AUTOTEXT AUTOTEXTLIST BARCODE ' +
    'BIBLIOGRAPHY BIDIOUTLINE CITATION COMMENTS COMPARE CREATEDATE DATABASE DATE DOCPROPERTY ' +
    'DOCVARIABLE EDITTIME EMBED EQ FILENAME FILESIZE FILLIN FORMCHECKBOX FORMDROPDOWN FORMTEXT ' +
    'GOTOBUTTON GREETINGLINE HYPERLINK IF INCLUDEPICTURE INCLUDETEXT INDEX INFO KEYWORDS ' +
    'LASTSAVEDBY LINK LISTNUM MACROBUTTON MERGEFIELD MERGEREC MERGESEQ NEXT NEXTIF NOTEREF ' +
    'NUMCHARS NUMPAGES NUMWORDS PAGE PAGEREF PRINT PRINTDATE PRIVATE QUOTE RD REF REVNUM SAVEDATE ' +
    'SECTION SECTIONPAGES SEQ SET SKIPIF STYLEREF SUBJECT SYMBOL TA TC TEMPLATE TIME TITLE TOA ' +
    'TOC USERADDRESS USERINITIALS USERNAME XE'
  ).split(' '),
);

/**
 * `instr` with bookmark names renamed: the argument of REF / PAGEREF / NOTEREF, the argument of
 * HYPERLINK's `\l` switch, and a bare bookmark name used as a field. Formatting (spacing,
 * quoting, switches) is kept.
 */
export function rewriteInstruction(instr: string, renames: ReadonlyMap<string, string>): string {
  const tokens = tokenize(instr);
  const first = tokens[0];
  if (first === undefined) return instr;
  const edits: { start: number; end: number; text: string }[] = [];
  const rename = (t: Token | undefined): void => {
    if (t === undefined || (!t.quoted && t.value.startsWith('\\'))) return;
    const to = renames.get(t.value.toLowerCase());
    if (to !== undefined)
      edits.push({ start: t.start, end: t.end, text: t.quoted ? `"${to}"` : to });
  };
  const keyword = first.quoted ? '' : first.value.toUpperCase();
  if (keyword === 'REF' || keyword === 'PAGEREF' || keyword === 'NOTEREF') {
    rename(tokens[1]);
  } else if (keyword === 'HYPERLINK') {
    for (let i = 1; i < tokens.length - 1; i++) {
      if (!tokens[i]!.quoted && tokens[i]!.value.toLowerCase() === '\\l') rename(tokens[i + 1]);
    }
  } else if (!first.quoted && !KEYWORDS.has(keyword)) {
    rename(first);
  }
  let out = instr;
  for (const e of edits.sort((a, b) => b.start - a.start)) {
    out = out.slice(0, e.start) + e.text + out.slice(e.end);
  }
  return out;
}
