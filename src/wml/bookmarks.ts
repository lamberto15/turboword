// Bookmarks: named ranges marked by a `w:bookmarkStart` / `w:bookmarkEnd` pair sharing a `w:id`.
//
// Names are unique document-wide (Word compares them case-insensitively), at most 40
// characters and without spaces; ids are document-wide numbers handed out by DocumentIds. A
// bookmark may span paragraphs, but both ends live in the same story. Names starting with `_`
// are hidden ones Word creates itself (`_GoBack`, `_Toc123`, `_Ref456`).
//
// Renaming also rewrites what points at the old name: internal hyperlinks (`w:anchor`) and the
// REF / PAGEREF / NOTEREF / `HYPERLINK \l` field instructions, so cross-references keep working.
// An instruction split over several `w:instrText` elements is only rewritten where the name
// sits in one piece.

import { UsageError } from '../core/errors.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { WordDocument } from './document.js';
import { DocumentIds } from './ids.js';
import { setWAttr, wElement } from './order.js';
import { Paragraph } from './paragraph.js';
import type { Story } from './story.js';

/** Longest bookmark name Word accepts. */
export const MAX_BOOKMARK_NAME = 40;

/** A bookmark: its start marker and the story holding it. */
export class Bookmark {
  /** The `w:bookmarkStart` element. */
  readonly start: XElement;
  /** The story containing the bookmark. */
  readonly story: Story;

  constructor(start: XElement, story: Story) {
    this.start = start;
    this.story = story;
  }

  /** `w:name`. */
  get name(): string {
    return this.start.attrNS(NS.w, 'name') ?? '';
  }

  /** `w:id` (shared by the start and end markers). */
  get id(): string {
    return this.start.attrNS(NS.w, 'id') ?? '';
  }

  /** True for names starting with `_` (hidden bookmarks Word manages: `_GoBack`, `_Toc…`). */
  get hidden(): boolean {
    return this.name.startsWith('_');
  }

  /** The matching `w:bookmarkEnd`, if present. */
  get end(): XElement | undefined {
    return findEnd(this.story, this.id);
  }

  /**
   * The paragraph the bookmark starts in (for a start marker between blocks, the next
   * paragraph), if any.
   */
  get paragraph(): Paragraph | undefined {
    const p = this.start.closest(NS.w, 'p');
    if (p !== undefined) return new Paragraph(p, this.story);
    for (let el = this.start.nextElement(); el !== undefined; el = el.nextElement()) {
      if (el.is(NS.w, 'p')) return new Paragraph(el, this.story);
      const inner = el.descendants(NS.w, 'p')[0];
      if (inner !== undefined) return new Paragraph(inner, this.story);
    }
    return undefined;
  }
}

/** A range of paragraphs to bookmark, `start` to `end` inclusive (same story). */
export interface BookmarkRange {
  start: Paragraph;
  end: Paragraph;
}

function findEnd(story: Story, id: string): XElement | undefined {
  const match = (el: XElement): boolean => el.attrNS(NS.w, 'id') === id;
  return (
    story.element.descendants(NS.w, 'bookmarkEnd').find(match) ??
    story.part.xml().root.descendants(NS.w, 'bookmarkEnd').find(match)
  );
}

/** Throws UsageError unless `name` is a legal bookmark name. */
function checkName(name: string): void {
  if (typeof name !== 'string' || name.length === 0) {
    throw new UsageError('Bookmark name must be a non-empty string');
  }
  if (name.length > MAX_BOOKMARK_NAME) {
    throw new UsageError(`Bookmark name "${name}" is longer than ${MAX_BOOKMARK_NAME} characters`);
  }
  if (/\s/.test(name)) throw new UsageError(`Bookmark name "${name}" contains whitespace`);
}

/** Root-to-node chain of `el`. */
function chain(el: XElement): XElement[] {
  const out: XElement[] = [];
  for (let e: XElement | null = el; e !== null; e = e.parent?.nodeType === 1 ? e.parent : null) {
    out.push(e);
  }
  return out.reverse();
}

/** Negative when `a` comes before `b` in document order (an ancestor comes first), 0 if same. */
function compareOrder(a: XElement, b: XElement): number {
  if (a === b) return 0;
  const ca = chain(a);
  const cb = chain(b);
  let i = 0;
  while (i < ca.length && i < cb.length && ca[i] === cb[i]) i++;
  if (i === ca.length) return -1;
  if (i === cb.length) return 1;
  return ca[i]!.index() - cb[i]!.index();
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** The bookmarks of a document (`doc.bookmarks`). A live view: every call reads the DOM. */
export class Bookmarks {
  readonly #doc: WordDocument;

  constructor(doc: WordDocument) {
    this.#doc = doc;
  }

  /**
   * Every bookmark in every story (body, headers/footers, notes, comments), in story order and
   * document order within a story. Hidden bookmarks are included (see {@link Bookmark.hidden}).
   */
  list(): Bookmark[] {
    const out: Bookmark[] = [];
    for (const story of this.#doc.stories()) {
      for (const el of story.element.descendants(NS.w, 'bookmarkStart')) {
        const fb = el.closest(NS.mc, 'Fallback');
        if (fb?.parent?.nodeType === 1 && fb.parent.element(NS.mc, 'Choice') !== undefined) {
          continue; // the Choice branch has the same bookmark
        }
        out.push(new Bookmark(el, story));
      }
    }
    return out;
  }

  /** The bookmark named `name` (case-insensitive, as in Word), if any. */
  get(name: string): Bookmark | undefined {
    const key = name.toLowerCase();
    return this.list().find((b) => b.name.toLowerCase() === key);
  }

  /** True when a bookmark named `name` exists (case-insensitive). */
  has(name: string): boolean {
    return this.get(name) !== undefined;
  }

  /**
   * Adds bookmark `name` around a whole paragraph, or around the paragraphs `start` to `end`
   * of one story. The start marker goes first in the start paragraph (after its properties),
   * the end marker last in the end paragraph. Throws UsageError for an invalid or duplicate
   * name, or a range whose ends are in different stories or out of order.
   */
  add(name: string, target: Paragraph | BookmarkRange): Bookmark {
    checkName(name);
    if (this.has(name)) throw new UsageError(`Bookmark "${name}" already exists`);
    const first = target instanceof Paragraph ? target : target.start;
    const last = target instanceof Paragraph ? target : target.end;
    const story = first.story;
    if (
      last.story.element !== story.element ||
      !chain(first.element).includes(story.element) ||
      !chain(last.element).includes(story.element)
    ) {
      throw new UsageError('Bookmark ends must be paragraphs of the same story');
    }
    if (compareOrder(first.element, last.element) > 0) {
      throw new UsageError('Bookmark end paragraph comes before its start paragraph');
    }
    const id = String(DocumentIds.of(this.#doc).next('bookmark'));
    const start = wElement(first.element, 'bookmarkStart', { id, name });
    const end = wElement(last.element, 'bookmarkEnd', { id });
    const pPr = first.element.element(NS.w, 'pPr');
    if (pPr !== undefined) pPr.after(start);
    else first.element.prepend(start);
    last.element.append(end);
    return new Bookmark(start, story);
  }

  /** Removes bookmark `name` (both markers; the content stays). Returns whether it existed. */
  remove(name: string): boolean {
    const bm = this.get(name);
    if (bm === undefined) return false;
    bm.end?.remove();
    bm.start.remove();
    return true;
  }

  /**
   * Renames bookmark `from` to `to` and, unless `updateReferences` is false, rewrites internal
   * hyperlinks and REF / PAGEREF / NOTEREF / `HYPERLINK \l` fields that point at it. Throws
   * UsageError when `from` does not exist or `to` is invalid or taken.
   */
  rename(from: string, to: string, o?: { updateReferences?: boolean }): Bookmark {
    const bm = this.get(from);
    if (bm === undefined) throw new UsageError(`No bookmark named "${from}"`);
    checkName(to);
    const clash = this.get(to);
    if (clash !== undefined && clash.start !== bm.start) {
      throw new UsageError(`Bookmark "${to}" already exists`);
    }
    const old = bm.name;
    setWAttr(bm.start, 'name', to);
    if (o?.updateReferences !== false) this.#retarget(old, to);
    return bm;
  }

  [Symbol.iterator](): Iterator<Bookmark> {
    return this.list()[Symbol.iterator]();
  }

  /** Points hyperlinks and field instructions naming bookmark `from` at `to`. */
  #retarget(from: string, to: string): void {
    const name = escapeRegExp(from);
    const tail = '(?="?(?:\\s|\\\\|$))';
    const ref = new RegExp(`(\\b(?:REF|PAGEREF|NOTEREF)\\s+"?)${name}${tail}`, 'gi');
    const link = new RegExp(`(\\\\l\\s+"?)${name}${tail}`, 'gi');
    const rewrite = (instr: string): string => {
      let out = instr.replace(ref, `$1${to}`);
      if (/^\s*HYPERLINK\b/i.test(out)) out = out.replace(link, `$1${to}`);
      return out;
    };
    const key = from.toLowerCase();
    for (const story of this.#doc.stories()) {
      for (const el of story.element.descendants()) {
        if (el.uri !== NS.w) continue;
        if (el.local === 'hyperlink') {
          const a = el.attrNS(NS.w, 'anchor');
          if (a?.toLowerCase() === key && el.attrNS(NS.r, 'id') === undefined) {
            setWAttr(el, 'anchor', to);
          }
        } else if (el.local === 'instrText') {
          const t = el.text;
          const n = rewrite(t);
          if (n !== t) el.text = n;
        } else if (el.local === 'fldSimple') {
          const t = el.attrNS(NS.w, 'instr');
          if (t === undefined) continue;
          const n = rewrite(t);
          if (n !== t) setWAttr(el, 'instr', n);
        }
      }
    }
  }
}
