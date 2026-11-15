// What field evaluators need to know about the document, computed lazily once per update:
// styles (names, basedOn chains, outline levels), bookmarks and the text they enclose,
// document variables, the settings part, section boundaries and note numbering.
//
// Bookmark names are case-insensitive, as in Word. Bookmark text follows the same visibility
// rules as the field index (instructions and deleted text excluded), with '\r' for paragraph
// marks; a text box is its own flow, so a bookmark outside one never collects its text.

import type { Part } from '../opc/part.js';
import { listLabel } from '../resolve/index.js';
import type { WordDocument } from '../wml/document.js';
import { Paragraph } from '../wml/paragraph.js';
import type { Story } from '../wml/story.js';
import { blockParagraphs, chooseAlternate } from '../wml/text.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { attached } from './scan.js';

/** One style of styles.xml. */
export interface StyleInfo {
  id: string;
  type: string;
  name: string;
  basedOn: string | undefined;
  /** `w:pPr/w:outlineLvl` (0-9) when the style sets it directly. */
  outline: number | undefined;
  isDefault: boolean;
}

/** One bookmark. */
export interface BookmarkInfo {
  name: string;
  id: string;
  start: XElement;
  story: Story;
  /** Text enclosed (paragraph marks '\r'); computed on demand. */
  text?: string;
  /** First footnote/endnote reference inside the bookmark. */
  noteRef?: XElement;
}

/** Visible text leaves, as in the field index. */
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
    default:
      return undefined;
  }
}

const HIDDEN: ReadonlySet<string> = new Set([
  'del',
  'moveFrom',
  'delText',
  'instrText',
  'delInstrText',
  'fldData',
]);

/** Lazily computed document facts for one update run. */
export class DocInfo {
  readonly doc: WordDocument;
  #styles: Map<string, StyleInfo> | undefined;
  #byName: Map<string, string> | undefined;
  #defaultPara: string | undefined;
  #bookmarks: Map<string, BookmarkInfo> | undefined;
  #order = new Map<Story, { version: number; map: Map<XElement, number> }>();

  constructor(doc: WordDocument) {
    this.doc = doc;
  }

  /** The styles part, if any. */
  stylesPart(): Part | undefined {
    return this.doc.mainPart.relatedByType('styles')[0];
  }

  /** The settings root element (`w:settings`), if there is a settings part. */
  settings(): XElement | undefined {
    const part = this.doc.mainPart.relatedByType('settings')[0];
    if (part === undefined) return undefined;
    for (const k of part.xml().children) if (k.nodeType === 1) return k;
    return undefined;
  }

  /** Forgets cached styles (after adding styles). */
  invalidateStyles(): void {
    this.#styles = undefined;
    this.#byName = undefined;
  }

  styles(): Map<string, StyleInfo> {
    if (this.#styles) return this.#styles;
    const map = new Map<string, StyleInfo>();
    const byName = new Map<string, string>();
    const part = this.stylesPart();
    let root: XElement | undefined;
    if (part) for (const k of part.xml().children) if (k.nodeType === 1) root = k;
    for (const s of root?.elements(NS.w, 'style') ?? []) {
      const id = s.attrNS(NS.w, 'styleId');
      if (id === undefined) continue;
      const type = s.attrNS(NS.w, 'type') ?? 'paragraph';
      const name = s.element(NS.w, 'name')?.attrNS(NS.w, 'val') ?? id;
      const lvl = s.element(NS.w, 'pPr')?.element(NS.w, 'outlineLvl')?.attrNS(NS.w, 'val');
      const n = lvl === undefined ? Number.NaN : Number.parseInt(lvl, 10);
      const info: StyleInfo = {
        id,
        type,
        name,
        basedOn: s.element(NS.w, 'basedOn')?.attrNS(NS.w, 'val'),
        outline: Number.isFinite(n) ? n : undefined,
        isDefault: /^(1|true|on)$/.test(s.attrNS(NS.w, 'default') ?? ''),
      };
      map.set(id, info);
      const key = `${type}|${name.toLowerCase()}`;
      if (!byName.has(key)) byName.set(key, id);
      if (info.isDefault && type === 'paragraph') this.#defaultPara = id;
    }
    this.#styles = map;
    this.#byName = byName;
    return map;
  }

  /** Style id for a style name (case-insensitive) or id, of `type` (default paragraph). */
  styleId(nameOrId: string, type = 'paragraph'): string | undefined {
    const styles = this.styles();
    const id = this.#byName?.get(`${type}|${nameOrId.toLowerCase()}`);
    if (id !== undefined) return id;
    const direct = styles.get(nameOrId);
    return direct?.type === type ? direct.id : undefined;
  }

  /** Default paragraph style id. */
  defaultParagraphStyle(): string | undefined {
    this.styles();
    return this.#defaultPara;
  }

  /** Paragraph style of `p` (direct `w:pStyle`, else the default paragraph style). */
  paragraphStyle(p: XElement): string | undefined {
    return (
      p.element(NS.w, 'pPr')?.element(NS.w, 'pStyle')?.attrNS(NS.w, 'val') ??
      this.defaultParagraphStyle()
    );
  }

  /**
   * Outline level (0-8; 9 = body text) a style gives, following basedOn; built-in "heading N"
   * styles have level N-1 even when styles.xml does not say so.
   */
  styleOutline(styleId: string | undefined): number | undefined {
    const styles = this.styles();
    let id = styleId;
    for (let hops = 0; id !== undefined && hops < 32; hops++) {
      const s = styles.get(id);
      if (s === undefined) break;
      if (s.outline !== undefined) return s.outline;
      const m = /^heading ([1-9])$/i.exec(s.name);
      if (m) return Number(m[1]) - 1;
      id = s.basedOn;
    }
    return undefined;
  }

  /** Direct outline level (`w:pPr/w:outlineLvl`). */
  directOutline(p: XElement): number | undefined {
    const v = p.element(NS.w, 'pPr')?.element(NS.w, 'outlineLvl')?.attrNS(NS.w, 'val');
    if (v === undefined) return undefined;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n : undefined;
  }

  /** Heading level 1-9 of `p` (undefined for body text); `direct` honours `w:outlineLvl` on p. */
  headingLevel(p: XElement, direct: boolean): number | undefined {
    const lvl =
      (direct ? this.directOutline(p) : undefined) ?? this.styleOutline(this.paragraphStyle(p));
    return lvl !== undefined && lvl >= 0 && lvl <= 8 ? lvl + 1 : undefined;
  }

  /** Document-order index of each paragraph of `story` (recomputed after edits). */
  paragraphOrder(story: Story): Map<XElement, number> {
    const version = story.part.xml().version;
    const hit = this.#order.get(story);
    if (hit !== undefined && hit.version === version) return hit.map;
    const map = new Map<XElement, number>();
    for (const p of blockParagraphs(story.element)) map.set(p, map.size);
    this.#order.set(story, { version, map });
    return map;
  }

  /** The list label of a paragraph ("1.2.a"), when the resolver is available. */
  listLabel(p: XElement, story: Story): string | undefined {
    try {
      return listLabel(new Paragraph(p, story));
    } catch {
      return undefined;
    }
  }

  /** Every bookmark of every story, by lower-case name. */
  bookmarks(): Map<string, BookmarkInfo> {
    if (this.#bookmarks) return this.#bookmarks;
    const map = new Map<string, BookmarkInfo>();
    for (const story of this.doc.stories()) {
      for (const b of story.element.descendants(NS.w, 'bookmarkStart')) {
        const name = b.attrNS(NS.w, 'name');
        const id = b.attrNS(NS.w, 'id');
        if (name === undefined || id === undefined) continue;
        const key = name.toLowerCase();
        if (!map.has(key)) map.set(key, { name, id, start: b, story });
      }
    }
    this.#bookmarks = map;
    return map;
  }

  /** Forgets the bookmark cache (after creating bookmarks). */
  invalidateBookmarks(): void {
    this.#bookmarks = undefined;
  }

  /**
   * A bookmark by name, with the text it encloses as the document reads now (earlier updates
   * included).
   */
  bookmark(name: string): BookmarkInfo | undefined {
    const key = name.toLowerCase();
    let b = this.bookmarks().get(key);
    if (b !== undefined && !attached(b.start, b.story.element)) {
      this.invalidateBookmarks();
      b = this.bookmarks().get(key);
    }
    if (b === undefined) return undefined;
    const { text, noteRef } = enclosed(b);
    b.text = text;
    if (noteRef) b.noteRef = noteRef;
    else delete b.noteRef;
    return b;
  }

  /** `w:docVars/w:docVar` value by name (case-insensitive). */
  docVar(name: string): string | undefined {
    const lower = name.toLowerCase();
    for (const v of this.settings()?.element(NS.w, 'docVars')?.elements(NS.w, 'docVar') ?? []) {
      if ((v.attrNS(NS.w, 'name') ?? '').toLowerCase() === lower) return v.attrNS(NS.w, 'val');
    }
    return undefined;
  }

  /** `w:listSeparator` / `w:decimalSymbol` from settings. */
  separatorsFromSettings(): { list?: string; decimal?: string } {
    const s = this.settings();
    const out: { list?: string; decimal?: string } = {};
    const list = s?.element(NS.w, 'listSeparator')?.attrNS(NS.w, 'val');
    const decimal = s?.element(NS.w, 'decimalSymbol')?.attrNS(NS.w, 'val');
    if (list !== undefined && list !== '') out.list = list;
    if (decimal !== undefined && decimal !== '') out.decimal = decimal;
    return out;
  }

  /** An element of docProps/app.xml by local name (Manager, HyperlinkBase, ...). */
  appValue(local: string): string | undefined {
    for (const rel of this.doc.package.rels.byType('extended-properties')) {
      const part = rel.partName === undefined ? undefined : this.doc.package.part(rel.partName);
      if (part === undefined) continue;
      for (const k of part.xml().children) {
        if (k.nodeType !== 1) continue;
        for (const e of k.elements()) if (e.local === local) return e.text;
      }
    }
    return undefined;
  }

  /** Body `w:sectPr` elements in document order (paragraph-level ones, then the final one). */
  sectionBreaks(): XElement[] {
    return this.doc.body.element
      .descendants(NS.w, 'sectPr')
      .filter((s) => (s.parent as XElement | null)?.local !== 'sectPrChange');
  }

  /** The `w:sectPr` governing body paragraph `p` (undefined when there is none). */
  sectionOf(p: XElement): { index: number; sectPr: XElement | undefined } {
    const order = this.paragraphOrder(this.doc.body);
    const at = order.get(p) ?? 0;
    let index = 0;
    for (const s of this.sectionBreaks()) {
      const owner = s.parent?.nodeType === 1 ? s.parent.parent : null;
      const ownerP = owner?.nodeType === 1 && owner.is(NS.w, 'p') ? owner : undefined;
      if (ownerP === undefined) return { index, sectPr: s };
      const pos = order.get(ownerP);
      if (pos !== undefined && pos >= at) return { index, sectPr: s };
      index++;
    }
    return { index, sectPr: undefined };
  }

  /** Width between the margins (twips) of the section governing `p`, minus Word's 10 twips. */
  textWidth(p: XElement): number {
    const s = this.sectionOf(p).sectPr;
    const w = Number(s?.element(NS.w, 'pgSz')?.attrNS(NS.w, 'w'));
    const mar = s?.element(NS.w, 'pgMar');
    const left = Number(mar?.attrNS(NS.w, 'left') ?? 0);
    const right = Number(mar?.attrNS(NS.w, 'right') ?? 0);
    const width = w - left - right - 10;
    return Number.isFinite(width) && width > 720 ? width : 9350;
  }
}

/** `w:` containers a bookmark's text never comes from (unless it starts inside one). */
const OPAQUE: ReadonlySet<string> = new Set(['drawing', 'pict', 'object', 'txbxContent']);

/** Upper bound on the elements one bookmark walk visits. */
const MAX_WALK = 1_000_000;

/**
 * Text between a bookmark's start and its end (document order, from the start), with '\r' for
 * paragraph marks, and the first note reference on the way. Stops at the end marker, at the
 * end of the story, or when leaving the text box the bookmark starts in.
 */
function enclosed(b: BookmarkInfo): { text: string; noteRef?: XElement } {
  const root = b.story.element;
  let text = '';
  let noteRef: XElement | undefined;
  let n: XElement = b.start;
  let descend = false;
  for (let steps = 0; steps < MAX_WALK; steps++) {
    let next: XElement | undefined = descend ? n.firstElement() : undefined;
    while (next === undefined) {
      const up = n.parent;
      const inBranch = up !== null && up.nodeType === 1 && up.is(NS.mc, 'AlternateContent');
      const sib = inBranch ? undefined : n.nextElement();
      if (sib) {
        next = sib;
        break;
      }
      const p = n.parent;
      if (p === null || p.nodeType !== 1 || p === root)
        return { text, ...(noteRef ? { noteRef } : {}) };
      if (p.is(NS.w, 'p')) text += '\r';
      if (p.is(NS.w, 'txbxContent')) return { text, ...(noteRef ? { noteRef } : {}) };
      n = p;
    }
    n = next;
    descend = false;
    if (n.uri === NS.w) {
      const local = n.local;
      if (local === 'bookmarkEnd' && n.attrNS(NS.w, 'id') === b.id) break;
      if (HIDDEN.has(local) || local.endsWith('Pr') || OPAQUE.has(local)) continue;
      const t = leafText(n);
      if (t !== undefined) {
        text += t;
        continue;
      }
      if (local === 'footnoteReference' || local === 'endnoteReference') {
        noteRef ??= n;
        continue;
      }
      descend = true;
    } else if (n.uri === NS.mc && n.local === 'AlternateContent') {
      // Only the branch readers show; climbing out of it leaves the AlternateContent.
      const branch = chooseAlternate(n);
      if (branch) {
        n = branch;
        descend = true;
      }
    } else {
      descend = true;
    }
  }
  return { text, ...(noteRef ? { noteRef } : {}) };
}
