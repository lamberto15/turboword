// Story: one independent flow of block content.
//
// Word keeps text in several separate flows: the main body, each header and footer part, each
// footnote, endnote and comment (text boxes too, later). They share the same block grammar
// (paragraphs, tables, content controls, ...), so one view type serves all of them; `kind` and
// `id` say which one it is. Like every view it holds only its container element.

import { UsageError } from '../core/errors.js';
import type { Part } from '../opc/part.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { WordDocument } from './document.js';
import { wElement } from './order.js';
import { Paragraph } from './paragraph.js';
import { newTable, Table } from './table.js';
import { blockParagraphs, chooseAlternate, paragraphText } from './text.js';

/** Which flow of the document a {@link Story} is. */
export type StoryKind = 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'comment';

/**
 * A block-level item of a story: a paragraph, a table, or any other block-level element as is
 * (`w:sdt`, `w:customXml`, `w:bookmarkStart`, `mc:AlternateContent`, ...).
 */
export type Block = Paragraph | Table | XElement;

/** One flow of block content: the body, a header/footer, a footnote/endnote or a comment. */
export class Story {
  /** The document the story belongs to. */
  readonly doc: WordDocument;
  readonly kind: StoryKind;
  /** The part holding the story (`word/document.xml`, `word/header1.xml`, ...). */
  readonly part: Part;
  /** The container element: `w:body`, `w:hdr`, `w:ftr`, `w:footnote`, `w:endnote`, `w:comment`. */
  readonly element: XElement;
  /** `w:id` of a footnote, endnote or comment; undefined for other stories. */
  readonly id: string | undefined;

  constructor(doc: WordDocument, kind: StoryKind, part: Part, element: XElement, id?: string) {
    this.doc = doc;
    this.kind = kind;
    this.part = part;
    this.element = element;
    this.id = id;
  }

  /**
   * The story's direct block-level children in order. The body's final `w:sectPr` (section
   * properties, not content) is left out; non-element nodes are skipped.
   */
  blocks(): Block[] {
    const out: Block[] = [];
    for (const k of this.element.children) {
      if (k.nodeType !== 1) continue;
      if (k.uri === NS.w) {
        if (k.local === 'p') {
          out.push(new Paragraph(k, this));
          continue;
        }
        if (k.local === 'tbl') {
          out.push(new Table(k, this));
          continue;
        }
        if (k.local === 'sectPr') continue;
      }
      out.push(k);
    }
    return out;
  }

  /**
   * Every paragraph of the story, deep and in document order: paragraphs inside tables,
   * content controls and custom XML included; text-box paragraphs excluded.
   */
  paragraphs(): Paragraph[] {
    return blockParagraphs(this.element).map((p) => new Paragraph(p, this));
  }

  /**
   * The story's top-level tables in order: direct children plus tables wrapped in block-level
   * content controls, custom XML or `mc:AlternateContent`. Nested tables (inside cells) are
   * reached through their parent table.
   */
  tables(): Table[] {
    const out: Table[] = [];
    const stack: XElement[] = [];
    pushChildren(stack, this.element);
    while (stack.length > 0) {
      const el = stack.pop()!;
      if (el.uri === NS.w) {
        if (el.local === 'tbl') {
          out.push(new Table(el, this));
          continue;
        }
        if (el.local !== 'sdt' && el.local !== 'sdtContent' && el.local !== 'customXml') continue;
      } else if (el.uri === NS.mc && el.local === 'AlternateContent') {
        const branch = chooseAlternate(el);
        if (branch) pushChildren(stack, branch);
        continue;
      } else {
        continue;
      }
      pushChildren(stack, el);
    }
    return out;
  }

  /** Visible text: the text of {@link paragraphs} joined with LF. */
  text(): string {
    return blockParagraphs(this.element).map(paragraphText).join('\n');
  }

  /**
   * Appends a paragraph at the end of the story (in the body: just before the final
   * `w:sectPr`, which must stay the last child), optionally with one run of `text` (see
   * {@link Paragraph.appendRun} for TAB/LF handling) and a paragraph style id.
   */
  appendParagraph(text?: string, o?: { style?: string }): Paragraph {
    const p = wElement(this.element, 'p');
    this.#appendBlock(p);
    return this.#fill(p, text, o);
  }

  /**
   * Inserts a paragraph at position `index` of {@link blocks} (0 = first; `blocks().length` =
   * at the end, like {@link appendParagraph}) and returns it. Throws RangeError for an index
   * out of range.
   */
  insertParagraph(index: number, text?: string, o?: { style?: string }): Paragraph {
    const blocks = this.#blockElements();
    if (!Number.isInteger(index) || index < 0 || index > blocks.length) {
      throw new RangeError(`insertParagraph(): index ${index} is outside 0..${blocks.length}`);
    }
    const p = wElement(this.element, 'p');
    const ref = blocks[index];
    if (ref === undefined) this.#appendBlock(p);
    else this.element.insertBefore(p, ref);
    return this.#fill(p, text, o);
  }

  /**
   * Appends a `rows` x `cols` table (see `newTable` in table.ts for the options: column widths
   * in twips, total width, table style) and returns it.
   */
  appendTable(
    rows: number,
    cols: number,
    o?: { widths?: readonly number[]; width?: number; style?: string },
  ): Table {
    const tbl = newTable(this, rows, cols, o);
    this.#appendBlock(tbl);
    return new Table(tbl, this);
  }

  /** Appends a paragraph holding a page break (`<w:br w:type="page"/>`) and returns it. */
  appendPageBreak(): Paragraph {
    const p = wElement(this.element, 'p');
    this.#appendBlock(p);
    const r = wElement(p, 'r');
    p.append(r);
    r.append(wElement(r, 'br', { type: 'page' }));
    return new Paragraph(p, this);
  }

  /**
   * Removes all content. The body keeps its final `w:sectPr` (the last section's page setup,
   * headers and footers); section breaks inside the content go with it. Headers, footers and
   * notes keep one empty paragraph, which Word requires.
   */
  clear(): void {
    const container = this.element;
    const last = container.lastElement();
    const keep = this.kind === 'body' && last?.is(NS.w, 'sectPr') ? last : undefined;
    for (const k of container.children.slice()) if (k !== keep) k.remove();
    const k = this.kind;
    if (k === 'header' || k === 'footer' || k === 'footnote' || k === 'endnote') {
      container.append(wElement(container, 'p'));
    }
  }

  /** Direct block-level child elements (what {@link blocks} wraps). */
  #blockElements(): XElement[] {
    const out: XElement[] = [];
    for (const k of this.element.children) {
      if (k.nodeType !== 1 || k.is(NS.w, 'sectPr')) continue;
      out.push(k);
    }
    return out;
  }

  /** Appends a block at the end (in the body: before the final `w:sectPr`, which stays last). */
  #appendBlock(el: XElement): void {
    const container = this.element;
    if (el.parent !== null) throw new UsageError('Block element is already attached');
    const last = container.lastElement();
    if (last?.is(NS.w, 'sectPr')) container.insertBefore(el, last);
    else container.append(el);
  }

  #fill(p: XElement, text: string | undefined, o: { style?: string } | undefined): Paragraph {
    const para = new Paragraph(p, this);
    if (o?.style !== undefined) para.style = o.style;
    if (text !== undefined && text !== '') para.appendRun(text);
    return para;
  }
}

function pushChildren(stack: XElement[], el: XElement): void {
  const kids = el.children;
  for (let i = kids.length - 1; i >= 0; i--) {
    const k = kids[i]!;
    if (k.nodeType === 1) stack.push(k);
  }
}
