// Paragraph and Run: thin views over `w:p` and `w:r` elements.
//
// A view holds nothing but its element; every read goes to the DOM and every write is a DOM
// mutation (which marks the part dirty), so views never go stale and anything they do not model
// survives untouched. New property children are placed in schema order (order.ts).
//
// Formatting getters report *direct* formatting only (what the paragraph's `w:pPr` or the run's
// own `w:rPr` says); the style cascade is the job of the resolver layer. Measurements are in
// points (the file stores twentieths of a point for paragraphs, half-points for font sizes).
//
// Structural edits keep the document valid: a table cell never loses its last paragraph, a
// section break (`w:pPr/w:sectPr`) is never silently lost, and copies get fresh drawing,
// content-control and revision ids so nothing is duplicated document-wide.

import { UsageError } from '../core/errors.js';
import { REL } from '../opc/rel-types.js';
import { type XElement, type XNode, XText } from '../xml/dom.js';
import { isStrictNs, NS } from '../xml/ns.js';
import { ensureHyperlinkStyle, Hyperlink } from './hyperlinks.js';
import { addImagePart, imageInfo, inlineDrawing, relTypeFor, renewIds } from './media.js';
import {
  type ChildOrder,
  ensureChild,
  ensureFirstChild,
  insertOrdered,
  PPR_ORDER,
  pruneEmpty,
  RPR_ORDER,
  setWAttr,
  wElement,
  wPrefix,
} from './order.js';
import type { Story } from './story.js';
import { chooseAlternate, paragraphText, walkVisible } from './text.js';

/** `w:u` values (ST_Underline). */
export type UnderlineStyle =
  | 'single'
  | 'words'
  | 'double'
  | 'thick'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashedHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dashDotHeavy'
  | 'dotDotDash'
  | 'dashDotDotHeavy'
  | 'wave'
  | 'wavyHeavy'
  | 'wavyDouble'
  | 'none';

/** `w:highlight` values (ST_HighlightColor). */
export type HighlightColor =
  | 'black'
  | 'blue'
  | 'cyan'
  | 'green'
  | 'magenta'
  | 'red'
  | 'yellow'
  | 'white'
  | 'darkBlue'
  | 'darkCyan'
  | 'darkGreen'
  | 'darkMagenta'
  | 'darkRed'
  | 'darkYellow'
  | 'darkGray'
  | 'lightGray'
  | 'none';

/** `w:jc` values for paragraphs (ST_Jc). Strict documents store left/right as start/end. */
export type ParagraphAlignment =
  | 'left'
  | 'center'
  | 'right'
  | 'both'
  | 'start'
  | 'end'
  | 'distribute'
  | 'thaiDistribute'
  | 'highKashida'
  | 'mediumKashida'
  | 'lowKashida'
  | 'numTab';

/**
 * Line spacing: `multiple` of single spacing (`value` 1.5 = one and a half lines), or a height
 * in points that is `exact` or a minimum (`atLeast`).
 */
export interface LineSpacing {
  value: number;
  rule: 'multiple' | 'exact' | 'atLeast';
}

/** Direct formatting accepted by {@link Paragraph.appendRun}. */
export interface RunProps {
  bold?: boolean;
  italic?: boolean;
  /** Character style id (`w:rStyle`). */
  style?: string;
  underline?: UnderlineStyle;
  /** `RRGGBB` hex (a leading `#` is accepted) or `auto`. */
  color?: string;
  /** Font size in points. */
  size?: number;
  /** Font name (ASCII and high-ANSI ranges). */
  font?: string;
  highlight?: HighlightColor;
}

/** Options of {@link Paragraph.appendImage}. */
export interface ImageOptions {
  /**
   * Displayed width in points. With only one of width/height the other follows the image's
   * aspect ratio; with neither, the natural size (pixels at the file's dpi, default 96).
   */
  width?: number;
  /** Displayed height in points. */
  height?: number;
  /** Drawing name (`wp:docPr/@name`; default `Picture <id>`). */
  name?: string;
  /** Alternative text (`wp:docPr/@descr`). */
  description?: string;
  /**
   * Raster fallback for an SVG image, shown by consumers without SVG support. Without it the
   * SVG part is referenced for both.
   */
  fallback?: Uint8Array;
}

/** Options of {@link Paragraph.appendHyperlink}: run formatting plus a tooltip. */
export interface HyperlinkProps extends RunProps {
  /** Screen tip (`w:tooltip`). */
  tooltip?: string;
}

/** `w:` containers whose runs are not part of the visible paragraph. */
const HIDDEN: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'del',
  'moveFrom',
  'drawing',
  'pict',
  'object',
]);

/** Block-level elements (what a cell, text box or story contains). */
const BLOCKS: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'p',
  'tbl',
  'sdt',
  'customXml',
  'altChunk',
]);

const EMU_PER_POINT = 12700;

/** A paragraph (`w:p`). */
export class Paragraph {
  /** The `w:p` element. */
  readonly element: XElement;
  /** The story the paragraph belongs to. */
  readonly story: Story;

  constructor(el: XElement, story: Story) {
    this.element = el;
    this.story = story;
  }

  /** Visible text (see text.ts for what counts). */
  get text(): string {
    return paragraphText(this.element);
  }

  /** Paragraph style id (`w:pPr/w:pStyle/@w:val`), undefined when none is set directly. */
  get style(): string | undefined {
    return this.element.element(NS.w, 'pPr')?.element(NS.w, 'pStyle')?.attrNS(NS.w, 'val');
  }

  /** Sets the paragraph style id; undefined removes `w:pStyle` (and an emptied `w:pPr`). */
  set style(v: string | undefined) {
    setValChild(this.element, 'pPr', 'pStyle', PPR_ORDER, v);
  }

  /**
   * The paragraph's visible runs in document order: direct children and runs inside
   * hyperlinks, insertions, smart tags, content controls, simple fields and custom XML.
   * Deleted and moved-away runs, and runs inside text boxes, are not included.
   */
  runs(): Run[] {
    const out: Run[] = [];
    const stack: XElement[] = [];
    pushChildren(stack, this.element);
    while (stack.length > 0) {
      const el = stack.pop()!;
      if (el.uri === NS.w) {
        if (el.local === 'r') {
          out.push(new Run(el, this));
          continue;
        }
        if (HIDDEN.has(el.local) || el.local === 'p') continue;
      } else if (el.uri === NS.mc && el.local === 'AlternateContent') {
        const branch = chooseAlternate(el);
        if (branch) stack.push(branch);
        continue;
      }
      if (el.local.endsWith('Pr')) continue;
      pushChildren(stack, el);
    }
    return out;
  }

  /**
   * Appends a run with `text` at the end of the paragraph. TAB becomes `w:tab`, LF (or CR/CRLF)
   * `w:br`, U+2011 `w:noBreakHyphen` and U+00AD `w:softHyphen`, so `run.text` reads back the same.
   */
  appendRun(text: string, props?: RunProps): Run {
    const r = wElement(this.element, 'r');
    this.element.append(r);
    const run = new Run(r, this);
    applyRunProps(run, props);
    run.text = text;
    return run;
  }

  /**
   * Appends an inline picture (in its own run) and returns that run. The image part is added
   * to the package (or reused when the same bytes are already there) with a relationship from
   * this paragraph's part, so it works in any story: body, headers, footers, notes, comments.
   * Throws UsageError for an unrecognized image format.
   */
  appendImage(bytes: Uint8Array, o?: ImageOptions): Run {
    const meta = imageInfo(bytes);
    if (meta === undefined) throw new UsageError('appendImage(): unrecognized image format');
    for (const v of [o?.width, o?.height]) {
      if (v !== undefined && !(v > 0 && Number.isFinite(v))) {
        throw new UsageError(`appendImage(): width/height must be positive points, got ${v}`);
      }
    }
    const naturalW = (meta.width * 72) / meta.dpiX;
    const naturalH = (meta.height * 72) / meta.dpiY;
    let w = o?.width;
    let h = o?.height;
    if (w === undefined && h === undefined) {
      w = naturalW;
      h = naturalH;
    } else if (h === undefined) h = ((w ?? 0) * naturalH) / naturalW;
    else if (w === undefined) w = (h * naturalW) / naturalH;

    const doc = this.story.doc;
    const owner = this.story.part;
    const added = addImagePart(doc, owner, bytes);
    let relId = added.relId;
    let svgRelId: string | undefined;
    if (meta.contentType === 'image/svg+xml') {
      svgRelId = relId;
      if (o?.fallback !== undefined) relId = addImagePart(doc, owner, o.fallback).relId;
    }
    const opts: { name?: string; description?: string; svgRelId?: string } = {};
    if (o?.name !== undefined) opts.name = o.name;
    if (o?.description !== undefined) opts.description = o.description;
    if (svgRelId !== undefined) opts.svgRelId = svgRelId;
    const drawing = inlineDrawing(
      doc,
      owner,
      relId,
      (w ?? 0) * EMU_PER_POINT,
      (h ?? 0) * EMU_PER_POINT,
      opts,
    );
    const r = wElement(this.element, 'r');
    this.element.append(r);
    r.append(drawing);
    return new Run(r, this);
  }

  /**
   * Appends a hyperlink with one run of `text` in the Hyperlink character style (created in
   * the styles part when missing; `props.style` overrides it). `target` is an external URL
   * (a relationship with TargetMode External is added to this paragraph's part) or
   * `{ anchor }`, the name of a bookmark in this document.
   */
  appendHyperlink(
    text: string,
    target: string | { anchor: string },
    props?: HyperlinkProps,
  ): Hyperlink {
    const doc = this.story.doc;
    const h = wElement(this.element, 'hyperlink');
    this.element.append(h);
    if (typeof target === 'string') {
      const rel = this.story.part.rels.add(relTypeFor(doc, REL.hyperlink), target, {
        external: true,
      });
      h.setAttrNS(NS.r, 'id', rel.id);
    } else {
      setWAttr(h, 'anchor', target.anchor);
    }
    if (props?.tooltip !== undefined) setWAttr(h, 'tooltip', props.tooltip);
    setWAttr(h, 'history', '1');
    const r = wElement(h, 'r');
    h.append(r);
    const run = new Run(r, this);
    const { tooltip: _tooltip, ...runProps } = props ?? {};
    applyRunProps(run, { ...runProps, style: props?.style ?? ensureHyperlinkStyle(doc) });
    run.text = text;
    return new Hyperlink(h, this.story);
  }

  /**
   * Splits the paragraph at visible-text offset `offset` (0..text.length) and returns the new
   * second half, inserted right after this one. Both halves keep the paragraph properties; a
   * section break moves to the second half (it ends the section). Runs, hyperlinks, content
   * controls and other inline wrappers spanning the offset are split in two; a simple field or
   * an equation is kept whole (in the second half). Throws RangeError for an invalid offset.
   */
  split(offset: number): Paragraph {
    const p = this.element;
    const leaves: [XElement, string][] = [];
    walkVisible(p, (node, t) => {
      leaves.push([node, t]);
    });
    let total = 0;
    for (const [, t] of leaves) total += t.length;
    if (!Number.isInteger(offset) || offset < 0 || offset > total) {
      throw new RangeError(`split(): offset ${offset} is outside 0..${total}`);
    }
    const doc = this.story.doc;
    let first: XNode | undefined;
    let at = 0;
    for (const [node, t] of leaves) {
      if (offset < at + t.length) {
        first = offset === at ? node : splitLeaf(node, offset - at);
        break;
      }
      at += t.length;
    }
    if (first !== undefined) {
      first = atomicAncestor(first, p);
      // Split every inline wrapper between the boundary and the paragraph.
      while (first.parent !== p) {
        const parent = first.parent as XElement;
        if (!contentBefore(first)) {
          first = parent;
          continue;
        }
        const copy = shallowCopy(parent);
        for (const k of parent.children) {
          if (k === first) break;
          if (k.nodeType === 1 && k.local.endsWith('Pr')) copy.append(k.clone());
        }
        for (const k of parent.children.slice(first.index())) copy.append(k);
        parent.after(copy);
        renewIds(doc, copy, false);
        first = copy;
      }
    }
    const np = shallowCopy(p);
    for (let i = np._attrs.length - 3; i >= 0; i -= 3) {
      if (np._attrs[i + 1] === NS.w14) np._attrs.splice(i, 3); // paraId/textId must stay unique
    }
    const pPr = p.element(NS.w, 'pPr');
    if (pPr !== undefined) {
      const copy = pPr.clone();
      renewIds(doc, copy, true);
      np.append(copy);
      const sectPr = pPr.element(NS.w, 'sectPr');
      if (sectPr !== undefined) {
        sectPr.remove();
        pruneEmpty(pPr);
      }
    }
    if (first !== undefined) for (const k of p.children.slice(first.index())) np.append(k);
    p.after(np);
    return new Paragraph(np, this.story);
  }

  /**
   * Inserts a paragraph right before this one and returns it: a new paragraph with one run of
   * `content` (and optional style), or an existing Paragraph, which is moved here (pass
   * `p.clone()` to copy; content referencing relationships must come from the same part).
   */
  insertBefore(content: string | Paragraph, o?: { style?: string }): Paragraph {
    return this.#insert(content, o, (el) => this.element.before(el));
  }

  /** Like {@link insertBefore}, right after this paragraph. */
  insertAfter(content: string | Paragraph, o?: { style?: string }): Paragraph {
    return this.#insert(content, o, (el) => this.element.after(el));
  }

  /**
   * Replaces this paragraph with `content` (see {@link insertBefore}) and returns the
   * replacement. A section break carried by this paragraph moves to the replacement.
   */
  replaceWith(content: string | Paragraph, o?: { style?: string }): Paragraph {
    const next = this.insertBefore(content, o);
    this.remove();
    return next;
  }

  /**
   * Removes the paragraph. A table cell, text box, header, footer or note left without
   * paragraphs gets an empty one (Word requires it). A section break carried by this paragraph
   * moves to the previous paragraph when that one has none; otherwise it is dropped and the
   * section merges into the next one, as when a section break is deleted in Word.
   */
  remove(): void {
    const p = this.element;
    const parent = p.parent;
    if (parent === null || parent.nodeType !== 1) return;
    const sectPr = p.element(NS.w, 'pPr')?.element(NS.w, 'sectPr');
    if (sectPr !== undefined) {
      const prev = p.previousElement();
      if (prev?.is(NS.w, 'p') && prev.element(NS.w, 'pPr')?.element(NS.w, 'sectPr') === undefined) {
        insertOrdered(ensureFirstChild(prev, 'pPr'), sectPr, PPR_ORDER);
      }
    }
    // The nearest container that must keep a paragraph.
    let box: XElement | undefined;
    for (let a: XElement | null = parent; a !== null; ) {
      if (a.uri === NS.w && (a.local === 'tc' || a.local === 'txbxContent')) {
        box = a;
        break;
      }
      if (a === this.story.element) {
        const k = this.story.kind;
        if (k === 'header' || k === 'footer' || k === 'footnote' || k === 'endnote') box = a;
        break;
      }
      a = a.parent?.nodeType === 1 ? a.parent : null;
    }
    p.remove();
    if (box === undefined) return;
    const last = lastBlock(box);
    if (box.local === 'tc' ? !last?.is(NS.w, 'p') : last === undefined) {
      box.append(wElement(box, 'p'));
    }
  }

  /**
   * A detached deep copy (insert it with {@link insertBefore}/{@link insertAfter} or
   * Story methods). Ids that must stay unique are renewed (drawings, content controls,
   * revisions) or dropped (`w14:paraId`/`textId`), and bookmarks and comment anchors are left
   * out, since the original keeps them.
   */
  clone(): Paragraph {
    const c = this.element.clone();
    const doc = this.story.doc;
    for (const el of [c, ...c.descendants()]) {
      if (el.uri === NS.w) {
        const l = el.local;
        if (
          l === 'bookmarkStart' ||
          l === 'bookmarkEnd' ||
          l === 'commentRangeStart' ||
          l === 'commentRangeEnd'
        ) {
          el.remove();
          continue;
        }
        if (l === 'commentReference' && el.parent?.nodeType === 1) {
          const run = el.parent;
          if (run.is(NS.w, 'r')) run.remove();
          else el.remove();
          continue;
        }
      }
      const a = el._attrs;
      for (let i = a.length - 3; i >= 0; i -= 3) {
        const u = a[i + 1];
        const local = a[i]!.slice(a[i]!.indexOf(':') + 1);
        if (
          (u === NS.w14 && (local === 'paraId' || local === 'textId')) ||
          (u === NS.wp14 && (local === 'anchorId' || local === 'editId'))
        ) {
          a.splice(i, 3);
        }
      }
    }
    renewIds(doc, c, true);
    return new Paragraph(c, this.story);
  }

  // ---- paragraph properties --------------------------------------------------------------

  /** Direct alignment (`w:jc`). */
  get alignment(): ParagraphAlignment | undefined {
    return pPrAttr(this.element, 'jc', 'val') as ParagraphAlignment | undefined;
  }

  /** Sets `w:jc` (in Strict documents left/right are written as start/end); undefined removes it. */
  set alignment(v: ParagraphAlignment | undefined) {
    let val: string | undefined = v;
    if (isStrict(this.element)) val = v === 'left' ? 'start' : v === 'right' ? 'end' : v;
    setValChild(this.element, 'pPr', 'jc', PPR_ORDER, val);
  }

  /** Left (start) indent in points (`w:ind/@w:left` or `@w:start`). */
  get indentLeft(): number | undefined {
    return points(pPrAttr(this.element, 'ind', 'start', 'left'));
  }

  set indentLeft(v: number | undefined) {
    this.#setInd(['left', 'start'], v);
  }

  /** Right (end) indent in points (`w:ind/@w:right` or `@w:end`). */
  get indentRight(): number | undefined {
    return points(pPrAttr(this.element, 'ind', 'end', 'right'));
  }

  set indentRight(v: number | undefined) {
    this.#setInd(['right', 'end'], v);
  }

  /** First-line indent in points: positive = `w:firstLine`, negative = hanging (`w:hanging`). */
  get firstLineIndent(): number | undefined {
    const hanging = points(pPrAttr(this.element, 'ind', 'hanging'));
    if (hanging !== undefined) return -hanging;
    return points(pPrAttr(this.element, 'ind', 'firstLine'));
  }

  set firstLineIndent(v: number | undefined) {
    const ind = this.#prop('ind', v !== undefined);
    if (ind === undefined) return;
    ind.removeAttrNS(NS.w, 'firstLineChars');
    ind.removeAttrNS(NS.w, 'hangingChars');
    if (v === undefined) {
      ind.removeAttrNS(NS.w, 'firstLine');
      ind.removeAttrNS(NS.w, 'hanging');
      pruneEmpty(ind);
    } else if (v < 0) {
      ind.removeAttrNS(NS.w, 'firstLine');
      setWAttr(ind, 'hanging', twips(-v));
    } else {
      ind.removeAttrNS(NS.w, 'hanging');
      setWAttr(ind, 'firstLine', twips(v));
    }
  }

  /** Space before the paragraph in points (`w:spacing/@w:before`). */
  get spaceBefore(): number | undefined {
    return points(pPrAttr(this.element, 'spacing', 'before'));
  }

  set spaceBefore(v: number | undefined) {
    this.#setSpacing('before', v);
  }

  /** Space after the paragraph in points (`w:spacing/@w:after`). */
  get spaceAfter(): number | undefined {
    return points(pPrAttr(this.element, 'spacing', 'after'));
  }

  set spaceAfter(v: number | undefined) {
    this.#setSpacing('after', v);
  }

  /** Line spacing (`w:spacing/@w:line` and `@w:lineRule`). */
  get lineSpacing(): LineSpacing | undefined {
    const line = pPrAttr(this.element, 'spacing', 'line');
    if (line === undefined) return undefined;
    const rule = pPrAttr(this.element, 'spacing', 'lineRule') ?? 'auto';
    if (rule === 'auto') {
      const n = Number(line);
      return Number.isFinite(n) ? { value: n / 240, rule: 'multiple' } : undefined;
    }
    const pt = points(line);
    return pt === undefined
      ? undefined
      : { value: pt, rule: rule === 'exact' ? 'exact' : 'atLeast' };
  }

  /** A number is a multiple of single spacing; undefined removes the line spacing. */
  set lineSpacing(v: LineSpacing | number | undefined) {
    const sp = this.#prop('spacing', v !== undefined);
    if (sp === undefined) return;
    if (v === undefined) {
      sp.removeAttrNS(NS.w, 'line');
      sp.removeAttrNS(NS.w, 'lineRule');
      pruneEmpty(sp);
      return;
    }
    const ls = typeof v === 'number' ? { value: v, rule: 'multiple' as const } : v;
    if (!(ls.value > 0 && Number.isFinite(ls.value))) {
      throw new UsageError(`lineSpacing must be positive, got ${ls.value}`);
    }
    if (ls.rule === 'multiple') {
      setWAttr(sp, 'line', String(Math.round(ls.value * 240)));
      setWAttr(sp, 'lineRule', 'auto');
    } else {
      setWAttr(sp, 'line', twips(ls.value));
      setWAttr(sp, 'lineRule', ls.rule);
    }
  }

  /** Keep with next (`w:keepNext`): tri-state like {@link Run.bold}. */
  get keepNext(): boolean | undefined {
    return readOnOff(this.element.element(NS.w, 'pPr')?.element(NS.w, 'keepNext'));
  }

  set keepNext(v: boolean | undefined) {
    setOnOff(this.element, 'pPr', 'keepNext', PPR_ORDER, v);
  }

  /** Keep lines together (`w:keepLines`). */
  get keepLines(): boolean | undefined {
    return readOnOff(this.element.element(NS.w, 'pPr')?.element(NS.w, 'keepLines'));
  }

  set keepLines(v: boolean | undefined) {
    setOnOff(this.element, 'pPr', 'keepLines', PPR_ORDER, v);
  }

  /** Page break before (`w:pageBreakBefore`). */
  get pageBreakBefore(): boolean | undefined {
    return readOnOff(this.element.element(NS.w, 'pPr')?.element(NS.w, 'pageBreakBefore'));
  }

  set pageBreakBefore(v: boolean | undefined) {
    setOnOff(this.element, 'pPr', 'pageBreakBefore', PPR_ORDER, v);
  }

  /** `w:pPr/w:<local>`, created (in schema order) when `create`. */
  #prop(local: string, create: boolean): XElement | undefined {
    if (create) return ensureChild(ensureFirstChild(this.element, 'pPr'), local, PPR_ORDER);
    return this.element.element(NS.w, 'pPr')?.element(NS.w, local);
  }

  /** Sets an indent attribute: transitional name, or the Strict/bidi-aware start/end form. */
  #setInd([transitional, logical]: [string, string], v: number | undefined): void {
    const ind = this.#prop('ind', v !== undefined);
    if (ind === undefined) return;
    ind.removeAttrNS(NS.w, `${transitional}Chars`);
    ind.removeAttrNS(NS.w, `${logical}Chars`);
    if (v === undefined) {
      ind.removeAttrNS(NS.w, transitional);
      ind.removeAttrNS(NS.w, logical);
      pruneEmpty(ind);
      return;
    }
    const useLogical = ind.attrNS(NS.w, logical) !== undefined || isStrict(this.element);
    ind.removeAttrNS(NS.w, useLogical ? transitional : logical);
    setWAttr(ind, useLogical ? logical : transitional, twips(v));
  }

  #setSpacing(which: 'before' | 'after', v: number | undefined): void {
    const sp = this.#prop('spacing', v !== undefined);
    if (sp === undefined) return;
    sp.removeAttrNS(NS.w, `${which}Lines`);
    sp.removeAttrNS(NS.w, `${which}Autospacing`);
    if (v === undefined) {
      sp.removeAttrNS(NS.w, which);
      pruneEmpty(sp);
    } else {
      setWAttr(sp, which, twips(v));
    }
  }

  #insert(
    content: string | Paragraph,
    o: { style?: string } | undefined,
    place: (el: XElement) => void,
  ): Paragraph {
    if (typeof content !== 'string') {
      if (content.element === this.element) {
        throw new UsageError('Cannot insert a paragraph next to itself');
      }
      place(content.element);
      const para = new Paragraph(content.element, this.story);
      if (o?.style !== undefined) para.style = o.style;
      return para;
    }
    const el = wElement(this.element, 'p');
    place(el);
    const para = new Paragraph(el, this.story);
    if (o?.style !== undefined) para.style = o.style;
    if (content !== '') para.appendRun(content);
    return para;
  }
}

/** A run (`w:r`): a stretch of content with uniform direct formatting. */
export class Run {
  /** The `w:r` element. */
  readonly element: XElement;
  /** The paragraph the run was reached from. */
  readonly paragraph: Paragraph;

  constructor(el: XElement, paragraph: Paragraph) {
    this.element = el;
    this.paragraph = paragraph;
  }

  /** Visible text of the run. */
  get text(): string {
    return paragraphText(this.element);
  }

  /**
   * Replaces the run's content (everything but `w:rPr`) with `value`, split into `w:t`,
   * `w:tab`, `w:br`, `w:noBreakHyphen` and `w:softHyphen` as in {@link Paragraph.appendRun}.
   */
  set text(value: string) {
    const r = this.element;
    for (const k of r.children.slice()) {
      if (k.nodeType === 1 && k.is(NS.w, 'rPr')) continue;
      k.remove();
    }
    for (const n of runContent(r, value)) r.append(n);
  }

  /** Character style id (`w:rPr/w:rStyle/@w:val`). */
  get style(): string | undefined {
    return this.element.element(NS.w, 'rPr')?.element(NS.w, 'rStyle')?.attrNS(NS.w, 'val');
  }

  /** Sets the character style id; undefined removes `w:rStyle` (and an emptied `w:rPr`). */
  set style(v: string | undefined) {
    setValChild(this.element, 'rPr', 'rStyle', RPR_ORDER, v);
  }

  /** Direct bold (`w:b`): true/false when set on the run, undefined when inherited. */
  get bold(): boolean | undefined {
    return readOnOff(this.element.element(NS.w, 'rPr')?.element(NS.w, 'b'));
  }

  /** true writes `<w:b/>`, false `<w:b w:val="0"/>`, undefined removes it. */
  set bold(v: boolean | undefined) {
    setOnOff(this.element, 'rPr', 'b', RPR_ORDER, v);
  }

  /** Direct italic (`w:i`): true/false when set on the run, undefined when inherited. */
  get italic(): boolean | undefined {
    return readOnOff(this.element.element(NS.w, 'rPr')?.element(NS.w, 'i'));
  }

  /** true writes `<w:i/>`, false `<w:i w:val="0"/>`, undefined removes it. */
  set italic(v: boolean | undefined) {
    setOnOff(this.element, 'rPr', 'i', RPR_ORDER, v);
  }

  /** Direct underline style (`w:u/@w:val`). */
  get underline(): UnderlineStyle | undefined {
    return this.#rPrAttr('u', 'val') as UnderlineStyle | undefined;
  }

  /** Sets `w:u/@w:val` (keeping an underline colour); undefined removes `w:u`. */
  set underline(v: UnderlineStyle | undefined) {
    setValChild(this.element, 'rPr', 'u', RPR_ORDER, v);
  }

  /** Direct text colour (`w:color/@w:val`): `RRGGBB` or `auto`. */
  get color(): string | undefined {
    return this.#rPrAttr('color', 'val');
  }

  /**
   * Sets `w:color` to an `RRGGBB` hex value (a leading `#` is accepted) or `auto`, dropping
   * theme colour attributes (they would override it); undefined removes it.
   */
  set color(v: string | undefined) {
    if (v === undefined) {
      setValChild(this.element, 'rPr', 'color', RPR_ORDER, undefined);
      return;
    }
    const hex = v.replace(/^#/, '');
    if (!/^(?:[0-9a-fA-F]{6}|auto)$/.test(hex)) {
      throw new UsageError(`color must be RRGGBB hex or "auto", got ${JSON.stringify(v)}`);
    }
    const el = ensureChild(ensureFirstChild(this.element, 'rPr'), 'color', RPR_ORDER);
    for (const a of ['themeColor', 'themeTint', 'themeShade']) el.removeAttrNS(NS.w, a);
    setWAttr(el, 'val', hex === 'auto' ? hex : hex.toUpperCase());
  }

  /** Direct font size in points (`w:sz`, stored in half-points). */
  get size(): number | undefined {
    const v = this.#rPrAttr('sz', 'val');
    const n = v === undefined ? Number.NaN : Number(v);
    return Number.isFinite(n) ? n / 2 : undefined;
  }

  /** Sets `w:sz` and `w:szCs` (complex scripts) to the nearest half-point; undefined removes both. */
  set size(v: number | undefined) {
    if (v !== undefined && !(v > 0 && v <= 1638)) {
      throw new UsageError(`size must be between 0.5 and 1638 points, got ${v}`);
    }
    const half = v === undefined ? undefined : String(Math.max(1, Math.round(v * 2)));
    setValChild(this.element, 'rPr', 'sz', RPR_ORDER, half);
    setValChild(this.element, 'rPr', 'szCs', RPR_ORDER, half);
  }

  /** Direct font name (`w:rFonts/@w:ascii`, else `@w:hAnsi`). */
  get font(): string | undefined {
    return this.#rPrAttr('rFonts', 'ascii') ?? this.#rPrAttr('rFonts', 'hAnsi');
  }

  /**
   * Sets the ASCII and high-ANSI fonts (`w:rFonts/@w:ascii`, `@w:hAnsi`), dropping their theme
   * fonts (which would override them); undefined removes them.
   */
  set font(v: string | undefined) {
    const rPr = this.element.element(NS.w, 'rPr');
    if (v === undefined) {
      const el = rPr?.element(NS.w, 'rFonts');
      if (el === undefined) return;
      for (const a of ['ascii', 'hAnsi', 'asciiTheme', 'hAnsiTheme']) el.removeAttrNS(NS.w, a);
      pruneEmpty(el);
      return;
    }
    const el = ensureChild(ensureFirstChild(this.element, 'rPr'), 'rFonts', RPR_ORDER);
    el.removeAttrNS(NS.w, 'asciiTheme');
    el.removeAttrNS(NS.w, 'hAnsiTheme');
    setWAttr(el, 'ascii', v);
    setWAttr(el, 'hAnsi', v);
  }

  /** Direct highlight (`w:highlight/@w:val`). */
  get highlight(): HighlightColor | undefined {
    return this.#rPrAttr('highlight', 'val') as HighlightColor | undefined;
  }

  /** Sets `w:highlight`; undefined removes it. */
  set highlight(v: HighlightColor | undefined) {
    setValChild(this.element, 'rPr', 'highlight', RPR_ORDER, v);
  }

  #rPrAttr(local: string, attr: string): string | undefined {
    return this.element.element(NS.w, 'rPr')?.element(NS.w, local)?.attrNS(NS.w, attr);
  }
}

function applyRunProps(run: Run, props: RunProps | undefined): void {
  if (props === undefined) return;
  if (props.style !== undefined) run.style = props.style;
  if (props.font !== undefined) run.font = props.font;
  if (props.bold !== undefined) run.bold = props.bold;
  if (props.italic !== undefined) run.italic = props.italic;
  if (props.color !== undefined) run.color = props.color;
  if (props.size !== undefined) run.size = props.size;
  if (props.highlight !== undefined) run.highlight = props.highlight;
  if (props.underline !== undefined) run.underline = props.underline;
}

function pushChildren(stack: XElement[], el: XElement): void {
  const kids = el.children;
  for (let i = kids.length - 1; i >= 0; i--) {
    const k = kids[i]!;
    if (k.nodeType === 1) stack.push(k);
  }
}

/** The document binds WordprocessingML to its Strict namespace URI. */
function isStrict(el: XElement): boolean {
  return isStrictNs(el.lookupNamespace(wPrefix(el)) ?? '');
}

/** Points → twentieths of a point, as the attribute string. */
function twips(pt: number): string {
  if (!Number.isFinite(pt)) throw new UsageError(`Expected a number of points, got ${pt}`);
  return String(Math.round(pt * 20));
}

/** Points per unit of an ST_UniversalMeasure. */
const UNIT_POINTS: Readonly<Record<string, number>> = {
  pt: 1,
  pc: 12,
  pi: 12,
  in: 72,
  cm: 72 / 2.54,
  mm: 72 / 25.4,
};

/** Twentieths of a point (or a universal measure such as `12pt`) → points. */
function points(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const m = /^\s*(-?\d+(?:\.\d+)?)\s*(pt|pc|pi|in|cm|mm)?\s*$/.exec(v);
  if (m === null) return undefined;
  const n = Number(m[1]);
  return m[2] === undefined ? n / 20 : n * UNIT_POINTS[m[2]]!;
}

/** First defined attribute among `attrs` of `w:pPr/w:<local>`. */
function pPrAttr(p: XElement, local: string, ...attrs: string[]): string | undefined {
  const el = p.element(NS.w, 'pPr')?.element(NS.w, local);
  if (el === undefined) return undefined;
  for (const a of attrs) {
    const v = el.attrNS(NS.w, a);
    if (v !== undefined) return v;
  }
  return undefined;
}

/** ST_OnOff (§17.17.4): absent `w:val` means on; `false`/`off`/`0` mean off. */
function readOnOff(el: XElement | undefined): boolean | undefined {
  if (el === undefined) return undefined;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined || !(v === 'false' || v === '0' || v === 'off');
}

/** Sets/removes the on/off property `owner/w:<props>/w:<local>`. */
function setOnOff(
  owner: XElement,
  props: string,
  local: string,
  order: ChildOrder,
  v: boolean | undefined,
): void {
  if (v === undefined) {
    const pr = owner.element(NS.w, props);
    const el = pr?.element(NS.w, local);
    if (pr && el) {
      el.remove();
      pruneEmpty(pr);
    }
    return;
  }
  const el = ensureChild(ensureFirstChild(owner, props), local, order);
  if (v) el.removeAttrNS(NS.w, 'val');
  else setWAttr(el, 'val', '0');
}

/** Sets/removes `owner/w:<props>/w:<local>/@w:val` (pStyle, rStyle, ...). */
function setValChild(
  owner: XElement,
  props: string,
  local: string,
  order: ChildOrder,
  v: string | undefined,
): void {
  if (v === undefined) {
    const pr = owner.element(NS.w, props);
    const el = pr?.element(NS.w, local);
    if (pr && el) {
      el.remove();
      pruneEmpty(pr);
    }
    return;
  }
  setWAttr(ensureChild(ensureFirstChild(owner, props), local, order), 'val', v);
}

/** Text that needs `xml:space="preserve"` to keep its whitespace. */
function needsPreserve(s: string): boolean {
  return /^\s|\s$|\s\s/.test(s);
}

/** Run content nodes for `text` (see {@link Paragraph.appendRun}), named in `scope`'s prefixes. */
function runContent(scope: XElement, text: string): XNode[] {
  const out: XNode[] = [];
  let start = 0;
  const flush = (end: number): void => {
    if (end <= start) return;
    const s = text.slice(start, end);
    const t = wElement(scope, 't');
    if (needsPreserve(s)) t.setAttr('xml:space', 'preserve');
    t.append(new XText(s));
    out.push(t);
  };
  for (let i = 0; i < text.length; i++) {
    let leaf: string | undefined;
    switch (text.charCodeAt(i)) {
      case 0x09:
        leaf = 'tab';
        break;
      case 0x0a:
        leaf = 'br';
        break;
      case 0x0d:
        leaf = 'br';
        break;
      case 0x2011:
        leaf = 'noBreakHyphen';
        break;
      case 0xad:
        leaf = 'softHyphen';
        break;
      default:
        continue;
    }
    flush(i);
    out.push(wElement(scope, leaf));
    if (text.charCodeAt(i) === 0x0d && text.charCodeAt(i + 1) === 0x0a) i++;
    start = i + 1;
  }
  flush(text.length);
  return out;
}

// ---- split helpers ------------------------------------------------------------------------

/** Detached copy of `el` with its attributes and no children. */
function shallowCopy(el: XElement): XElement {
  return el.clone(false);
}

/**
 * Splits the text leaf `leaf` (a `w:t`) at `k` (0 < k < length) and returns the second half,
 * inserted right after it. Leaves that cannot be split (symbols, math) return themselves, so
 * the boundary falls before them; a surrogate pair is never cut.
 */
function splitLeaf(leaf: XElement, k: number): XElement {
  if (!leaf.is(NS.w, 't')) return leaf;
  const s = leaf.text;
  let at = k;
  const c = s.charCodeAt(at - 1);
  if (c >= 0xd800 && c <= 0xdbff) at--;
  if (at <= 0) return leaf;
  const second = shallowCopy(leaf);
  const a = s.slice(0, at);
  const b = s.slice(at);
  leaf.text = a;
  second.text = b;
  for (const [el, v] of [
    [leaf, a],
    [second, b],
  ] as const) {
    if (needsPreserve(v)) el.setAttr('xml:space', 'preserve');
  }
  leaf.after(second);
  return second;
}

/** The outermost simple field or equation containing `node` below `p`, else `node`. */
function atomicAncestor(node: XNode, p: XElement): XNode {
  let out = node;
  for (let a = node.parent; a !== null && a !== p && a.nodeType === 1; a = a.parent) {
    if (a.is(NS.w, 'fldSimple') || a.is(NS.m, 'oMath') || a.is(NS.m, 'oMathPara')) out = a;
  }
  return out;
}

/** Something other than property elements and whitespace precedes `node` among its siblings. */
function contentBefore(node: XNode): boolean {
  const kids = node.parent?.children ?? [];
  for (const k of kids) {
    if (k === node) return false;
    if (k.nodeType === 1) {
      if (!k.local.endsWith('Pr')) return true;
    } else if (k.nodeType === 3) {
      if (k.text.trim() !== '') return true;
    }
  }
  return false;
}

/**
 * Last block-level element of a container, looking through block-level content controls and
 * custom XML (the rule validate.ts checks for cells).
 */
function lastBlock(container: XElement): XElement | undefined {
  let c: XElement = container;
  for (;;) {
    let found: XElement | undefined;
    const kids = c.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1 && k.uri === NS.w && BLOCKS.has(k.local)) {
        found = k;
        break;
      }
    }
    if (found === undefined) return undefined;
    if (found.local === 'sdt') {
      const content = found.element(NS.w, 'sdtContent');
      if (content === undefined) return undefined;
      c = content;
    } else if (found.local === 'customXml') {
      c = found;
    } else {
      return found;
    }
  }
}
