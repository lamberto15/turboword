// One `w:p` → IRParagraph: resolved properties, the list label and the inline content.
//
// The walk follows the same visibility rules as the text flattener (wml/text.ts): deleted and
// moved-away content is skipped, insertions, smart tags, content controls, custom XML, `w:dir`/
// `w:bdo` and simple fields are traversed, property elements are ignored. In addition it keeps
// what layout needs and text() drops: run formatting, tabs and breaks as such, drawings, math,
// note references, page-dependent fields, hyperlinks and bookmarks. Hidden runs (`w:vanish`) show
// nothing but still delimit fields. Explicit stacks only.

import type { Part } from '../../opc/part.js';
import { listInfo, resolveParagraph, resolveRun } from '../../resolve/index.js';
import type { ResolvedParagraphProps, ResolvedRunProps } from '../../resolve/types.js';
import { Paragraph, Run } from '../../wml/paragraph.js';
import type { Story } from '../../wml/story.js';
import { paragraphText } from '../../wml/text.js';
import type { XElement, XNode } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import type { IRAnchor, IRInline, IRLink, IRParagraph } from '../types.js';
import type { IRInlineExtras, IRParagraphExtras, LineCompat, PtabInfo } from './extras.js';
import { type DrawingContext, drawingToIR } from './ir-drawing.js';
import {
  type FieldInline,
  type FieldState,
  hyperlinkTarget,
  pageFieldInline,
  parseInstruction,
} from './ir-fields.js';
import type { NoteMarks } from './ir-notes.js';
import { layoutAlternate, onOffValue } from './ir-util.js';

/** Everything the paragraph builder needs from the document-level builder. */
export interface ParagraphContext {
  story: Story;
  /** Part holding the paragraph (for hyperlink and image relationships). */
  part: Part;
  /** Complex-field state of the story (fields can span paragraphs). */
  fields: FieldState;
  notes: NoteMarks;
  drawing: DrawingContext;
  compat: LineCompat;
  warn(code: string, message: string): void;
}

interface RunCtx {
  el: XElement;
  props: ResolvedRunProps;
  /** Text after a `customMarkFollows` reference is that reference's mark, not content. */
  skipText: boolean;
  /** The noteRef waiting for its custom mark text. */
  customRef?: Extract<IRInline, { kind: 'noteRef' }>;
}

interface Frame {
  kids: readonly XNode[];
  i: number;
  link: IRLink | undefined;
  run: RunCtx | undefined;
}

/** Character of a `w:sym` (hex `w:char`; '' when invalid). */
function symChar(el: XElement): string {
  const hex = el.attrNS(NS.w, 'char');
  if (hex === undefined || !/^[0-9A-Fa-f]{1,6}$/.test(hex)) return '';
  let cp = Number.parseInt(hex, 16);
  // Symbol fonts are addressed in the F000 private-use page; plain 00xx codes mean the same glyph.
  if (cp > 0 && cp < 0x100 && el.attrNS(NS.w, 'font') !== undefined) cp += 0xf000;
  return cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : '';
}

/** Props with every font slot set to `font` (for `w:sym`). */
function withFont(props: ResolvedRunProps, font: string): ResolvedRunProps {
  return { ...props, fonts: { ascii: font, hAnsi: font, eastAsia: font, cs: font } };
}

function hyperlinkOf(el: XElement, part: Part): IRLink | undefined {
  const link: IRLink = {};
  const id = el.attrNS(NS.r, 'id');
  if (id !== undefined) {
    const rel = part.rels.get(id);
    if (rel !== undefined) link.url = rel.target;
  }
  const anchor = el.attrNS(NS.w, 'anchor');
  if (anchor !== undefined && anchor !== '') link.anchor = anchor;
  return link.url === undefined && link.anchor === undefined ? undefined : link;
}

function ptabOf(el: XElement): PtabInfo {
  const a = el.attrNS(NS.w, 'alignment');
  const leader = el.attrNS(NS.w, 'leader');
  const out: PtabInfo = {
    alignment: a === 'center' || a === 'right' ? a : 'left',
    relativeTo: el.attrNS(NS.w, 'relativeTo') === 'indent' ? 'indent' : 'margin',
  };
  if (
    leader === 'dot' ||
    leader === 'hyphen' ||
    leader === 'underscore' ||
    leader === 'middleDot'
  ) {
    out.leader = leader;
  }
  return out;
}

/** Builds the IR of one paragraph. `pendingBookmarks` (block-level starts before it) are moved in. */
export function buildParagraph(
  pEl: XElement,
  ctx: ParagraphContext,
  pendingBookmarks: string[],
): IRParagraph {
  const para = new Paragraph(pEl, ctx.story);
  const props: ResolvedParagraphProps = resolveParagraph(para);
  const inlines: IRInline[] = [];
  const anchors: IRAnchor[] = [];
  const bookmarks: string[] = pendingBookmarks.splice(0);
  const bookmarkAt: number[] = bookmarks.map(() => 0);
  const out: IRParagraph & IRParagraphExtras = {
    kind: 'paragraph',
    source: pEl,
    props,
    inlines,
    anchors,
    bookmarks,
    bookmarkAt,
    compat: ctx.compat,
  };

  if (props.numbering !== undefined) {
    const li = listInfo(para);
    if (li !== undefined && (li.label !== '' || li.suffix !== 'nothing')) {
      out.label = {
        text: li.label,
        props: li.runProps,
        suffix: li.suffix,
        alignment: li.alignment,
      };
    }
  }

  const fields = ctx.fields;
  /** Last text inline, to extend while text continues in the same run with the same link. */
  let lastText: Extract<IRInline, { kind: 'text' }> | undefined;
  const push = (inline: IRInline): void => {
    inlines.push(inline);
    lastText = inline.kind === 'text' ? inline : undefined;
  };
  const emitText = (
    text: string,
    run: RunCtx,
    linkCtx: IRLink | undefined,
    props2?: ResolvedRunProps,
  ): void => {
    if (text === '') return;
    const r = fields.route();
    if (r.to === 'instr') return;
    if (r.to === 'fallback') {
      r.field.fallback += text;
      if (r.first) {
        r.field.props = run.props;
        fields.tookProps();
      }
      return;
    }
    if (run.props.hidden) return;
    const link = r.link ?? linkCtx;
    const p = props2 ?? run.props;
    if (
      lastText !== undefined &&
      lastText.source === run.el &&
      lastText.props === p &&
      lastText.link === link
    ) {
      lastText.text += text;
      return;
    }
    const t: Extract<IRInline, { kind: 'text' }> = { kind: 'text', text, props: p, source: run.el };
    if (link !== undefined) t.link = link;
    push(t);
  };
  /** True when non-text content (tabs, breaks, drawings, ...) is shown at this point. */
  const shown = (run: RunCtx | undefined): boolean =>
    fields.route().to === 'show' && run?.props.hidden !== true;
  const pushField = (f: FieldInline | undefined): void => {
    if (f !== undefined) push(f);
  };

  const stack: Frame[] = [{ kids: pEl.children, i: 0, link: undefined, run: undefined }];
  while (stack.length > 0) {
    const f = stack[stack.length - 1]!;
    if (f.i >= f.kids.length) {
      stack.pop();
      continue;
    }
    const k = f.kids[f.i++]!;
    if (k.nodeType !== 1) continue;
    const el = k as XElement;
    const run = f.run;

    if (el.uri === NS.mc) {
      if (el.local === 'AlternateContent') {
        const branch = layoutAlternate(el);
        if (branch !== undefined) stack.push({ kids: branch.children, i: 0, link: f.link, run });
      }
      continue;
    }
    if (el.uri === NS.m) {
      if ((el.local === 'oMath' || el.local === 'oMathPara') && shown(run)) {
        push({ kind: 'math', source: el, props: run?.props ?? props.runDefaults });
      }
      continue;
    }
    if (el.uri !== NS.w) continue;

    switch (el.local) {
      // ---- containers -------------------------------------------------------------------
      case 'r': {
        const rp = resolveRun(new Run(el, para), props);
        stack.push({
          kids: el.children,
          i: 0,
          link: f.link,
          run: { el, props: rp, skipText: false },
        });
        continue;
      }
      case 'hyperlink':
        stack.push({ kids: el.children, i: 0, link: hyperlinkOf(el, ctx.part) ?? f.link, run });
        continue;
      case 'ins':
      case 'moveTo':
      case 'smartTag':
      case 'customXml':
      case 'dir':
      case 'bdo':
      case 'sdtContent':
      case 'rubyBase':
      case 'ruby':
        stack.push({ kids: el.children, i: 0, link: f.link, run });
        continue;
      case 'sdt': {
        const content = el.element(NS.w, 'sdtContent');
        if (content !== undefined) stack.push({ kids: content.children, i: 0, link: f.link, run });
        continue;
      }
      case 'fldSimple': {
        const ins = parseInstruction(el.attrNS(NS.w, 'instr') ?? '');
        const firstRun = el.descendants(NS.w, 'r')[0];
        const rp = firstRun ? resolveRun(new Run(firstRun, para), props) : props.runDefaults;
        const page =
          fields.route().to === 'show' ? pageFieldInline(ins, rp, paragraphText(el)) : undefined;
        if (page !== undefined) {
          if (!rp.hidden) push(page);
          continue;
        }
        const link = ins.type === 'HYPERLINK' ? (hyperlinkTarget(ins) ?? f.link) : f.link;
        stack.push({ kids: el.children, i: 0, link, run });
        continue;
      }
      case 'bookmarkStart': {
        const name = el.attrNS(NS.w, 'name');
        if (name !== undefined && name !== '') {
          bookmarks.push(name);
          bookmarkAt.push(inlines.length);
        }
        continue;
      }
      case 'pPr':
      case 'rPr':
      case 'del':
      case 'moveFrom':
        continue;
    }

    // ---- run content (only meaningful inside a w:r) ---------------------------------------
    if (run === undefined) continue;
    switch (el.local) {
      case 't':
        if (run.skipText) {
          if (run.customRef !== undefined) run.customRef.text += el.text;
          continue;
        }
        emitText(el.text, run, f.link);
        continue;
      case 'noBreakHyphen':
        if (!run.skipText) emitText('\u2011', run, f.link);
        continue;
      case 'softHyphen':
        if (!run.skipText) emitText('\u00ad', run, f.link);
        continue;
      case 'sym': {
        const ch = symChar(el);
        const font = el.attrNS(NS.w, 'font');
        if (run.skipText) {
          if (run.customRef !== undefined) run.customRef.text += ch;
          continue;
        }
        emitText(ch, run, f.link, font !== undefined ? withFont(run.props, font) : undefined);
        continue;
      }
      case 'instrText':
        fields.instrText(el.text);
        continue;
      case 'fldChar': {
        const type = el.attrNS(NS.w, 'fldCharType');
        if (type === 'begin') fields.begin();
        else if (type === 'separate') pushField(fields.separate(run.props));
        else if (type === 'end') pushField(fields.end(run.props));
        lastText = undefined;
        continue;
      }
      case 'tab':
      case 'ptab':
        if (shown(run)) {
          const tab: IRInline & IRInlineExtras = { kind: 'tab', props: run.props };
          if (el.local === 'ptab') tab.ptab = ptabOf(el);
          push(tab);
        }
        continue;
      case 'br':
      case 'cr': {
        if (!shown(run)) continue;
        const t = el.attrNS(NS.w, 'type');
        const type = t === 'page' ? 'page' : t === 'column' ? 'column' : 'line';
        const br: Extract<IRInline, { kind: 'break' }> = { kind: 'break', type };
        const clear = el.attrNS(NS.w, 'clear');
        if (clear === 'left' || clear === 'right' || clear === 'all') br.clear = clear;
        push(br);
        continue;
      }
      case 'pgNum':
        if (shown(run)) push({ kind: 'field', type: 'PAGE', fallback: '', props: run.props });
        continue;
      case 'footnoteReference':
      case 'endnoteReference': {
        const id = el.attrNS(NS.w, 'id');
        if (id === undefined || !shown(run)) continue;
        const note = el.local === 'footnoteReference' ? 'footnote' : 'endnote';
        const custom = onOffValue(el.attrNS(NS.w, 'customMarkFollows'), false);
        const ref: Extract<IRInline, { kind: 'noteRef' }> = {
          kind: 'noteRef',
          note,
          id,
          props: run.props,
          text: custom ? '' : (ctx.notes.ofReference(el) ?? id),
        };
        if (custom) {
          run.skipText = true;
          run.customRef = ref;
        }
        push(ref);
        continue;
      }
      case 'footnoteRef':
      case 'endnoteRef': {
        const kind = el.local === 'footnoteRef' ? 'footnote' : 'endnote';
        const id = ctx.story.id;
        const mark = id === undefined ? undefined : ctx.notes.ofNote(kind, id);
        emitText(mark ?? id ?? '', run, f.link);
        continue;
      }
      case 'drawing':
      case 'pict':
      case 'object': {
        if (!shown(run)) continue;
        const r = fields.route();
        const link = r.to === 'show' ? (r.link ?? f.link) : undefined;
        const res = drawingToIR(el, run.props, link, ctx.drawing);
        if (res === undefined) continue;
        if ('inline' in res) push(res.inline);
        else anchors.push(res.anchor);
        continue;
      }
      default:
        continue;
    }
  }
  return out;
}
