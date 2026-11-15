// Paragraph inlines → "pieces": uniformly styled spans in logical order, with the character
// transformations and size/baseline rules of Word applied.
//
// - caps: upper-case; small caps: lower-case letters upper-cased at 0.8 × size (as Word/LO do);
// - superscript/subscript: 2/3 of the size, baseline raised by 1/3 of the size (super) or lowered
//   by 0.14 of it (sub); `w:position` raises/lowers by its half-points on top;
// - the list label (and its suffix) comes first; a field becomes its fallback text (the paginator
//   replaces it per page); a note reference becomes its mark;
// - inline images, drawings and math become boxed objects with their display items.
// Each piece also contributes characters to the paragraph's logical text, which feeds UAX #14 and
// UAX #9: text as is, TAB for tabs, LF for breaks, U+FFFC for objects.

import type { ResolvedRunProps } from '../../resolve/types.js';
import type { LayoutEnv } from '../engine.js';
import { drawingItems, layoutMath, metafileImageItems } from '../graphics/index.js';
import type { DisplayItem, IRInline, IRLink, IRParagraph } from '../types.js';
import { inlineExtras, type PtabInfo } from './extras.js';
import type { FieldInline } from './ir-fields.js';

/** Index used for pieces of the list label and its suffix. */
export const LABEL = -1;

export interface TextPiece {
  kind: 'text';
  text: string;
  props: ResolvedRunProps;
  /** Size factor (super/subscript 2/3, small caps 0.8). */
  factor: number;
  /** Baseline shift upwards, points. */
  rise: number;
  /** Index in IRParagraph.inlines (LABEL for the label). */
  inline: number;
  link?: IRLink;
  field?: FieldInline;
  /** Note id of a footnote/endnote reference. */
  note?: string;
  /** The space after a list label (`w:suff="space"`). */
  labelSuffix?: boolean;
}

export interface TabPiece {
  kind: 'tab';
  props: ResolvedRunProps;
  inline: number;
  ptab?: PtabInfo;
  /** The tab after a list label. */
  listTab: boolean;
}

export interface BreakPiece {
  kind: 'break';
  type: 'line' | 'page' | 'column';
  props: ResolvedRunProps;
  inline: number;
}

export interface ObjectPiece {
  kind: 'object';
  props: ResolvedRunProps;
  inline: number;
  width: number;
  ascent: number;
  descent: number;
  /** Relative to (left edge, baseline). */
  items: DisplayItem[];
  link?: IRLink;
}

export type Piece = TextPiece | TabPiece | BreakPiece | ObjectPiece;

const SMALL_CAPS = 0.8;
const SCRIPT_SIZE = 2 / 3;
const SUPER_RISE = 1 / 3;
const SUB_DROP = 0.14;

/** Upper-cases code point by code point, keeping a char when its upper case changes the length. */
function upper(s: string): string {
  let out = '';
  for (const ch of s) {
    const u = ch.toUpperCase();
    out += u.length === ch.length ? u : ch;
  }
  return out;
}

function isLower(ch: string): boolean {
  const u = ch.toUpperCase();
  return u !== ch && u.length === ch.length;
}

/** Base size factor and rise of a run's vertical alignment and position. */
function vertical(props: ResolvedRunProps): { factor: number; rise: number } {
  let factor = 1;
  let rise = props.position / 2;
  if (props.vertAlign === 'superscript') {
    factor = SCRIPT_SIZE;
    rise += props.size * SUPER_RISE;
  } else if (props.vertAlign === 'subscript') {
    factor = SCRIPT_SIZE;
    rise -= props.size * SUB_DROP;
  }
  return { factor, rise };
}

/** Text pieces for `text` in `props` (caps/small caps split), appended to `out`. */
function pushText(
  out: Piece[],
  text: string,
  props: ResolvedRunProps,
  inline: number,
  extra: { link?: IRLink; field?: FieldInline; note?: string },
): void {
  if (text === '') return;
  const { factor, rise } = vertical(props);
  const base = (t: string, f: number): TextPiece => {
    const p: TextPiece = { kind: 'text', text: t, props, factor: f, rise, inline };
    if (extra.link !== undefined) p.link = extra.link;
    if (extra.field !== undefined) p.field = extra.field;
    if (extra.note !== undefined) p.note = extra.note;
    return p;
  };
  if (props.caps || !props.smallCaps || extra.field !== undefined) {
    out.push(base(props.caps || props.smallCaps ? upper(text) : text, factor));
    return;
  }
  // Small caps: alternate runs of lower-case (shown upper-case, smaller) and everything else.
  let run = '';
  let runLower = false;
  for (const ch of text) {
    const low = isLower(ch);
    if (run !== '' && low !== runLower) {
      out.push(base(runLower ? upper(run) : run, runLower ? factor * SMALL_CAPS : factor));
      run = '';
    }
    run += ch;
    runLower = low;
  }
  if (run !== '')
    out.push(base(runLower ? upper(run) : run, runLower ? factor * SMALL_CAPS : factor));
}

/** Text a piece contributes to the paragraph's logical text. */
export function pieceText(p: Piece): string {
  switch (p.kind) {
    case 'text':
      return p.text;
    case 'tab':
      return '\t';
    case 'break':
      return '\n';
    default:
      return '\ufffc';
  }
}

function objectPiece(
  inline: Extract<IRInline, { kind: 'image' | 'drawing' | 'math' }>,
  index: number,
  env: LayoutEnv,
): ObjectPiece {
  if (inline.kind === 'math') {
    const box = layoutMath(inline.source, inline.props, env);
    return {
      kind: 'object',
      props: inline.props,
      inline: index,
      width: box.width,
      ascent: box.ascent,
      descent: box.descent,
      items: box.items,
    };
  }
  const { width, height } = inline;
  let items: DisplayItem[];
  if (inline.kind === 'image') {
    // EMF/WMF pictures are converted to vector items (renderers cannot embed metafiles).
    const res = env.images.get(inline.image);
    const meta =
      res !== undefined && (res.contentType === 'image/x-emf' || res.contentType === 'image/x-wmf')
        ? metafileImageItems(res, width, height, env)
        : undefined;
    if (meta !== undefined) {
      items = meta.map((it) => translate(it, 0, -height));
    } else {
      const img: DisplayItem = {
        kind: 'image',
        x: 0,
        y: -height,
        width,
        height,
        image: inline.image,
      };
      const crop = inlineExtras(inline).crop;
      if (crop !== undefined) img.crop = crop;
      items = [img];
    }
  } else {
    const drawn = drawingItems(inline.source, width, height, env);
    if (drawn === undefined) {
      env.warn('LAYOUT_DRAWING', `Unsupported inline drawing <${inline.source.name}> left blank`);
      items = [];
    } else {
      items = drawn.map((it) => translate(it, 0, -height));
    }
  }
  const piece: ObjectPiece = {
    kind: 'object',
    props: inline.props,
    inline: index,
    width,
    ascent: height,
    descent: 0,
    items,
  };
  if (inline.kind === 'image' && inline.link !== undefined) piece.link = inline.link;
  return piece;
}

/** A display item moved by (dx, dy). */
export function translate(it: DisplayItem, dx: number, dy: number): DisplayItem {
  switch (it.kind) {
    case 'line':
      return { ...it, x1: it.x1 + dx, y1: it.y1 + dy, x2: it.x2 + dx, y2: it.y2 + dy };
    case 'path':
      return dx === 0 && dy === 0 ? it : { ...it, d: translatePath(it.d, dx, dy) };
    default:
      return { ...it, x: it.x + dx, y: it.y + dy };
  }
}

/**
 * Translates SVG path data made of absolute commands (M L C Q Z, H V, A); relative commands are
 * left as they are (they move with their starting point).
 */
function translatePath(d: string, dx: number, dy: number): string {
  const tokens = d.match(/[A-Za-z]|-?\d*\.?\d+(?:e[-+]?\d+)?/g) ?? [];
  let out = '';
  let cmd = '';
  let k = 0;
  for (const t of tokens) {
    if (/[A-Za-z]/.test(t)) {
      cmd = t;
      k = 0;
      out += `${out === '' ? '' : ' '}${t}`;
      continue;
    }
    let v = Number(t);
    switch (cmd) {
      case 'M':
      case 'L':
      case 'C':
      case 'Q':
      case 'S':
      case 'T':
        v += k % 2 === 0 ? dx : dy;
        break;
      case 'H':
        v += dx;
        break;
      case 'V':
        v += dy;
        break;
      case 'A':
        if (k % 7 === 5) v += dx;
        else if (k % 7 === 6) v += dy;
        break;
    }
    k++;
    out += ` ${Math.round(v * 1000) / 1000}`;
  }
  return out;
}

/** Builds the pieces of a paragraph (label first) and the paragraph's logical text. */
export function buildPieces(p: IRParagraph, env: LayoutEnv): { pieces: Piece[]; text: string } {
  const pieces: Piece[] = [];
  if (p.label !== undefined) {
    pushText(pieces, p.label.text, p.label.props, LABEL, {});
    if (p.label.suffix === 'tab') {
      pieces.push({ kind: 'tab', props: p.label.props, inline: LABEL, listTab: true });
    } else if (p.label.suffix === 'space') {
      pieces.push({
        kind: 'text',
        text: ' ',
        props: p.label.props,
        factor: 1,
        rise: 0,
        inline: LABEL,
        labelSuffix: true,
      });
    }
  }
  let lastProps: ResolvedRunProps = p.props.markRun;
  for (let i = 0; i < p.inlines.length; i++) {
    const inline = p.inlines[i]!;
    switch (inline.kind) {
      case 'text': {
        const extra: { link?: IRLink } = {};
        if (inline.link !== undefined) extra.link = inline.link;
        pushText(pieces, inline.text, inline.props, i, extra);
        lastProps = inline.props;
        break;
      }
      case 'tab': {
        const t: TabPiece = { kind: 'tab', props: inline.props, inline: i, listTab: false };
        const ptab = inlineExtras(inline).ptab;
        if (ptab !== undefined) t.ptab = ptab;
        pieces.push(t);
        lastProps = inline.props;
        break;
      }
      case 'break':
        pieces.push({ kind: 'break', type: inline.type, props: lastProps, inline: i });
        break;
      case 'field':
        pushText(pieces, inline.fallback === '' ? '0' : inline.fallback, inline.props, i, {
          field: inline,
        });
        lastProps = inline.props;
        break;
      case 'noteRef':
        pushText(pieces, inline.text === '' ? inline.id : inline.text, inline.props, i, {
          note: inline.id,
        });
        lastProps = inline.props;
        break;
      default:
        pieces.push(objectPiece(inline, i, env));
        lastProps = inline.props;
    }
  }
  let text = '';
  for (const piece of pieces) text += pieceText(piece);
  return { pieces, text };
}
