// Pieces → shaped glyph runs → atoms.
//
// Itemization: each text piece is split wherever the face (FontManager.fontFor, which applies the
// rFonts slot rules and per-character fallback), the script (scriptOf, with Common/Inherited
// characters joining their neighbours) or the bidi level changes; each item is shaped once. The
// complex-script size/flags (`w:szCs`) apply to items in complex scripts or runs marked cs/rtl.
//
// Atoms are what line breaking and positioning work on: one per glyph cluster (a character, or a
// ligature/combining sequence), one per tab, break and inline object. Widths are in points and
// include the run's character spacing (`w:spacing`) and horizontal scale (`w:w`).

import type { ResolvedRunProps } from '../../resolve/types.js';
import type { LayoutEnv } from '../engine.js';
import { scriptOf } from '../text/index.js';
import type { Font, IRLink } from '../types.js';
import type { Piece, TextPiece } from './pieces.js';

/** A shaped item: glyphs of one face/size/style, in logical order. */
export interface GlyphRun {
  font: Font;
  /** Font size, points. */
  size: number;
  /** Horizontal scale percent. */
  scale: number;
  /** Baseline shift upwards, points. */
  rise: number;
  color: string;
  piece: number;
  rtl: boolean;
  /** True when the shaper returned visual order and the arrays were reversed to logical order. */
  reversed: boolean;
  glyphs: number[];
  /** Advance per glyph, points (spacing and scale included). */
  advances: number[];
  fakeBold: boolean;
  fakeItalic: boolean;
}

export type AtomKind = 'glyphs' | 'tab' | 'break' | 'object';

export interface Atom {
  kind: AtomKind;
  /** UTF-16 range in the paragraph text. */
  start: number;
  end: number;
  piece: number;
  /** Natural width, points (0 for tabs and breaks: tabs are sized during positioning). */
  width: number;
  /** Bidi embedding level. */
  level: number;
  run?: GlyphRun;
  /** Glyph range of the cluster in `run`. */
  g0: number;
  g1: number;
  /** A U+0020 space: a justification opportunity. */
  space: boolean;
  /** Whitespace that hangs past the line end. */
  white: boolean;
  /** A soft hyphen (shown only when the line breaks after it). */
  soft: boolean;
  /** Kind of break of a break atom. */
  breakType?: 'line' | 'page' | 'column';
}

/** Scripts whose text uses the complex-script formatting (`w:szCs`, `w:bCs`, `w:iCs`). */
const CS_SCRIPTS: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'Arab',
  'Hebr',
  'Syrc',
  'Thaa',
  'Nkoo',
  'Thai',
  'Laoo',
  'Khmr',
  'Deva',
  'Beng',
  'Guru',
  'Gujr',
  'Orya',
  'Taml',
  'Telu',
  'Knda',
  'Mlym',
  'Sinh',
  'Tibt',
  'Mymr',
]);

const NEUTRAL_SCRIPTS: ReadonlySet<string> = /* @__PURE__ */ new Set(['Zyyy', 'Zinh', 'Zzzz', '']);

/** Characters that take the face of the character before them (never trigger font fallback). */
function inheritsFont(cp: number): boolean {
  return (
    cp === 0xad || cp === 0x200b || cp === 0x200c || cp === 0x200d || cp === 0x2060 || cp === 0xfeff
  );
}

/** Effective point size of a text piece for an item in `script`. */
function sizeFor(piece: TextPiece, cs: boolean): number {
  return (cs ? piece.props.sizeCs || piece.props.size : piece.props.size) * piece.factor;
}

/** Text colour: `auto` is black, or white on a dark shading. */
export function textColor(props: ResolvedRunProps): string {
  if (props.color !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(props.color))
    return props.color.toUpperCase();
  const fill = props.shading;
  if (fill !== undefined && /^[0-9A-Fa-f]{6}$/.test(fill)) {
    const r = Number.parseInt(fill.slice(0, 2), 16);
    const g = Number.parseInt(fill.slice(2, 4), 16);
    const b = Number.parseInt(fill.slice(4, 6), 16);
    if (0.299 * r + 0.587 * g + 0.114 * b < 128) return 'FFFFFF';
  }
  return '000000';
}

interface Item {
  start: number;
  end: number;
  font: Font;
  script: string;
  level: number;
  cs: boolean;
}

/** Splits text piece `p` (at `offset` in the paragraph text) into items. */
function itemize(
  p: TextPiece,
  offset: number,
  levels: Uint8Array | undefined,
  env: LayoutEnv,
): Item[] {
  const items: Item[] = [];
  const text = p.text;
  // Scripts per code point, neutral ones resolved to their neighbours.
  const cps: { i: number; cp: number; script: string }[] = [];
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    cps.push({ i, cp, script: scriptOf(cp) });
    i += cp > 0xffff ? 2 : 1;
  }
  let prev = '';
  for (const c of cps) {
    if (NEUTRAL_SCRIPTS.has(c.script)) c.script = prev;
    else prev = c.script;
  }
  let next = 'Latn';
  for (let k = cps.length - 1; k >= 0; k--) {
    const c = cps[k]!;
    if (c.script === '') c.script = next;
    else next = c.script;
  }
  const forcedCs = p.props.cs || p.props.rtl;
  let cur: Item | undefined;
  for (const c of cps) {
    const level = levels === undefined ? 0 : levels[offset + c.i]!;
    const font =
      cur !== undefined && inheritsFont(c.cp) ? cur.font : env.fonts.fontFor(p.props, c.cp);
    const cs = forcedCs || CS_SCRIPTS.has(c.script);
    const len = c.cp > 0xffff ? 2 : 1;
    if (
      cur !== undefined &&
      cur.font === font &&
      cur.script === c.script &&
      cur.level === level &&
      cur.cs === cs
    ) {
      cur.end = c.i + len;
      continue;
    }
    cur = { start: c.i, end: c.i + len, font, script: c.script, level, cs };
    items.push(cur);
  }
  return items;
}

/**
 * Whether the renderer must embolden/slant `font` for `props`: the font manager's answer when it
 * offers one (LayoutFontManager.synthetic), else "the run asks for a style the face lacks".
 */
function syntheticStyle(
  env: LayoutEnv,
  props: ResolvedRunProps,
  font: Font,
  cp: number,
  cs: boolean,
): { bold: boolean; italic: boolean } {
  const fm = env.fonts as LayoutEnv['fonts'] & {
    synthetic?: (p: ResolvedRunProps, f: Font, c?: number) => { bold: boolean; italic: boolean };
  };
  if (typeof fm.synthetic === 'function') return fm.synthetic(props, font, cp);
  const bold = cs ? props.boldCs : props.bold;
  const italic = cs ? props.italicCs : props.italic;
  return { bold: bold && !font.bold, italic: italic && !font.italic };
}

/** Text handed to the shaper: characters the face lacks are replaced by look-alikes of the same length. */
function shapingText(s: string, font: Font): string {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c === 0x2011 && font.glyphId(0x2011) === 0) out += '-';
    else if (c === 0xa0 && font.glyphId(0xa0) === 0) out += ' ';
    else if (c === 0x202f && font.glyphId(0x202f) === 0) out += ' ';
    else out += s[i];
  }
  return out;
}

function shapeText(
  p: TextPiece,
  pieceIndex: number,
  offset: number,
  levels: Uint8Array | undefined,
  env: LayoutEnv,
  atoms: Atom[],
): void {
  const spacing = p.props.spacing / 20;
  for (const item of itemize(p, offset, levels, env)) {
    const size = sizeFor(p, item.cs);
    const rtl = item.level % 2 === 1;
    const sub = p.text.slice(item.start, item.end);
    const kerning = p.props.kern > 0 && size >= p.props.kern;
    const lang = rtl ? (p.props.langBidi ?? p.props.lang) : p.props.lang;
    const o: { rtl: boolean; kerning: boolean; script?: string; lang?: string } = { rtl, kerning };
    if (item.script !== '') o.script = item.script;
    if (lang !== undefined) o.lang = lang;
    const shaped = env.shaper.shape(item.font, shapingText(sub, item.font), o);
    let reversed = false;
    if (shaped.length > 1 && shaped[0]!.cluster > shaped[shaped.length - 1]!.cluster) {
      shaped.reverse();
      reversed = true;
    }
    const upm = item.font.metrics.unitsPerEm > 0 ? item.font.metrics.unitsPerEm : 1000;
    const scale = p.props.scale > 0 ? p.props.scale : 100;
    const k = (size / upm) * (scale / 100);
    const synth = syntheticStyle(env, p.props, item.font, sub.codePointAt(0) ?? 0x20, item.cs);
    const run: GlyphRun = {
      font: item.font,
      size,
      scale,
      rise: p.rise,
      color: textColor(p.props),
      piece: pieceIndex,
      rtl,
      reversed,
      glyphs: shaped.map((g) => g.id),
      advances: shaped.map((g) => g.advance * k),
      fakeBold: synth.bold,
      fakeItalic: synth.italic,
    };
    // Group glyphs into clusters (cluster values are indices into `sub`).
    let g = 0;
    const n = shaped.length;
    if (n === 0) continue;
    while (g < n) {
      const c = shaped[g]!.cluster;
      let h = g + 1;
      while (h < n && shaped[h]!.cluster === c) h++;
      const cStart = g === 0 ? 0 : c;
      const cEnd = h < n ? shaped[h]!.cluster : sub.length;
      const chars = sub.slice(cStart, Math.max(cStart, cEnd));
      const soft = chars.endsWith('\u00ad');
      let width = 0;
      for (let q = g; q < h; q++) width += run.advances[q]!;
      if (chars === '\u00ad') {
        // A shaper that draws the soft hyphen: it only shows at a line end (emit.ts).
        for (let q = g; q < h; q++) run.advances[q] = 0;
        width = 0;
      } else if (spacing !== 0) {
        run.advances[h - 1]! += spacing;
        width += spacing;
      }
      const start = offset + item.start + cStart;
      atoms.push({
        kind: 'glyphs',
        start,
        end: offset + item.start + Math.max(cStart, cEnd),
        piece: pieceIndex,
        width,
        level: item.level,
        run,
        g0: g,
        g1: h,
        space: chars === ' ',
        white: chars === ' ' || chars === '\u3000',
        soft,
      });
      g = h;
    }
  }
}

/** Shapes every piece and returns the atoms in logical order. */
export function buildAtoms(
  pieces: readonly Piece[],
  levels: Uint8Array | undefined,
  env: LayoutEnv,
): Atom[] {
  const atoms: Atom[] = [];
  let offset = 0;
  for (let i = 0; i < pieces.length; i++) {
    const p = pieces[i]!;
    const level = levels === undefined ? 0 : levels[offset]!;
    const base = {
      start: offset,
      piece: i,
      level,
      g0: 0,
      g1: 0,
      space: false,
      white: false,
      soft: false,
    };
    switch (p.kind) {
      case 'text':
        shapeText(p, i, offset, levels, env, atoms);
        offset += p.text.length;
        continue;
      case 'tab':
        atoms.push({ ...base, kind: 'tab', end: offset + 1, width: 0 });
        break;
      case 'break':
        atoms.push({ ...base, kind: 'break', end: offset + 1, width: 0, breakType: p.type });
        break;
      default:
        atoms.push({ ...base, kind: 'object', end: offset + 1, width: p.width });
    }
    offset += 1;
  }
  return atoms;
}

/** The hyperlink of a piece, if any. */
export function linkOf(p: Piece): IRLink | undefined {
  return p.kind === 'text' || p.kind === 'object' ? p.link : undefined;
}
