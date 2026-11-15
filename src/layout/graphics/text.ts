// Text for graphics (chart labels, metafile text, math runs): builds ResolvedRunProps for text
// that does not come from a w:r, and shapes a string into TextItems through the layout's
// FontManager and Shaper. The string is split where the face changes (fontFor picks a face per
// code point, falling back per glyph), giving one TextItem per face.

import type { ResolvedRunProps } from '../../resolve/types.js';
import type { LayoutEnv } from '../engine.js';
import type { Font, TextItem } from '../types.js';

/** Options for {@link runProps}. */
export interface RunPropsInit {
  /** Font family for every rFonts slot. */
  font?: string;
  /** Points. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  /** RRGGBB. */
  color?: string;
  lang?: string;
}

/** Plain ResolvedRunProps for synthetic text (chart labels, metafile text, math). */
export function runProps(o: RunPropsInit = {}, base?: ResolvedRunProps): ResolvedRunProps {
  const font = o.font;
  const fonts =
    font === undefined
      ? (base?.fonts ?? {})
      : { ascii: font, hAnsi: font, eastAsia: font, cs: font };
  const size = o.size ?? base?.size ?? 10;
  const bold = o.bold ?? base?.bold ?? false;
  const italic = o.italic ?? base?.italic ?? false;
  const props: ResolvedRunProps = {
    fonts: { ...fonts },
    size,
    sizeCs: size,
    bold,
    italic,
    boldCs: bold,
    italicCs: italic,
    strike: false,
    doubleStrike: false,
    color: o.color ?? base?.color ?? '000000',
    caps: false,
    smallCaps: false,
    hidden: false,
    vertAlign: 'baseline',
    spacing: 0,
    position: 0,
    scale: 100,
    kern: 0,
    rtl: false,
    cs: false,
  };
  const lang = o.lang ?? base?.lang;
  if (lang !== undefined) props.lang = lang;
  return props;
}

/** A shaped single line of text, relative to (start x, baseline). */
export interface ShapedText {
  items: TextItem[];
  width: number;
  /** Above the baseline, points (positive). */
  ascent: number;
  /** Below the baseline, points (positive). */
  descent: number;
}

/** Ascent/descent of a face at `size` points. */
export function fontExtents(font: Font, size: number): { ascent: number; descent: number } {
  const upm = font.metrics.unitsPerEm || 1000;
  return {
    ascent: (font.metrics.ascender * size) / upm,
    descent: (-font.metrics.descender * size) / upm,
  };
}

/** RRGGBB for a run colour ('auto' → black). */
function textColor(c: string): string {
  return /^[0-9A-Fa-f]{6}$/.test(c) ? c.toUpperCase() : '000000';
}

/**
 * Shapes `text` (one line, no wrapping) with the run's faces. Items start at x = 0 on the
 * baseline y = 0; `color` overrides props.color.
 */
export function shapeText(
  text: string,
  props: ResolvedRunProps,
  env: Pick<LayoutEnv, 'fonts' | 'shaper'>,
  color?: string,
): ShapedText {
  const size = props.size;
  const items: TextItem[] = [];
  let x = 0;
  let ascent = 0;
  let descent = 0;
  const s = props.caps ? text.toUpperCase() : text;
  // Split into segments of consecutive code points mapped to the same face.
  const segs: { font: Font; text: string }[] = [];
  for (const ch of s) {
    const font = env.fonts.fontFor(props, ch.codePointAt(0) ?? 32);
    const last = segs[segs.length - 1];
    if (last !== undefined && last.font === font) last.text += ch;
    else segs.push({ font, text: ch });
  }
  if (segs.length === 0) {
    const font = env.fonts.fontFor(props, 32);
    const e = fontExtents(font, size);
    return { items, width: 0, ascent: e.ascent, descent: e.descent };
  }
  const spacing = props.spacing / 20;
  const fill = textColor(color ?? props.color);
  for (const seg of segs) {
    const font = seg.font;
    const upm = font.metrics.unitsPerEm || 1000;
    const k = (size * (props.scale / 100)) / upm;
    const glyphs = env.shaper.shape(font, seg.text, { rtl: false, kerning: props.kern > 0 });
    const item: TextItem = {
      kind: 'text',
      x,
      y: 0,
      font: env.fonts.indexOf(font),
      size,
      color: fill,
      glyphs: glyphs.map((g) => g.id),
      advances: glyphs.map((g) => g.advance * k + spacing),
      text: seg.text,
      clusters: glyphs.map((g) => g.cluster),
    };
    if (glyphs.some((g) => g.dy !== 0)) item.dy = glyphs.map((g) => (-g.dy * size) / upm);
    if (props.scale !== 100) item.scale = props.scale;
    if (props.bold && !font.bold) item.fakeBold = true;
    if (props.italic && !font.italic) item.fakeItalic = true;
    for (const a of item.advances) x += a;
    items.push(item);
    const e = fontExtents(font, size);
    if (e.ascent > ascent) ascent = e.ascent;
    if (e.descent > descent) descent = e.descent;
  }
  return { items, width: x, ascent, descent };
}

/** Moves items by (dx, dy) in place and returns them. */
export function offsetItems<T extends { kind: string }>(items: T[], dx: number, dy: number): T[] {
  for (const it of items as unknown as Record<string, unknown>[]) {
    if (dx === 0 && dy === 0) break;
    switch (it.kind) {
      case 'text':
      case 'image':
      case 'rect':
      case 'link':
        it.x = (it.x as number) + dx;
        it.y = (it.y as number) + dy;
        break;
      case 'line':
        it.x1 = (it.x1 as number) + dx;
        it.y1 = (it.y1 as number) + dy;
        it.x2 = (it.x2 as number) + dx;
        it.y2 = (it.y2 as number) + dy;
        break;
      case 'path':
        it.d = translatePathData(it.d as string, dx, dy);
        break;
    }
  }
  return items;
}

/** Translates SVG path data made of absolute M/L/C/Z commands (as produced by Path). */
export function translatePathData(d: string, dx: number, dy: number): string {
  let i = 0;
  return d.replace(/[MLCZ]|-?[\d.]+(?:e-?\d+)?/g, (tok) => {
    if (tok === 'M' || tok === 'L' || tok === 'C' || tok === 'Z') {
      i = 0;
      return tok;
    }
    const v = Number(tok) + (i++ % 2 === 0 ? dx : dy);
    const r = Math.round(v * 1000) / 1000;
    return r === 0 ? '0' : String(r);
  });
}
