// The built-in shaper: one glyph per character from the cmap, advances from hmtx, optional pair
// kerning (the font's kern table / GPOS pairs via Font.kerning). No ligatures, contextual forms
// or mark positioning: complex scripts need the HarfBuzz plugin, which implements the same
// Shaper interface.
//
// Glyphs stay in logical order; the line layout reorders bidi runs. What this shaper does beyond
// a bare cmap lookup is what Word users visibly depend on:
// - default-ignorable characters (ZWJ, ZWNJ, soft hyphen, bidi controls, variation selectors) and
//   C0/C1 controls produce no glyph, so they never print as .notdef boxes;
// - in right-to-left runs, Bidi_Mirrored characters use their mirror glyph (UAX #9 L4);
// - characters the face lacks get the fallbacks HarfBuzz applies: the special spaces become the
//   space glyph with their nominal width, the non-breaking hyphen and U+2010 become '-'.
// Standard-14 faces work unchanged: their glyphId returns the WinAnsi code.

import type { Font, ShapedGlyph, ShapeOptions, Shaper } from '../types.js';
import { BIDI_NSM, bidiClassId, mirrorOf } from './bidi.js';
import { isDefaultIgnorable } from './script.js';

/** Width of spaces the face may lack, as a fraction of the em (0 = the face's space width). */
const SPACE_WIDTHS: ReadonlyMap<number, number> = /* @__PURE__ */ new Map([
  [0x00a0, 0], // no-break space
  [0x1680, 0], // ogham space mark
  [0x2000, 1 / 2], // en quad
  [0x2001, 1], // em quad
  [0x2002, 1 / 2], // en space
  [0x2003, 1], // em space
  [0x2004, 1 / 3], // three-per-em space
  [0x2005, 1 / 4], // four-per-em space
  [0x2006, 1 / 6], // six-per-em space
  [0x2007, -1], // figure space: width of the digit zero
  [0x2008, -2], // punctuation space: width of the full stop
  [0x2009, 1 / 5], // thin space
  [0x200a, 1 / 16], // hair space
  [0x202f, -3], // narrow no-break space: half the space
  [0x205f, 4 / 18], // medium mathematical space
  [0x3000, 1], // ideographic space
]);

/** Hyphens the face may lack, with the characters tried instead. */
const HYPHENS: ReadonlyMap<number, readonly number[]> = /* @__PURE__ */ new Map([
  [0x2010, [0x2d]], // hyphen
  [0x2011, [0x2010, 0x2d]], // non-breaking hyphen (w:noBreakHyphen)
  [0x2212, [0x2d]], // minus sign
]);

/** Glyph id and advance for a code point the face has no glyph for. */
function fallback(font: Font, cp: number): { id: number; advance: number } | undefined {
  const width = SPACE_WIDTHS.get(cp);
  if (width !== undefined) {
    const space = font.glyphId(0x20);
    if (space === 0) return undefined;
    let advance = font.advance(space);
    if (width > 0) advance = Math.round(width * font.metrics.unitsPerEm);
    else if (width === -1) advance = font.advance(font.glyphId(0x30)) || advance;
    else if (width === -2) advance = font.advance(font.glyphId(0x2e)) || advance;
    else if (width === -3) advance = Math.round(advance / 2);
    return { id: space, advance };
  }
  for (const alt of HYPHENS.get(cp) ?? []) {
    const id = font.glyphId(alt);
    if (id !== 0) return { id, advance: font.advance(id) };
  }
  return undefined;
}

/** True for characters that never produce a glyph. */
function invisible(cp: number): boolean {
  return cp < 0x20 || (cp >= 0x7f && cp < 0xa0) || isDefaultIgnorable(cp);
}

/** cmap + pair-kerning shaper (no ligatures/contextual forms; see the HarfBuzz plugin for that). */
export const basicShaper: Shaper = {
  shape(font: Font, text: string, o: ShapeOptions): ShapedGlyph[] {
    const glyphs: ShapedGlyph[] = [];
    const marks: boolean[] = [];
    for (let i = 0; i < text.length; ) {
      const cp = text.codePointAt(i)!;
      const len = cp > 0xffff ? 2 : 1;
      if (!invisible(cp)) {
        let ch = cp;
        if (o.rtl) {
          const m = mirrorOf(cp);
          if (m !== undefined && font.glyphId(m) !== 0) ch = m;
        }
        let id = font.glyphId(ch);
        let advance: number;
        const alt = id === 0 ? fallback(font, ch) : undefined;
        if (alt) {
          id = alt.id;
          advance = alt.advance;
        } else advance = font.advance(id);
        glyphs.push({ id, advance, dx: 0, dy: 0, cluster: i });
        marks.push(bidiClassId(cp) === BIDI_NSM);
      }
      i += len;
    }
    if (o.kerning) kern(font, glyphs, marks, o.rtl);
    return glyphs;
  },
};

/**
 * Pair kerning between consecutive base glyphs (marks are skipped, so a kerned pair keeps its
 * marks on their base). In logical order the adjustment goes on the glyph just before the right
 * glyph of the visual pair: for LTR the last glyph (base or mark) before the next base; for RTL
 * the next base itself, which sits visually left of the current base and its marks.
 */
function kern(font: Font, glyphs: ShapedGlyph[], marks: boolean[], rtl: boolean): void {
  let prev = -1;
  for (let k = 0; k < glyphs.length; k++) {
    if (marks[k]) continue;
    if (prev >= 0) {
      const left = rtl ? glyphs[k]! : glyphs[prev]!;
      const right = rtl ? glyphs[prev]! : glyphs[k]!;
      const adjust = font.kerning(left.id, right.id);
      if (adjust !== 0) glyphs[rtl ? k : k - 1]!.advance += adjust;
    }
    prev = k;
  }
}
