// Vertical metrics of fonts and lines, following Word.
//
// Font height. Word (a GDI application at heart) sizes a line from the Windows metrics of each
// font: usWinAscent above the baseline, usWinDescent below, plus the GDI "external leading"
// max(0, hhea.lineGap - ((winAscent + winDescent) - (hhea.ascender - hhea.descender))). That is
// why Calibri 11 pt is 13.43 pt single-spaced (winAscent 1950 + winDescent 550, no extra leading)
// and Times New Roman 12 pt is 13.8 pt (its hhea lineGap of 87 units survives as leading). We
// use these OS/2 win metrics, not hhea or typo metrics, and put the external leading above the
// ascent. Faces without win metrics (some standard-14 tables) fall back to hhea.
//
// Line spacing (`w:spacing/@w:line`, `@w:lineRule`):
// - auto: the text height (largest text ascent + largest text descent on the line) × line/240;
//   inline objects (pictures, charts, drawings, math) are NOT scaled: whatever they rise above
//   the text ascent (or drop below its descent) is added unscaled, as Word and LibreOffice do
//   (two 320 pt charts at 1.08 spacing fit an A4 page; a 128 pt picture line at 1.15 is ~131 pt);
// - exact: exactly line/20 pt, whatever the content;
// - atLeast: max(natural, line/20), natural = largest ascent + largest descent of everything.
// Word puts any extra space *above* the content (a double-spaced first line starts lower) and,
// when the height is smaller than the content, clips the top: the baseline sits the largest
// descent above the bottom of the line.

import type { ResolvedParagraphProps } from '../../resolve/types.js';
import type { Font } from '../types.js';

/** Ascent (incl. external leading) and descent of `font` at `size` points, per Word. */
export function fontExtents(font: Font, size: number): { ascent: number; descent: number } {
  const m = font.metrics;
  const upm = m.unitsPerEm > 0 ? m.unitsPerEm : 1000;
  let a = m.winAscent;
  let d = m.winDescent;
  if (a + d <= 0) {
    a = m.ascender;
    d = -m.descender;
  }
  const hhea = m.ascender - m.descender;
  const leading = Math.max(0, m.lineGap - (a + d - hhea));
  return { ascent: ((a + leading) * size) / upm, descent: (d * size) / upm };
}

/**
 * Height and baseline of a line whose text spans `ascent` above and `descent` below the
 * baseline and whose inline objects span `objectAscent`/`objectDescent` (0 when none).
 */
export function lineBox(
  ascent: number,
  descent: number,
  spacing: ResolvedParagraphProps['spacing'],
  objectAscent = 0,
  objectDescent = 0,
): { height: number; baseline: number } {
  const above = Math.max(0, objectAscent - ascent);
  const below = Math.max(0, objectDescent - descent);
  const bottom = descent + below;
  switch (spacing.lineRule) {
    case 'exact': {
      const height = spacing.line / 20;
      return { height, baseline: height - bottom };
    }
    case 'atLeast': {
      const height = Math.max(ascent + above + bottom, spacing.line / 20);
      return { height, baseline: height - bottom };
    }
    default: {
      const text = ((ascent + descent) * (spacing.line > 0 ? spacing.line : 240)) / 240;
      const height = text + above + below;
      return { height, baseline: height - bottom };
    }
  }
}
