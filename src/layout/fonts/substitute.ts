// A face standing in for another family, reporting that family's vertical metrics.
//
// Metric-compatible substitutes match their originals' advance widths, but not always their
// vertical metrics (Caladea for Cambria, Gelasio for Georgia), and line height in Word comes from
// the vertical metrics. SubstituteFont wraps the face actually used and replaces its ascender,
// descender, line gap, win metrics and underline values with the original family's (from
// SUBSTITUTE_METRICS), scaled to the face's unitsPerEm. The font manager only wraps faces whose
// line metrics differ from the original's. Glyphs, advances, kerning and bytes are the wrapped
// face's: renderers embed `base` / baseFont(font) (several wrappers may share one base).

import type { Font, FontMetrics } from '../types.js';
import type { FontDetails } from './sfnt.js';
import type { VerticalMetrics } from './substitutions.js';

/** A face reporting the vertical metrics of the family it substitutes for. */
export class SubstituteFont implements Font, FontDetails {
  /** The face whose glyphs, advances and bytes are used. */
  readonly base: Font;
  /** The family it stands in for (whose vertical metrics it reports). */
  readonly substituteFor: string;
  readonly family: string;
  readonly subfamily: string;
  readonly postscriptName: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly metrics: FontMetrics;
  readonly outlines: Font['outlines'];
  readonly bytes: Uint8Array;
  readonly numGlyphs: number;
  readonly bbox: readonly [number, number, number, number];
  readonly weight: number;
  readonly fixedPitch: boolean;
  readonly symbolic: boolean;
  readonly fsType: number;
  readonly typographicFamily: string;
  readonly fullName: string;
  readonly faceIndex: number;

  constructor(base: Font, substituteFor: string, original: VerticalMetrics) {
    this.base = base;
    this.substituteFor = substituteFor;
    this.family = base.family;
    this.subfamily = base.subfamily;
    this.postscriptName = base.postscriptName;
    this.bold = base.bold;
    this.italic = base.italic;
    this.outlines = base.outlines;
    this.bytes = base.bytes;
    this.numGlyphs = base.numGlyphs;
    const m = base.metrics;
    const k = m.unitsPerEm / original.unitsPerEm;
    this.metrics = {
      ...m,
      ascender: original.ascender * k,
      descender: original.descender * k,
      lineGap: original.lineGap * k,
      winAscent: original.winAscent * k,
      winDescent: original.winDescent * k,
      underlinePosition:
        original.underlinePosition === undefined
          ? m.underlinePosition
          : original.underlinePosition * k,
      underlineThickness:
        original.underlineThickness === undefined
          ? m.underlineThickness
          : original.underlineThickness * k,
    };
    const d = base as Partial<FontDetails>;
    this.bbox = d.bbox ?? [0, m.descender, m.unitsPerEm, m.ascender];
    this.weight = d.weight ?? (base.bold ? 700 : 400);
    this.fixedPitch = d.fixedPitch ?? false;
    this.symbolic = d.symbolic ?? false;
    this.fsType = d.fsType ?? 0;
    this.typographicFamily = d.typographicFamily ?? base.family;
    this.fullName = d.fullName ?? `${base.family} ${base.subfamily}`;
    this.faceIndex = d.faceIndex ?? 0;
  }

  glyphId(codePoint: number): number {
    return this.base.glyphId(codePoint);
  }

  advance(glyphId: number): number {
    return this.base.advance(glyphId);
  }

  kerning(left: number, right: number): number {
    return this.base.kerning(left, right);
  }

  /** As SfntFont.hasTable (false for standard-14 bases). */
  hasTable(tag: string): boolean {
    return (this.base as { hasTable?: (t: string) => boolean }).hasTable?.(tag) ?? false;
  }

  /** As SfntFont.tableRange. */
  tableRange(tag: string): { offset: number; length: number } | undefined {
    const b = this.base as {
      tableRange?: (t: string) => { offset: number; length: number } | undefined;
    };
    return b.tableRange?.(tag);
  }
}

/** The face renderers should embed for `font` (unwraps SubstituteFont). */
export function baseFont(font: Font): Font {
  return font instanceof SubstituteFont ? font.base : font;
}
