// Fonts for layout (plan decision D7: no font files in core).
//
// Faces come, in order, from the fonts embedded in the docx, the caller's FontResolvers, the
// metric-compatible substitution table and finally the PDF standard-14 faces (metrics only).
//
//   sfnt.ts           loadFont: TrueType/OpenType/TTC parser (reader.ts, cmap.ts, names.ts,
//                     kerning.ts)
//   standard14.ts     standardFont: base-14 metrics (data generated into standard14-data.ts)
//   embedded.ts       embeddedFonts / fontTable: w:embed* parts, odttf deobfuscation
//   substitutions.ts  FONT_SUBSTITUTIONS, SUBSTITUTE_METRICS, generic classes, standard-14 fallbacks
//   substitute.ts     SubstituteFont: a stand-in face reporting the original's vertical metrics
//   slots.ts          Word's w:rFonts slot rules per character
//   manager.ts        createFontManager (the FontManager of engine.ts + synthetic())

export {
  deobfuscateFont,
  type EmbeddedFace,
  type EmbedStyle,
  embeddedFaces,
  embeddedFonts,
  type FontTableEntry,
  fontKeyBytes,
  fontTable,
  obfuscateFont,
} from './embedded.js';
export {
  createFontManager,
  EAST_ASIAN_FALLBACKS,
  type LayoutFontManager,
  type SyntheticStyle,
} from './manager.js';
export { type FontDetails, loadFont, SfntFont } from './sfnt.js';
export { type FontSlot, type SlotProps, slotFor } from './slots.js';
export {
  STANDARD14_NAMES,
  StandardFont,
  standard14For,
  standardFont,
  winAnsiCode,
} from './standard14.js';
export { baseFont, SubstituteFont } from './substitute.js';
export {
  FONT_SUBSTITUTIONS,
  type GenericFamily,
  genericFamily,
  isSymbolFamily,
  STANDARD14_FALLBACKS,
  SUBSTITUTE_METRICS,
  substituteMetricsOf,
  substitutesOf,
  type VerticalMetrics,
} from './substitutions.js';
