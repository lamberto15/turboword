// Which w:rFonts slot (ascii, hAnsi, eastAsia, cs) formats a character.
//
// Word's rules (ECMA-376 Part 1 §17.3.2.26 and the character-range table behind it):
//   - a run with w:cs or w:rtl uses the cs font for every character;
//   - U+0000-U+007F uses ascii;
//   - complex-script letters (Hebrew, Arabic, Syriac, Thaana, N'Ko, Indic, Thai, Lao, Tibetan,
//     Myanmar, Khmer and their presentation forms) use cs;
//   - CJK, kana, Hangul, Yi, full-width forms and the CJK planes use eastAsia;
//   - "ambiguous" ranges (Latin-1 punctuation, Latin extensions, IPA, Greek, Cyrillic, general
//     punctuation, symbols, private use) use hAnsi, or eastAsia when the run has
//     w:rFonts/@w:hint="eastAsia" (for a few Latin letters only when the East Asian language is
//     Chinese), or cs when the hint is "cs";
//   - everything else uses hAnsi.
// The hint is read from an optional `fonts.hint` on the resolved run properties.

import type { ResolvedRunProps } from '../../resolve/types.js';

export type FontSlot = 'ascii' | 'hAnsi' | 'eastAsia' | 'cs';

/** Run properties as far as slot selection is concerned. */
export type SlotProps = Pick<ResolvedRunProps, 'rtl' | 'cs' | 'langEastAsia'> & {
  fonts: ResolvedRunProps['fonts'] & { hint?: string };
};

// Range kinds. Ambiguous: hAnsi, eastAsia with hint eastAsia. AmbiguousZh: likewise, but only
// when the East Asian language is Chinese.
const Kind = { HAnsi: 0, EastAsia: 1, Cs: 2, Ambiguous: 3, AmbiguousZh: 4 } as const;
type Kind = (typeof Kind)[keyof typeof Kind];

// Sorted, non-overlapping [start, end, kind] ranges above U+007F; gaps are hAnsi.
const RANGES: readonly (readonly [number, number, Kind])[] = [
  [0x0100, 0x024f, Kind.AmbiguousZh], // Latin Extended-A/B
  [0x0250, 0x036f, Kind.Ambiguous], // IPA, spacing modifiers, combining diacritics
  [0x0370, 0x03ff, Kind.Ambiguous], // Greek
  [0x0400, 0x04ff, Kind.Ambiguous], // Cyrillic
  [0x0590, 0x08ff, Kind.Cs], // Hebrew, Arabic, Syriac, Thaana, N'Ko, Samaritan, Mandaic, Arabic ext.
  [0x0900, 0x0dff, Kind.Cs], // Indic scripts, Sinhala
  [0x0e00, 0x0fff, Kind.Cs], // Thai, Lao, Tibetan
  [0x1000, 0x109f, Kind.Cs], // Myanmar
  [0x1100, 0x11ff, Kind.EastAsia], // Hangul Jamo
  [0x1780, 0x17ff, Kind.Cs], // Khmer
  [0x19e0, 0x19ff, Kind.Cs], // Khmer symbols
  [0x1e00, 0x1eff, Kind.AmbiguousZh], // Latin Extended Additional
  [0x1f00, 0x1fff, Kind.Ambiguous], // Greek Extended
  [0x2000, 0x27bf, Kind.Ambiguous], // punctuation, currency, letterlike, arrows, math, dingbats
  [0x2e80, 0x2fff, Kind.EastAsia], // CJK radicals, Kangxi, IDC
  [0x3000, 0x33ff, Kind.EastAsia], // CJK punctuation, kana, Bopomofo, Hangul compat, enclosed
  [0x3400, 0x4dbf, Kind.EastAsia], // CJK Extension A
  [0x4e00, 0x9fff, Kind.EastAsia], // CJK Unified Ideographs
  [0xa000, 0xa4cf, Kind.EastAsia], // Yi
  [0xa960, 0xa97f, Kind.EastAsia], // Hangul Jamo Extended-A
  [0xa980, 0xa9df, Kind.Cs], // Javanese
  [0xac00, 0xd7ff, Kind.EastAsia], // Hangul syllables, Jamo Extended-B
  [0xe000, 0xf8ff, Kind.Ambiguous], // private use
  [0xf900, 0xfaff, Kind.EastAsia], // CJK compatibility ideographs
  [0xfb00, 0xfb1c, Kind.Ambiguous], // Latin/Armenian ligatures
  [0xfb1d, 0xfdff, Kind.Cs], // Hebrew and Arabic presentation forms A
  [0xfe10, 0xfe1f, Kind.EastAsia], // vertical forms
  [0xfe30, 0xfe6f, Kind.EastAsia], // CJK compatibility forms, small forms
  [0xfe70, 0xfeff, Kind.Cs], // Arabic presentation forms B
  [0xff00, 0xffef, Kind.EastAsia], // half-width and full-width forms
  [0x10800, 0x10fff, Kind.Cs], // historic right-to-left scripts
  [0x16fe0, 0x18aff, Kind.EastAsia], // Tangut, ideographic symbols
  [0x1b000, 0x1b2ff, Kind.EastAsia], // kana supplement/extended, Nushu
  [0x1e800, 0x1efff, Kind.Cs], // Mende Kikakui, Adlam, Arabic mathematical
  [0x1f200, 0x1f2ff, Kind.EastAsia], // enclosed ideographic supplement
  [0x20000, 0x3ffff, Kind.EastAsia], // CJK Extensions B-H, compatibility supplement
];

/** Latin-1 characters that follow the East Asian font under hint="eastAsia". */
const LATIN1_EAST_ASIA = new Set([
  0xa1, 0xa4, 0xa7, 0xa8, 0xaa, 0xad, 0xaf, 0xb0, 0xb1, 0xb2, 0xb3, 0xb4, 0xb6, 0xb7, 0xb8, 0xb9,
  0xba, 0xbc, 0xbd, 0xbe, 0xbf, 0xd7, 0xf7,
]);
/** Latin-1 letters that follow it only when the East Asian language is Chinese. */
const LATIN1_EAST_ASIA_ZH = new Set([
  0xe0, 0xe1, 0xe8, 0xe9, 0xea, 0xec, 0xed, 0xf2, 0xf3, 0xf9, 0xfa, 0xfc,
]);

function kindOf(cp: number): Kind {
  if (cp >= 0xa0 && cp <= 0xff) {
    if (LATIN1_EAST_ASIA.has(cp)) return Kind.Ambiguous;
    if (LATIN1_EAST_ASIA_ZH.has(cp)) return Kind.AmbiguousZh;
    return Kind.HAnsi;
  }
  let lo = 0;
  let hi = RANGES.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const r = RANGES[mid]!;
    if (cp < r[0]) hi = mid - 1;
    else if (cp > r[1]) lo = mid + 1;
    else return r[2];
  }
  return Kind.HAnsi;
}

/** The rFonts slot that formats `codePoint` in a run with properties `props`. */
export function slotFor(props: SlotProps, codePoint: number): FontSlot {
  if (props.cs || props.rtl) return 'cs';
  if (codePoint <= 0x7f) return 'ascii';
  const kind = kindOf(codePoint);
  switch (kind) {
    case Kind.EastAsia:
      return 'eastAsia';
    case Kind.Cs:
      return 'cs';
    case Kind.HAnsi:
      return 'hAnsi';
    default: {
      const hint = props.fonts.hint;
      if (hint === 'cs') return 'cs';
      if (hint !== 'eastAsia') return 'hAnsi';
      if (kind === Kind.Ambiguous) return 'eastAsia';
      return /^zh\b/i.test(props.langEastAsia ?? '') ? 'eastAsia' : 'hAnsi';
    }
  }
}
