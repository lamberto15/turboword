// Theme fonts and colours (DrawingML theme part + the settings that steer it).
//
// Runs can name fonts and colours indirectly: `w:rFonts/@w:asciiTheme="minorHAnsi"` means "the
// theme's minor Latin font", `w:color/@w:themeColor="accent1"` + `w:themeShade="BF"` means "the
// theme's accent 1, darkened". Two settings take part:
// - `w:themeFontLang` picks the theme's per-script font (`<a:font script="Jpan">`) for the East
//   Asian and complex-script slots when the theme leaves `a:ea`/`a:cs` empty or a script font
//   exists for the document's language;
// - `w:clrSchemeMapping` maps the logical colours (text1, background1, ...) to theme slots.
// Tints and shades are applied to the HSL luminance, as Office does.

import type { XDocument, XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { ThemeInfo } from './types.js';
import { hexColor } from './values.js';

/** One font scheme (major or minor) of a theme. */
export interface FontScheme {
  latin?: string;
  ea?: string;
  cs?: string;
  /** `a:font/@script` → typeface. */
  scripts: Map<string, string>;
}

/** `w:themeFontLang`: languages selecting the theme's script fonts. */
export interface ThemeFontLang {
  val?: string;
  eastAsia?: string;
  bidi?: string;
}

/** Parsed theme part. */
export interface ThemeData {
  major: FontScheme;
  minor: FontScheme;
  /** Theme slot (dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink) → RRGGBB. */
  colors: Map<string, string>;
}

/** Settings that influence formatting resolution. */
export interface SettingsData {
  themeFontLang: ThemeFontLang;
  /** `w:clrSchemeMapping` attributes (bg1, t1, ..., accent1, hyperlink) → logical name. */
  colorMapping: Map<string, string>;
}

const EMPTY_SCHEME = (): FontScheme => ({ scripts: new Map() });

function nonEmpty(v: string | undefined): string | undefined {
  return v === undefined || v.trim() === '' ? undefined : v;
}

function readScheme(el: XElement | undefined): FontScheme {
  const out = EMPTY_SCHEME();
  if (el === undefined) return out;
  const latin = nonEmpty(el.element(NS.a, 'latin')?.attr('typeface'));
  const ea = nonEmpty(el.element(NS.a, 'ea')?.attr('typeface'));
  const cs = nonEmpty(el.element(NS.a, 'cs')?.attr('typeface'));
  if (latin !== undefined) out.latin = latin;
  if (ea !== undefined) out.ea = ea;
  if (cs !== undefined) out.cs = cs;
  for (const f of el.elements(NS.a, 'font')) {
    const script = f.attr('script');
    const face = nonEmpty(f.attr('typeface'));
    if (script !== undefined && face !== undefined) out.scripts.set(script, face);
  }
  return out;
}

/** RRGGBB of a DrawingML colour choice element (srgbClr, sysClr, scrgbClr). */
function drawingColor(el: XElement | undefined): string | undefined {
  if (el === undefined) return undefined;
  for (const c of el.elements(NS.a)) {
    switch (c.local) {
      case 'srgbClr':
        return hexColor(c.attr('val'));
      case 'sysClr': {
        const last = hexColor(c.attr('lastClr'));
        if (last !== undefined && last !== 'auto') return last;
        const v = c.attr('val');
        return v === 'window' || v === 'highlight' ? 'FFFFFF' : '000000';
      }
      case 'scrgbClr': {
        const ch = (n: string): string => {
          const v = Number(c.attr(n) ?? 0) / 100000;
          const lin = Math.min(1, Math.max(0, Number.isFinite(v) ? v : 0));
          // scRGB is linear; convert to sRGB.
          const s = lin <= 0.0031308 ? lin * 12.92 : 1.055 * lin ** (1 / 2.4) - 0.055;
          return Math.round(s * 255)
            .toString(16)
            .padStart(2, '0');
        };
        return `${ch('r')}${ch('g')}${ch('b')}`.toUpperCase();
      }
      default:
    }
  }
  return undefined;
}

const COLOR_SLOTS = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

/** Parses a theme part (empty data when there is none). */
export function parseTheme(doc: XDocument | undefined): ThemeData {
  const out: ThemeData = { major: EMPTY_SCHEME(), minor: EMPTY_SCHEME(), colors: new Map() };
  if (doc === undefined) return out;
  let root: XElement;
  try {
    root = doc.root;
  } catch {
    return out;
  }
  const elements = root.element(NS.a, 'themeElements');
  const scheme = elements?.element(NS.a, 'clrScheme');
  if (scheme !== undefined) {
    for (const slot of COLOR_SLOTS) {
      const c = drawingColor(scheme.element(NS.a, slot));
      if (c !== undefined) out.colors.set(slot, c);
    }
  }
  const fonts = elements?.element(NS.a, 'fontScheme');
  out.major = readScheme(fonts?.element(NS.a, 'majorFont'));
  out.minor = readScheme(fonts?.element(NS.a, 'minorFont'));
  return out;
}

/** Parses the settings that matter for resolution. */
export function parseSettings(doc: XDocument | undefined): SettingsData {
  const out: SettingsData = { themeFontLang: {}, colorMapping: new Map() };
  if (doc === undefined) return out;
  let root: XElement;
  try {
    root = doc.root;
  } catch {
    return out;
  }
  const lang = root.element(NS.w, 'themeFontLang');
  if (lang !== undefined) {
    const val = lang.attrNS(NS.w, 'val');
    const ea = lang.attrNS(NS.w, 'eastAsia');
    const bidi = lang.attrNS(NS.w, 'bidi');
    if (val !== undefined) out.themeFontLang.val = val;
    if (ea !== undefined) out.themeFontLang.eastAsia = ea;
    if (bidi !== undefined) out.themeFontLang.bidi = bidi;
  }
  const map = root.element(NS.w, 'clrSchemeMapping');
  if (map !== undefined) {
    for (const a of map.attributes()) {
      if (a.uri === NS.w) out.colorMapping.set(a.name.slice(a.name.indexOf(':') + 1), a.value);
    }
  }
  return out;
}

// ---- fonts ----------------------------------------------------------------------------------

/** Language (BCP 47 primary subtag, or full tag for Chinese) → theme script tag. */
const SCRIPT_OF_LANG: Readonly<Record<string, string>> = {
  ja: 'Jpan',
  ko: 'Hang',
  'zh-cn': 'Hans',
  'zh-sg': 'Hans',
  'zh-hans': 'Hans',
  'zh-tw': 'Hant',
  'zh-hk': 'Hant',
  'zh-mo': 'Hant',
  'zh-hant': 'Hant',
  zh: 'Hans',
  ar: 'Arab',
  fa: 'Arab',
  ur: 'Arab',
  ps: 'Arab',
  sd: 'Arab',
  ug: 'Uigh',
  he: 'Hebr',
  yi: 'Hebr',
  th: 'Thai',
  hi: 'Deva',
  mr: 'Deva',
  sa: 'Deva',
  ne: 'Deva',
  kok: 'Deva',
  bn: 'Beng',
  as: 'Beng',
  gu: 'Gujr',
  pa: 'Guru',
  ta: 'Taml',
  te: 'Telu',
  kn: 'Knda',
  ml: 'Mlym',
  or: 'Orya',
  si: 'Sinh',
  km: 'Khmr',
  lo: 'Laoo',
  my: 'Mymr',
  bo: 'Tibt',
  mn: 'Mong',
  am: 'Ethi',
  ti: 'Ethi',
  syr: 'Syrc',
  dv: 'Thaa',
  iu: 'Cans',
  chr: 'Cher',
  vi: 'Viet',
  ka: 'Geor',
  hy: 'Armn',
};

/** Theme script tag for a language tag (`ja-JP` → `Jpan`), or undefined. */
export function scriptOfLang(lang: string | undefined): string | undefined {
  if (lang === undefined) return undefined;
  const l = lang.toLowerCase();
  return SCRIPT_OF_LANG[l] ?? SCRIPT_OF_LANG[l.split('-')[0]!];
}

/** The effective typeface of one slot of a font scheme. */
export function schemeFont(
  scheme: FontScheme,
  slot: 'latin' | 'ea' | 'cs',
  lang: ThemeFontLang,
): string | undefined {
  if (slot === 'latin') {
    return scheme.latin ?? scheme.scripts.get(scriptOfLang(lang.val) ?? '');
  }
  const script = scriptOfLang(slot === 'ea' ? lang.eastAsia : lang.bidi);
  const byScript = script === undefined ? undefined : scheme.scripts.get(script);
  return byScript ?? (slot === 'ea' ? scheme.ea : scheme.cs);
}

/** Resolves an ST_Theme font reference (majorAscii, minorHAnsi, majorEastAsia, minorBidi, ...). */
export function themeFont(
  theme: ThemeData,
  settings: SettingsData,
  ref: string,
): string | undefined {
  const scheme = ref.startsWith('major')
    ? theme.major
    : ref.startsWith('minor')
      ? theme.minor
      : undefined;
  if (scheme === undefined) return undefined;
  const rest = ref.slice(5);
  const slot = rest === 'EastAsia' ? 'ea' : rest === 'Bidi' ? 'cs' : 'latin';
  return schemeFont(scheme, slot, settings.themeFontLang);
}

// ---- colours --------------------------------------------------------------------------------

/** ST_ThemeColor (WordprocessingML) → theme slot name, before the colour mapping. */
const WML_THEME_SLOT: Readonly<Record<string, string>> = {
  dark1: 'dk1',
  light1: 'lt1',
  dark2: 'dk2',
  light2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hyperlink: 'hlink',
  followedHyperlink: 'folHlink',
};

/** Logical colour → (clrSchemeMapping attribute, default target). */
const MAPPED: Readonly<Record<string, readonly [string, string]>> = {
  text1: ['t1', 'dark1'],
  background1: ['bg1', 'light1'],
  text2: ['t2', 'dark2'],
  background2: ['bg2', 'light2'],
  accent1: ['accent1', 'accent1'],
  accent2: ['accent2', 'accent2'],
  accent3: ['accent3', 'accent3'],
  accent4: ['accent4', 'accent4'],
  accent5: ['accent5', 'accent5'],
  accent6: ['accent6', 'accent6'],
  hyperlink: ['hyperlink', 'hyperlink'],
  followedHyperlink: ['followedHyperlink', 'followedHyperlink'],
};

/** Theme slot for a `w:themeColor` value, following `w:clrSchemeMapping`. */
export function themeSlot(settings: SettingsData, name: string): string | undefined {
  const mapped = MAPPED[name];
  const target = mapped === undefined ? name : (settings.colorMapping.get(mapped[0]) ?? mapped[1]);
  return WML_THEME_SLOT[target] ?? WML_THEME_SLOT[name];
}

function rgbToHsl(hex: string): [number, number, number] {
  const r = Number.parseInt(hex.slice(0, 2), 16) / 255;
  const g = Number.parseInt(hex.slice(2, 4), 16) / 255;
  const b = Number.parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): string {
  const hue = (p: number, q: number, t0: number): number => {
    let t = t0;
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r: number;
  let g: number;
  let b: number;
  if (s === 0) {
    r = g = b = l;
  } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue(p, q, h + 1 / 3);
    g = hue(p, q, h);
    b = hue(p, q, h - 1 / 3);
  }
  const hex = (x: number): string =>
    Math.round(Math.min(1, Math.max(0, x)) * 255)
      .toString(16)
      .padStart(2, '0');
  return `${hex(r)}${hex(g)}${hex(b)}`.toUpperCase();
}

/**
 * Applies `w:themeTint` / `w:themeShade` (hex bytes, 00-FF) to an RRGGBB colour: a tint moves the
 * HSL luminance towards white (`L·t + (1 − t)`), a shade towards black (`L·s`).
 */
export function tintShade(hex: string, tint?: string, shade?: string): string {
  const t = tint === undefined ? undefined : Number.parseInt(tint, 16);
  const s = shade === undefined ? undefined : Number.parseInt(shade, 16);
  const tOk = t !== undefined && Number.isFinite(t) && t >= 0 && t < 255;
  const sOk = s !== undefined && Number.isFinite(s) && s >= 0 && s < 255;
  if (!tOk && !sOk) return hex;
  let [h, sat, l] = rgbToHsl(hex);
  if (tOk) l = l * (t / 255) + (1 - t / 255);
  if (sOk) l *= s / 255;
  return hslToRgb(h, sat, l);
}

/**
 * Resolves a colour given as `val` (RRGGBB/auto) and optional theme colour + tint/shade.
 * The theme colour wins when the theme defines it; otherwise `val` is used.
 */
export function resolveColor(
  theme: ThemeData,
  settings: SettingsData,
  val: string | undefined,
  themeColor: string | undefined,
  tint: string | undefined,
  shade: string | undefined,
): string | undefined {
  if (themeColor !== undefined && themeColor !== 'none') {
    const slot = themeSlot(settings, themeColor);
    const base = slot === undefined ? undefined : theme.colors.get(slot);
    if (base !== undefined) return tintShade(base, tint, shade);
  }
  return hexColor(val);
}

/** The public {@link ThemeInfo} view of a theme. */
export function themeInfo(theme: ThemeData, settings: SettingsData): ThemeInfo {
  const fonts = (s: FontScheme): { latin?: string; eastAsia?: string; cs?: string } => {
    const out: { latin?: string; eastAsia?: string; cs?: string } = {};
    const latin = schemeFont(s, 'latin', settings.themeFontLang);
    const ea = schemeFont(s, 'ea', settings.themeFontLang);
    const cs = schemeFont(s, 'cs', settings.themeFontLang);
    if (latin !== undefined) out.latin = latin;
    if (ea !== undefined) out.eastAsia = ea;
    if (cs !== undefined) out.cs = cs;
    return out;
  };
  const colors: Record<string, string> = {};
  for (const [k, v] of theme.colors) colors[k] = v;
  for (const name of Object.keys(WML_THEME_SLOT).concat(Object.keys(MAPPED))) {
    const slot = themeSlot(settings, name);
    const c = slot === undefined ? undefined : theme.colors.get(slot);
    // Slot names (accent1, ...) keep the theme's own value; mapped names only add.
    if (c !== undefined && colors[name] === undefined) colors[name] = c;
  }
  return { fonts: { major: fonts(theme.major), minor: fonts(theme.minor) }, colors };
}
