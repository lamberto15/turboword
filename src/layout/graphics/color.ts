// DrawingML colours: srgbClr / schemeClr / prstClr / sysClr / scrgbClr / hslClr plus the colour
// transforms Word writes most (lumMod, lumOff, tint, shade, satMod, alpha, ...), and the bits of
// the theme graphics need (colour scheme, line widths of the style matrix, major/minor fonts).
//
// The PageModel has no transparency, so alpha is approximated by blending toward white (pages
// are white). Tint and shade are applied in sRGB space; Office uses linear RGB, which differs
// slightly in mid tones. Results are uppercase RRGGBB strings.

import type { WordDocument } from '../../wml/document.js';
import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';

/** What graphics read from the theme part. */
export interface GraphicsTheme {
  /** Scheme colour name (dk1, lt1, accent1, ..., plus the mapped bg1/tx1/bg2/tx2) → RRGGBB. */
  colors: Record<string, string>;
  /** a:lnStyleLst widths in points (style matrix lnRef idx 1..3). */
  lineWidths: number[];
  /** Theme font families. */
  fonts: { major?: string; minor?: string };
}

/** Office 2013+ default theme colours (used when the document has no theme). */
const DEFAULT_COLORS: Record<string, string> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

const SCHEME_NAMES = Object.keys(DEFAULT_COLORS);

const cache = new WeakMap<WordDocument, { key: string; theme: GraphicsTheme }>();

/** The document's theme as graphics need it (cached; re-read when the theme part changes). */
export function graphicsTheme(doc: WordDocument): GraphicsTheme {
  const part = doc.mainPart.relatedByType('theme')[0];
  const settings = doc.mainPart.relatedByType('settings')[0];
  let root: XElement | undefined;
  let mapEl: XElement | undefined;
  try {
    root = part?.xml().root;
    mapEl = settings?.xml().root.element(NS.w, 'clrSchemeMapping');
  } catch {
    root = undefined;
  }
  const key = `${part?.name ?? ''}:${part === undefined ? 0 : part.xml().version}:${
    settings === undefined ? 0 : settings.xml().version
  }`;
  const hit = cache.get(doc);
  if (hit !== undefined && hit.key === key) return hit.theme;
  const theme = parseTheme(root, mapEl);
  cache.set(doc, { key, theme });
  return theme;
}

/** Parses a:theme (undefined → Office defaults) with an optional w:clrSchemeMapping. */
export function parseTheme(root: XElement | undefined, mapping?: XElement): GraphicsTheme {
  const colors: Record<string, string> = { ...DEFAULT_COLORS };
  const lineWidths: number[] = [0.5, 1, 1.5];
  const fonts: { major?: string; minor?: string } = {};
  const elems = root?.element(NS.a, 'themeElements');
  const scheme = elems?.element(NS.a, 'clrScheme');
  if (scheme !== undefined) {
    for (const name of SCHEME_NAMES) {
      const c = scheme.element(NS.a, name)?.firstElement();
      if (c === undefined) continue;
      const v = c.is(NS.a, 'sysClr')
        ? (c.attr('lastClr') ?? (c.attr('val') === 'window' ? 'FFFFFF' : '000000'))
        : c.attr('val');
      if (v !== undefined && /^[0-9A-Fa-f]{6}$/.test(v)) colors[name] = v.toUpperCase();
    }
  }
  const lns = elems?.element(NS.a, 'fmtScheme')?.element(NS.a, 'lnStyleLst')?.elements(NS.a, 'ln');
  if (lns !== undefined && lns.length > 0) {
    lns.forEach((ln, i) => {
      const w = Number(ln.attr('w'));
      if (Number.isFinite(w) && w >= 0) lineWidths[i] = w / 12700;
    });
  }
  const fs = elems?.element(NS.a, 'fontScheme');
  const major = fs?.element(NS.a, 'majorFont')?.element(NS.a, 'latin')?.attr('typeface');
  const minor = fs?.element(NS.a, 'minorFont')?.element(NS.a, 'latin')?.attr('typeface');
  if (major) fonts.major = major;
  if (minor) fonts.minor = minor;
  // Logical names: bg1/tx1/bg2/tx2 through the settings mapping (default light1/dark1/...).
  const map: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
  if (mapping !== undefined) {
    for (const k of ['bg1', 'bg2', 't1', 't2'] as const) {
      const v = mapping.attrNS(NS.w, k);
      if (v === undefined) continue;
      const target = v.replace(/^light/, 'lt').replace(/^dark/, 'dk');
      map[k === 't1' ? 'tx1' : k === 't2' ? 'tx2' : k] = target;
    }
  }
  for (const [k, v] of Object.entries(map)) colors[k] = colors[v] ?? DEFAULT_COLORS[v] ?? '000000';
  return { colors, lineWidths, fonts };
}

/** DrawingML preset colours (a:prstClr) — the common subset; others fall back to black. */
const PRESET: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  magenta: 'FF00FF',
  gray: '808080',
  grey: '808080',
  silver: 'C0C0C0',
  maroon: '800000',
  olive: '808000',
  lime: '00FF00',
  aqua: '00FFFF',
  teal: '008080',
  navy: '000080',
  purple: '800080',
  fuchsia: 'FF00FF',
  orange: 'FFA500',
  brown: 'A52A2A',
  pink: 'FFC0CB',
  gold: 'FFD700',
  dkGray: 'A9A9A9',
  dkGrey: 'A9A9A9',
  ltGray: 'D3D3D3',
  ltGrey: 'D3D3D3',
  dkBlue: '00008B',
  dkRed: '8B0000',
  dkGreen: '006400',
  dkCyan: '008B8B',
  dkMagenta: '8B008B',
  dkOrange: 'FF8C00',
  dkYellow: '808000',
  ltBlue: 'ADD8E6',
  ltGreen: '90EE90',
  ltYellow: 'FFFFE0',
  ltCyan: 'E0FFFF',
  medBlue: '0000CD',
  violet: 'EE82EE',
  indigo: '4B0082',
  beige: 'F5F5DC',
  tan: 'D2B48C',
  coral: 'FF7F50',
  salmon: 'FA8072',
  crimson: 'DC143C',
  khaki: 'F0E68C',
  turquoise: '40E0D0',
};

/** System colours (a:sysClr without lastClr). */
const SYSTEM: Record<string, string> = {
  windowText: '000000',
  window: 'FFFFFF',
  btnFace: 'F0F0F0',
  btnText: '000000',
  highlight: '0078D7',
  highlightText: 'FFFFFF',
  grayText: '6D6D6D',
  menu: 'F0F0F0',
  menuText: '000000',
  infoBk: 'FFFFE1',
  infoText: '000000',
  btnShadow: 'A0A0A0',
  btnHighlight: 'FFFFFF',
  '3dDkShadow': '696969',
  '3dLight': 'E3E3E3',
};

/** A resolved colour with its opacity (0..1). */
export interface RGBA {
  r: number;
  g: number;
  b: number;
  a: number;
}

/** RRGGBB → RGBA (opaque). */
export function hexToRgba(hex: string): RGBA {
  const n = Number.parseInt(hex, 16) || 0;
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: 1 };
}

const hex2 = (v: number): string =>
  Math.max(0, Math.min(255, Math.round(v)))
    .toString(16)
    .padStart(2, '0')
    .toUpperCase();

/** RGBA → RRGGBB, alpha blended over white. */
export function rgbaToHex(c: RGBA): string {
  const a = Math.max(0, Math.min(1, c.a));
  const mix = (v: number): number => v * a + 255 * (1 - a);
  return hex2(mix(c.r)) + hex2(mix(c.g)) + hex2(mix(c.b));
}

function rgbToHsl(c: RGBA): [number, number, number] {
  const r = c.r / 255;
  const g = c.g / 255;
  const b = c.b / 255;
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

function hslToRgb(h: number, s: number, l: number, a: number): RGBA {
  if (s === 0) return { r: l * 255, g: l * 255, b: l * 255, a };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number): number => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return { r: f(h + 1 / 3) * 255, g: f(h) * 255, b: f(h - 1 / 3) * 255, a };
}

const clamp01 = (v: number): number => Math.max(0, Math.min(1, v));

/** Applies the colour transform children of a colour element. */
function transform(base: RGBA, el: XElement): RGBA {
  let c = { ...base };
  for (const m of el.elements(NS.a)) {
    const val = Number(m.attr('val') ?? 0) / 100000;
    const local = m.name.slice(m.name.indexOf(':') + 1);
    switch (local) {
      case 'lumMod':
      case 'lumOff':
      case 'satMod':
      case 'satOff':
      case 'hueMod':
      case 'hueOff':
      case 'lum':
      case 'sat':
      case 'hue': {
        let [h, s, l] = rgbToHsl(c);
        if (local === 'lumMod') l = clamp01(l * val);
        else if (local === 'lumOff') l = clamp01(l + val);
        else if (local === 'lum') l = clamp01(val);
        else if (local === 'satMod') s = clamp01(s * val);
        else if (local === 'satOff') s = clamp01(s + val);
        else if (local === 'sat') s = clamp01(val);
        else if (local === 'hueMod') h = (((h * val) % 1) + 1) % 1;
        else if (local === 'hueOff') h = (((h + (val * 100000) / 21600000) % 1) + 1) % 1;
        else h = (val * 100000) / 21600000;
        c = hslToRgb(h, s, l, c.a);
        break;
      }
      case 'tint':
        c = {
          r: 255 - (255 - c.r) * val,
          g: 255 - (255 - c.g) * val,
          b: 255 - (255 - c.b) * val,
          a: c.a,
        };
        break;
      case 'shade':
        c = { r: c.r * val, g: c.g * val, b: c.b * val, a: c.a };
        break;
      case 'alpha':
        c.a = clamp01(val);
        break;
      case 'alphaMod':
        c.a = clamp01(c.a * val);
        break;
      case 'alphaOff':
        c.a = clamp01(c.a + val);
        break;
      case 'inv':
        c = { r: 255 - c.r, g: 255 - c.g, b: 255 - c.b, a: c.a };
        break;
      case 'gray': {
        const y = 0.299 * c.r + 0.587 * c.g + 0.114 * c.b;
        c = { r: y, g: y, b: y, a: c.a };
        break;
      }
      case 'comp': {
        const [h, s, l] = rgbToHsl(c);
        c = hslToRgb((h + 0.5) % 1, s, l, c.a);
        break;
      }
      case 'red':
        c.r = 255 * clamp01(val);
        break;
      case 'green':
        c.g = 255 * clamp01(val);
        break;
      case 'blue':
        c.b = 255 * clamp01(val);
        break;
    }
  }
  return c;
}

/**
 * A DrawingML colour element (srgbClr, schemeClr, ...) → RGBA; `phClr` is the placeholder colour
 * from a style reference. Undefined for anything that is not a colour element.
 */
export function colorRgba(
  el: XElement | undefined,
  theme: GraphicsTheme,
  phClr?: RGBA,
): RGBA | undefined {
  if (el === undefined || el.uri !== NS.a) return undefined;
  const local = el.name.slice(el.name.indexOf(':') + 1);
  const val = el.attr('val') ?? '';
  let base: RGBA | undefined;
  switch (local) {
    case 'srgbClr':
      base = hexToRgba(/^[0-9A-Fa-f]{6}$/.test(val) ? val : '000000');
      break;
    case 'schemeClr':
      if (val === 'phClr') base = phClr ?? hexToRgba('000000');
      else base = hexToRgba(theme.colors[val] ?? '000000');
      break;
    case 'prstClr':
      base = hexToRgba(PRESET[val] ?? '000000');
      break;
    case 'sysClr':
      base = hexToRgba(el.attr('lastClr') ?? SYSTEM[val] ?? '000000');
      break;
    case 'scrgbClr': {
      const lin = (k: string): number => {
        const v = clamp01(Number(el.attr(k) ?? 0) / 100000);
        return 255 * (v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055);
      };
      base = { r: lin('r'), g: lin('g'), b: lin('b'), a: 1 };
      break;
    }
    case 'hslClr':
      base = hslToRgb(
        Number(el.attr('hue') ?? 0) / 21600000,
        clamp01(Number(el.attr('sat') ?? 0) / 100000),
        clamp01(Number(el.attr('lum') ?? 0) / 100000),
        1,
      );
      break;
    default:
      return undefined;
  }
  return transform(base, el);
}

/** The first colour child of `parent` (e.g. a:solidFill, a:fgClr, a:lnRef) → RRGGBB. */
export function colorOf(
  parent: XElement | undefined,
  theme: GraphicsTheme,
  phClr?: RGBA,
): string | undefined {
  if (parent === undefined) return undefined;
  for (const c of parent.elements(NS.a)) {
    const rgba = colorRgba(c, theme, phClr);
    if (rgba !== undefined) return rgbaToHex(rgba);
  }
  return undefined;
}

/** Like {@link colorOf} but keeps the alpha channel. */
export function colorRgbaOf(
  parent: XElement | undefined,
  theme: GraphicsTheme,
  phClr?: RGBA,
): RGBA | undefined {
  if (parent === undefined) return undefined;
  for (const c of parent.elements(NS.a)) {
    const rgba = colorRgba(c, theme, phClr);
    if (rgba !== undefined) return rgba;
  }
  return undefined;
}

/** Parses a CSS/VML-style colour (#rgb, #rrggbb, names, "rgb(r,g,b)") → RRGGBB. */
export function cssColor(v: string | undefined): string | undefined {
  if (v === undefined) return undefined;
  const s = v.trim().split(/\s+/)[0] ?? '';
  let m = /^#?([0-9a-fA-F]{6})$/.exec(s);
  if (m?.[1]) return m[1].toUpperCase();
  m = /^#([0-9a-fA-F]{3})$/.exec(s);
  if (m?.[1])
    return [...m[1]]
      .map((ch) => ch + ch)
      .join('')
      .toUpperCase();
  m = /^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/.exec(s);
  if (m) return hex2(Number(m[1])) + hex2(Number(m[2])) + hex2(Number(m[3]));
  const named = PRESET[s] ?? PRESET[s.toLowerCase()] ?? SYSTEM[s];
  return named;
}

/** Averages colours (gradient approximation). */
export function averageColor(colors: readonly RGBA[]): RGBA | undefined {
  if (colors.length === 0) return undefined;
  const n = colors.length;
  let r = 0;
  let g = 0;
  let b = 0;
  let a = 0;
  for (const c of colors) {
    r += c.r;
    g += c.g;
    b += c.b;
    a += c.a;
  }
  return { r: r / n, g: g / n, b: b / n, a: a / n };
}

/** Darkens/lightens RRGGBB (for preset path shading: darken, darkenLess, lighten, lightenLess). */
export function shadeHex(hex: string, mode: string): string {
  const c = hexToRgba(hex);
  const f = mode === 'darken' ? 0.6 : mode === 'darkenLess' ? 0.8 : 1;
  if (f !== 1) return rgbaToHex({ r: c.r * f, g: c.g * f, b: c.b * f, a: 1 });
  const t = mode === 'lighten' ? 0.6 : mode === 'lightenLess' ? 0.8 : 1;
  return rgbaToHex({
    r: 255 - (255 - c.r) * t,
    g: 255 - (255 - c.g) * t,
    b: 255 - (255 - c.b) * t,
    a: 1,
  });
}
