// Run properties: reading one `w:rPr`, merging levels, and the final resolved form.
//
// A `RunLayer` is what one level of the cascade says (undefined = says nothing). Levels combine
// in two ways (ECMA-376 §17.7.2-17.7.3):
// - within one style chain (basedOn), and for every non-toggle property, a later level simply
//   overrides an earlier one;
// - toggle properties (b, i, caps, strike, vanish, ...) behave differently *between* the style
//   types: the table style, the paragraph style and the character style each contribute their
//   (chain-resolved) value, and every one that says "on" flips the state reached so far; "off"
//   leaves it alone. Document defaults give the initial state and direct formatting is absolute.
// Colours are resolved against the theme as they are read, fonts are kept as references (name or
// theme slot) until the end so a later level's plain name can replace an earlier theme font.

import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { SettingsData, ThemeData } from './theme.js';
import { resolveColor, themeFont } from './theme.js';
import type { ResolvedRunProps } from './types.js';
import { halfPoints, onOff, twips, wAttr } from './values.js';

/** Toggle properties (§17.7.3), plus `dstrike`, which Word treats the same way. */
export const TOGGLES = [
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'vanish',
] as const;
export type Toggle = (typeof TOGGLES)[number];

/** A font slot value: a plain typeface name or a theme font reference. */
export interface FontRef {
  name?: string;
  theme?: string;
}

/** What one cascade level says about run formatting. */
export interface RunLayer {
  fontAscii?: FontRef;
  fontHAnsi?: FontRef;
  fontEastAsia?: FontRef;
  fontCs?: FontRef;
  /** `w:rFonts/@w:hint`. */
  fontHint?: 'default' | 'eastAsia' | 'cs';
  /** Half-points. */
  sz?: number;
  szCs?: number;
  b?: boolean;
  bCs?: boolean;
  i?: boolean;
  iCs?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  strike?: boolean;
  dstrike?: boolean;
  outline?: boolean;
  shadow?: boolean;
  emboss?: boolean;
  imprint?: boolean;
  vanish?: boolean;
  /** Underline style; 'none' turns it off. */
  u?: string;
  uColor?: string;
  color?: string;
  /** Highlight colour name; 'none' turns it off. */
  highlight?: string;
  /** Shading fill; 'none' = explicitly no shading. */
  shd?: string;
  vertAlign?: 'baseline' | 'superscript' | 'subscript';
  spacing?: number;
  position?: number;
  w?: number;
  /** Half-points. */
  kern?: number;
  rtl?: boolean;
  cs?: boolean;
  lang?: string;
  langEastAsia?: string;
  langBidi?: string;
}

/** Theme + settings needed to read colours. */
export interface ThemeCtx {
  theme: ThemeData;
  settings: SettingsData;
}

const FONT_SLOTS = [
  ['fontAscii', 'ascii', 'asciiTheme'],
  ['fontHAnsi', 'hAnsi', 'hAnsiTheme'],
  ['fontEastAsia', 'eastAsia', 'eastAsiaTheme'],
  ['fontCs', 'cs', 'cstheme'],
] as const;

/**
 * Shading fill of a `w:shd` (CT_Shd): the fill colour for `clear`, the pattern colour for
 * `solid`, a mix for `pctNN` patterns; 'none' for an explicit no-shading, undefined when the
 * element says nothing usable.
 */
export function shadingFill(t: ThemeCtx, shd: XElement | undefined): string | undefined {
  if (shd === undefined) return undefined;
  const val = wAttr(shd, 'val') ?? 'clear';
  const fill = resolveColor(
    t.theme,
    t.settings,
    wAttr(shd, 'fill'),
    wAttr(shd, 'themeFill'),
    wAttr(shd, 'themeFillTint'),
    wAttr(shd, 'themeFillShade'),
  );
  const color = resolveColor(
    t.theme,
    t.settings,
    wAttr(shd, 'color'),
    wAttr(shd, 'themeColor'),
    wAttr(shd, 'themeTint'),
    wAttr(shd, 'themeShade'),
  );
  if (val === 'nil') return 'none';
  if (val === 'solid') return color === undefined || color === 'auto' ? '000000' : color;
  const pct = /^pct(\d+)$/.exec(val);
  if (pct !== null && color !== undefined && color !== 'auto') {
    const f = Math.min(100, Number(pct[1])) / 100;
    const bg = fill === undefined || fill === 'auto' ? 'FFFFFF' : fill;
    let out = '';
    for (let k = 0; k < 6; k += 2) {
      const a = Number.parseInt(bg.slice(k, k + 2), 16);
      const b = Number.parseInt(color.slice(k, k + 2), 16);
      out += Math.round(a + (b - a) * f)
        .toString(16)
        .padStart(2, '0');
    }
    return out.toUpperCase();
  }
  if (fill === undefined || fill === 'auto') return val === 'clear' ? 'none' : undefined;
  return fill;
}

/** Reads one `w:rPr` (or `w:rPrDefault/w:rPr`, a style's `w:rPr`, ...) into a layer. */
export function readRPr(t: ThemeCtx, rPr: XElement | undefined): RunLayer {
  const out: RunLayer = {};
  if (rPr === undefined) return out;
  for (const el of rPr.elements(NS.w)) {
    const k = el.local;
    switch (k) {
      case 'rFonts':
        for (const [key, attr, themeAttr] of FONT_SLOTS) {
          const th = wAttr(el, themeAttr);
          const name = wAttr(el, attr);
          if (th !== undefined) out[key] = name === undefined ? { theme: th } : { theme: th, name };
          else if (name !== undefined) out[key] = { name };
        }
        {
          const hint = wAttr(el, 'hint');
          if (hint === 'default' || hint === 'eastAsia' || hint === 'cs') out.fontHint = hint;
        }
        break;
      case 'sz':
      case 'szCs': {
        const v = halfPoints(wAttr(el, 'val'));
        if (v !== undefined && v > 0) out[k] = v;
        break;
      }
      case 'b':
      case 'bCs':
      case 'i':
      case 'iCs':
      case 'caps':
      case 'smallCaps':
      case 'strike':
      case 'dstrike':
      case 'outline':
      case 'shadow':
      case 'emboss':
      case 'imprint':
      case 'vanish':
        out[k] = onOff(el) as boolean;
        break;
      case 'rtl':
      case 'cs':
        out[k] = onOff(el) as boolean;
        break;
      case 'u': {
        const v = wAttr(el, 'val') ?? 'single';
        out.u = v;
        const c = resolveColor(
          t.theme,
          t.settings,
          wAttr(el, 'color'),
          wAttr(el, 'themeColor'),
          wAttr(el, 'themeTint'),
          wAttr(el, 'themeShade'),
        );
        if (c !== undefined) out.uColor = c;
        break;
      }
      case 'color': {
        const c = resolveColor(
          t.theme,
          t.settings,
          wAttr(el, 'val'),
          wAttr(el, 'themeColor'),
          wAttr(el, 'themeTint'),
          wAttr(el, 'themeShade'),
        );
        if (c !== undefined) out.color = c;
        break;
      }
      case 'highlight': {
        const v = wAttr(el, 'val');
        if (v !== undefined) out.highlight = v;
        break;
      }
      case 'shd': {
        const v = shadingFill(t, el);
        if (v !== undefined) out.shd = v;
        break;
      }
      case 'vertAlign': {
        const v = wAttr(el, 'val');
        if (v === 'superscript' || v === 'subscript' || v === 'baseline') out.vertAlign = v;
        break;
      }
      case 'spacing': {
        const v = twips(wAttr(el, 'val'));
        if (v !== undefined) out.spacing = v;
        break;
      }
      case 'position': {
        const v = halfPoints(wAttr(el, 'val'));
        if (v !== undefined) out.position = v;
        break;
      }
      case 'w': {
        const raw = wAttr(el, 'val');
        const v = raw === undefined ? undefined : Number.parseFloat(raw);
        if (v !== undefined && Number.isFinite(v) && v > 0) out.w = v;
        break;
      }
      case 'kern': {
        const v = halfPoints(wAttr(el, 'val'));
        if (v !== undefined && v >= 0) out.kern = v;
        break;
      }
      case 'lang': {
        const v = wAttr(el, 'val');
        const ea = wAttr(el, 'eastAsia');
        const bidi = wAttr(el, 'bidi');
        if (v !== undefined) out.lang = v;
        if (ea !== undefined) out.langEastAsia = ea;
        if (bidi !== undefined) out.langBidi = bidi;
        break;
      }
      default:
    }
  }
  return out;
}

/** `into` overridden by every property `from` specifies (plain override, all properties). */
export function mergeRun(into: RunLayer, from: RunLayer): RunLayer {
  const out: RunLayer = { ...into };
  for (const key of Object.keys(from) as (keyof RunLayer)[]) {
    const v = from[key];
    if (v !== undefined) (out as Record<string, unknown>)[key] = v;
  }
  return out;
}

const TOGGLE_SET: ReadonlySet<string> = new Set(TOGGLES);

/** How a level is combined into a run state. */
export type Mode =
  /** Direct formatting / defaults: every property overrides, toggles absolutely. */
  | 'abs'
  /** A style type's resolved chain: non-toggles override, toggles flip when on. */
  | 'style'
  /** Only the non-toggle properties (the toggles are applied by a separate 'toggles' level). */
  | 'plain'
  /** Only the toggles, flipping when on. */
  | 'toggles';

/** A complete run state: a layer where every toggle is decided. */
export type RunState = RunLayer & Record<Toggle, boolean>;

/** The initial state (Word's defaults before docDefaults). */
export function initialRunState(): RunState {
  const out = {} as RunState;
  for (const k of TOGGLES) out[k] = false;
  return out;
}

/** Applies one level to a state. */
export function applyRun(state: RunState, layer: RunLayer, mode: Mode): RunState {
  const out: RunState = { ...state };
  for (const key of Object.keys(layer) as (keyof RunLayer)[]) {
    const v = layer[key];
    if (v === undefined) continue;
    if (TOGGLE_SET.has(key)) {
      const k = key as Toggle;
      if (mode === 'abs') out[k] = v as boolean;
      else if ((mode === 'style' || mode === 'toggles') && v === true) out[k] = !out[k];
    } else if (mode !== 'toggles') {
      (out as Record<string, unknown>)[key] = v;
    }
  }
  return out;
}

function fontName(t: ThemeCtx, ref: FontRef | undefined): string | undefined {
  if (ref === undefined) return undefined;
  if (ref.theme !== undefined) return themeFont(t.theme, t.settings, ref.theme) ?? ref.name;
  return ref.name;
}

/** Word's font when nothing names one. */
const DEFAULT_FONT = 'Times New Roman';
/** Word's size when nothing sets one: 10 pt (§17.3.2.38). */
const DEFAULT_HALF_POINTS = 20;

/** The public resolved form of a state. */
export function finishRun(t: ThemeCtx, s: RunState, styleId: string | undefined): ResolvedRunProps {
  const fonts: ResolvedRunProps['fonts'] = {};
  const ascii = fontName(t, s.fontAscii) ?? DEFAULT_FONT;
  const hAnsi = fontName(t, s.fontHAnsi) ?? ascii;
  const eastAsia = fontName(t, s.fontEastAsia);
  const cs = fontName(t, s.fontCs) ?? DEFAULT_FONT;
  fonts.ascii = ascii;
  fonts.hAnsi = hAnsi;
  if (eastAsia !== undefined) fonts.eastAsia = eastAsia;
  fonts.cs = cs;
  if (s.fontHint !== undefined) fonts.hint = s.fontHint;
  const sz = s.sz ?? DEFAULT_HALF_POINTS;
  const out: ResolvedRunProps = {
    fonts,
    size: sz / 2,
    sizeCs: (s.szCs ?? sz) / 2,
    bold: s.b,
    italic: s.i,
    boldCs: s.bCs,
    italicCs: s.iCs,
    strike: s.strike,
    doubleStrike: s.dstrike,
    color: s.color ?? 'auto',
    caps: s.caps,
    smallCaps: s.smallCaps,
    hidden: s.vanish,
    vertAlign: s.vertAlign ?? 'baseline',
    spacing: s.spacing ?? 0,
    position: s.position ?? 0,
    scale: s.w ?? 100,
    kern: (s.kern ?? 0) / 2,
    rtl: s.rtl ?? false,
    cs: s.cs ?? false,
  };
  if (s.u !== undefined && s.u !== 'none') {
    out.underline = s.u;
    if (s.uColor !== undefined) out.underlineColor = s.uColor;
  }
  if (s.highlight !== undefined && s.highlight !== 'none') out.highlight = s.highlight;
  if (s.shd !== undefined && s.shd !== 'none') out.shading = s.shd;
  if (s.lang !== undefined) out.lang = s.lang;
  if (s.langEastAsia !== undefined) out.langEastAsia = s.langEastAsia;
  if (s.langBidi !== undefined) out.langBidi = s.langBidi;
  if (styleId !== undefined) out.styleId = styleId;
  return out;
}

/** `w:rStyle/@w:val` of an rPr. */
export function rStyleOf(rPr: XElement | undefined): string | undefined {
  return rPr?.element(NS.w, 'rStyle')?.attrNS(NS.w, 'val');
}
