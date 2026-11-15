// Attribute value parsing shared by the resolver: ST_OnOff, integers, measures and colours.
//
// WordprocessingML spells the same quantity several ways: ST_OnOff accepts true/false/1/0/on/off
// (and an omitted `w:val` means on), lengths may be plain twips or a "universal measure" such as
// `1.5in` (ISO 29500 transitional and Strict), percentages may be fiftieths of a percent or `50%`.
// Everything is normalized here so the cascade code only ever sees numbers and booleans.

import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';

/** `w:<local>` attribute of `el` (undefined when `el` or the attribute is missing). */
export function wAttr(el: XElement | undefined, local: string): string | undefined {
  return el?.attrNS(NS.w, local);
}

/** ST_OnOff text: true/1/on → true, false/0/off → false, anything else → undefined. */
export function onOffValue(v: string | undefined): boolean | undefined {
  switch (v?.trim().toLowerCase()) {
    case 'true':
    case '1':
    case 'on':
      return true;
    case 'false':
    case '0':
    case 'off':
      return false;
    default:
      return undefined;
  }
}

/**
 * An ST_OnOff property element (`<w:b/>`, `<w:b w:val="0"/>`): undefined when absent, true when
 * present without `w:val` (or with an unrecognized one), else the parsed value.
 */
export function onOff(el: XElement | undefined): boolean | undefined {
  if (el === undefined) return undefined;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined ? true : (onOffValue(v) ?? true);
}

/** Integer value of a decimal string, or undefined. */
export function int(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const n = Number.parseInt(v.trim(), 10);
  return Number.isFinite(n) ? n : undefined;
}

/** `w:val` of `el` as an integer. */
export function intVal(el: XElement | undefined): number | undefined {
  return int(el?.attrNS(NS.w, 'val'));
}

/** Twips per universal-measure unit. */
const TWIPS_PER_UNIT: Readonly<Record<string, number>> = {
  mm: 1440 / 25.4,
  cm: 1440 / 2.54,
  in: 1440,
  pt: 20,
  pc: 240,
  pi: 240,
};

const MEASURE = /^([-+]?\d+(?:\.\d+)?|[-+]?\.\d+)\s*(mm|cm|in|pt|pc|pi)?$/;

/** A length in twips: plain twips (`720`) or a universal measure (`0.5in`, `12.7mm`). */
export function twips(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const m = MEASURE.exec(v.trim());
  if (m === null) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  return m[2] === undefined ? Math.round(n) : Math.round(n * TWIPS_PER_UNIT[m[2]]!);
}

/** A font size in half-points: plain half-points (`24`) or a universal measure (`12pt`). */
export function halfPoints(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const m = MEASURE.exec(v.trim());
  if (m === null) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  return m[2] === undefined ? n : (n * TWIPS_PER_UNIT[m[2]]!) / 10;
}

/**
 * A table width (`w:w` + `w:type`, CT_TblWidth). Percentages are returned in fiftieths of a
 * percent (`5000` = 100%), whether written that way or as `100%`.
 */
export function tableWidth(
  el: XElement | undefined,
): { value: number; type: 'dxa' | 'pct' | 'auto' | 'nil' } | undefined {
  if (el === undefined) return undefined;
  const raw = el.attrNS(NS.w, 'w');
  const t = el.attrNS(NS.w, 'type');
  const type: 'dxa' | 'pct' | 'auto' | 'nil' =
    t === 'pct' || t === 'auto' || t === 'nil' ? t : raw?.trim().endsWith('%') ? 'pct' : 'dxa';
  let value = 0;
  if (raw !== undefined) {
    const s = raw.trim();
    if (s.endsWith('%')) {
      const n = Number(s.slice(0, -1));
      value = Number.isFinite(n) ? Math.round(n * 50) : 0;
    } else if (type === 'pct') {
      const n = Number(s);
      value = Number.isFinite(n) ? Math.round(n) : 0;
    } else {
      value = twips(s) ?? 0;
    }
  }
  return { value, type };
}

/** Twips of a CT_TblWidth used as a length (margins, spacing, indent); non-dxa types give 0. */
export function widthTwips(el: XElement | undefined): number | undefined {
  const w = tableWidth(el);
  if (w === undefined) return undefined;
  return w.type === 'dxa' ? w.value : 0;
}

/** RRGGBB (upper case) or 'auto'; undefined for anything else. */
export function hexColor(v: string | undefined): string | undefined {
  if (v === undefined) return undefined;
  const s = v.trim();
  if (s.toLowerCase() === 'auto') return 'auto';
  if (/^[0-9A-Fa-f]{6}$/.test(s)) return s.toUpperCase();
  return undefined;
}
