// Table and cell property layers (`w:tblPr`, `w:tcPr`), read and merged like pPr/rPr layers.
//
// Table formatting comes from the table style chain (whole table plus conditional regions) and
// the table's own `w:tblPr`; cell formatting from the style's `w:tcPr` (whole table and regions)
// and the cell's own `w:tcPr`. Everything overrides side by side / attribute by attribute.

import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { BorderSide } from './ppr.js';
import { readBorders } from './ppr.js';
import type { ThemeCtx } from './rpr.js';
import { shadingFill } from './rpr.js';
import type { BorderSpec, ResolvedTableProps } from './types.js';
import { intVal, onOff, onOffValue, tableWidth, wAttr, widthTwips } from './values.js';

export type Look = ResolvedTableProps['look'];
export type MarginSide = 'top' | 'left' | 'bottom' | 'right';
export type Width = { value: number; type: 'dxa' | 'pct' | 'auto' | 'nil' };

/** What one level says about table formatting. */
export interface TblLayer {
  width?: Width;
  jc?: 'left' | 'center' | 'right';
  ind?: number;
  layout?: 'fixed' | 'autofit';
  cellMar?: Partial<Record<MarginSide, number>>;
  cellSpacing?: number;
  borders?: Partial<Record<BorderSide, BorderSpec | null>>;
  look?: Look;
  bidiVisual?: boolean;
  rowBand?: number;
  colBand?: number;
  shd?: string;
}

/** What one level says about cell formatting. */
export interface TcLayer {
  width?: Width;
  borders?: Partial<Record<BorderSide, BorderSpec | null>>;
  shd?: string;
  vAlign?: 'top' | 'center' | 'bottom';
  mar?: Partial<Record<MarginSide, number>>;
  noWrap?: boolean;
  textDirection?: string;
}

const MARGIN_ALIASES: Readonly<Record<string, MarginSide>> = {
  top: 'top',
  left: 'left',
  start: 'left',
  bottom: 'bottom',
  right: 'right',
  end: 'right',
};

function readMargins(el: XElement): Partial<Record<MarginSide, number>> {
  const out: Partial<Record<MarginSide, number>> = {};
  for (const k of el.elements(NS.w)) {
    const side = MARGIN_ALIASES[k.local];
    if (side === undefined) continue;
    const v = widthTwips(k);
    if (v !== undefined) out[side] = v;
  }
  return out;
}

const LOOK_BITS: readonly (readonly [keyof Look, number])[] = [
  ['firstRow', 0x0020],
  ['lastRow', 0x0040],
  ['firstColumn', 0x0080],
  ['lastColumn', 0x0100],
  ['noHBand', 0x0200],
  ['noVBand', 0x0400],
];

/** `w:tblLook`: the hex `w:val` bitmask, overridden by the explicit (w: or w14:) attributes. */
export function readLook(el: XElement): Look {
  const look: Look = {
    firstRow: false,
    lastRow: false,
    firstColumn: false,
    lastColumn: false,
    noHBand: false,
    noVBand: false,
  };
  const val = wAttr(el, 'val');
  const bits = val === undefined ? 0 : Number.parseInt(val, 16);
  if (Number.isFinite(bits)) for (const [k, bit] of LOOK_BITS) look[k] = (bits & bit) !== 0;
  for (const [k] of LOOK_BITS) {
    const v = onOffValue(el.attrNS(NS.w, k) ?? el.attrNS(NS.w14, k));
    if (v !== undefined) look[k] = v;
  }
  return look;
}

function jcOf(v: string | undefined): 'left' | 'center' | 'right' | undefined {
  switch (v) {
    case 'left':
    case 'start':
      return 'left';
    case 'center':
      return 'center';
    case 'right':
    case 'end':
      return 'right';
    default:
      return undefined;
  }
}

/** Reads a `w:tblPr` (or a style's / a `w:tblPrEx`) into a layer. */
export function readTblPr(t: ThemeCtx, el: XElement | undefined): TblLayer {
  const out: TblLayer = {};
  if (el === undefined) return out;
  for (const k of el.elements(NS.w)) {
    switch (k.local) {
      case 'tblW': {
        const w = tableWidth(k);
        if (w !== undefined) out.width = w;
        break;
      }
      case 'jc': {
        const v = jcOf(wAttr(k, 'val'));
        if (v !== undefined) out.jc = v;
        break;
      }
      case 'tblInd': {
        const v = widthTwips(k);
        if (v !== undefined) out.ind = v;
        break;
      }
      case 'tblLayout':
        out.layout = wAttr(k, 'type') === 'fixed' ? 'fixed' : 'autofit';
        break;
      case 'tblCellMar':
        out.cellMar = readMargins(k);
        break;
      case 'tblCellSpacing': {
        const v = widthTwips(k);
        if (v !== undefined) out.cellSpacing = v;
        break;
      }
      case 'tblBorders': {
        const b = readBorders(t, k);
        if (b !== undefined) out.borders = b;
        break;
      }
      case 'tblLook':
        out.look = readLook(k);
        break;
      case 'bidiVisual':
        out.bidiVisual = onOff(k) as boolean;
        break;
      case 'tblStyleRowBandSize': {
        const v = intVal(k);
        if (v !== undefined && v > 0) out.rowBand = v;
        break;
      }
      case 'tblStyleColBandSize': {
        const v = intVal(k);
        if (v !== undefined && v > 0) out.colBand = v;
        break;
      }
      case 'shd': {
        const v = shadingFill(t, k);
        if (v !== undefined) out.shd = v;
        break;
      }
      default:
    }
  }
  return out;
}

/** Reads a `w:tcPr` into a layer. */
export function readTcPr(t: ThemeCtx, el: XElement | undefined): TcLayer {
  const out: TcLayer = {};
  if (el === undefined) return out;
  for (const k of el.elements(NS.w)) {
    switch (k.local) {
      case 'tcW': {
        const w = tableWidth(k);
        if (w !== undefined) out.width = w;
        break;
      }
      case 'tcBorders': {
        const b = readBorders(t, k);
        if (b !== undefined) out.borders = b;
        break;
      }
      case 'shd': {
        const v = shadingFill(t, k);
        if (v !== undefined) out.shd = v;
        break;
      }
      case 'vAlign': {
        const v = wAttr(k, 'val');
        if (v === 'top' || v === 'center' || v === 'bottom') out.vAlign = v;
        else if (v === 'both') out.vAlign = 'center';
        break;
      }
      case 'tcMar':
        out.mar = readMargins(k);
        break;
      case 'noWrap':
        out.noWrap = onOff(k) as boolean;
        break;
      case 'textDirection': {
        const v = wAttr(k, 'val');
        if (v !== undefined) out.textDirection = v;
        break;
      }
      default:
    }
  }
  return out;
}

/** `into` overridden by `from`, margins and border sides merged one by one. */
export function mergeTbl(into: TblLayer, from: TblLayer): TblLayer {
  const out: TblLayer = { ...into };
  for (const key of Object.keys(from) as (keyof TblLayer)[]) {
    const v = from[key];
    if (v === undefined) continue;
    if (key === 'cellMar') out.cellMar = { ...(into.cellMar ?? {}), ...(v as TblLayer['cellMar']) };
    else if (key === 'borders') {
      out.borders = { ...(into.borders ?? {}), ...(v as TblLayer['borders']) };
    } else (out as Record<string, unknown>)[key] = v;
  }
  return out;
}

/** `into` overridden by `from`, margins and border sides merged one by one. */
export function mergeTc(into: TcLayer, from: TcLayer): TcLayer {
  const out: TcLayer = { ...into };
  for (const key of Object.keys(from) as (keyof TcLayer)[]) {
    const v = from[key];
    if (v === undefined) continue;
    if (key === 'mar') out.mar = { ...(into.mar ?? {}), ...(v as TcLayer['mar']) };
    else if (key === 'borders') {
      out.borders = { ...(into.borders ?? {}), ...(v as TcLayer['borders']) };
    } else (out as Record<string, unknown>)[key] = v;
  }
  return out;
}
