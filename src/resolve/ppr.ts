// Paragraph properties: reading one `w:pPr`, merging levels, and borders.
//
// Every paragraph property is a plain override between levels (there are no toggles in pPr),
// with three refinements Word applies:
// - indentation attributes override one by one, and `w:firstLine`/`w:hanging` are one signed
//   quantity (a level setting either replaces the other; `hanging` wins inside one element);
// - tab stops accumulate: each level adds stops or clears inherited ones at a position;
// - border sides override one by one (`nil`/`none` removes an inherited side).

import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { ThemeCtx } from './rpr.js';
import { shadingFill } from './rpr.js';
import { resolveColor } from './theme.js';
import type { Alignment, BorderSpec, Borders, TabStop } from './types.js';
import { int, intVal, onOff, onOffValue, twips, wAttr } from './values.js';

export type BorderSide = keyof Borders;

/** What one level says about paragraph formatting. */
export interface ParaLayer {
  jc?: string;
  /** Logical start/end indents (twips). */
  indStart?: number;
  indEnd?: number;
  /** Signed: positive = first-line indent, negative = hanging. */
  firstLine?: number;
  before?: number;
  after?: number;
  beforeAuto?: boolean;
  afterAuto?: boolean;
  line?: number;
  lineRule?: 'auto' | 'exact' | 'atLeast';
  contextualSpacing?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  suppressAutoHyphens?: boolean;
  bidi?: boolean;
  outlineLvl?: number;
  numId?: number;
  ilvl?: number;
  /** Position → stop ('clear' stops kept until the end so they can cancel inherited ones). */
  tabs?: Map<number, TabStop>;
  /** null = side explicitly removed. */
  borders?: Partial<Record<BorderSide, BorderSpec | null>>;
  /** Fill RRGGBB, or 'none'. */
  shd?: string;
  framePr?: Record<string, string>;
}

/**
 * One border element (CT_Border): undefined when absent, null for `nil`/`none` (an explicit
 * "no border", which still overrides), else the spec with its colour resolved.
 */
export function readBorder(t: ThemeCtx, el: XElement | undefined): BorderSpec | null | undefined {
  if (el === undefined) return undefined;
  const style = wAttr(el, 'val') ?? 'none';
  if (style === 'nil' || style === 'none') return null;
  const size = int(wAttr(el, 'sz')) ?? 4;
  const space = int(wAttr(el, 'space')) ?? 0;
  const color =
    resolveColor(
      t.theme,
      t.settings,
      wAttr(el, 'color'),
      wAttr(el, 'themeColor'),
      wAttr(el, 'themeTint'),
      wAttr(el, 'themeShade'),
    ) ?? 'auto';
  return { style, size, space, color };
}

const SIDE_ALIASES: Readonly<Record<string, BorderSide>> = {
  top: 'top',
  left: 'left',
  start: 'left',
  bottom: 'bottom',
  right: 'right',
  end: 'right',
  between: 'between',
  bar: 'bar',
  insideH: 'insideH',
  insideV: 'insideV',
};

/** A borders container (`w:pBdr`, `w:tblBorders`, `w:tcBorders`) as per-side overrides. */
export function readBorders(
  t: ThemeCtx,
  el: XElement | undefined,
): Partial<Record<BorderSide, BorderSpec | null>> | undefined {
  if (el === undefined) return undefined;
  const out: Partial<Record<BorderSide, BorderSpec | null>> = {};
  for (const k of el.elements(NS.w)) {
    const side = SIDE_ALIASES[k.local];
    if (side === undefined) continue;
    const b = readBorder(t, k);
    if (b !== undefined) out[side] = b;
  }
  return out;
}

const TAB_VALS: Readonly<Record<string, TabStop['val']>> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  decimal: 'decimal',
  bar: 'bar',
  num: 'num',
  clear: 'clear',
};

const LEADERS: ReadonlySet<string> = new Set(['dot', 'hyphen', 'underscore', 'heavy', 'middleDot']);

function readTabs(el: XElement): Map<number, TabStop> {
  const out = new Map<number, TabStop>();
  for (const tab of el.elements(NS.w, 'tab')) {
    const val = TAB_VALS[wAttr(tab, 'val') ?? ''];
    const pos = twips(wAttr(tab, 'pos'));
    if (val === undefined || pos === undefined) continue;
    const stop: TabStop = { pos, val };
    const leader = wAttr(tab, 'leader');
    if (leader !== undefined && LEADERS.has(leader)) {
      stop.leader = leader as NonNullable<TabStop['leader']>;
    }
    out.set(pos, stop);
  }
  return out;
}

/** Reads one `w:pPr` into a layer. */
export function readPPr(t: ThemeCtx, pPr: XElement | undefined): ParaLayer {
  const out: ParaLayer = {};
  if (pPr === undefined) return out;
  for (const el of pPr.elements(NS.w)) {
    switch (el.local) {
      case 'jc': {
        const v = wAttr(el, 'val');
        if (v !== undefined) out.jc = v;
        break;
      }
      case 'ind': {
        const start = twips(wAttr(el, 'start') ?? wAttr(el, 'left'));
        const end = twips(wAttr(el, 'end') ?? wAttr(el, 'right'));
        const hanging = twips(wAttr(el, 'hanging'));
        const first = twips(wAttr(el, 'firstLine'));
        if (start !== undefined) out.indStart = start;
        if (end !== undefined) out.indEnd = end;
        if (hanging !== undefined) out.firstLine = -hanging;
        else if (first !== undefined) out.firstLine = first;
        break;
      }
      case 'spacing': {
        const before = twips(wAttr(el, 'before'));
        const after = twips(wAttr(el, 'after'));
        const line = twips(wAttr(el, 'line'));
        const rule = wAttr(el, 'lineRule');
        const ba = wAttr(el, 'beforeAutospacing');
        const aa = wAttr(el, 'afterAutospacing');
        if (before !== undefined) out.before = before;
        if (after !== undefined) out.after = after;
        if (ba !== undefined) out.beforeAuto = onOffValue(ba) ?? true;
        if (aa !== undefined) out.afterAuto = onOffValue(aa) ?? true;
        if (line !== undefined) {
          out.line = line;
          out.lineRule = rule === 'exact' || rule === 'atLeast' ? rule : 'auto';
        } else if (rule === 'exact' || rule === 'atLeast' || rule === 'auto') {
          out.lineRule = rule;
        }
        break;
      }
      case 'contextualSpacing':
      case 'keepNext':
      case 'keepLines':
      case 'pageBreakBefore':
      case 'widowControl':
      case 'suppressAutoHyphens':
      case 'bidi':
        out[el.local] = onOff(el) as boolean;
        break;
      case 'outlineLvl': {
        const v = intVal(el);
        if (v !== undefined) out.outlineLvl = v;
        break;
      }
      case 'numPr': {
        const numId = intVal(el.element(NS.w, 'numId'));
        const ilvl = intVal(el.element(NS.w, 'ilvl'));
        if (numId !== undefined) out.numId = numId;
        if (ilvl !== undefined) out.ilvl = ilvl;
        break;
      }
      case 'tabs':
        out.tabs = readTabs(el);
        break;
      case 'pBdr': {
        const b = readBorders(t, el);
        if (b !== undefined) out.borders = b;
        break;
      }
      case 'shd': {
        const v = shadingFill(t, el);
        if (v !== undefined) out.shd = v;
        break;
      }
      case 'framePr': {
        const rec: Record<string, string> = {};
        for (const a of el.attributes()) {
          if (a.uri === NS.w) rec[a.name.slice(a.name.indexOf(':') + 1)] = a.value;
        }
        out.framePr = rec;
        break;
      }
      default:
    }
  }
  return out;
}

/** `into` overridden by `from` (tabs and border sides merged entry by entry). */
export function mergePara(into: ParaLayer, from: ParaLayer): ParaLayer {
  const out: ParaLayer = { ...into };
  for (const key of Object.keys(from) as (keyof ParaLayer)[]) {
    const v = from[key];
    if (v === undefined) continue;
    if (key === 'tabs') {
      const m = new Map(into.tabs ?? []);
      for (const [pos, stop] of v as Map<number, TabStop>) m.set(pos, stop);
      out.tabs = m;
    } else if (key === 'borders') {
      out.borders = { ...(into.borders ?? {}), ...(v as ParaLayer['borders']) };
    } else {
      (out as Record<string, unknown>)[key] = v;
    }
  }
  return out;
}

/** Borders with removed sides dropped. */
export function finishBorders(
  b: Partial<Record<BorderSide, BorderSpec | null>> | undefined,
): Borders {
  const out: Borders = {};
  if (b === undefined) return out;
  for (const side of Object.keys(b) as BorderSide[]) {
    const v = b[side];
    if (v) out[side] = v;
  }
  return out;
}

const ALIGNMENTS: ReadonlySet<string> = new Set([
  'left',
  'center',
  'right',
  'both',
  'distribute',
  'mediumKashida',
  'highKashida',
  'lowKashida',
  'thaiDistribute',
]);

/**
 * `w:jc` → {@link Alignment}. Word reads `left`/`right` as logical (start/end) in a
 * right-to-left paragraph, like `start`/`end`, so those are mirrored when `bidi` is on.
 */
export function alignmentOf(jc: string | undefined, bidi: boolean): Alignment {
  const v = jc === 'justify' ? 'both' : (jc ?? 'start');
  if (v === 'start' || v === 'left') return bidi ? 'right' : 'left';
  if (v === 'end' || v === 'right') return bidi ? 'left' : 'right';
  return ALIGNMENTS.has(v) ? (v as Alignment) : bidi ? 'right' : 'left';
}

/** Effective tab stops: sorted by position, `clear` stops removed. */
export function finishTabs(tabs: Map<number, TabStop> | undefined): TabStop[] {
  if (tabs === undefined) return [];
  return [...tabs.values()].filter((s) => s.val !== 'clear').sort((a, b) => a.pos - b.pos);
}
