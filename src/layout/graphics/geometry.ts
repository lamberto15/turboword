// DrawingML shape geometry: evaluates guide formulas (ECMA-376 §20.1.9.11) and turns preset
// (a:prstGeom) or custom (a:custGeom) path lists into Paths in shape coordinates.
//
// Shape coordinates are the shape's extents in EMU (formulas such as `*/ ss a 100000` are
// scale-free, but a few use absolute EMU constants); the caller maps them to points with the
// shape's transform. Angles are in 60000ths of a degree, as in the XML.

import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import { Path } from './path.js';
import { PRESETS } from './presets.js';

/** One geometry path, already evaluated. */
export interface GeomPath {
  path: Path;
  /** Fill mode: 'norm', 'none' or a shading (darken, darkenLess, lighten, lightenLess). */
  fill: string;
  stroke: boolean;
}

interface PathDef {
  w: number;
  h: number;
  fill: string;
  stroke: boolean;
  /** Commands: [op, ...args] with args as guide names or numbers. */
  cmds: string[][];
}

interface GeomDef {
  av: [string, string][];
  gd: [string, string][];
  paths: PathDef[];
}

const FILL_CODE: Record<string, string> = {
  N: 'norm',
  n: 'none',
  d: 'darken',
  D: 'darkenLess',
  l: 'lighten',
  L: 'lightenLess',
};
const ARGS: Record<string, number> = { M: 2, L: 2, A: 4, Q: 4, C: 6, Z: 0 };

const parsed = new Map<string, GeomDef>();

function parsePreset(name: string): GeomDef | undefined {
  const hit = parsed.get(name);
  if (hit !== undefined) return hit;
  const src = PRESETS[name];
  if (src === undefined) return undefined;
  const [av = '', gd = '', ps = ''] = src.split('|');
  const pair = (s: string): [string, string] => {
    const i = s.indexOf(' ');
    return [s.slice(0, i), s.slice(i + 1)];
  };
  const def: GeomDef = {
    av:
      av === ''
        ? []
        : av.split(';').map((s) => {
            const [n, v] = pair(s);
            return [n, `val ${v}`];
          }),
    gd: gd === '' ? [] : gd.split(';').map(pair),
    paths: ps.split('!').map((p) => {
      const t = p.split(' ');
      const cmds: string[][] = [];
      let i = 4;
      while (i < t.length) {
        const op = t[i] ?? 'Z';
        const n = ARGS[op] ?? 0;
        cmds.push(t.slice(i, i + 1 + n));
        i += 1 + n;
      }
      return {
        w: Number(t[0]),
        h: Number(t[1]),
        fill: FILL_CODE[t[2] ?? 'N'] ?? 'norm',
        stroke: t[3] !== '0',
        cmds,
      };
    }),
  };
  parsed.set(name, def);
  return def;
}

/** Builds the custom geometry definition of an a:custGeom element. */
function parseCustom(el: XElement): GeomDef {
  const gds = (list: XElement | undefined): [string, string][] =>
    (list?.elements(NS.a, 'gd') ?? []).map((g) => [
      g.attr('name') ?? '',
      g.attr('fmla') ?? 'val 0',
    ]);
  const paths: PathDef[] = [];
  for (const p of el.element(NS.a, 'pathLst')?.elements(NS.a, 'path') ?? []) {
    const cmds: string[][] = [];
    for (const c of p.elements(NS.a)) {
      const local = c.name.slice(c.name.indexOf(':') + 1);
      const pts = c.elements(NS.a, 'pt').flatMap((q) => [q.attr('x') ?? '0', q.attr('y') ?? '0']);
      if (local === 'moveTo') cmds.push(['M', ...pts.slice(0, 2)]);
      else if (local === 'lnTo') cmds.push(['L', ...pts.slice(0, 2)]);
      else if (local === 'quadBezTo') cmds.push(['Q', ...pts.slice(0, 4)]);
      else if (local === 'cubicBezTo') cmds.push(['C', ...pts.slice(0, 6)]);
      else if (local === 'close') cmds.push(['Z']);
      else if (local === 'arcTo') {
        cmds.push([
          'A',
          c.attr('wR') ?? '0',
          c.attr('hR') ?? '0',
          c.attr('stAng') ?? '0',
          c.attr('swAng') ?? '0',
        ]);
      }
    }
    const fill = p.attr('fill') ?? 'norm';
    paths.push({
      w: Number(p.attr('w') ?? 0) || 0,
      h: Number(p.attr('h') ?? 0) || 0,
      fill,
      stroke: !['false', '0', 'off'].includes(p.attr('stroke') ?? 'true'),
      cmds,
    });
  }
  return { av: gds(el.element(NS.a, 'avLst')), gd: gds(el.element(NS.a, 'gdLst')), paths };
}

const DEG = Math.PI / 180 / 60000;

/** Built-in guide values for a w × h shape. */
function builtins(w: number, h: number): Map<string, number> {
  const ss = Math.min(w, h);
  const v = new Map<string, number>([
    ['w', w],
    ['h', h],
    ['l', 0],
    ['t', 0],
    ['r', w],
    ['b', h],
    ['hc', w / 2],
    ['vc', h / 2],
    ['ss', ss],
    ['ls', Math.max(w, h)],
    ['cd2', 10800000],
    ['cd4', 5400000],
    ['cd8', 2700000],
    ['3cd4', 16200000],
    ['3cd8', 8100000],
    ['5cd8', 13500000],
    ['7cd8', 18900000],
  ]);
  for (const d of [2, 3, 4, 5, 6, 8, 10, 12, 32]) {
    v.set(`wd${d}`, w / d);
    v.set(`hd${d}`, h / d);
  }
  for (const d of [2, 4, 6, 8, 16, 32]) v.set(`ssd${d}`, ss / d);
  return v;
}

/** Evaluates one guide formula. */
export function evalFormula(fmla: string, vars: ReadonlyMap<string, number>): number {
  const t = fmla.trim().split(/\s+/);
  const arg = (i: number): number => {
    const s = t[i];
    if (s === undefined) return 0;
    const n = Number(s);
    return Number.isNaN(n) ? (vars.get(s) ?? 0) : n;
  };
  const a = arg(1);
  const b = arg(2);
  const c = arg(3);
  let r: number;
  switch (t[0]) {
    case '*/':
      r = c === 0 ? 0 : (a * b) / c;
      break;
    case '+-':
      r = a + b - c;
      break;
    case '+/':
      r = c === 0 ? 0 : (a + b) / c;
      break;
    case '?:':
      r = a > 0 ? b : c;
      break;
    case 'abs':
      r = Math.abs(a);
      break;
    case 'at2':
      r = (Math.atan2(b, a) * 180 * 60000) / Math.PI;
      break;
    case 'cat2':
      r = a * Math.cos(Math.atan2(c, b));
      break;
    case 'cos':
      r = a * Math.cos(b * DEG);
      break;
    case 'max':
      r = Math.max(a, b);
      break;
    case 'min':
      r = Math.min(a, b);
      break;
    case 'mod':
      r = Math.sqrt(a * a + b * b + c * c);
      break;
    case 'pin':
      r = b < a ? a : Math.min(b, c);
      break;
    case 'sat2':
      r = a * Math.sin(Math.atan2(c, b));
      break;
    case 'sin':
      r = a * Math.sin(b * DEG);
      break;
    case 'sqrt':
      r = Math.sqrt(Math.max(0, a));
      break;
    case 'tan':
      r = a * Math.tan(b * DEG);
      break;
    default:
      r = a;
  }
  return Number.isFinite(r) ? r : 0;
}

function evaluate(
  def: GeomDef,
  w: number,
  h: number,
  adj: ReadonlyMap<string, string>,
): GeomPath[] {
  const vars = builtins(w, h);
  for (const [name, fmla] of def.av) vars.set(name, evalFormula(adj.get(name) ?? fmla, vars));
  for (const [name, fmla] of def.gd) vars.set(name, evalFormula(fmla, vars));
  const val = (s: string | undefined): number => {
    if (s === undefined) return 0;
    const n = Number(s);
    return Number.isNaN(n) ? (vars.get(s) ?? 0) : n;
  };
  const out: GeomPath[] = [];
  for (const p of def.paths) {
    const sx = p.w > 0 ? w / p.w : 1;
    const sy = p.h > 0 ? h / p.h : 1;
    const path = new Path();
    for (const c of p.cmds) {
      const x = (i: number): number => val(c[i]) * sx;
      const y = (i: number): number => val(c[i]) * sy;
      switch (c[0]) {
        case 'M':
          path.moveTo(x(1), y(2));
          break;
        case 'L':
          path.lineTo(x(1), y(2));
          break;
        case 'Q':
          path.quadTo(x(1), y(2), x(3), y(4));
          break;
        case 'C':
          path.cubicTo(x(1), y(2), x(3), y(4), x(5), y(6));
          break;
        case 'A':
          path.arcTo(x(1), y(2), val(c[3]) * DEG, val(c[4]) * DEG);
          break;
        case 'Z':
          path.close();
          break;
      }
    }
    out.push({ path, fill: p.fill, stroke: p.stroke });
  }
  return out;
}

/** Names of the supported presets. */
export function presetNames(): string[] {
  return Object.keys(PRESETS);
}

/**
 * Paths of preset `name` for a w × h shape (EMU), with adjust values from `adj` (guide name →
 * formula, from a:avLst). Undefined when the preset is not supported.
 */
export function presetGeometry(
  name: string,
  w: number,
  h: number,
  adj: ReadonlyMap<string, string> = new Map(),
): GeomPath[] | undefined {
  const def = parsePreset(name);
  return def === undefined ? undefined : evaluate(def, w, h, adj);
}

/** Paths of an a:custGeom for a w × h shape (EMU). */
export function customGeometry(el: XElement, w: number, h: number): GeomPath[] {
  return evaluate(parseCustom(el), w, h, new Map());
}

/** Adjust values of an a:prstGeom (its a:avLst). */
export function adjustValues(prstGeom: XElement): Map<string, string> {
  const m = new Map<string, string>();
  for (const g of prstGeom.element(NS.a, 'avLst')?.elements(NS.a, 'gd') ?? []) {
    const n = g.attr('name');
    const f = g.attr('fmla');
    if (n !== undefined && f !== undefined) m.set(n, f);
  }
  return m;
}
