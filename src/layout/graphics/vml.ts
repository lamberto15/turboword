// Legacy VML shapes (w:pict, w:object and mc:Fallback content of Word 2007-era files): v:rect,
// v:roundrect, v:oval, v:line, v:polyline, v:shape (path strings, with v:shapetype definitions,
// adjust values and v:formulas), v:group (coordsize/coordorigin child spaces) and v:imagedata.
//
// CSS `style` gives each shape's box in its parent's coordinate space (points, with units, at the
// top level; unitless group units inside a v:group). A shape's path lives in its own coordsize
// space, mapped onto that box, then rotated/flipped about the box centre. Groups are walked with
// an explicit stack. Text in v:textbox is laid out by the layout engine, not here.

import type { XDocument, XElement } from '../../xml/dom.js';
import { localOf, NS } from '../../xml/ns.js';
import type { LayoutEnv } from '../engine.js';
import type { DisplayItem, ImageItem } from '../types.js';
import { cssColor, hexToRgba, rgbaToHex } from './color.js';
import { addImage, blipPart, imageContentType } from './images.js';
import { type ArrowSpec, type LineStyle, paintPath } from './paint.js';
import { apply, boxTransform, type Matrix, multiply, Path, scale, translate } from './path.js';
import type { ShapeHooks } from './shape.js';
import { offsetItems } from './text.js';

const DRAWABLE = new Set([
  'group',
  'shape',
  'rect',
  'roundrect',
  'oval',
  'line',
  'polyline',
  'image',
]);

/** Parses a CSS style attribute into a map. */
function parseStyle(s: string | undefined): Map<string, string> {
  const m = new Map<string, string>();
  for (const part of (s ?? '').split(';')) {
    const i = part.indexOf(':');
    if (i > 0) m.set(part.slice(0, i).trim().toLowerCase(), part.slice(i + 1).trim());
  }
  return m;
}

const UNIT: Record<string, number> = {
  pt: 1,
  in: 72,
  cm: 72 / 2.54,
  mm: 72 / 25.4,
  px: 0.75,
  pc: 12,
  em: 12,
  emu: 1 / 12700,
};

/** A CSS/VML length → points (`unitless` multiplier applies to bare numbers). */
export function vmlLength(v: string | undefined, unitless = 1): number | undefined {
  if (v === undefined) return undefined;
  const m = /^\s*(-?[\d.]+(?:e-?\d+)?)\s*([a-z]*)\s*$/i.exec(v);
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  const u = (m[2] ?? '').toLowerCase();
  if (u === '') return n * unitless;
  if (u === 'f') return n / 65536;
  return n * (UNIT[u] ?? 1);
}

/** A VML fraction/number value ("0.5", "32768f", "50%"). */
function vmlFraction(v: string | undefined, d: number): number {
  if (v === undefined || v.trim() === '') return d;
  const s = v.trim();
  if (s.endsWith('f')) return Number(s.slice(0, -1)) / 65536;
  if (s.endsWith('%')) return Number(s.slice(0, -1)) / 100;
  const n = Number(s);
  return Number.isFinite(n) ? n : d;
}

const on = (v: string | undefined, d: boolean): boolean =>
  v === undefined ? d : !['f', 'false', '0', 'off'].includes(v.trim().toLowerCase());

const pair = (v: string | undefined, d: [number, number], unitless = 1): [number, number] => {
  if (v === undefined) return d;
  const [a, b] = v.split(',');
  return [
    a === undefined || a.trim() === '' ? 0 : (vmlLength(a, unitless) ?? d[0]),
    b === undefined || b.trim() === '' ? 0 : (vmlLength(b, unitless) ?? d[1]),
  ];
};

// ---- shapetypes and formulas --------------------------------------------------------------------

const typeCache = new WeakMap<XDocument, { version: number; map: Map<string, XElement> }>();

/** The v:shapetype a `type="#id"` refers to (searched in the element's own part). */
function shapetype(el: XElement): XElement | undefined {
  const ref = el.attr('type');
  if (ref === undefined || !ref.startsWith('#')) return undefined;
  const doc = el.ownerDocument();
  if (doc === undefined) return undefined;
  let hit = typeCache.get(doc);
  if (hit === undefined || hit.version !== doc.version) {
    const map = new Map<string, XElement>();
    for (const t of doc.root.descendants(NS.v, 'shapetype')) {
      const id = t.attr('id');
      if (id !== undefined && !map.has(id)) map.set(id, t);
    }
    hit = { version: doc.version, map };
    typeCache.set(doc, hit);
  }
  return hit.map.get(ref.slice(1));
}

const FD = Math.PI / 180 / 65536;

/** Evaluates a shapetype's v:formulas for the given adjust values. */
function formulas(
  list: XElement | undefined,
  adj: number[],
  vars: ReadonlyMap<string, number>,
): number[] {
  const out: number[] = [];
  for (const f of list?.elements(NS.v, 'f') ?? []) {
    const t = (f.attr('eqn') ?? '').trim().split(/[\s,]+/);
    const arg = (i: number): number => {
      const s = t[i] ?? '0';
      if (s.startsWith('#')) return adj[Number(s.slice(1))] ?? 0;
      if (s.startsWith('@')) return out[Number(s.slice(1))] ?? 0;
      const n = Number(s);
      return Number.isNaN(n) ? (vars.get(s) ?? 0) : n;
    };
    const a = arg(1);
    const b = arg(2);
    const c = arg(3);
    let r: number;
    switch (t[0]) {
      case 'sum':
        r = a + b - c;
        break;
      case 'product':
        r = c === 0 ? 0 : (a * b) / c;
        break;
      case 'mid':
        r = (a + b) / 2;
        break;
      case 'abs':
        r = Math.abs(a);
        break;
      case 'min':
        r = Math.min(a, b);
        break;
      case 'max':
        r = Math.max(a, b);
        break;
      case 'if':
        r = a > 0 ? b : c;
        break;
      case 'mod':
        r = Math.sqrt(a * a + b * b + c * c);
        break;
      case 'atan2':
        r = (Math.atan2(b, a) * 180 * 65536) / Math.PI;
        break;
      case 'sin':
        r = a * Math.sin(b * FD);
        break;
      case 'cos':
        r = a * Math.cos(b * FD);
        break;
      case 'tan':
        r = a * Math.tan(b * FD);
        break;
      case 'cosatan2':
        r = a * Math.cos(Math.atan2(c, b));
        break;
      case 'sinatan2':
        r = a * Math.sin(Math.atan2(c, b));
        break;
      case 'sqrt':
        r = Math.sqrt(Math.max(0, a));
        break;
      case 'sumangle':
        r = a + b * 65536 - c * 65536;
        break;
      case 'ellipse':
        r = b === 0 ? 0 : c * Math.sqrt(Math.max(0, 1 - (a / b) ** 2));
        break;
      default:
        r = a;
    }
    out.push(Number.isFinite(r) ? r : 0);
  }
  return out;
}

// ---- path strings ------------------------------------------------------------------------------

const TWO = new Set(['nf', 'ns', 'ae', 'al', 'at', 'ar', 'wa', 'wr', 'qx', 'qy', 'qb']);

/** Tokenises a VML path into commands with raw parameter strings (empty slots → '0'). */
function tokenize(s: string): { cmd: string; args: string[] }[] {
  const out: { cmd: string; args: string[] }[] = [];
  let cur: { cmd: string; args: string[] } | undefined;
  let lastValue = false;
  let i = 0;
  while (i < s.length) {
    const ch = s[i] ?? '';
    if (/[a-z]/i.test(ch)) {
      const two = s.slice(i, i + 2).toLowerCase();
      const cmd = TWO.has(two) ? two : /^h[a-i]$/.test(two) ? two : ch.toLowerCase();
      i += cmd.length;
      cur = { cmd, args: [] };
      out.push(cur);
      lastValue = false;
    } else if (ch === ',') {
      if (!lastValue) cur?.args.push('0');
      lastValue = false;
      i++;
    } else if (/[-+\d.@#]/.test(ch)) {
      const m = /^(?:[@#]\d+|[-+]?\d*\.?\d+(?:e[-+]?\d+)?)/i.exec(s.slice(i));
      if (!m || m[0] === '') {
        i++;
        continue;
      }
      cur?.args.push(m[0]);
      lastValue = true;
      i += m[0].length;
    } else i++;
  }
  return out;
}

interface SubPath {
  path: Path;
  fill: boolean;
  stroke: boolean;
}

/** Builds paths from a VML path string; values resolve through `val`. */
function vmlPath(s: string, val: (raw: string) => number): SubPath[] {
  const groups: SubPath[] = [];
  let g: SubPath = { path: new Path(), fill: true, stroke: true };
  const K = 0.5522847498;
  for (const { cmd, args } of tokenize(s)) {
    const v = args.map(val);
    const p = g.path;
    const n = (i: number): number => v[i] ?? 0;
    switch (cmd) {
      case 'm':
        p.moveTo(n(0), n(1));
        break;
      case 't':
        p.moveTo(p.x + n(0), p.y + n(1));
        break;
      case 'l':
        for (let i = 0; i + 1 < Math.max(2, v.length); i += 2) p.lineTo(n(i), n(i + 1));
        break;
      case 'r':
        for (let i = 0; i + 1 < Math.max(2, v.length); i += 2) p.lineTo(p.x + n(i), p.y + n(i + 1));
        break;
      case 'c':
        for (let i = 0; i + 5 < Math.max(6, v.length); i += 6) {
          p.cubicTo(n(i), n(i + 1), n(i + 2), n(i + 3), n(i + 4), n(i + 5));
        }
        break;
      case 'v':
        for (let i = 0; i + 5 < Math.max(6, v.length); i += 6) {
          const x = p.x;
          const y = p.y;
          p.cubicTo(x + n(i), y + n(i + 1), x + n(i + 2), y + n(i + 3), x + n(i + 4), y + n(i + 5));
        }
        break;
      case 'x':
        p.close();
        break;
      case 'e':
        if (!p.empty) groups.push(g);
        g = { path: new Path(), fill: true, stroke: true };
        break;
      case 'nf':
        g.fill = false;
        break;
      case 'ns':
        g.stroke = false;
        break;
      case 'ae':
      case 'al':
        for (let i = 0; i + 5 < Math.max(6, v.length); i += 6) {
          const t0 = n(i + 4) * FD;
          const sx = n(i) + n(i + 2) * Math.cos(t0);
          const sy = n(i + 1) + n(i + 3) * Math.sin(t0);
          if (cmd === 'al' && i === 0) p.moveTo(sx, sy);
          p.ellipseArc(n(i), n(i + 1), n(i + 2), n(i + 3), t0, n(i + 5) * FD);
        }
        break;
      case 'at':
      case 'ar':
      case 'wa':
      case 'wr':
        for (let i = 0; i + 7 < Math.max(8, v.length); i += 8) {
          const cx = (n(i) + n(i + 2)) / 2;
          const cy = (n(i + 1) + n(i + 3)) / 2;
          const rx = Math.abs(n(i + 2) - n(i)) / 2 || 1e-9;
          const ry = Math.abs(n(i + 3) - n(i + 1)) / 2 || 1e-9;
          const t1 = Math.atan2((n(i + 5) - cy) / ry, (n(i + 4) - cx) / rx);
          const t2 = Math.atan2((n(i + 7) - cy) / ry, (n(i + 6) - cx) / rx);
          const cw = cmd === 'wa' || cmd === 'wr';
          let dt = t2 - t1;
          if (cw) while (dt <= 0) dt += 2 * Math.PI;
          else while (dt >= 0) dt -= 2 * Math.PI;
          if (cmd === 'ar' || cmd === 'wr') {
            p.close();
            p.moveTo(cx + rx * Math.cos(t1), cy + ry * Math.sin(t1));
          }
          p.ellipseArc(cx, cy, rx, ry, t1, dt);
        }
        break;
      case 'qx':
      case 'qy': {
        let horiz = cmd === 'qx';
        for (let i = 0; i + 1 < Math.max(2, v.length); i += 2) {
          const x0 = p.x;
          const y0 = p.y;
          const x = n(i);
          const y = n(i + 1);
          if (horiz) p.cubicTo(x0 + K * (x - x0), y0, x, y - K * (y - y0), x, y);
          else p.cubicTo(x0, y0 + K * (y - y0), x - K * (x - x0), y, x, y);
          horiz = !horiz;
        }
        break;
      }
      case 'qb': {
        const pts: [number, number][] = [];
        for (let i = 0; i + 1 < v.length; i += 2) pts.push([n(i), n(i + 1)]);
        const end = pts.pop();
        if (end === undefined) break;
        if (pts.length === 0) p.lineTo(end[0], end[1]);
        for (let k = 0; k < pts.length; k++) {
          const [cx, cy] = pts[k] as [number, number];
          const next = pts[k + 1];
          const [ex, ey] = next === undefined ? end : [(cx + next[0]) / 2, (cy + next[1]) / 2];
          p.quadTo(cx, cy, ex, ey);
        }
        break;
      }
    }
  }
  if (!g.path.empty) groups.push(g);
  return groups;
}

// ---- paint ------------------------------------------------------------------------------------

const VML_DASH: Record<string, number[]> = {
  shortdash: [3, 1],
  shortdot: [1, 1],
  shortdashdot: [3, 1, 1, 1],
  shortdashdotdot: [3, 1, 1, 1, 1, 1],
  dot: [1, 3],
  dash: [4, 3],
  longdash: [8, 3],
  dashdot: [4, 3, 1, 3],
  longdashdot: [8, 3, 1, 3],
  longdashdotdot: [8, 3, 1, 3, 1, 3],
};
const VML_ARROW: Record<string, ArrowSpec['type']> = {
  block: 'triangle',
  classic: 'stealth',
  diamond: 'diamond',
  oval: 'oval',
  open: 'arrow',
};
const ARROW_SIZE: Record<string, number> = { narrow: 2, short: 2, medium: 3, wide: 5, long: 5 };

function vmlArrow(stroke: XElement | undefined, which: 'start' | 'end'): ArrowSpec | undefined {
  const t = VML_ARROW[stroke?.attr(`${which}arrow`) ?? 'none'];
  if (t === undefined) return undefined;
  return {
    type: t,
    w: ARROW_SIZE[stroke?.attr(`${which}arrowwidth`) ?? 'medium'] ?? 3,
    len: ARROW_SIZE[stroke?.attr(`${which}arrowlength`) ?? 'medium'] ?? 3,
  };
}

/** Fill colour of a VML shape (null = not filled). */
function vmlFill(el: XElement, type: XElement | undefined): string | null {
  const f = el.element(NS.v, 'fill') ?? type?.element(NS.v, 'fill');
  if (!on(f?.attr('on'), on(el.attr('filled') ?? type?.attr('filled'), true))) return null;
  let c = cssColor(f?.attr('color') ?? el.attr('fillcolor') ?? type?.attr('fillcolor')) ?? 'FFFFFF';
  const kind = f?.attr('type') ?? 'solid';
  if (kind.startsWith('gradient')) {
    const c2 = cssColor(f?.attr('color2'));
    if (c2 !== undefined) {
      const a = hexToRgba(c);
      const b = hexToRgba(c2);
      c = rgbaToHex({ r: (a.r + b.r) / 2, g: (a.g + b.g) / 2, b: (a.b + b.b) / 2, a: 1 });
    }
  }
  const op = f?.attr('opacity');
  if (op !== undefined) c = rgbaToHex({ ...hexToRgba(c), a: vmlFraction(op, 1) });
  return c;
}

/** Outline of a VML shape (null = not stroked). */
function vmlLine(el: XElement, type: XElement | undefined): LineStyle | null {
  const s = el.element(NS.v, 'stroke') ?? type?.element(NS.v, 'stroke');
  if (!on(s?.attr('on'), on(el.attr('stroked') ?? type?.attr('stroked'), true))) return null;
  const color =
    cssColor(s?.attr('color') ?? el.attr('strokecolor') ?? type?.attr('strokecolor')) ?? '000000';
  const width =
    vmlLength(
      s?.attr('weight') ?? el.attr('strokeweight') ?? type?.attr('strokeweight'),
      1 / 12700,
    ) ?? 0.75;
  const line: LineStyle = { color, width };
  const ds = s?.attr('dashstyle');
  if (ds !== undefined && ds !== 'solid') {
    const nums = ds.trim().split(/\s+/).map(Number);
    const pattern = VML_DASH[ds] ?? (nums.every((x) => Number.isFinite(x)) ? nums : undefined);
    if (pattern !== undefined && pattern.length >= 2) line.dash = pattern;
  }
  const head = vmlArrow(s, 'start');
  const tail = vmlArrow(s, 'end');
  if (head !== undefined) line.head = head;
  if (tail !== undefined) line.tail = tail;
  return line;
}

// ---- walking ----------------------------------------------------------------------------------

/** Box of a VML element in its parent's coordinate space (`unitless` for bare numbers). */
function cssBox(el: XElement, unitless: number) {
  const st = parseStyle(el.attr('style'));
  const len = (k: string): number | undefined => vmlLength(st.get(k), unitless);
  const x = len('left') ?? 0;
  const y = len('top') ?? 0;
  const rotRaw = st.get('rotation');
  const flip = st.get('flip') ?? '';
  return {
    x: x + (len('margin-left') ?? 0),
    y: y + (len('margin-top') ?? 0),
    w: len('width') ?? 0,
    h: len('height') ?? 0,
    rot: rotRaw === undefined ? 0 : vmlFraction(rotRaw.replace(/fd$/, 'f'), 0),
    flipH: flip.includes('x'),
    flipV: flip.includes('y'),
  };
}

/** Points of v:line / v:polyline in parent coordinates. */
function linePoints(el: XElement, unitless: number): number[] {
  const local = localOf(el.name);
  if (local === 'line') {
    return [...pair(el.attr('from'), [0, 0], unitless), ...pair(el.attr('to'), [10, 10], unitless)];
  }
  const raw = (el.attr('points') ?? '').split(/[\s,]+/).filter((t) => t !== '');
  return raw.map((t) => vmlLength(t, unitless) ?? 0);
}

/** VML items for a container (w:pict, w:object) or a single v:* element. */
export function vmlItems(
  root: XElement,
  width: number,
  height: number,
  env: LayoutEnv,
  hooks: ShapeHooks,
): DisplayItem[] | undefined {
  const tops =
    root.uri === NS.v ? [root] : root.elements(NS.v).filter((e) => DRAWABLE.has(localOf(e.name)));
  if (tops.length === 0) return undefined;
  // Fit the top-level boxes (points) onto the (width × height) box.
  const boxes = tops.map((t) => {
    const local = localOf(t.name);
    if (local === 'line' || local === 'polyline') {
      const pts = linePoints(t, 0.75);
      const xs = pts.filter((_, i) => i % 2 === 0);
      const ys = pts.filter((_, i) => i % 2 === 1);
      const x = Math.min(...xs);
      const y = Math.min(...ys);
      return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
    }
    return cssBox(t, 0.75);
  });
  const x0 = Math.min(...boxes.map((b) => b.x));
  const y0 = Math.min(...boxes.map((b) => b.y));
  const bw = Math.max(...boxes.map((b) => b.x + b.w)) - x0;
  const bh = Math.max(...boxes.map((b) => b.y + b.h)) - y0;
  const m0 = multiply(
    translate(-x0, -y0),
    scale(bw > 0 && width > 0 ? width / bw : 1, bh > 0 && height > 0 ? height / bh : 1),
  );
  const out: DisplayItem[] = [];
  const stack: { el: XElement; m: Matrix; unitless: number }[] = [];
  for (let i = tops.length - 1; i >= 0; i--)
    stack.push({ el: tops[i] as XElement, m: m0, unitless: 0.75 });
  let guard = 0;
  while (stack.length > 0 && guard++ < 100000) {
    const { el, m, unitless } = stack.pop() as (typeof stack)[number];
    const local = localOf(el.name);
    if (local === 'line' || local === 'polyline') {
      const pts = linePoints(el, unitless);
      const path = new Path().poly(pts, false).transform(m);
      const closed =
        local === 'polyline' &&
        pts.length > 4 &&
        pts[0] === pts[pts.length - 2] &&
        pts[1] === pts[pts.length - 1];
      out.push(...paintPath(path, closed ? vmlFill(el, undefined) : null, vmlLine(el, undefined)));
      continue;
    }
    const b = cssBox(el, unitless);
    const place = multiply(boxTransform(b.x, b.y, b.w, b.h, b.rot, b.flipH, b.flipV), m);
    const type = shapetype(el);
    const [cw, ch] = pair(el.attr('coordsize') ?? type?.attr('coordsize'), [1000, 1000]);
    const [ox, oy] = pair(el.attr('coordorigin') ?? type?.attr('coordorigin'), [0, 0]);
    const inner = multiply(
      multiply(translate(-ox, -oy), scale(cw !== 0 ? b.w / cw : 1, ch !== 0 ? b.h / ch : 1)),
      place,
    );
    if (local === 'group') {
      const kids = el.elements(NS.v).filter((e) => DRAWABLE.has(localOf(e.name)));
      for (let i = kids.length - 1; i >= 0; i--)
        stack.push({ el: kids[i] as XElement, m: inner, unitless: 1 });
      continue;
    }
    const img = el.element(NS.v, 'imagedata');
    if (img !== undefined || local === 'image') {
      drawImage(img ?? el, place, b.w, b.h, env, hooks, out);
      if (local === 'image') continue;
    }
    const fill = vmlFill(el, type);
    const line = vmlLine(el, type);
    if (local === 'rect' || local === 'oval' || local === 'roundrect') {
      const p = new Path();
      if (local === 'rect') p.rect(0, 0, b.w, b.h);
      else if (local === 'oval') p.ellipse(0, 0, b.w, b.h);
      else roundRect(p, b.w, b.h, vmlFraction(el.attr('arcsize'), 0.2) * Math.min(b.w, b.h));
      out.push(...paintPath(p.transform(place), fill, line));
      continue;
    }
    // v:shape: a path (own, v:path child or shapetype) or the bounding rectangle.
    const src =
      el.attr('path') ??
      el.element(NS.v, 'path')?.attr('v') ??
      type?.attr('path') ??
      type?.element(NS.v, 'path')?.attr('v');
    if (src === undefined) {
      if (img === undefined)
        out.push(...paintPath(new Path().rect(0, 0, b.w, b.h).transform(place), fill, line));
      continue;
    }
    const defAdj = (type?.attr('adj') ?? '').split(',');
    const ownAdj = (el.attr('adj') ?? '').split(',');
    const adj: number[] = [];
    for (let i = 0; i < Math.max(defAdj.length, ownAdj.length); i++) {
      const s = ownAdj[i]?.trim() || defAdj[i]?.trim() || '0';
      adj.push(Number(s) || 0);
    }
    const vars = new Map<string, number>([
      ['width', cw],
      ['height', ch],
      ['xcenter', ox + cw / 2],
      ['ycenter', oy + ch / 2],
      ['hasfill', fill === null ? 0 : 1],
      ['hasstroke', line === null ? 0 : 1],
      ['pixelWidth', b.w / 0.75],
      ['pixelHeight', b.h / 0.75],
      ['emuWidth', b.w * 12700],
      ['emuHeight', b.h * 12700],
      ['emuWidth2', b.w * 6350],
      ['emuHeight2', b.h * 6350],
      ['pixelLineWidth', (line?.width ?? 0.75) / 0.75],
      ['lineDrawn', line === null ? 0 : 1],
    ]);
    const fv = formulas(el.element(NS.v, 'formulas') ?? type?.element(NS.v, 'formulas'), adj, vars);
    const val = (raw: string): number => {
      if (raw.startsWith('@')) return fv[Number(raw.slice(1))] ?? 0;
      if (raw.startsWith('#')) return adj[Number(raw.slice(1))] ?? 0;
      return Number(raw) || 0;
    };
    for (const g of vmlPath(src, val)) {
      out.push(...paintPath(g.path.transform(inner), g.fill ? fill : null, g.stroke ? line : null));
    }
  }
  return out;
}

function roundRect(p: Path, w: number, h: number, r0: number): void {
  const r = Math.max(0, Math.min(r0, w / 2, h / 2));
  p.moveTo(r, 0)
    .lineTo(w - r, 0)
    .arcTo(r, r, -Math.PI / 2, Math.PI / 2)
    .lineTo(w, h - r)
    .arcTo(r, r, 0, Math.PI / 2)
    .lineTo(r, h)
    .arcTo(r, r, Math.PI / 2, Math.PI / 2)
    .lineTo(0, r)
    .arcTo(r, r, Math.PI, Math.PI / 2)
    .close();
}

/** v:imagedata (or v:image) filling the shape box. */
function drawImage(
  el: XElement,
  place: Matrix,
  w: number,
  h: number,
  env: LayoutEnv,
  hooks: ShapeHooks,
  out: DisplayItem[],
): void {
  const id = el.attrNS(NS.o, 'relid') ?? el.attrNS(NS.r, 'id') ?? el.attrNS(NS.r, 'pict');
  const part = id === undefined ? undefined : (blipPart(el, env) ?? env.partOf(el)?.related(id));
  if (part === undefined) return;
  const [x0, y0] = apply(place, 0, 0);
  const [x1, y1] = apply(place, w, 0);
  const [x2, y2] = apply(place, 0, h);
  const [cx, cy] = apply(place, w / 2, h / 2);
  const bw = Math.hypot(x1 - x0, y1 - y0);
  const bh = Math.hypot(x2 - x0, y2 - y0);
  const ct = imageContentType(part);
  if (ct === 'image/x-emf' || ct === 'image/x-wmf') {
    const items = hooks.metafile?.(part.bytes(), ct, bw, bh, env, part.name);
    if (items !== undefined) out.push(...offsetItems(items, cx - bw / 2, cy - bh / 2));
    return;
  }
  const image = addImage(env, part.name, ct, part.bytes());
  const item: ImageItem = {
    kind: 'image',
    x: cx - bw / 2,
    y: cy - bh / 2,
    width: bw,
    height: bh,
    image,
  };
  const crop = ['cropleft', 'croptop', 'cropright', 'cropbottom'].map((k) =>
    vmlFraction(el.attr(k), 0),
  );
  if (crop.some((c) => c !== 0)) {
    item.crop = {
      left: crop[0] ?? 0,
      top: crop[1] ?? 0,
      right: crop[2] ?? 0,
      bottom: crop[3] ?? 0,
    };
  }
  const rot = Math.round(((Math.atan2(y1 - y0, x1 - x0) * 180) / Math.PI) * 1000) / 1000;
  if (rot !== 0) item.rotation = rot;
  out.push(item);
}
