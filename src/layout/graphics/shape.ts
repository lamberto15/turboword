// DrawingML shapes → display items: word-processing shapes (wps:wsp), groups (wpg:wgp and
// nested grpSp, with chOff/chExt child coordinate spaces), canvases (wpc:wpc), pictures inside
// groups, charts in group frames and SmartArt drawings (dsp:sp in the diagram drawing part).
//
// Geometry is evaluated in EMU in each shape's own box, then mapped through the chain of
// shape/group transforms to points relative to the drawing's box. Groups are walked with an
// explicit stack (documents may nest groups arbitrarily deep), in document order = z-order.
// DrawingML text bodies (a:txBody: SmartArt, shapes from other producers) are drawn with a
// simple greedy line breaker; Word's own text boxes (wps:txbx, w:p content) are laid out by the
// layout engine, not here.

import type { XElement } from '../../xml/dom.js';
import { alternateContent, selectBranch } from '../../xml/mc.js';
import { hasLocal, localOf, NS } from '../../xml/ns.js';
import type { LayoutEnv } from '../engine.js';
import type { DisplayItem, ImageItem } from '../types.js';
import { colorOf, type GraphicsTheme, graphicsTheme, shadeHex } from './color.js';
import { adjustValues, customGeometry, type GeomPath, presetGeometry } from './geometry.js';
import { addImage, blipPart, imageContentType } from './images.js';
import { fillOf, type LineStyle, lineOf, paintPath } from './paint.js';
import { apply, boxTransform, type Matrix, multiply, Path, scale, translate } from './path.js';
import { offsetItems, runProps, shapeText } from './text.js';

/** Hooks for content handled by sibling modules (charts, metafiles), injected by index.ts. */
export interface ShapeHooks {
  chart?: (
    chartRef: XElement,
    width: number,
    height: number,
    env: LayoutEnv,
  ) => DisplayItem[] | undefined;
  metafile?: (
    bytes: Uint8Array,
    contentType: string,
    width: number,
    height: number,
    env: LayoutEnv,
    id: string,
  ) => DisplayItem[] | undefined;
}

/** Namespaces whose mc:Choice branches graphics can read. */
export const UNDERSTOOD: ReadonlySet<string> = new Set<string>([
  NS.w,
  NS.wp,
  NS.a,
  NS.pic,
  NS.c,
  NS.wps,
  NS.wpg,
  NS.wpc,
  NS.wp14,
  NS.w14,
  NS.v,
  NS.o,
  NS.w10,
]);

/** First child element with local name `local` (any namespace). */
export function kid(el: XElement | undefined, local: string): XElement | undefined {
  if (el === undefined) return undefined;
  for (const c of el.children) {
    if (c.nodeType === 1 && hasLocal(c.name, local)) return c;
  }
  return undefined;
}

/** The mc:AlternateContent branch to read (its children). */
export function pickBranch(ac: XElement): XElement | undefined {
  return selectBranch(alternateContent(ac), (uri) => UNDERSTOOD.has(uri));
}

interface Xfrm {
  x: number;
  y: number;
  cx: number;
  cy: number;
  rot: number;
  flipH: boolean;
  flipV: boolean;
  chX: number;
  chY: number;
  chCx: number;
  chCy: number;
}

const num = (v: string | undefined, d = 0): number => {
  const n = Number(v);
  return v === undefined || !Number.isFinite(n) ? d : n;
};
const flag = (v: string | undefined): boolean => v === '1' || v === 'true' || v === 'on';

/** Reads an a:xfrm (or wpg:xfrm / p:xfrm). */
export function readXfrm(el: XElement | undefined): Xfrm {
  const off = kid(el, 'off');
  const ext = kid(el, 'ext');
  const chOff = kid(el, 'chOff');
  const chExt = kid(el, 'chExt');
  const cx = num(ext?.attr('cx'));
  const cy = num(ext?.attr('cy'));
  return {
    x: num(off?.attr('x')),
    y: num(off?.attr('y')),
    cx,
    cy,
    rot: num(el?.attr('rot')) / 60000,
    flipH: flag(el?.attr('flipH')),
    flipV: flag(el?.attr('flipV')),
    chX: num(chOff?.attr('x')),
    chY: num(chOff?.attr('y')),
    chCx: num(chExt?.attr('cx'), cx),
    chCy: num(chExt?.attr('cy'), cy),
  };
}

/** The transform-bearing properties element of a shape/group/picture/frame. */
function xfrmOf(el: XElement): Xfrm {
  const local = localOf(el.name);
  if (local === 'wgp' || local === 'grpSp') return readXfrm(kid(kid(el, 'grpSpPr'), 'xfrm'));
  if (local === 'graphicFrame') return readXfrm(kid(el, 'xfrm'));
  return readXfrm(kid(kid(el, 'spPr'), 'xfrm'));
}

/** Matrix mapping a root object's EMU coordinates onto the (width × height) point box. */
export function rootMatrix(root: XElement, width: number, height: number): Matrix {
  const local = localOf(root.name);
  if (local === 'wpc' || local === 'spTree' || local === 'drawing')
    return scale(1 / 12700, 1 / 12700);
  const xf = xfrmOf(root);
  const sx = xf.cx > 0 && width > 0 ? width / xf.cx : 1 / 12700;
  const sy = xf.cy > 0 && height > 0 ? height / xf.cy : 1 / 12700;
  return multiply(translate(-xf.x, -xf.y), scale(sx, sy));
}

/** Context shared while drawing one drawing. */
interface Ctx {
  env: LayoutEnv;
  theme: GraphicsTheme;
  hooks: ShapeHooks;
  out: DisplayItem[];
}

/**
 * Display items for a DrawingML object (wps:wsp, wpg:wgp, wpc:wpc, pic:pic, dsp:spTree), with
 * `m` mapping its parent coordinates (EMU) to points; default: fit the object to the box.
 */
export function drawingMLItems(
  root: XElement,
  width: number,
  height: number,
  env: LayoutEnv,
  hooks: ShapeHooks,
  m: Matrix = rootMatrix(root, width, height),
): DisplayItem[] {
  const ctx: Ctx = { env, theme: graphicsTheme(env.doc), hooks, out: [] };
  if (localOf(root.name) === 'wpc') {
    // Canvas background and outline cover the whole box.
    const bg = fillOf(kid(root, 'bg'), ctx.theme);
    const ln = lineOf(kid(kid(root, 'whole'), 'ln'), ctx.theme);
    ctx.out.push(...paintPath(new Path().rect(0, 0, width, height), bg ?? null, ln));
  }
  const stack: { el: XElement; m: Matrix; grpFill: string | null | undefined }[] = [
    { el: root, m, grpFill: undefined },
  ];
  let guard = 0;
  while (stack.length > 0 && guard++ < 100000) {
    const { el, m: pm, grpFill } = stack.pop() as (typeof stack)[number];
    const local = localOf(el.name);
    switch (local) {
      case 'wsp':
      case 'sp':
      case 'cxnSp':
        drawShape(el, pm, grpFill, ctx);
        break;
      case 'pic':
        drawPicture(el, pm, ctx);
        break;
      case 'graphicFrame':
        drawFrame(el, pm, ctx);
        break;
      case 'wgp':
      case 'grpSp':
      case 'wpc':
      case 'spTree':
      case 'drawing': {
        let cm = pm;
        let fill = grpFill;
        if (local === 'wgp' || local === 'grpSp') {
          const xf = xfrmOf(el);
          const sx = xf.chCx > 0 ? xf.cx / xf.chCx : 1;
          const sy = xf.chCy > 0 ? xf.cy / xf.chCy : 1;
          cm = multiply(
            multiply(translate(-xf.chX, -xf.chY), scale(sx, sy)),
            multiply(boxTransform(xf.x, xf.y, xf.cx, xf.cy, xf.rot, xf.flipH, xf.flipV), pm),
          );
          const f = fillOf(kid(el, 'grpSpPr'), ctx.theme, undefined, grpFill);
          if (f !== undefined) fill = f;
        }
        const kids = el.elements();
        for (let i = kids.length - 1; i >= 0; i--) {
          const k = kids[i] as XElement;
          stack.push({ el: k, m: cm, grpFill: fill });
        }
        break;
      }
      case 'AlternateContent': {
        const branch = el.uri === NS.mc ? pickBranch(el) : undefined;
        const kids = branch?.elements() ?? [];
        for (let i = kids.length - 1; i >= 0; i--) {
          stack.push({ el: kids[i] as XElement, m: pm, grpFill });
        }
        break;
      }
    }
  }
  return ctx.out;
}

/** Style-matrix fallbacks (wps:style / dsp:style): fill colour, outline, font colour. */
function styleRefs(style: XElement | undefined, theme: GraphicsTheme) {
  const fillRef = kid(style, 'fillRef');
  const lnRef = kid(style, 'lnRef');
  const fontRef = kid(style, 'fontRef');
  const fIdx = num(fillRef?.attr('idx'));
  const lIdx = num(lnRef?.attr('idx'));
  const fill =
    fillRef === undefined ? undefined : fIdx > 0 ? (colorOf(fillRef, theme) ?? null) : null;
  let line: LineStyle | null | undefined;
  if (lnRef !== undefined) {
    const c = colorOf(lnRef, theme);
    line =
      lIdx > 0 && c !== undefined ? { color: c, width: theme.lineWidths[lIdx - 1] ?? 0.75 } : null;
  }
  return { fill, line, font: colorOf(fontRef, theme) };
}

/** Geometry paths of a shape's spPr in its own EMU box (rect when unknown). */
export function shapeGeometry(
  spPr: XElement | undefined,
  cx: number,
  cy: number,
  warn?: (code: string, message: string) => void,
): GeomPath[] {
  const cust = kid(spPr, 'custGeom');
  if (cust !== undefined) return customGeometry(cust, cx, cy);
  const prst = kid(spPr, 'prstGeom');
  const name = prst?.attr('prst') ?? 'rect';
  const g = presetGeometry(name, cx, cy, prst === undefined ? undefined : adjustValues(prst));
  if (g !== undefined) return g;
  warn?.('UNSUPPORTED_PRESET_SHAPE', `Preset shape "${name}" drawn as a rectangle`);
  return presetGeometry('rect', cx, cy) ?? [];
}

function drawShape(el: XElement, pm: Matrix, grpFill: string | null | undefined, ctx: Ctx): void {
  const spPr = kid(el, 'spPr');
  const xf = xfrmOf(el);
  const m = multiply(boxTransform(xf.x, xf.y, xf.cx, xf.cy, xf.rot, xf.flipH, xf.flipV), pm);
  const refs = styleRefs(kid(el, 'style'), ctx.theme);
  let fill = fillOf(spPr, ctx.theme, undefined, grpFill);
  if (fill === undefined) fill = refs.fill ?? null;
  const line = lineOf(kid(spPr, 'ln'), ctx.theme, refs.line ?? null);
  for (const g of shapeGeometry(spPr, xf.cx, xf.cy, ctx.env.warn)) {
    const path = g.path.transform(m);
    const f =
      g.fill === 'none' || fill == null ? null : g.fill === 'norm' ? fill : shadeHex(fill, g.fill);
    ctx.out.push(...paintPath(path, f, g.stroke ? line : null));
  }
  const txBody = kid(el, 'txBody');
  if (txBody !== undefined) {
    const box = placedBox(m, xf.cx, xf.cy);
    ctx.out.push(...textBodyItems(txBody, box, ctx.theme, refs.font, ctx.env));
  }
}

/** Axis-aligned placement (x, y, w, h) and rotation of an EMU box under `m`, in points. */
function placedBox(m: Matrix, cx: number, cy: number) {
  const [x0, y0] = apply(m, 0, 0);
  const [x1, y1] = apply(m, cx, 0);
  const [x2, y2] = apply(m, 0, cy);
  const [mx, my] = apply(m, cx / 2, cy / 2);
  const w = Math.hypot(x1 - x0, y1 - y0);
  const h = Math.hypot(x2 - x0, y2 - y0);
  const rot = (Math.atan2(y1 - y0, x1 - x0) * 180) / Math.PI;
  return { x: mx - w / 2, y: my - h / 2, w, h, rot };
}

function drawPicture(el: XElement, pm: Matrix, ctx: Ctx): void {
  const xf = xfrmOf(el);
  const m = multiply(boxTransform(xf.x, xf.y, xf.cx, xf.cy, xf.rot, xf.flipH, xf.flipV), pm);
  const box = placedBox(m, xf.cx, xf.cy);
  const blipFill = kid(el, 'blipFill');
  const blip = kid(blipFill, 'blip');
  const part = blip === undefined ? undefined : blipPart(blip, ctx.env);
  if (part === undefined) return;
  const ct = imageContentType(part);
  if (ct === 'image/x-emf' || ct === 'image/x-wmf') {
    const items = ctx.hooks.metafile?.(part.bytes(), ct, box.w, box.h, ctx.env, part.name);
    if (items !== undefined) ctx.out.push(...offsetItems(items, box.x, box.y));
    return;
  }
  const id = addImage(ctx.env, part.name, ct, part.bytes());
  const item: ImageItem = {
    kind: 'image',
    x: box.x,
    y: box.y,
    width: box.w,
    height: box.h,
    image: id,
  };
  const src = kid(blipFill, 'srcRect');
  if (src !== undefined) {
    const f = (k: string): number => num(src.attr(k)) / 100000;
    item.crop = { left: f('l'), top: f('t'), right: f('r'), bottom: f('b') };
  }
  const rot = Math.round(box.rot * 1000) / 1000;
  if (rot !== 0) item.rotation = rot;
  ctx.out.push(item);
}

function drawFrame(el: XElement, pm: Matrix, ctx: Ctx): void {
  const xf = xfrmOf(el);
  const m = multiply(boxTransform(xf.x, xf.y, xf.cx, xf.cy, xf.rot, xf.flipH, xf.flipV), pm);
  const box = placedBox(m, xf.cx, xf.cy);
  const data = kid(kid(el, 'graphic'), 'graphicData');
  const chart = data?.element(NS.c, 'chart');
  if (chart !== undefined) {
    const items = ctx.hooks.chart?.(chart, box.w, box.h, ctx.env);
    if (items !== undefined) ctx.out.push(...offsetItems(items, box.x, box.y));
  }
}

// ---- DrawingML text bodies -------------------------------------------------------------------

interface Seg {
  text: string;
  props: ReturnType<typeof runProps>;
  color: string;
}

/**
 * Lays out an a:txBody (paragraphs of runs) inside `box` (points): insets, anchoring, alignment,
 * greedy word wrap. Text is not rotated with the shape.
 */
export function textBodyItems(
  txBody: XElement,
  box: { x: number; y: number; w: number; h: number },
  theme: GraphicsTheme,
  fontColor: string | undefined,
  env: LayoutEnv,
): DisplayItem[] {
  const bodyPr = kid(txBody, 'bodyPr');
  const ins = (k: string, d: number): number => num(bodyPr?.attr(k), d) / 12700;
  const left = box.x + ins('lIns', 91440);
  const right = box.x + box.w - ins('rIns', 91440);
  const top = box.y + ins('tIns', 45720);
  const bottom = box.y + box.h - ins('bIns', 45720);
  const avail = Math.max(1, right - left);
  const wrap = bodyPr?.attr('wrap') !== 'none';
  const scalePct = num(kid(bodyPr, 'normAutofit')?.attr('fontScale'), 100000) / 100000;
  const themeFont = (f: string | undefined): string | undefined =>
    f === '+mn-lt' ? theme.fonts.minor : f === '+mj-lt' ? theme.fonts.major : f;
  const lines: {
    items: DisplayItem[];
    width: number;
    ascent: number;
    descent: number;
    algn: string;
  }[] = [];
  for (const p of txBody.elements(NS.a, 'p')) {
    const algn = kid(p, 'pPr')?.attr('algn') ?? 'l';
    const segs: (Seg | 'br')[] = [];
    let endProps: Seg['props'] | undefined;
    for (const r of p.elements(NS.a)) {
      const local = localOf(r.name);
      if (local === 'br') {
        segs.push('br');
        continue;
      }
      const rPr =
        kid(r, local === 'endParaRPr' ? 'endParaRPr' : 'rPr') ??
        (local === 'endParaRPr' ? r : undefined);
      const font = themeFont(kid(rPr, 'latin')?.attr('typeface')) ?? theme.fonts.minor ?? 'Calibri';
      const props = runProps({
        font,
        size: (num(rPr?.attr('sz'), 1800) / 100) * scalePct,
        bold: flag(rPr?.attr('b')),
        italic: flag(rPr?.attr('i')),
      });
      if (local === 'endParaRPr') {
        endProps = props;
        continue;
      }
      if (local !== 'r' && local !== 'fld') continue;
      const color = colorOf(kid(rPr, 'solidFill'), theme) ?? fontColor ?? '000000';
      segs.push({ text: kid(r, 't')?.text ?? '', props, color });
    }
    // Greedy line filling, word by word.
    let line: (typeof lines)[number] = { items: [], width: 0, ascent: 0, descent: 0, algn };
    const flush = (): void => {
      lines.push(line);
      line = { items: [], width: 0, ascent: 0, descent: 0, algn };
    };
    let any = false;
    for (const s of segs) {
      if (s === 'br') {
        flush();
        continue;
      }
      for (const word of s.text.match(/\S+\s*|\s+/g) ?? []) {
        const shaped = shapeText(word, s.props, env, s.color);
        const trimmed = word.trimEnd();
        const ink = trimmed === word ? shaped.width : shapeText(trimmed, s.props, env).width;
        if (wrap && line.items.length > 0 && line.width + ink > avail) flush();
        line.items.push(...offsetItems(shaped.items, line.width, 0));
        line.width += shaped.width;
        line.ascent = Math.max(line.ascent, shaped.ascent);
        line.descent = Math.max(line.descent, shaped.descent);
        any = true;
      }
    }
    if (!any && line.items.length === 0) {
      const sz = endProps?.size ?? 18 * scalePct;
      line.ascent = sz * 0.8;
      line.descent = sz * 0.2;
    }
    flush();
  }
  const pitch = (l: (typeof lines)[number]): number => (l.ascent + l.descent) * 1.2;
  const total = lines.reduce((a, l) => a + pitch(l), 0);
  const anchor = bodyPr?.attr('anchor') ?? 't';
  let y = anchor === 'ctr' ? (top + bottom - total) / 2 : anchor === 'b' ? bottom - total : top;
  const out: DisplayItem[] = [];
  for (const l of lines) {
    const base = y + l.ascent + (pitch(l) - l.ascent - l.descent) / 2;
    const w = l.width;
    const x = l.algn === 'ctr' ? left + (avail - w) / 2 : l.algn === 'r' ? right - w : left;
    out.push(...offsetItems(l.items, x, base));
    y += pitch(l);
  }
  return out;
}
