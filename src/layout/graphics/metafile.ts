// The GDI drawing state shared by the EMF and WMF players: device context (mapping, world
// transform, selected pen/brush/font, text settings, current position), the save/restore stack,
// path brackets (BEGINPATH … FILLPATH) and the conversion of GDI primitives into DisplayItems.
//
// Primitives arrive in logical coordinates. `matrix()` maps logical → points through the
// player-supplied mapping (EMF: world → page → device → picture frame; WMF: window → frame),
// so ellipses and arcs are built in logical space and transformed exactly. Consecutive LINETO /
// POLYLINETO strokes with the same pen are merged into one path until another record arrives.

import type { LayoutEnv } from '../engine.js';
import type { DisplayItem, ImageItem, TextItem } from '../types.js';
import { averageDibColor, type DecodedDib } from './dib.js';
import { addImage } from './images.js';
import { type LineStyle, paintPath, strokeItems } from './paint.js';
import { apply, IDENTITY, type Matrix, matrixScale, multiply, Path } from './path.js';
import { encodePng } from './png.js';
import { offsetItems, runProps, shapeText } from './text.js';

/** Bounds-checked little-endian reader over one record (out-of-range reads return 0). */
export class Reader {
  readonly b: Uint8Array;
  readonly #dv: DataView;
  constructor(b: Uint8Array) {
    this.b = b;
    this.#dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  }
  get length(): number {
    return this.b.length;
  }
  u8(o: number): number {
    return o >= 0 && o < this.b.length ? (this.b[o] ?? 0) : 0;
  }
  u16(o: number): number {
    return o >= 0 && o + 2 <= this.b.length ? this.#dv.getUint16(o, true) : 0;
  }
  i16(o: number): number {
    return o >= 0 && o + 2 <= this.b.length ? this.#dv.getInt16(o, true) : 0;
  }
  u32(o: number): number {
    return o >= 0 && o + 4 <= this.b.length ? this.#dv.getUint32(o, true) : 0;
  }
  i32(o: number): number {
    return o >= 0 && o + 4 <= this.b.length ? this.#dv.getInt32(o, true) : 0;
  }
  f32(o: number): number {
    const v = o >= 0 && o + 4 <= this.b.length ? this.#dv.getFloat32(o, true) : 0;
    return Number.isFinite(v) ? v : 0;
  }
  /** Bytes [o, o + n) clamped to the record. */
  slice(o: number, n: number): Uint8Array {
    const a = Math.max(0, Math.min(o, this.b.length));
    return this.b.subarray(a, Math.max(a, Math.min(a + Math.max(0, n), this.b.length)));
  }
}

/** COLORREF (0x00BBGGRR) → RRGGBB. */
export function colorRef(v: number): string {
  const r = v & 255;
  const g = (v >>> 8) & 255;
  const b = (v >>> 16) & 255;
  return ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0').toUpperCase();
}

export interface GdiPen {
  kind: 'pen';
  /** PS_* line style (low 4 bits). */
  style: number;
  /** Logical units; 0 = one device pixel. */
  width: number;
  /** Cosmetic pens are always one device pixel wide. */
  cosmetic: boolean;
  color: string;
}

export interface GdiBrush {
  kind: 'brush';
  /** null = hollow. */
  color: string | null;
}

export interface GdiFont {
  kind: 'font';
  /** LOGFONT height in logical units (< 0: character height, > 0: cell height, 0: default). */
  height: number;
  weight: number;
  italic: boolean;
  face: string;
}

export type GdiObject = GdiPen | GdiBrush | GdiFont | { kind: 'other' };

export const PS_NULL = 5;
const BS_SOLID = 0;
const BS_NULL = 1;
const BS_HATCHED = 2;

/** LOGBRUSH → brush (hatched brushes are approximated by their colour). */
export function makeBrush(style: number, color: number): GdiBrush {
  if (style === BS_NULL) return { kind: 'brush', color: null };
  if (style === BS_SOLID || style === BS_HATCHED) return { kind: 'brush', color: colorRef(color) };
  return { kind: 'brush', color: null };
}

/** A pattern brush approximated by the average colour of its bitmap. */
export function patternBrush(dib: DecodedDib | undefined): GdiBrush {
  const c = dib === undefined ? undefined : averageDibColor(dib);
  return { kind: 'brush', color: c ?? null };
}

/** EMF stock objects (index & 0x7FFFFFFF). */
export function stockObject(i: number): GdiObject | undefined {
  switch (i) {
    case 0:
      return { kind: 'brush', color: 'FFFFFF' };
    case 1:
      return { kind: 'brush', color: 'C0C0C0' };
    case 2:
      return { kind: 'brush', color: '808080' };
    case 3:
      return { kind: 'brush', color: '404040' };
    case 4:
      return { kind: 'brush', color: '000000' };
    case 5:
      return { kind: 'brush', color: null };
    case 6:
      return { kind: 'pen', style: 0, width: 0, cosmetic: true, color: 'FFFFFF' };
    case 7:
      return { kind: 'pen', style: 0, width: 0, cosmetic: true, color: '000000' };
    case 8:
      return { kind: 'pen', style: PS_NULL, width: 0, cosmetic: true, color: '000000' };
    case 10:
    case 11:
    case 16:
      return { kind: 'font', height: 0, weight: 400, italic: false, face: 'Courier New' };
    case 12:
    case 13:
    case 14:
    case 17:
      return { kind: 'font', height: 0, weight: 400, italic: false, face: 'Arial' };
    case 18:
      return { kind: 'brush', color: 'FFFFFF' };
    case 19:
      return { kind: 'pen', style: 0, width: 0, cosmetic: true, color: '000000' };
    default:
      return undefined;
  }
}

/** Mapping mode constants. */
export const MM_TEXT = 1;
export const MM_ISOTROPIC = 7;
export const MM_ANISOTROPIC = 8;

/** One device context. */
export interface DcState {
  winOrg: [number, number];
  winExt: [number, number];
  winExtSet: boolean;
  vpOrg: [number, number];
  vpExt: [number, number];
  mapMode: number;
  world: Matrix;
  pen: GdiPen;
  brush: GdiBrush;
  font: GdiFont;
  textColor: string;
  bkColor: string;
  bkMode: number;
  textAlign: number;
  /** 1 = ALTERNATE (even-odd), 2 = WINDING. */
  fillMode: number;
  cur: [number, number];
}

function defaultDc(): DcState {
  return {
    winOrg: [0, 0],
    winExt: [1, 1],
    winExtSet: false,
    vpOrg: [0, 0],
    vpExt: [1, 1],
    mapMode: MM_TEXT,
    world: IDENTITY,
    pen: { kind: 'pen', style: 0, width: 0, cosmetic: true, color: '000000' },
    brush: { kind: 'brush', color: 'FFFFFF' },
    font: { kind: 'font', height: 0, weight: 400, italic: false, face: 'Arial' },
    textColor: '000000',
    bkColor: 'FFFFFF',
    bkMode: 2,
    textAlign: 0,
    fillMode: 1,
    cur: [0, 0],
  };
}

function cloneDc(d: DcState): DcState {
  return {
    ...d,
    winOrg: [d.winOrg[0], d.winOrg[1]],
    winExt: [d.winExt[0], d.winExt[1]],
    vpOrg: [d.vpOrg[0], d.vpOrg[1]],
    vpExt: [d.vpExt[0], d.vpExt[1]],
    cur: [d.cur[0], d.cur[1]],
  };
}

/** Text alignment flags. */
const TA_UPDATECP = 1;
const TA_RIGHT = 2;
const TA_CENTER = 6;
const TA_BOTTOM = 8;
const TA_BASELINE = 24;

/** Maximum items one metafile may produce. */
const MAX_ITEMS = 200_000;

/** GDI playback state producing DisplayItems in points. */
export class Gdi {
  readonly items: DisplayItem[] = [];
  dc: DcState = defaultDc();
  readonly #stack: DcState[] = [];
  readonly #env: LayoutEnv;
  readonly #id: string;
  readonly #mapping: (dc: DcState) => Matrix;
  /** One device pixel in points (hairline pens). */
  hairline: number;
  /** When true nothing is registered or shaped; only bounds are collected. */
  readonly measure: boolean;
  /** Bounds of everything drawn, in points. */
  bounds: [number, number, number, number] = [Infinity, Infinity, -Infinity, -Infinity];
  #bracket: Path | undefined;
  #inBracket = false;
  #figureOpen = false;
  #pending: Path | undefined;
  #pendingStyle: LineStyle | null = null;
  #dibs = 0;
  #textFailed = false;

  constructor(
    env: LayoutEnv,
    id: string,
    mapping: (dc: DcState) => Matrix,
    hairline: number,
    measure = false,
  ) {
    this.#env = env;
    this.#id = id;
    this.#mapping = mapping;
    this.hairline = hairline;
    this.measure = measure;
  }

  /** Logical → points for the current state. */
  matrix(): Matrix {
    return this.#mapping(this.dc);
  }

  save(): void {
    this.flush();
    if (this.#stack.length < 1000) this.#stack.push(cloneDc(this.dc));
  }

  /** RESTOREDC: negative = relative to the top, positive = absolute 1-based level. */
  restore(n: number): void {
    this.flush();
    const len = this.#stack.length;
    const target = n < 0 ? len + n : n - 1;
    if (target < 0 || target >= len) return;
    const s = this.#stack[target];
    if (s === undefined) return;
    this.dc = s;
    this.#stack.length = target;
  }

  select(obj: GdiObject | undefined): void {
    if (obj === undefined) return;
    if (obj.kind === 'pen') this.dc.pen = obj;
    else if (obj.kind === 'brush') this.dc.brush = obj;
    else if (obj.kind === 'font') this.dc.font = obj;
  }

  /** The current pen as a LineStyle in points (null for PS_NULL). */
  lineStyle(m: Matrix = this.matrix()): LineStyle | null {
    const pen = this.dc.pen;
    if (pen.style === PS_NULL) return null;
    let width = this.hairline;
    if (!pen.cosmetic && pen.width > 0) width = Math.max(pen.width * matrixScale(m), this.hairline);
    const out: LineStyle = { color: pen.color, width };
    const dash: Record<number, number[]> = {
      1: [3, 1],
      2: [1, 1],
      3: [3, 1, 1, 1],
      4: [3, 1, 1, 1, 1, 1],
    };
    const d = dash[pen.style];
    if (d !== undefined) out.dash = d;
    return out;
  }

  get fill(): string | null {
    return this.dc.brush.color;
  }

  #track(x: number, y: number): void {
    const b = this.bounds;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;
    if (x < b[0]) b[0] = x;
    if (y < b[1]) b[1] = y;
    if (x > b[2]) b[2] = x;
    if (y > b[3]) b[3] = y;
  }

  #emit(items: DisplayItem[]): void {
    for (const it of items) if (this.items.length < MAX_ITEMS) this.items.push(it);
  }

  #emitPath(p: Path, fill: string | null, line: LineStyle | null): void {
    const bb = p.bounds();
    if (bb === undefined) return;
    if (fill === null && line === null) return;
    this.#track(bb[0], bb[1]);
    this.#track(bb[2], bb[3]);
    if (!this.measure) this.#emit(paintPath(p, fill, line, this.dc.fillMode !== 2));
  }

  /** Draws a closed or open figure built in logical coordinates. */
  draw(p: Path, fill: boolean, stroke: boolean): void {
    this.flush();
    const m = this.matrix();
    const t = p.transform(m);
    if (this.#inBracket) {
      this.#bracket ??= new Path();
      this.#bracket.append(t);
      this.#figureOpen = false;
      return;
    }
    this.#emitPath(t, fill ? this.fill : null, stroke ? this.lineStyle(m) : null);
  }

  moveTo(x: number, y: number): void {
    this.dc.cur = [x, y];
    const [px, py] = apply(this.matrix(), x, y);
    if (this.#inBracket) {
      this.#bracket ??= new Path();
      this.#bracket.moveTo(px, py);
      this.#figureOpen = true;
    } else if (this.#pending !== undefined) {
      this.#pending.moveTo(px, py);
    }
  }

  #target(): Path | undefined {
    const m = this.matrix();
    const [cx, cy] = apply(m, this.dc.cur[0], this.dc.cur[1]);
    if (this.#inBracket) {
      this.#bracket ??= new Path();
      if (!this.#figureOpen) {
        this.#bracket.moveTo(cx, cy);
        this.#figureOpen = true;
      }
      return this.#bracket;
    }
    const style = this.lineStyle(m);
    if (style === null) return undefined;
    if (this.#pending === undefined || !sameStyle(style, this.#pendingStyle)) {
      this.flush();
      this.#pending = new Path().moveTo(cx, cy);
      this.#pendingStyle = style;
    } else if (this.#pending.x !== cx || this.#pending.y !== cy) {
      this.#pending.moveTo(cx, cy);
    }
    return this.#pending;
  }

  lineTo(x: number, y: number): void {
    const t = this.#target();
    const [px, py] = apply(this.matrix(), x, y);
    t?.lineTo(px, py);
    this.dc.cur = [x, y];
  }

  /** POLYLINETO / POLYBEZIERTO from the current position (`pts` logical x,y pairs). */
  polyTo(pts: number[], bezier: boolean): void {
    if (pts.length < 2) return;
    const m = this.matrix();
    const t = this.#target();
    if (bezier) {
      for (let i = 0; i + 5 < pts.length; i += 6) {
        const [x1, y1] = apply(m, pts[i] ?? 0, pts[i + 1] ?? 0);
        const [x2, y2] = apply(m, pts[i + 2] ?? 0, pts[i + 3] ?? 0);
        const [x3, y3] = apply(m, pts[i + 4] ?? 0, pts[i + 5] ?? 0);
        t?.cubicTo(x1, y1, x2, y2, x3, y3);
        this.dc.cur = [pts[i + 4] ?? 0, pts[i + 5] ?? 0];
      }
    } else {
      for (let i = 0; i + 1 < pts.length; i += 2) {
        const [x, y] = apply(m, pts[i] ?? 0, pts[i + 1] ?? 0);
        t?.lineTo(x, y);
        this.dc.cur = [pts[i] ?? 0, pts[i + 1] ?? 0];
      }
    }
  }

  /** Emits the merged LINETO strokes. */
  flush(): void {
    const p = this.#pending;
    const style = this.#pendingStyle;
    this.#pending = undefined;
    this.#pendingStyle = null;
    if (p === undefined || style === null) return;
    if (!p.cmds.some((c) => c.op !== 'M')) return;
    const bb = p.bounds();
    if (bb !== undefined) {
      this.#track(bb[0], bb[1]);
      this.#track(bb[2], bb[3]);
    }
    if (!this.measure) this.#emit(strokeItems(p, style));
  }

  beginPath(): void {
    this.flush();
    this.#inBracket = true;
    this.#bracket = new Path();
    this.#figureOpen = false;
  }

  endPath(): void {
    this.#inBracket = false;
  }

  closeFigure(): void {
    if (this.#inBracket && this.#bracket !== undefined) {
      this.#bracket.close();
      this.#figureOpen = false;
    }
  }

  abortPath(): void {
    this.#inBracket = false;
    this.#bracket = undefined;
  }

  /** FILLPATH / STROKEPATH / STROKEANDFILLPATH. */
  usePath(fill: boolean, stroke: boolean): void {
    this.flush();
    const p = this.#bracket;
    this.#bracket = undefined;
    this.#inBracket = false;
    if (p === undefined || p.empty) return;
    if (fill) {
      // Filling closes every open figure.
      const closed = new Path();
      for (const sub of splitSubpaths(p)) closed.append(sub.close());
      this.#emitPath(closed, this.fill, stroke ? this.lineStyle() : null);
    } else if (stroke) this.#emitPath(p, null, this.lineStyle());
  }

  /** Fills a logical rectangle with a colour (PATBLT, ETO_OPAQUE, BLACKNESS...). */
  fillRect(x: number, y: number, w: number, h: number, color: string | null): void {
    if (color === null) return;
    this.flush();
    const t = new Path().rect(x, y, w, h).transform(this.matrix());
    this.#emitPath(t, color, null);
  }

  /**
   * Text at logical (x, y) (or the current position with TA_UPDATECP); `dx` = per UTF-16 unit
   * advances in logical units, when given.
   */
  text(x: number, y: number, str: string, dx?: number[]): void {
    this.flush();
    if (str === '') return;
    const m = this.matrix();
    const align = this.dc.textAlign;
    const useCp = (align & TA_UPDATECP) !== 0;
    const [px, py] = apply(m, useCp ? this.dc.cur[0] : x, useCp ? this.dc.cur[1] : y);
    const font = this.dc.font;
    const sy = Math.hypot(m[2], m[3]);
    const sx = Math.hypot(m[0], m[1]);
    let size = font.height === 0 ? 10 : Math.abs(font.height) * sy;
    if (font.height > 0) size /= 1.15; // cell height → em height
    if (!(size > 0.1) || !Number.isFinite(size)) size = 10;
    if (this.measure) {
      this.#track(px, py);
      return;
    }
    if (this.#textFailed) return;
    let shaped: ReturnType<typeof shapeText>;
    try {
      const props = runProps({
        font: font.face || 'Arial',
        size,
        bold: font.weight >= 600,
        italic: font.italic,
        color: this.dc.textColor,
      });
      shaped = shapeText(str, props, this.#env, this.dc.textColor);
    } catch (e) {
      this.#textFailed = true;
      this.#env.warn('METAFILE_TEXT', `metafile text skipped: ${String(e)}`);
      return;
    }
    let width = shaped.width;
    if (dx !== undefined && dx.length > 0)
      width = applyDx(
        shaped.items,
        dx.map((v) => v * sx),
      );
    let ox = px;
    if ((align & TA_CENTER) === TA_CENTER) ox -= width / 2;
    else if (align & TA_RIGHT) ox -= width;
    let oy = py;
    if ((align & TA_BASELINE) === TA_BASELINE) oy = py;
    else if (align & TA_BOTTOM) oy = py - shaped.descent;
    else oy = py + shaped.ascent;
    this.#track(ox, oy - shaped.ascent);
    this.#track(ox + width, oy + shaped.descent);
    this.#emit(offsetItems(shaped.items, ox, oy));
    if (useCp && sx > 0) this.dc.cur = [this.dc.cur[0] + width / (m[0] || sx), this.dc.cur[1]];
  }

  /**
   * A bitmap drawn into the logical rectangle (x, y, w, h); `src` = source rectangle in bitmap
   * pixels (top-left origin) for cropping.
   */
  image(
    dib: DecodedDib | undefined,
    x: number,
    y: number,
    w: number,
    h: number,
    src?: [number, number, number, number],
  ): void {
    this.flush();
    if (dib === undefined) return;
    const m = this.matrix();
    const corners = [
      apply(m, x, y),
      apply(m, x + w, y),
      apply(m, x, y + h),
      apply(m, x + w, y + h),
    ];
    const xs = corners.map((c) => c[0]);
    const ys = corners.map((c) => c[1]);
    const x0 = Math.min(...xs);
    const y0 = Math.min(...ys);
    const x1 = Math.max(...xs);
    const y1 = Math.max(...ys);
    if (!(x1 - x0 > 0 && y1 - y0 > 0)) return;
    this.#track(x0, y0);
    this.#track(x1, y1);
    if (this.measure) return;
    let bytes = dib.bytes;
    let type = dib.contentType;
    if (bytes === undefined && dib.rgba !== undefined) {
      bytes = encodePng(dib.width, dib.height, dib.rgba);
      type = 'image/png';
    }
    if (bytes === undefined || type === undefined) return;
    const id = addImage(
      this.#env,
      `${this.#id}#dib${this.#dibs++}`,
      type,
      bytes,
      dib.width,
      dib.height,
    );
    const item: ImageItem = {
      kind: 'image',
      x: x0,
      y: y0,
      width: x1 - x0,
      height: y1 - y0,
      image: id,
    };
    if (src !== undefined && dib.width > 0 && dib.height > 0) {
      const [sx, sy, sw, sh] = src;
      const left = Math.max(0, sx / dib.width);
      const top = Math.max(0, sy / dib.height);
      const right = Math.max(0, 1 - (sx + Math.abs(sw)) / dib.width);
      const bottom = Math.max(0, 1 - (sy + Math.abs(sh)) / dib.height);
      if (left > 1e-6 || top > 1e-6 || right > 1e-6 || bottom > 1e-6) {
        item.crop = { left, top, right, bottom };
      }
    }
    this.#emit([item]);
  }
}

function sameStyle(a: LineStyle | null, b: LineStyle | null): boolean {
  if (a === null || b === null) return a === b;
  return a.color === b.color && a.width === b.width && String(a.dash) === String(b.dash);
}

/** The subpaths of a path (each starting with its M). */
function splitSubpaths(p: Path): Path[] {
  const out: Path[] = [];
  let cur: Path | undefined;
  for (const c of p.cmds) {
    if (c.op === 'M') {
      cur = new Path();
      out.push(cur);
      cur.moveTo(c.p[0] ?? 0, c.p[1] ?? 0);
    } else if (cur === undefined) continue;
    else if (c.op === 'L') cur.lineTo(c.p[0] ?? 0, c.p[1] ?? 0);
    else if (c.op === 'C') {
      cur.cubicTo(c.p[0] ?? 0, c.p[1] ?? 0, c.p[2] ?? 0, c.p[3] ?? 0, c.p[4] ?? 0, c.p[5] ?? 0);
    } else cur.close();
  }
  return out;
}

/**
 * Overrides glyph advances with per-UTF-16-unit advances (points) and re-chains the items' x
 * positions. Returns the total width.
 */
export function applyDx(items: TextItem[], dx: readonly number[]): number {
  let offset = 0; // index into the full string of the item's first unit
  let x = items[0]?.x ?? 0;
  for (const it of items) {
    it.x = x;
    const n = it.text.length;
    it.advances = it.clusters.map((c, i) => {
      const end = it.clusters[i + 1] ?? n;
      let a = 0;
      for (let k = c; k < Math.max(end, c + 1); k++) a += dx[offset + k] ?? 0;
      return a;
    });
    for (const a of it.advances) x += a;
    offset += n;
  }
  return x - (items[0]?.x ?? 0);
}

/** Matrix mapping the rectangle (x0, y0)-(x1, y1) onto the box (0, 0)-(width, height). */
export function frameMatrix(
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  width: number,
  height: number,
): Matrix | undefined {
  const fw = x1 - x0;
  const fh = y1 - y0;
  if (!(Math.abs(fw) > 1e-9 && Math.abs(fh) > 1e-9)) return undefined;
  return multiply([1, 0, 0, 1, -x0, -y0], [width / fw, 0, 0, height / fh, 0, 0]);
}
