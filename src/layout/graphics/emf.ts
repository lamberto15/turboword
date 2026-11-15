// Enhanced Metafile (EMF, [MS-EMF]) playback into DisplayItems: the GDI drawing records Word
// documents use (shapes, polygons, paths, pens/brushes/fonts, text, DIB bitmaps, world
// transforms). The picture frame from the header (rclFrame, 0.01 mm) is mapped onto the box.
// EMF+ (records inside GDI comments) is ignored: dual-mode files carry the GDI records too.
// Every read is bounds-checked; a malformed record ends playback with what was drawn so far.

import type { LayoutEnv } from '../engine.js';
import type { DisplayItem } from '../types.js';
import { type DecodedDib, decodeDib } from './dib.js';
import {
  colorRef,
  type DcState,
  frameMatrix,
  Gdi,
  type GdiObject,
  MM_ISOTROPIC,
  makeBrush,
  PS_NULL,
  patternBrush,
  Reader,
  stockObject,
} from './metafile.js';
import { IDENTITY, type Matrix, matrixScale, multiply, Path } from './path.js';

const EMF_SIGNATURE = 0x464d4520;
const MAX_RECORDS = 1_000_000;

/** Page (logical, after the world transform) → device pixels. */
function pageToDevice(dc: DcState, ppmX: number, ppmY: number): Matrix {
  let sx = 1;
  let sy = 1;
  const mm = (k: number): void => {
    sx = k * ppmX;
    sy = -k * ppmY;
  };
  switch (dc.mapMode) {
    case 2: // MM_LOMETRIC
      mm(0.1);
      break;
    case 3: // MM_HIMETRIC
      mm(0.01);
      break;
    case 4: // MM_LOENGLISH
      mm(0.254);
      break;
    case 5: // MM_HIENGLISH
      mm(0.0254);
      break;
    case 6: // MM_TWIPS
      mm(25.4 / 1440);
      break;
    case 7: // MM_ISOTROPIC
    case 8: {
      sx = dc.vpExt[0] / (dc.winExt[0] || 1);
      sy = dc.vpExt[1] / (dc.winExt[1] || 1);
      if (dc.mapMode === MM_ISOTROPIC) {
        const k = Math.min(Math.abs(sx), Math.abs(sy));
        sx = (Math.sign(sx) || 1) * k;
        sy = (Math.sign(sy) || 1) * k;
      }
      break;
    }
  }
  if (!Number.isFinite(sx) || sx === 0) sx = 1;
  if (!Number.isFinite(sy) || sy === 0) sy = 1;
  return [sx, 0, 0, sy, dc.vpOrg[0] - dc.winOrg[0] * sx, dc.vpOrg[1] - dc.winOrg[1] * sy];
}

/** `n` points at `at` (32-bit or 16-bit pairs), clamped to the record. */
function points(r: Reader, at: number, n: number, small: boolean): number[] {
  const size = small ? 4 : 8;
  const max = Math.max(0, Math.floor((r.length - at) / size));
  const count = Math.min(n, max);
  const out = new Array<number>(count * 2);
  for (let i = 0; i < count; i++) {
    const o = at + i * size;
    out[2 * i] = small ? r.i16(o) : r.i32(o);
    out[2 * i + 1] = small ? r.i16(o + 2) : r.i32(o + 4);
  }
  return out;
}

/** UTF-16LE string of `n` code units at `at` (NUL-terminated when `stopAtNul`). */
function utf16(r: Reader, at: number, n: number, stopAtNul: boolean): string {
  let s = '';
  for (let i = 0; i < n; i++) {
    const o = at + 2 * i;
    if (o + 2 > r.length) break;
    const c = r.u16(o);
    if (c === 0 && stopAtNul) break;
    s += String.fromCharCode(c);
  }
  return s;
}

/** A logical rectangle record field (l, t, r, b) → x, y, w, h. */
function box(r: Reader, at: number): [number, number, number, number] {
  const l = r.i32(at);
  const t = r.i32(at + 4);
  return [l, t, r.i32(at + 8) - l, r.i32(at + 12) - t];
}

/** Arc through a box from the ray toward (sx, sy) to the ray toward (ex, ey), counter-clockwise. */
export function arcPath(
  b: [number, number, number, number],
  sx: number,
  sy: number,
  ex: number,
  ey: number,
  kind: 'arc' | 'chord' | 'pie',
  from?: [number, number],
): Path {
  const [x, y, w, h] = b;
  const rx = w / 2;
  const ry = h / 2;
  const cx = x + rx;
  const cy = y + ry;
  const p = new Path();
  if (rx === 0 || ry === 0) return p;
  const t0 = Math.atan2((sy - cy) / ry, (sx - cx) / rx);
  const t1 = Math.atan2((ey - cy) / ry, (ex - cx) / rx);
  let dt = t1 - t0;
  while (dt >= 0) dt -= 2 * Math.PI;
  if (Math.abs(dt) < 1e-9) dt = -2 * Math.PI;
  if (kind === 'pie') p.moveTo(cx, cy);
  else if (from !== undefined) p.moveTo(from[0], from[1]);
  p.ellipseArc(cx, cy, Math.abs(rx), Math.abs(ry), t0, dt);
  if (kind !== 'arc') p.close();
  return p;
}

/** Rounded rectangle with corner ellipse (cw, ch). */
export function roundRectPath(b: [number, number, number, number], cw: number, ch: number): Path {
  const [x, y, w, h] = b;
  const rx = Math.min(Math.abs(cw) / 2, Math.abs(w) / 2);
  const ry = Math.min(Math.abs(ch) / 2, Math.abs(h) / 2);
  const x0 = Math.min(x, x + w);
  const y0 = Math.min(y, y + h);
  const x1 = Math.max(x, x + w);
  const y1 = Math.max(y, y + h);
  if (rx === 0 || ry === 0) return new Path().rect(x0, y0, x1 - x0, y1 - y0);
  const p = new Path().moveTo(x0 + rx, y0).lineTo(x1 - rx, y0);
  p.ellipseArc(x1 - rx, y0 + ry, rx, ry, -Math.PI / 2, Math.PI / 2);
  p.lineTo(x1, y1 - ry);
  p.ellipseArc(x1 - rx, y1 - ry, rx, ry, 0, Math.PI / 2);
  p.lineTo(x0 + rx, y1);
  p.ellipseArc(x0 + rx, y1 - ry, rx, ry, Math.PI / 2, Math.PI / 2);
  p.lineTo(x0, y0 + ry);
  p.ellipseArc(x0 + rx, y0 + ry, rx, ry, Math.PI, Math.PI / 2);
  return p.close();
}

/** Polygons/polylines of a poly-poly record. */
function polyPoly(r: Reader, small: boolean): number[][] {
  const n = Math.min(r.u32(24), Math.floor(r.length / 4));
  const counts: number[] = [];
  for (let i = 0; i < n; i++) counts.push(r.u32(32 + 4 * i));
  let at = 32 + 4 * n;
  const out: number[][] = [];
  for (const c of counts) {
    const pts = points(r, at, c, small);
    out.push(pts);
    at += c * (small ? 4 : 8);
    if (at > r.length) break;
  }
  return out;
}

/** True when the DIB in `bmi` is stored bottom-up. */
function bottomUp(bmi: Uint8Array): boolean {
  const r = new Reader(bmi);
  return r.u32(0) === 12 ? r.i16(6) > 0 : r.i32(8) > 0;
}

/** A DIB referenced by offset/size fields of a bitmap record. */
function recordDib(
  r: Reader,
  offBmi: number,
  cbBmi: number,
  offBits: number,
  cbBits: number,
  usage: number,
): { dib: DecodedDib | undefined; bottomUp: boolean } | undefined {
  if (cbBmi === 0 || offBmi === 0) return undefined;
  const bmi = r.slice(offBmi, cbBmi);
  return { dib: decodeDib(bmi, r.slice(offBits, cbBits), usage), bottomUp: bottomUp(bmi) };
}

const SRCCOPY_NOP = 0x00aa0029;
const PATCOPY = 0x00f00021;
const BLACKNESS = 0x00000042;
const WHITENESS = 0x00ff0062;

/** Plays the records into `gdi`. */
function play(r: Reader, gdi: Gdi): void {
  const objects = new Map<number, GdiObject>();
  const len = r.length;
  let off = 0;
  for (let count = 0; count < MAX_RECORDS && off + 8 <= len; count++) {
    const type = r.u32(off);
    const size = r.u32(off + 4);
    if (size < 8 || off + size > len) break;
    const rec = new Reader(r.slice(off, size));
    off += size;
    if (type !== 27 && type !== 54 && type !== 5 && type !== 6 && type !== 88 && type !== 89) {
      gdi.flush();
    }
    const dc = gdi.dc;
    switch (type) {
      case 14: // EOF
        return;
      case 9: // SETWINDOWEXTEX
        dc.winExt = [rec.i32(8), rec.i32(12)];
        dc.winExtSet = true;
        break;
      case 10: // SETWINDOWORGEX
        dc.winOrg = [rec.i32(8), rec.i32(12)];
        break;
      case 11: // SETVIEWPORTEXTEX
        dc.vpExt = [rec.i32(8), rec.i32(12)];
        break;
      case 12: // SETVIEWPORTORGEX
        dc.vpOrg = [rec.i32(8), rec.i32(12)];
        break;
      case 31: // SCALEVIEWPORTEXTEX
      case 32: {
        const ext = type === 31 ? dc.vpExt : dc.winExt;
        const xd = rec.i32(12) || 1;
        const yd = rec.i32(20) || 1;
        ext[0] = (ext[0] * rec.i32(8)) / xd;
        ext[1] = (ext[1] * rec.i32(16)) / yd;
        break;
      }
      case 17: // SETMAPMODE
        dc.mapMode = rec.u32(8);
        break;
      case 18: // SETBKMODE
        dc.bkMode = rec.u32(8);
        break;
      case 19: // SETPOLYFILLMODE
        dc.fillMode = rec.u32(8);
        break;
      case 22: // SETTEXTALIGN
        dc.textAlign = rec.u32(8);
        break;
      case 24: // SETTEXTCOLOR
        dc.textColor = colorRef(rec.u32(8));
        break;
      case 25: // SETBKCOLOR
        dc.bkColor = colorRef(rec.u32(8));
        break;
      case 27: // MOVETOEX
        gdi.moveTo(rec.i32(8), rec.i32(12));
        break;
      case 54: // LINETO
        gdi.lineTo(rec.i32(8), rec.i32(12));
        break;
      case 33: // SAVEDC
        gdi.save();
        break;
      case 34: // RESTOREDC
        gdi.restore(rec.i32(8));
        break;
      case 35: // SETWORLDTRANSFORM
        dc.world = xform(rec, 8);
        break;
      case 36: {
        // MODIFYWORLDTRANSFORM
        const m = xform(rec, 8);
        const mode = rec.u32(32);
        if (mode === 1) dc.world = IDENTITY;
        else if (mode === 2) dc.world = multiply(m, dc.world);
        else if (mode === 3) dc.world = multiply(dc.world, m);
        else if (mode === 4) dc.world = m;
        break;
      }
      case 37: {
        // SELECTOBJECT
        const ih = rec.u32(8);
        gdi.select(ih & 0x80000000 ? stockObject(ih & 0x7fffffff) : objects.get(ih));
        break;
      }
      case 40: // DELETEOBJECT
        objects.delete(rec.u32(8));
        break;
      case 38: {
        // CREATEPEN
        const style = rec.u32(12) & 0x0f;
        objects.set(rec.u32(8), {
          kind: 'pen',
          style,
          width: Math.abs(rec.i32(16)),
          cosmetic: false,
          color: colorRef(rec.u32(24)),
        });
        break;
      }
      case 95: {
        // EXTCREATEPEN
        const ps = rec.u32(28);
        let style = ps & 0x0f;
        if (style === 8) style = 2; // PS_ALTERNATE → dotted
        if (style === 7) style = 0; // PS_USERSTYLE → solid
        if (rec.u32(36) === 1) style = PS_NULL; // BS_NULL brush
        objects.set(rec.u32(8), {
          kind: 'pen',
          style,
          width: rec.u32(32),
          cosmetic: (ps & 0xf0000) === 0,
          color: colorRef(rec.u32(40)),
        });
        break;
      }
      case 39: // CREATEBRUSHINDIRECT
        objects.set(rec.u32(8), makeBrush(rec.u32(12), rec.u32(16)));
        break;
      case 93: // CREATEMONOBRUSH
      case 94: {
        // CREATEDIBPATTERNBRUSHPT
        const d = recordDib(rec, rec.u32(16), rec.u32(20), rec.u32(24), rec.u32(28), rec.u32(12));
        objects.set(rec.u32(8), patternBrush(d?.dib));
        break;
      }
      case 82: {
        // EXTCREATEFONTINDIRECTW
        objects.set(rec.u32(8), {
          kind: 'font',
          height: rec.i32(12),
          weight: rec.i32(28),
          italic: rec.u8(32) !== 0,
          face: utf16(rec, 40, 32, true).trim(),
        });
        break;
      }
      case 49: // CREATEPALETTE
        objects.set(rec.u32(8), { kind: 'other' });
        break;
      case 43: {
        // RECTANGLE
        const [x, y, w, h] = box(rec, 8);
        gdi.draw(new Path().rect(x, y, w, h), true, true);
        break;
      }
      case 44: // ROUNDRECT
        gdi.draw(roundRectPath(box(rec, 8), rec.i32(24), rec.i32(28)), true, true);
        break;
      case 42: {
        // ELLIPSE
        const [x, y, w, h] = box(rec, 8);
        gdi.draw(new Path().ellipse(x, y, w, h), true, true);
        break;
      }
      case 45: // ARC
      case 46: // CHORD
      case 47: // PIE
      case 55: {
        // ARCTO
        const kind = type === 46 ? 'chord' : type === 47 ? 'pie' : 'arc';
        const ex = rec.i32(32);
        const ey = rec.i32(36);
        const from: [number, number] | undefined = type === 55 ? [dc.cur[0], dc.cur[1]] : undefined;
        const p = arcPath(box(rec, 8), rec.i32(24), rec.i32(28), ex, ey, kind, from);
        gdi.draw(p, kind !== 'arc', true);
        if (type === 55) dc.cur = [p.x, p.y];
        break;
      }
      case 41: {
        // ANGLEARC
        const cx = rec.i32(8);
        const cy = rec.i32(12);
        const rad = rec.u32(16);
        const a0 = (rec.f32(20) * Math.PI) / 180;
        const sw = (rec.f32(24) * Math.PI) / 180;
        const p = new Path().moveTo(dc.cur[0], dc.cur[1]);
        p.ellipseArc(cx, cy, rad, rad, -a0, -sw);
        gdi.draw(p, false, true);
        dc.cur = [p.x, p.y];
        break;
      }
      case 3: // POLYGON
      case 86: // POLYGON16
      case 4: // POLYLINE
      case 87: {
        // POLYLINE16
        const pts = points(rec, 28, rec.u32(24), type === 86 || type === 87);
        const closed = type === 3 || type === 86;
        gdi.draw(new Path().poly(pts, closed), closed, true);
        break;
      }
      case 2: // POLYBEZIER
      case 85: {
        // POLYBEZIER16
        const pts = points(rec, 28, rec.u32(24), type === 85);
        if (pts.length < 2) break;
        const p = new Path().moveTo(pts[0] ?? 0, pts[1] ?? 0);
        for (let i = 2; i + 5 < pts.length; i += 6) {
          p.cubicTo(
            pts[i] ?? 0,
            pts[i + 1] ?? 0,
            pts[i + 2] ?? 0,
            pts[i + 3] ?? 0,
            pts[i + 4] ?? 0,
            pts[i + 5] ?? 0,
          );
        }
        gdi.draw(p, false, true);
        break;
      }
      case 5: // POLYBEZIERTO
      case 88:
      case 6: // POLYLINETO
      case 89: {
        const small = type === 88 || type === 89;
        gdi.polyTo(points(rec, 28, rec.u32(24), small), type === 5 || type === 88);
        break;
      }
      case 7: // POLYPOLYLINE
      case 90:
      case 8: // POLYPOLYGON
      case 91: {
        const small = type === 90 || type === 91;
        const closed = type === 8 || type === 91;
        const p = new Path();
        for (const pts of polyPoly(rec, small)) p.poly(pts, closed);
        gdi.draw(p, closed, true);
        break;
      }
      case 56: // POLYDRAW
      case 92: {
        // POLYDRAW16
        const n = rec.u32(24);
        const small = type === 92;
        const pts = points(rec, 28, n, small);
        const typesAt = 28 + n * (small ? 4 : 8);
        const p = new Path().moveTo(dc.cur[0], dc.cur[1]);
        for (let i = 0; i < pts.length / 2; i++) {
          const t = rec.u8(typesAt + i);
          const x = pts[2 * i] ?? 0;
          const y = pts[2 * i + 1] ?? 0;
          if ((t & 6) === 6) p.moveTo(x, y);
          else if ((t & 6) === 4 && i + 2 < pts.length / 2) {
            p.cubicTo(
              x,
              y,
              pts[2 * i + 2] ?? 0,
              pts[2 * i + 3] ?? 0,
              pts[2 * i + 4] ?? 0,
              pts[2 * i + 5] ?? 0,
            );
            i += 2;
          } else p.lineTo(x, y);
          if (t & 1) p.close();
        }
        gdi.draw(p, false, true);
        dc.cur = [p.x, p.y];
        break;
      }
      case 59: // BEGINPATH
        gdi.beginPath();
        break;
      case 60: // ENDPATH
        gdi.endPath();
        break;
      case 61: // CLOSEFIGURE
        gdi.closeFigure();
        break;
      case 62: // FILLPATH
        gdi.usePath(true, false);
        break;
      case 63: // STROKEANDFILLPATH
        gdi.usePath(true, true);
        break;
      case 64: // STROKEPATH
        gdi.usePath(false, true);
        break;
      case 67: // SELECTCLIPPATH (clipping is not supported: the path is consumed)
      case 68: // ABORTPATH
        gdi.abortPath();
        break;
      case 83: // EXTTEXTOUTA
      case 84: // EXTTEXTOUTW
        text(rec, gdi, type === 84);
        break;
      case 81: {
        // STRETCHDIBITS
        const rop = rec.u32(68);
        if (rop === SRCCOPY_NOP) break;
        const d = recordDib(rec, rec.u32(48), rec.u32(52), rec.u32(56), rec.u32(60), rec.u32(64));
        if (d?.dib === undefined) break;
        const cySrc = rec.i32(44);
        const ySrc = d.bottomUp ? d.dib.height - rec.i32(36) - cySrc : rec.i32(36);
        gdi.image(d.dib, rec.i32(24), rec.i32(28), rec.i32(72), rec.i32(76), [
          rec.i32(32),
          ySrc,
          rec.i32(40),
          cySrc,
        ]);
        break;
      }
      case 76: // BITBLT
      case 77: {
        // STRETCHBLT
        const x = rec.i32(24);
        const y = rec.i32(28);
        const w = rec.i32(32);
        const h = rec.i32(36);
        const rop = rec.u32(40);
        const d = recordDib(rec, rec.u32(84), rec.u32(88), rec.u32(92), rec.u32(96), rec.u32(80));
        if (d === undefined) {
          if (rop === PATCOPY) gdi.fillRect(x, y, w, h, gdi.fill);
          else if (rop === BLACKNESS) gdi.fillRect(x, y, w, h, '000000');
          else if (rop === WHITENESS) gdi.fillRect(x, y, w, h, 'FFFFFF');
          break;
        }
        if (rop === SRCCOPY_NOP || d.dib === undefined) break;
        let src: [number, number, number, number] | undefined;
        if (type === 77) {
          const cySrc = rec.i32(104);
          const ySrc = d.bottomUp ? d.dib.height - rec.i32(48) - cySrc : rec.i32(48);
          src = [rec.i32(44), ySrc, rec.i32(100), cySrc];
        }
        gdi.image(d.dib, x, y, w, h, src);
        break;
      }
      case 80: {
        // SETDIBITSTODEVICE
        const d = recordDib(rec, rec.u32(48), rec.u32(52), rec.u32(56), rec.u32(60), rec.u32(64));
        if (d?.dib === undefined) break;
        gdi.image(d.dib, rec.i32(24), rec.i32(28), rec.i32(40), rec.i32(44));
        break;
      }
      default:
        // GDICOMMENT (EMF+), clipping, palettes, ROP2, ... are ignored.
        break;
    }
  }
}

/** XFORM at `at` → Matrix. */
function xform(r: Reader, at: number): Matrix {
  const m: Matrix = [
    r.f32(at),
    r.f32(at + 4),
    r.f32(at + 8),
    r.f32(at + 12),
    r.f32(at + 16),
    r.f32(at + 20),
  ];
  return m.every((v) => Number.isFinite(v)) ? m : IDENTITY;
}

const ETO_OPAQUE = 0x2;
const ETO_GLYPH_INDEX = 0x10;
const ETO_NO_RECT = 0x100;
const ETO_PDY = 0x2000;

/** EXTTEXTOUTW / EXTTEXTOUTA. */
function text(rec: Reader, gdi: Gdi, wide: boolean): void {
  const x = rec.i32(36);
  const y = rec.i32(40);
  const n = Math.min(rec.u32(44), 32768);
  const offString = rec.u32(48);
  const opts = rec.u32(52);
  const hasRect = (opts & ETO_NO_RECT) === 0;
  const offDx = rec.u32(hasRect ? 72 : 56);
  if (hasRect && opts & ETO_OPAQUE) {
    const [rx, ry, rw, rh] = box(rec, 56);
    if (rw !== 0 && rh !== 0) gdi.fillRect(rx, ry, rw, rh, gdi.dc.bkColor);
  }
  if (opts & ETO_GLYPH_INDEX || n === 0) return;
  let s = '';
  if (wide) s = utf16(rec, offString, n, false);
  else {
    const bytes = rec.slice(offString, n);
    for (const b of bytes) s += String.fromCharCode(b);
  }
  let dx: number[] | undefined;
  if (offDx > 0 && offDx + 4 * n <= rec.length) {
    const step = opts & ETO_PDY ? 8 : 4;
    if (offDx + step * n <= rec.length) {
      dx = [];
      for (let i = 0; i < n; i++) dx.push(rec.i32(offDx + step * i));
    }
  }
  gdi.text(x, y, s.replace(/\0+$/, ''), dx);
}

/** Parsed EMF header fields needed for the frame mapping. */
interface EmfHeader {
  frame: Matrix | undefined;
  ppmX: number;
  ppmY: number;
  devicePx: number;
}

function header(r: Reader, width: number, height: number): EmfHeader | undefined {
  if (r.length < 88 || r.u32(0) !== 1 || r.u32(40) !== EMF_SIGNATURE) return undefined;
  // The header proper ends where the description / pixel format data starts, if earlier.
  let hsize = r.u32(4);
  const offDesc = r.u32(64);
  if (r.u32(60) > 0 && offDesc >= 88 && offDesc < hsize) hsize = offDesc;
  const offPix = r.u32(92);
  if (hsize >= 100 && r.u32(88) > 0 && offPix >= 88 && offPix < hsize) hsize = offPix;
  const devX = r.i32(72);
  const devY = r.i32(76);
  let mmX = r.i32(80);
  let mmY = r.i32(84);
  if (hsize >= 108 && r.i32(100) > 0 && r.i32(104) > 0) {
    mmX = r.i32(100) / 1000;
    mmY = r.i32(104) / 1000;
  }
  const ppmX = devX > 0 && mmX > 0 ? devX / mmX : 96 / 25.4;
  const ppmY = devY > 0 && mmY > 0 ? devY / mmY : 96 / 25.4;
  // rclFrame (0.01 mm) in device pixels, else rclBounds (inclusive device pixels).
  let frame = frameMatrix(
    (r.i32(24) / 100) * ppmX,
    (r.i32(28) / 100) * ppmY,
    (r.i32(32) / 100) * ppmX,
    (r.i32(36) / 100) * ppmY,
    width,
    height,
  );
  frame ??= frameMatrix(r.i32(8), r.i32(12), r.i32(16) + 1, r.i32(20) + 1, width, height);
  return { frame, ppmX, ppmY, devicePx: frame === undefined ? 0.75 : matrixScale(frame) };
}

/**
 * Display items for an EMF picture drawn into a (width × height) point box; `id` is the image
 * id of the metafile (bitmaps inside get `${id}#dibN`). Undefined when `bytes` is not an EMF.
 */
export function emfItems(
  bytes: Uint8Array,
  width: number,
  height: number,
  env: LayoutEnv,
  id: string,
): DisplayItem[] | undefined {
  if (!(width > 0 && height > 0)) return undefined;
  const r = new Reader(bytes);
  const h = header(r, width, height);
  if (h === undefined) return undefined;
  const hairline = Math.min(Math.max(h.devicePx, 0.25), 1);
  let frame = h.frame;
  if (frame === undefined) {
    // No usable frame: measure the drawing in device pixels, then fit it into the box.
    const probe = new Gdi(
      env,
      id,
      (dc) => multiply(dc.world, pageToDevice(dc, h.ppmX, h.ppmY)),
      1,
      true,
    );
    play(r, probe);
    probe.flush();
    const [x0, y0, x1, y1] = probe.bounds;
    frame = Number.isFinite(x0) ? frameMatrix(x0, y0, x1, y1, width, height) : undefined;
    if (frame === undefined) return [];
  }
  const base = frame;
  const gdi = new Gdi(
    env,
    id,
    (dc) => multiply(multiply(dc.world, pageToDevice(dc, h.ppmX, h.ppmY)), base),
    hairline,
  );
  play(r, gdi);
  gdi.flush();
  return gdi.items;
}
