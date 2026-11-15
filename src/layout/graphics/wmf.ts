// Windows Metafile (WMF, [MS-WMF]) playback into DisplayItems, sharing the GDI state of
// metafile.ts with the EMF player. The logical window (SETWINDOWORG/EXT) is mapped onto the box,
// as Word plays a WMF into an anisotropic viewport; without a window, the placeable header's
// bounding box is used, and without either the drawing's own bounds. Objects live in the WMF
// object table (each create takes the lowest free slot). Reads are bounds-checked.

import type { LayoutEnv } from '../engine.js';
import type { DisplayItem } from '../types.js';
import { decodePackedDib } from './dib.js';
import { arcPath, roundRectPath } from './emf.js';
import {
  colorRef,
  type DcState,
  frameMatrix,
  Gdi,
  type GdiObject,
  makeBrush,
  patternBrush,
  Reader,
} from './metafile.js';
import { IDENTITY, type Matrix, Path } from './path.js';

const PLACEABLE_KEY = 0x9ac6cdd7;
const MAX_RECORDS = 1_000_000;
const MAX_OBJECTS = 65535;

let decoder: TextDecoder | undefined | null;

/** ANSI (Windows-1252) bytes → string, one character per byte. */
function ansi(bytes: Uint8Array): string {
  if (decoder === undefined) {
    try {
      decoder = new TextDecoder('windows-1252');
    } catch {
      decoder = null;
    }
  }
  if (decoder !== null) return decoder.decode(bytes);
  let s = '';
  for (const b of bytes) s += String.fromCharCode(b);
  return s;
}

/** `n` 16-bit points at `at`. */
function points(r: Reader, at: number, n: number): number[] {
  const count = Math.max(0, Math.min(n, Math.floor((r.length - at) / 4)));
  const out = new Array<number>(count * 2);
  for (let i = 0; i < count; i++) {
    out[2 * i] = r.i16(at + 4 * i);
    out[2 * i + 1] = r.i16(at + 4 * i + 2);
  }
  return out;
}

/** (bottom, right, top, left) at `at` → x, y, w, h. */
function box(r: Reader, at: number): [number, number, number, number] {
  const b = r.i16(at);
  const rt = r.i16(at + 2);
  const t = r.i16(at + 4);
  const l = r.i16(at + 6);
  return [l, t, rt - l, b - t];
}

const PATCOPY = 0x00f00021;
const BLACKNESS = 0x00000042;
const WHITENESS = 0x00ff0062;

function play(r: Reader, start: number, gdi: Gdi): void {
  const objs: (GdiObject | undefined)[] = [];
  const create = (o: GdiObject): void => {
    let i = objs.indexOf(undefined);
    if (i < 0) i = objs.length;
    if (i < MAX_OBJECTS) objs[i] = o;
  };
  let off = start;
  for (let count = 0; count < MAX_RECORDS && off + 6 <= r.length; count++) {
    const words = r.u32(off);
    const fn = r.u16(off + 4);
    if (words < 3 || off + words * 2 > r.length) break;
    const rec = new Reader(r.slice(off, words * 2));
    off += words * 2;
    if (fn !== 0x0214 && fn !== 0x0213) gdi.flush();
    const dc = gdi.dc;
    switch (fn) {
      case 0x0000: // EOF
        return;
      case 0x020b: // SETWINDOWORG
        dc.winOrg = [rec.i16(8), rec.i16(6)];
        break;
      case 0x020c: // SETWINDOWEXT
        dc.winExt = [rec.i16(8), rec.i16(6)];
        dc.winExtSet = true;
        break;
      case 0x020f: // OFFSETWINDOWORG
        dc.winOrg = [dc.winOrg[0] + rec.i16(8), dc.winOrg[1] + rec.i16(6)];
        break;
      case 0x0410: // SCALEWINDOWEXT
        dc.winExt = [
          (dc.winExt[0] * rec.i16(12)) / (rec.i16(10) || 1),
          (dc.winExt[1] * rec.i16(8)) / (rec.i16(6) || 1),
        ];
        break;
      case 0x020d: // SETVIEWPORTORG
        dc.vpOrg = [rec.i16(8), rec.i16(6)];
        break;
      case 0x020e: // SETVIEWPORTEXT
        dc.vpExt = [rec.i16(8), rec.i16(6)];
        break;
      case 0x0103: // SETMAPMODE
        dc.mapMode = rec.u16(6);
        break;
      case 0x0102: // SETBKMODE
        dc.bkMode = rec.u16(6);
        break;
      case 0x0106: // SETPOLYFILLMODE
        dc.fillMode = rec.u16(6);
        break;
      case 0x012e: // SETTEXTALIGN
        dc.textAlign = rec.u16(6);
        break;
      case 0x0209: // SETTEXTCOLOR
        dc.textColor = colorRef(rec.u32(6));
        break;
      case 0x0201: // SETBKCOLOR
        dc.bkColor = colorRef(rec.u32(6));
        break;
      case 0x001e: // SAVEDC
        gdi.save();
        break;
      case 0x0127: // RESTOREDC
        gdi.restore(rec.i16(6));
        break;
      case 0x012d: // SELECTOBJECT
        gdi.select(objs[rec.u16(6)]);
        break;
      case 0x01f0: // DELETEOBJECT
        if (rec.u16(6) < objs.length) objs[rec.u16(6)] = undefined;
        break;
      case 0x02fa: // CREATEPENINDIRECT
        create({
          kind: 'pen',
          style: rec.u16(6) & 0x0f,
          width: Math.abs(rec.i16(8)),
          cosmetic: false,
          color: colorRef(rec.u32(12)),
        });
        break;
      case 0x02fc: // CREATEBRUSHINDIRECT
        create(makeBrush(rec.u16(6), rec.u32(8)));
        break;
      case 0x0142: // DIBCREATEPATTERNBRUSH
        create(patternBrush(decodePackedDib(rec.slice(10, rec.length), rec.u16(8))));
        break;
      case 0x02fb: {
        // CREATEFONTINDIRECT
        const face = rec.slice(24, 32);
        const nul = face.indexOf(0);
        create({
          kind: 'font',
          height: rec.i16(6),
          weight: rec.i16(14),
          italic: rec.u8(16) !== 0,
          face: ansi(nul < 0 ? face : face.subarray(0, nul)).trim(),
        });
        break;
      }
      case 0x00f7: // CREATEPALETTE
      case 0x01f9: // CREATEPATTERNBRUSH
      case 0x06ff: // CREATEREGION
        create({ kind: 'other' });
        break;
      case 0x0214: // MOVETO
        gdi.moveTo(rec.i16(8), rec.i16(6));
        break;
      case 0x0213: // LINETO
        gdi.lineTo(rec.i16(8), rec.i16(6));
        break;
      case 0x0324: // POLYGON
      case 0x0325: {
        // POLYLINE
        const closed = fn === 0x0324;
        gdi.draw(new Path().poly(points(rec, 8, rec.i16(6)), closed), closed, true);
        break;
      }
      case 0x0538: {
        // POLYPOLYGON
        const n = rec.u16(6);
        let at = 8 + 2 * n;
        const p = new Path();
        for (let i = 0; i < n && at < rec.length; i++) {
          const c = rec.u16(8 + 2 * i);
          p.poly(points(rec, at, c), true);
          at += 4 * c;
        }
        gdi.draw(p, true, true);
        break;
      }
      case 0x041b: {
        // RECTANGLE
        const [x, y, w, h] = box(rec, 6);
        gdi.draw(new Path().rect(x, y, w, h), true, true);
        break;
      }
      case 0x061c: // ROUNDRECT
        gdi.draw(roundRectPath(box(rec, 10), rec.i16(8), rec.i16(6)), true, true);
        break;
      case 0x0418: {
        // ELLIPSE
        const [x, y, w, h] = box(rec, 6);
        gdi.draw(new Path().ellipse(x, y, w, h), true, true);
        break;
      }
      case 0x0817: // ARC
      case 0x081a: // PIE
      case 0x0830: {
        // CHORD
        const kind = fn === 0x081a ? 'pie' : fn === 0x0830 ? 'chord' : 'arc';
        const p = arcPath(box(rec, 14), rec.i16(12), rec.i16(10), rec.i16(8), rec.i16(6), kind);
        gdi.draw(p, kind !== 'arc', true);
        break;
      }
      case 0x061d: {
        // PATBLT
        const rop = rec.u32(6);
        const [h, w, y, x] = [rec.i16(10), rec.i16(12), rec.i16(14), rec.i16(16)];
        if (rop === PATCOPY) gdi.fillRect(x, y, w, h, gdi.fill);
        else if (rop === BLACKNESS) gdi.fillRect(x, y, w, h, '000000');
        else if (rop === WHITENESS) gdi.fillRect(x, y, w, h, 'FFFFFF');
        break;
      }
      case 0x0521: {
        // TEXTOUT
        const n = rec.u16(6);
        const s = ansi(rec.slice(8, n));
        const at = 8 + n + (n & 1);
        gdi.text(rec.i16(at + 2), rec.i16(at), s);
        break;
      }
      case 0x0a32: {
        // EXTTEXTOUT
        const n = rec.u16(10);
        const opts = rec.u16(12);
        const hasRect = (opts & 0x6) !== 0;
        const at = hasRect ? 22 : 14;
        if (hasRect && opts & 0x2) {
          const l = rec.i16(14);
          const t = rec.i16(16);
          gdi.fillRect(l, t, rec.i16(18) - l, rec.i16(20) - t, dc.bkColor);
        }
        const s = ansi(rec.slice(at, n));
        const dxAt = at + n + (n & 1);
        let dx: number[] | undefined;
        if (dxAt + 2 * n <= rec.length && n > 0) {
          dx = [];
          for (let i = 0; i < n; i++) dx.push(rec.i16(dxAt + 2 * i));
        }
        gdi.text(rec.i16(8), rec.i16(6), s, dx);
        break;
      }
      case 0x0f43: {
        // STRETCHDIB
        const dib = decodePackedDib(rec.slice(28, rec.length), rec.u16(10));
        if (dib === undefined) break;
        const srcH = rec.i16(12);
        const ySrc = dib.height - rec.i16(16) - srcH;
        gdi.image(dib, rec.i16(26), rec.i16(24), rec.i16(22), rec.i16(20), [
          rec.i16(18),
          ySrc,
          rec.i16(14),
          srcH,
        ]);
        break;
      }
      case 0x0b41: {
        // DIBSTRETCHBLT
        if (rec.length <= 30) break;
        const dib = decodePackedDib(rec.slice(26, rec.length));
        if (dib === undefined) break;
        const srcH = rec.i16(10);
        const ySrc = dib.height - rec.i16(14) - srcH;
        gdi.image(dib, rec.i16(24), rec.i16(22), rec.i16(20), rec.i16(18), [
          rec.i16(16),
          ySrc,
          rec.i16(12),
          srcH,
        ]);
        break;
      }
      case 0x0940: {
        // DIBBITBLT
        if (rec.length <= 26) break;
        const dib = decodePackedDib(rec.slice(22, rec.length));
        if (dib === undefined) break;
        gdi.image(dib, rec.i16(20), rec.i16(18), rec.i16(16), rec.i16(14));
        break;
      }
      case 0x0d33: {
        // SETDIBTODEV
        const dib = decodePackedDib(rec.slice(24, rec.length), rec.u16(6));
        if (dib === undefined) break;
        gdi.image(dib, rec.i16(22), rec.i16(20), rec.i16(18), rec.i16(16));
        break;
      }
      default:
        break;
    }
  }
}

/**
 * Display items for a WMF picture drawn into a (width × height) point box; `id` is the image id
 * of the metafile (bitmaps inside get `${id}#dibN`). Undefined when `bytes` is not a WMF.
 */
export function wmfItems(
  bytes: Uint8Array,
  width: number,
  height: number,
  env: LayoutEnv,
  id: string,
): DisplayItem[] | undefined {
  if (!(width > 0 && height > 0)) return undefined;
  const r = new Reader(bytes);
  let at = 0;
  let bbox: Matrix | undefined;
  if (r.u32(0) === PLACEABLE_KEY) {
    bbox = frameMatrix(r.i16(6), r.i16(8), r.i16(10), r.i16(12), width, height);
    at = 22;
  }
  const type = r.u16(at);
  const hsize = r.u16(at + 2);
  if ((type !== 1 && type !== 2) || hsize < 9 || at + hsize * 2 > r.length) return undefined;
  const start = at + hsize * 2;
  const windowMatrix = (dc: DcState): Matrix | undefined =>
    dc.winExtSet
      ? frameMatrix(
          dc.winOrg[0],
          dc.winOrg[1],
          dc.winOrg[0] + dc.winExt[0],
          dc.winOrg[1] + dc.winExt[1],
          width,
          height,
        )
      : undefined;
  let fallback = bbox;
  if (fallback === undefined) {
    // No placeable header: measure in logical units so drawing before SETWINDOWEXT still fits.
    const probe = new Gdi(env, id, () => IDENTITY, 1, true);
    play(r, start, probe);
    probe.flush();
    const [x0, y0, x1, y1] = probe.bounds;
    fallback = Number.isFinite(x0) ? frameMatrix(x0, y0, x1, y1, width, height) : undefined;
  }
  const base = fallback ?? IDENTITY;
  const gdi = new Gdi(env, id, (dc) => windowMatrix(dc) ?? base, 0.5);
  play(r, start, gdi);
  gdi.flush();
  return gdi.items;
}
