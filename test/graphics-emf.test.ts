// EMF/WMF playback (src/layout/graphics/{emf,wmf,metafile,dib,png}.ts): real metafiles from the
// Apache POI corpus (VariousPictures.docx) must give finite items inside the box, and small
// hand-built metafiles must give exact coordinates (mapping, world transform, save/restore,
// object tables, paths, text alignment, bitmaps). Malformed input must never throw.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { createDocx } from '../src/index.js';
import { decodeDib, decodePackedDib } from '../src/layout/graphics/dib.js';
import { emfItems } from '../src/layout/graphics/emf.js';
import { encodePng } from '../src/layout/graphics/png.js';
import { wmfItems } from '../src/layout/graphics/wmf.js';
import type { DisplayItem, ImageItem, PathItem, TextItem } from '../src/layout/types.js';
import { fakeEnv } from './graphics-env.js';

const media = unzipSync(
  new Uint8Array(
    readFileSync(
      fileURLToPath(
        new URL('./fixtures/graphics/apache-poi/VariousPictures.docx', import.meta.url),
      ),
    ),
  ),
  { filter: (f) => f.name.startsWith('word/media/') },
);

const env = () => fakeEnv(createDocx());

/** Every number appearing in the items (path data included). */
function numbers(items: DisplayItem[]): number[] {
  const out: number[] = [];
  for (const it of items) {
    for (const [k, v] of Object.entries(it)) {
      if (typeof v === 'number') out.push(v);
      else if (Array.isArray(v)) {
        for (const n of v) if (typeof n === 'number') out.push(n);
      } else if (k === 'd' && typeof v === 'string') {
        for (const t of v.split(/[\sMLCZ]+/)) if (t !== '') out.push(Number(t));
      }
    }
  }
  return out;
}

/** Coordinates of path data as [x, y] pairs. */
function pathPoints(d: string): [number, number][] {
  const nums = d
    .split(/[\sMLCZ]+/)
    .filter((t) => t !== '')
    .map(Number);
  const out: [number, number][] = [];
  for (let i = 0; i + 1 < nums.length; i += 2) out.push([nums[i]!, nums[i + 1]!]);
  return out;
}

function expectInside(items: DisplayItem[], w: number, h: number, tol: number): void {
  for (const it of items) {
    if (it.kind !== 'path') continue;
    for (const [x, y] of pathPoints(it.d)) {
      expect(x).toBeGreaterThanOrEqual(-tol);
      expect(x).toBeLessThanOrEqual(w + tol);
      expect(y).toBeGreaterThanOrEqual(-tol);
      expect(y).toBeLessThanOrEqual(h + tol);
    }
  }
}

// ---- synthetic metafile builders ----------------------------------------------------------

type Field = ['u8' | 'u16' | 'i16' | 'u32' | 'i32' | 'f32', number] | ['bytes', Uint8Array];

function pack(fields: Field[]): Uint8Array {
  let n = 0;
  for (const f of fields) {
    n +=
      f[0] === 'bytes' ? f[1].length : f[0] === 'u8' ? 1 : f[0] === 'u16' || f[0] === 'i16' ? 2 : 4;
  }
  const b = new Uint8Array(n);
  const dv = new DataView(b.buffer);
  let o = 0;
  for (const f of fields) {
    switch (f[0]) {
      case 'u8':
        dv.setUint8(o, f[1]);
        o += 1;
        break;
      case 'u16':
        dv.setUint16(o, f[1], true);
        o += 2;
        break;
      case 'i16':
        dv.setInt16(o, f[1], true);
        o += 2;
        break;
      case 'u32':
        dv.setUint32(o, f[1] >>> 0, true);
        o += 4;
        break;
      case 'i32':
        dv.setInt32(o, f[1], true);
        o += 4;
        break;
      case 'f32':
        dv.setFloat32(o, f[1], true);
        o += 4;
        break;
      case 'bytes':
        b.set(f[1], o);
        o += f[1].length;
        break;
    }
  }
  return b;
}

const u32 = (v: number): Field => ['u32', v];
const i32 = (v: number): Field => ['i32', v];
const i16 = (v: number): Field => ['i16', v];
const u16 = (v: number): Field => ['u16', v];

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

/** An EMF record: type, size (padded to 4), body. */
function emr(type: number, ...body: Field[]): Uint8Array {
  const b = pack(body);
  const pad = (4 - (b.length % 4)) % 4;
  return concat([pack([u32(type), u32(8 + b.length + pad)]), b, new Uint8Array(pad)]);
}

/**
 * An EMF whose frame is 100 × 50 mm on a 100 × 50 px / 100 × 50 mm device (1 px per mm), so a
 * 200 × 100 pt box maps one logical unit (MM_TEXT) to 2 pt.
 */
function emf(records: Uint8Array[]): Uint8Array {
  const recs = [...records, emr(14, u32(0), u32(0), u32(20))];
  const total = 108 + recs.reduce((n, r) => n + r.length, 0);
  const header = pack([
    u32(1),
    u32(108),
    ...[0, 0, 99, 49].map(i32),
    ...[0, 0, 10000, 5000].map(i32),
    u32(0x464d4520),
    u32(0x10000),
    u32(total),
    u32(recs.length + 1),
    u16(8),
    u16(0),
    u32(0),
    u32(0),
    u32(0),
    i32(100),
    i32(50),
    i32(100),
    i32(50),
    u32(0),
    u32(0),
    u32(0),
    i32(100000),
    i32(50000),
  ]);
  return concat([header, ...recs]);
}

const box = (l: number, t: number, r: number, b: number): Field[] => [l, t, r, b].map(i32);

function utf16(s: string, chars: number): Uint8Array {
  const b = new Uint8Array(chars * 2);
  for (let i = 0; i < s.length && i < chars; i++) {
    b[2 * i] = s.charCodeAt(i) & 255;
    b[2 * i + 1] = s.charCodeAt(i) >> 8;
  }
  return b;
}

/** EMR_EXTCREATEFONTINDIRECTW with a LOGFONTW. */
function font(ih: number, height: number, face: string, weight = 400, italic = 0): Uint8Array {
  return emr(
    82,
    u32(ih),
    i32(height),
    i32(0),
    i32(0),
    i32(0),
    i32(weight),
    ['u8', italic],
    ['u8', 0],
    ['u8', 0],
    ['u8', 0],
    ['u8', 0],
    ['u8', 0],
    ['u8', 0],
    ['u8', 0],
    ['bytes', utf16(face, 32)],
  );
}

/** EMR_EXTTEXTOUTW at (x, y) with optional Dx. */
function textOut(x: number, y: number, s: string, dx?: number[]): Uint8Array {
  const offString = 76;
  const strBytes = utf16(s, s.length + (s.length % 2));
  const offDx = dx === undefined ? 0 : offString + strBytes.length;
  return emr(
    84,
    ...box(0, 0, 0, 0),
    u32(1),
    ['f32', 1],
    ['f32', 1],
    i32(x),
    i32(y),
    u32(s.length),
    u32(offString),
    u32(0),
    ...box(0, 0, 0, 0),
    u32(offDx),
    ['bytes', strBytes],
    ...(dx ?? []).map(i32),
  );
}

/** A 2 × 2 24-bpp bottom-up DIB: BITMAPINFOHEADER + bits (rows padded to 4 bytes). */
function dib2x2(): { bmi: Uint8Array; bits: Uint8Array } {
  const bmi = pack([
    u32(40),
    i32(2),
    i32(2),
    u16(1),
    u16(24),
    u32(0),
    u32(16),
    i32(0),
    i32(0),
    u32(0),
    u32(0),
  ]);
  // bottom row: blue, green; top row: red, white (BGR order)
  const bits = new Uint8Array([255, 0, 0, 0, 255, 0, 0, 0, 0, 0, 255, 255, 255, 255, 0, 0]);
  return { bmi, bits };
}

const paths = (items: DisplayItem[]) => items.filter((i): i is PathItem => i.kind === 'path');
const texts = (items: DisplayItem[]) => items.filter((i): i is TextItem => i.kind === 'text');

describe('EMF playback (synthetic)', () => {
  it('maps the frame, objects, world transform and save/restore exactly', () => {
    const bytes = emf([
      emr(39, u32(1), u32(0), u32(0x000000ff), u32(0)), // red solid brush
      emr(37, u32(1)),
      emr(38, u32(2), u32(0), i32(0), i32(0), u32(0x00ff0000)), // blue hairline pen
      emr(37, u32(2)),
      emr(43, ...box(10, 10, 30, 20)), // RECTANGLE
      emr(33), // SAVEDC
      emr(35, ['f32', 1], ['f32', 0], ['f32', 0], ['f32', 1], ['f32', 50], ['f32', 0]),
      emr(86, ...box(0, 0, 0, 0), u32(3), i16(0), i16(0), i16(10), i16(0), i16(10), i16(10)),
      emr(34, i32(-1)), // RESTOREDC
      emr(37, u32(0x80000005)), // NULL_BRUSH
      emr(86, ...box(0, 0, 0, 0), u32(3), i16(0), i16(0), i16(10), i16(0), i16(10), i16(10)),
    ]);
    const items = emfItems(bytes, 200, 100, env(), 'word/media/x.emf');
    expect(items).toBeDefined();
    const p = paths(items!);
    expect(p).toHaveLength(3);
    expect(p[0]).toMatchObject({
      d: 'M 20 20 L 60 20 L 60 40 L 20 40 Z',
      fill: 'FF0000',
      stroke: '0000FF',
    });
    expect(p[0]!.strokeWidth).toBeCloseTo(1, 5); // one device pixel = 2 pt, clamped to 1
    expect(p[1]).toMatchObject({ d: 'M 100 0 L 120 0 L 120 20 Z', fill: 'FF0000' });
    expect(p[2]).toMatchObject({ d: 'M 0 0 L 20 0 L 20 20 Z', stroke: '0000FF' });
    expect(p[2]!.fill).toBeUndefined();
    expect(p[1]!.evenOdd).toBe(true); // ALTERNATE is the default fill mode
  });

  it('honours window/viewport mapping (MM_ANISOTROPIC) and MODIFYWORLDTRANSFORM', () => {
    const bytes = emf([
      emr(17, u32(8)), // MM_ANISOTROPIC
      emr(9, i32(1000), i32(500)), // window ext
      emr(11, i32(100), i32(50)), // viewport ext → 0.1 px per unit
      emr(36, ['f32', 2], ['f32', 0], ['f32', 0], ['f32', 2], ['f32', 0], ['f32', 0], u32(4)),
      emr(43, ...box(100, 100, 200, 150)),
    ]);
    const items = emfItems(bytes, 200, 100, env(), 'x')!;
    // world ×2 → page (200,200)-(400,300) → device ×0.1 → (20,20)-(40,30) → points ×2
    expect(paths(items)[0]!.d).toBe('M 40 40 L 80 40 L 80 60 L 40 60 Z');
  });

  it('fills and strokes path brackets, and merges LINETO runs into one stroke', () => {
    const bytes = emf([
      emr(59), // BEGINPATH
      emr(27, i32(0), i32(0)),
      emr(54, i32(10), i32(0)),
      emr(54, i32(10), i32(10)),
      emr(61), // CLOSEFIGURE
      emr(60), // ENDPATH
      emr(62, ...box(0, 0, 10, 10)), // FILLPATH (white stock brush)
      emr(27, i32(0), i32(20)),
      emr(54, i32(10), i32(20)),
      emr(54, i32(20), i32(20)),
      emr(54, i32(20), i32(25)),
    ]);
    const p = paths(emfItems(bytes, 200, 100, env(), 'x')!);
    expect(p).toHaveLength(2);
    expect(p[0]).toMatchObject({ d: 'M 0 0 L 20 0 L 20 20 Z', fill: 'FFFFFF' });
    expect(p[0]!.stroke).toBeUndefined();
    expect(p[1]).toMatchObject({ d: 'M 0 40 L 20 40 L 40 40 L 40 50', stroke: '000000' });
  });

  it('draws text with the selected font, colour, alignment and Dx advances', () => {
    const e = env();
    const bytes = emf([
      font(3, -10, 'Arial', 700),
      emr(37, u32(3)),
      emr(24, u32(0x0000ff00)), // green
      emr(22, u32(24)), // TA_BASELINE
      textOut(10, 40, 'Hi'),
      emr(22, u32(0)), // TA_TOP
      textOut(10, 10, 'AB', [5, 7]),
    ]);
    const t = texts(emfItems(bytes, 200, 100, e, 'x')!);
    expect(t).toHaveLength(2);
    expect(t[0]).toMatchObject({ x: 20, y: 80, size: 20, color: '00FF00', text: 'Hi' });
    expect(t[0]!.glyphs).toEqual([72, 105]);
    expect(t[0]!.advances).toEqual([10, 10]); // 500/1000 em at 20 pt
    expect(e.fonts.fonts()[t[0]!.font]?.family).toBe('Arial');
    expect(e.fonts.fonts()[t[0]!.font]?.bold).toBe(true);
    // TA_TOP: baseline = y + ascent (0.8 em); Dx in logical units × 2
    expect(t[1]).toMatchObject({ x: 20, text: 'AB' });
    expect(t[1]!.y).toBeCloseTo(20 + 16, 6);
    expect(t[1]!.advances).toEqual([10, 14]);
  });

  it('decodes STRETCHDIBITS bitmaps into PNG image resources', () => {
    const e = env();
    const { bmi, bits } = dib2x2();
    const offBmi = 80;
    const offBits = offBmi + bmi.length;
    const bytes = emf([
      emr(
        81,
        ...box(0, 0, 3, 3),
        i32(5), // xDest
        i32(5), // yDest
        i32(0),
        i32(0),
        i32(2),
        i32(2),
        u32(offBmi),
        u32(bmi.length),
        u32(offBits),
        u32(bits.length),
        u32(0),
        u32(0x00cc0020),
        i32(4),
        i32(4),
        ['bytes', bmi],
        ['bytes', bits],
      ),
    ]);
    const items = emfItems(bytes, 200, 100, e, 'word/media/p.emf')!;
    const img = items.find((i): i is ImageItem => i.kind === 'image');
    expect(img).toMatchObject({
      x: 10,
      y: 10,
      width: 8,
      height: 8,
      image: 'word/media/p.emf#dib0',
    });
    expect(img!.crop).toBeUndefined();
    const res = e.images.get('word/media/p.emf#dib0');
    expect(res?.contentType).toBe('image/png');
    expect([...res!.bytes.subarray(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10]);
    expect(res).toMatchObject({ pixelWidth: 2, pixelHeight: 2 });
  });

  it('draws ellipses, round rects, pies and bezier curves as finite closed paths', () => {
    const bytes = emf([
      emr(42, ...box(0, 0, 20, 10)),
      emr(44, ...box(30, 0, 60, 20), i32(10), i32(10)),
      emr(47, ...box(0, 20, 20, 40), i32(20), i32(30), i32(10), i32(20)),
      emr(
        85,
        ...box(0, 0, 0, 0),
        u32(4),
        i16(0),
        i16(45),
        i16(10),
        i16(40),
        i16(20),
        i16(50),
        i16(30),
        i16(45),
      ),
    ]);
    const p = paths(emfItems(bytes, 200, 100, env(), 'x')!);
    expect(p).toHaveLength(4);
    for (const it of p.slice(0, 3)) expect(it.d.endsWith('Z')).toBe(true);
    expect(p[0]!.d.startsWith('M 40 10 C')).toBe(true); // ellipse starts at its right edge
    expect(p[3]!.d).toBe('M 0 90 C 20 80 40 100 60 90');
    expect(numbers(p).every(Number.isFinite)).toBe(true);
  });
});

describe('EMF playback (real files)', () => {
  for (const [name, w, h] of [
    ['word/media/image3.emf', 191.45, 133.54],
    ['word/media/image4.emf', 168.33, 101.88],
  ] as const) {
    it(`${name}: finite items inside the box`, () => {
      const bytes = media[name]!;
      expect(bytes).toBeDefined();
      const items = emfItems(bytes, w, h, env(), name);
      expect(items).toBeDefined();
      expect(items!.length).toBeGreaterThan(10);
      expect(numbers(items!).every(Number.isFinite)).toBe(true);
      expect(paths(items!).every((p) => p.fill !== undefined || p.stroke !== undefined)).toBe(true);
      expectInside(items!, w, h, Math.max(w, h) * 0.05);
      // The drawing covers a good part of the box (not collapsed into a corner).
      const pts = paths(items!).flatMap((p) => pathPoints(p.d));
      const xs = pts.map((p) => p[0]);
      const ys = pts.map((p) => p[1]);
      expect(Math.max(...xs) - Math.min(...xs)).toBeGreaterThan(w * 0.5);
      expect(Math.max(...ys) - Math.min(...ys)).toBeGreaterThan(h * 0.5);
    });
  }

  it('is deterministic', () => {
    const bytes = media['word/media/image4.emf']!;
    expect(JSON.stringify(emfItems(bytes, 100, 60, env(), 'a'))).toBe(
      JSON.stringify(emfItems(bytes, 100, 60, env(), 'a')),
    );
  });
});

// ---- WMF ------------------------------------------------------------------------------------

/** A WMF record: size in words, function, params. */
function wmr(fn: number, ...params: Field[]): Uint8Array {
  const b = pack(params);
  const pad = b.length % 2;
  return concat([pack([u32((6 + b.length + pad) / 2), u16(fn)]), b, new Uint8Array(pad)]);
}

function wmf(records: Uint8Array[], placeable?: [number, number, number, number]): Uint8Array {
  const recs = [...records, wmr(0)];
  const words = 9 + recs.reduce((n, r) => n + r.length / 2, 0);
  const header = pack([u16(1), u16(9), u16(0x300), u32(words), u16(4), u32(32), u16(0)]);
  const parts = [header, ...recs];
  if (placeable !== undefined) {
    parts.unshift(
      pack([u32(0x9ac6cdd7), u16(0), ...placeable.map(i16), u16(1440), u32(0), u16(0)]),
    );
  }
  return concat(parts);
}

function ansiBytes(s: string): Uint8Array {
  const b = new Uint8Array(s.length + (s.length % 2));
  for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i);
  return b;
}

describe('WMF playback', () => {
  it('maps the window onto the box and reuses freed object slots', () => {
    const e = env();
    const bytes = wmf(
      [
        wmr(0x020b, i16(0), i16(0)), // SETWINDOWORG (y, x)
        wmr(0x020c, i16(500), i16(1000)), // SETWINDOWEXT (y, x)
        wmr(0x02fc, u16(0), u32(0x000000ff), u16(0)), // slot 0: red brush
        wmr(0x02fa, u16(5), i16(0), i16(0), u32(0)), // slot 1: null pen
        wmr(0x012d, u16(0)),
        wmr(0x012d, u16(1)),
        wmr(0x041b, i16(200), i16(300), i16(100), i16(100)), // RECTANGLE b, r, t, l
        wmr(0x01f0, u16(0)), // delete slot 0
        wmr(0x02fc, u16(0), u32(0x0000ff00), u16(0)), // slot 0 again: green
        wmr(0x012d, u16(0)),
        wmr(0x0324, i16(3), i16(0), i16(0), i16(100), i16(0), i16(100), i16(100)),
        wmr(
          0x02fb,
          i16(-100),
          i16(0),
          i16(0),
          i16(0),
          i16(400),
          ['u8', 1],
          ['u8', 0],
          ['u8', 0],
          ['u8', 0],
          u32(0),
          ['bytes', ansiBytes('Times New Roman\0')],
        ),
        wmr(0x012d, u16(2)),
        wmr(0x012e, u16(24)), // TA_BASELINE
        wmr(0x0521, u16(3), ['bytes', ansiBytes('Ab\xe9')], i16(250), i16(500)), // TEXTOUT
      ],
      [0, 0, 1000, 500],
    );
    const items = wmfItems(bytes, 200, 100, e, 'w')!;
    const p = paths(items);
    expect(p[0]).toMatchObject({ d: 'M 20 20 L 60 20 L 60 40 L 20 40 Z', fill: 'FF0000' });
    expect(p[0]!.stroke).toBeUndefined();
    expect(p[1]).toMatchObject({ d: 'M 0 0 L 20 0 L 20 20 Z', fill: '00FF00' });
    const t = texts(items);
    expect(t).toHaveLength(1);
    expect(t[0]).toMatchObject({ x: 100, y: 50, size: 20, text: 'Abé' });
    expect(e.fonts.fonts()[t[0]!.font]).toMatchObject({ family: 'Times New Roman', italic: true });
  });

  it('falls back to the placeable bounding box without a window', () => {
    const bytes = wmf([wmr(0x041b, i16(1000), i16(2000), i16(0), i16(0))], [0, 0, 2000, 1000]);
    expect(paths(wmfItems(bytes, 100, 50, env(), 'w')!)[0]!.d).toBe(
      'M 0 0 L 100 0 L 100 50 L 0 50 Z',
    );
  });

  it('plays the real WMF of VariousPictures.docx', () => {
    const bytes = media['word/media/image1.wmf']!;
    const items = wmfItems(bytes, 150, 120, env(), 'word/media/image1.wmf');
    expect(items).toBeDefined();
    expect(items!.length).toBeGreaterThan(50);
    expect(numbers(items!).every(Number.isFinite)).toBe(true);
    expectInside(items!, 150, 120, 10);
  });

  it('decodes DIBSTRETCHBLT bitmaps', () => {
    const e = env();
    const { bmi, bits } = dib2x2();
    const bytes = wmf([
      wmr(0x020c, i16(100), i16(100)),
      wmr(
        0x0b41,
        u32(0x00cc0020),
        i16(2),
        i16(2),
        i16(0),
        i16(0),
        i16(50),
        i16(50),
        i16(0),
        i16(0),
        ['bytes', bmi],
        ['bytes', bits],
      ),
    ]);
    const img = wmfItems(bytes, 100, 100, e, 'w')!.find((i) => i.kind === 'image');
    expect(img).toMatchObject({ x: 0, y: 0, width: 50, height: 50, image: 'w#dib0' });
    expect(e.images.get('w#dib0')?.contentType).toBe('image/png');
  });
});

// ---- DIB / PNG -------------------------------------------------------------------------------

describe('DIB decoding and PNG encoding', () => {
  it('decodes 24-bpp bottom-up rows into top-first RGBA', () => {
    const { bmi, bits } = dib2x2();
    const d = decodeDib(bmi, bits)!;
    expect(d).toMatchObject({ width: 2, height: 2 });
    expect([...d.rgba!]).toEqual([
      255, 0, 0, 255, 255, 255, 255, 255, 0, 0, 255, 255, 0, 255, 0, 255,
    ]);
  });

  it('decodes palettes (1 and 8 bpp), 16-bpp 5-5-5, 32-bpp bitfields and top-down DIBs', () => {
    const hdr = (w: number, h: number, bpp: number, comp: number, colors: number) =>
      pack([
        u32(40),
        i32(w),
        i32(h),
        u16(1),
        u16(bpp),
        u32(comp),
        u32(0),
        i32(0),
        i32(0),
        u32(colors),
        u32(0),
      ]);
    // 1 bpp, 3 × 1, palette black/white, bits 101
    const one = decodePackedDib(
      concat([
        hdr(3, 1, 1, 0, 2),
        pack([u32(0), u32(0xffffff)]),
        new Uint8Array([0b10100000, 0, 0, 0]),
      ]),
    )!;
    expect([...one.rgba!].filter((_, i) => i % 4 === 0)).toEqual([255, 0, 255]);
    // 8 bpp top-down 1 × 2 with a 2-entry palette
    const eight = decodePackedDib(
      concat([
        hdr(1, -2, 8, 0, 2),
        pack([u32(0x0000ff), u32(0x00ff00)]),
        new Uint8Array([1, 0, 0, 0, 0, 0, 0, 0]),
      ]),
    )!;
    expect([...eight.rgba!]).toEqual([0, 255, 0, 255, 0, 0, 255, 255]);
    // 16 bpp 5-5-5 pure red
    const sixteen = decodePackedDib(
      concat([hdr(1, 1, 16, 0, 0), new Uint8Array([0x00, 0x7c, 0, 0])]),
    )!;
    expect([...sixteen.rgba!]).toEqual([255, 0, 0, 255]);
    // 32 bpp BI_BITFIELDS with masks R=0xff, G=0xff00, B=0xff0000
    const bf = decodePackedDib(
      concat([
        hdr(1, 1, 32, 3, 0),
        pack([u32(0xff), u32(0xff00), u32(0xff0000)]),
        new Uint8Array([10, 20, 30, 0]),
      ]),
    )!;
    expect([...bf.rgba!]).toEqual([10, 20, 30, 255]);
  });

  it('rejects unsupported or hostile headers without throwing', () => {
    expect(decodeDib(new Uint8Array(4), new Uint8Array(0))).toBeUndefined();
    const huge = pack([
      u32(40),
      i32(100000),
      i32(100000),
      u16(1),
      u16(24),
      u32(0),
      u32(0),
      i32(0),
      i32(0),
      u32(0),
      u32(0),
    ]);
    expect(decodeDib(huge, new Uint8Array(10))).toBeUndefined();
    const rle = pack([
      u32(40),
      i32(2),
      i32(2),
      u16(1),
      u16(8),
      u32(1),
      u32(0),
      i32(0),
      i32(0),
      u32(0),
      u32(0),
    ]);
    expect(decodeDib(rle, new Uint8Array(10))).toBeUndefined();
  });

  it('writes a PNG with signature, IHDR and RGB/RGBA colour types', () => {
    const opaque = encodePng(2, 1, new Uint8Array([1, 2, 3, 255, 4, 5, 6, 255]));
    expect([...opaque.subarray(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10]);
    expect(String.fromCharCode(...opaque.subarray(12, 16))).toBe('IHDR');
    const dv = new DataView(opaque.buffer, opaque.byteOffset);
    expect([dv.getUint32(16), dv.getUint32(20), opaque[24], opaque[25]]).toEqual([2, 1, 8, 2]);
    const alpha = encodePng(1, 1, new Uint8Array([1, 2, 3, 4]));
    expect(alpha[25]).toBe(6);
    expect(String.fromCharCode(...alpha.subarray(alpha.length - 8, alpha.length - 4))).toBe('IEND');
  });
});

describe('malformed metafiles', () => {
  it('returns undefined for non-metafiles and survives truncation and corruption', () => {
    expect(emfItems(new Uint8Array(10), 10, 10, env(), 'x')).toBeUndefined();
    expect(wmfItems(new Uint8Array([1, 2, 3]), 10, 10, env(), 'x')).toBeUndefined();
    expect(emfItems(media['word/media/image4.emf']!, 0, 10, env(), 'x')).toBeUndefined();
    for (const name of [
      'word/media/image3.emf',
      'word/media/image4.emf',
      'word/media/image1.wmf',
    ]) {
      const src = media[name]!;
      const play = (b: Uint8Array) =>
        name.endsWith('.wmf')
          ? wmfItems(b, 100, 100, env(), 'x')
          : emfItems(b, 100, 100, env(), 'x');
      for (const cut of [120, 500, 1000, Math.floor(src.length / 2)]) {
        expect(() => play(src.slice(0, cut))).not.toThrow();
      }
      // Deterministic corruption (LCG): flip bytes after the header.
      let seed = 12345;
      for (let round = 0; round < 30; round++) {
        const b = src.slice();
        for (let k = 0; k < 40; k++) {
          seed = (seed * 1103515245 + 12345) >>> 0;
          const at = 120 + (seed % (b.length - 120));
          b[at] = (seed >>> 8) & 255;
        }
        const items = play(b);
        if (items !== undefined) expect(numbers(items).every(Number.isFinite)).toBe(true);
      }
    }
  });
});
