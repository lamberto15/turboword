// PDF images: decoders for every supported format (fed by small encoders written here, so every
// colour type, bit depth, filter, interlace and transparency variant is covered), plus rendering
// through poppler to check orientation, colour, alpha, crop and rotation end to end.

import { deflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import type { ImageItem, ImageResource } from '../src/layout/types.js';
import { decodeImage } from '../src/render/pdf/images/index.js';
import { renderPdf } from '../src/render/pdf/render.js';
import {
  dump,
  HAS_POPPLER,
  HAS_QPDF,
  layoutOf,
  qpdfCheck,
  rasterize,
  runOn,
} from './pdf-helpers.js';

// ---- tiny encoders ---------------------------------------------------------------------------

const CRC = new Uint32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
function crc32(b: Uint8Array): number {
  let c = 0xffffffff;
  for (const v of b) c = CRC[(c ^ v) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
const u32be = (v: number): number[] => [v >>> 24, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
const u16le = (v: number): number[] => [v & 255, (v >>> 8) & 255];
const u32le = (v: number): number[] => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, v >>> 24];

function chunk(type: string, data: ArrayLike<number>): number[] {
  const td = Uint8Array.from([...type].map((c) => c.charCodeAt(0)).concat(Array.from(data)));
  return [...u32be(data.length), ...td, ...u32be(crc32(td))];
}

const CHANNELS: Record<number, number> = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };
const ADAM7 = [
  [0, 0, 8, 8],
  [4, 0, 8, 8],
  [0, 4, 4, 8],
  [2, 0, 4, 4],
  [0, 2, 2, 4],
  [1, 0, 2, 2],
  [0, 1, 1, 2],
];

function getBits(b: Uint8Array, bit: number, n: number): number {
  let v = 0;
  for (let i = 0; i < n; i++) v = (v << 1) | ((b[(bit + i) >> 3]! >> (7 - ((bit + i) & 7))) & 1);
  return v;
}
function setBits(b: Uint8Array, bit: number, n: number, v: number): void {
  for (let i = 0; i < n; i++) {
    const bitv = (v >> (n - 1 - i)) & 1;
    if (bitv) b[(bit + i) >> 3]! |= 0x80 >> ((bit + i) & 7);
  }
}

/** PNG from packed rows (the decoder's expected output), cycling all five filters. */
function png(
  w: number,
  h: number,
  ctype: number,
  depth: number,
  rows: Uint8Array,
  o: { palette?: number[]; trns?: number[]; interlace?: boolean; split?: boolean } = {},
): Uint8Array {
  const bitsPP = CHANNELS[ctype]! * depth;
  const rb = Math.ceil((w * bitsPP) / 8);
  const bpp = Math.max(1, bitsPP >> 3);
  const filtered: number[] = [];
  let ft = 0;
  const emit = (img: Uint8Array, pw: number, ph: number): void => {
    const prb = Math.ceil((pw * bitsPP) / 8);
    for (let y = 0; y < ph; y++) {
      const f = ft++ % 5;
      filtered.push(f);
      for (let x = 0; x < prb; x++) {
        const raw = img[y * prb + x]!;
        const a = x >= bpp ? img[y * prb + x - bpp]! : 0;
        const b = y > 0 ? img[(y - 1) * prb + x]! : 0;
        const c = x >= bpp && y > 0 ? img[(y - 1) * prb + x - bpp]! : 0;
        const pa = Math.abs(b - c);
        const pb = Math.abs(a - c);
        const pc = Math.abs(a + b - 2 * c);
        const pred = [0, a, b, (a + b) >> 1, pa <= pb && pa <= pc ? a : pb <= pc ? b : c][f]!;
        filtered.push((raw - pred) & 255);
      }
    }
  };
  if (!o.interlace) emit(rows, w, h);
  else {
    for (const [x0, y0, dx, dy] of ADAM7) {
      const pw = Math.ceil((w - x0!) / dx!);
      const ph = Math.ceil((h - y0!) / dy!);
      if (pw <= 0 || ph <= 0) continue;
      const prb = Math.ceil((pw * bitsPP) / 8);
      const pass = new Uint8Array(prb * ph);
      for (let j = 0; j < ph; j++) {
        for (let i = 0; i < pw; i++) {
          const v = getBits(
            rows,
            (y0! + j * dy!) * rb * 8 + (x0! + i * dx!) * bitsPP,
            Math.min(bitsPP, 24),
          );
          setBits(pass, j * prb * 8 + i * bitsPP, Math.min(bitsPP, 24), v);
          if (bitsPP > 24) {
            for (let k = 24; k < bitsPP; k += 8) {
              const v2 = getBits(rows, (y0! + j * dy!) * rb * 8 + (x0! + i * dx!) * bitsPP + k, 8);
              setBits(pass, j * prb * 8 + i * bitsPP + k, 8, v2);
            }
          }
        }
      }
      emit(pass, pw, ph);
    }
  }
  const z = deflateSync(Uint8Array.from(filtered));
  const ihdr = [...u32be(w), ...u32be(h), depth, ctype, 0, 0, o.interlace ? 1 : 0];
  const idat = o.split
    ? [chunk('IDAT', z.subarray(0, 5)), chunk('IDAT', z.subarray(5))].flat()
    : chunk('IDAT', z);
  return Uint8Array.from([
    0x89,
    0x50,
    0x4e,
    0x47,
    0x0d,
    0x0a,
    0x1a,
    0x0a,
    ...chunk('IHDR', ihdr),
    ...(o.palette ? chunk('PLTE', o.palette) : []),
    ...(o.trns ? chunk('tRNS', o.trns) : []),
    ...idat,
    ...chunk('IEND', []),
  ]);
}

/** Deterministic pseudo-random bytes. */
function bytes(n: number, seed: number): Uint8Array {
  const out = new Uint8Array(n);
  let s = seed;
  for (let i = 0; i < n; i++) {
    s = (s * 1103515245 + 12345) >>> 0;
    out[i] = s >>> 24;
  }
  return out;
}

/** GIF with literal-only LZW codes (a clear code before the code width would grow). */
function gif(
  w: number,
  h: number,
  palette: number[],
  idx: Uint8Array,
  o: { transparent?: number; interlace?: boolean; local?: boolean } = {},
): Uint8Array {
  const n = Math.max(1, Math.ceil(Math.log2(palette.length / 3)));
  const table = [...palette, ...new Array(3 * (1 << n) - palette.length).fill(0)];
  const min = Math.max(2, n);
  const clear = 1 << min;
  let rows = idx;
  if (o.interlace) {
    const order: number[] = [];
    for (const [s, st] of [
      [0, 8],
      [4, 8],
      [2, 4],
      [1, 2],
    ] as const)
      for (let y = s; y < h; y += st) order.push(y);
    rows = new Uint8Array(idx.length);
    order.forEach((y, i) => {
      rows.set(idx.subarray(y * w, y * w + w), i * w);
    });
  }
  const codes = [clear];
  let since = 0;
  for (const v of rows) {
    if (since === clear - 2) {
      codes.push(clear);
      since = 0;
    }
    codes.push(v);
    since++;
  }
  codes.push(clear + 1);
  const data: number[] = [];
  let acc = 0;
  let bits = 0;
  for (const c of codes) {
    acc |= c << bits;
    bits += min + 1;
    while (bits >= 8) {
      data.push(acc & 255);
      acc >>>= 8;
      bits -= 8;
    }
  }
  if (bits) data.push(acc & 255);
  const blocks: number[] = [];
  for (let i = 0; i < data.length; i += 255) {
    const part = data.slice(i, i + 255);
    blocks.push(part.length, ...part);
  }
  const packed = 0x80 | 0x70 | (n - 1);
  return Uint8Array.from([
    ...[...'GIF89a'].map((c) => c.charCodeAt(0)),
    ...u16le(w),
    ...u16le(h),
    o.local ? 0 : packed,
    0,
    0,
    ...(o.local ? [] : table),
    ...(o.transparent !== undefined ? [0x21, 0xf9, 4, 1, 0, 0, o.transparent, 0] : []),
    0x21,
    0xfe,
    3,
    0x61,
    0x62,
    0x63,
    0, // a comment extension to skip
    0x2c,
    0,
    0,
    0,
    0,
    ...u16le(w),
    ...u16le(h),
    (o.interlace ? 0x40 : 0) | (o.local ? 0x80 | (n - 1) : 0),
    ...(o.local ? table : []),
    min,
    ...blocks,
    0,
    0x3b,
  ]);
}

/** BMP writer: `px` is RGBA per pixel; palettes take indices. */
function bmp(
  w: number,
  h: number,
  bpp: number,
  o: {
    rgba?: number[][];
    idx?: number[];
    palette?: number[][];
    topDown?: boolean;
    hdr?: 12 | 40 | 108;
    masks?: number[];
    comp?: number;
  },
): Uint8Array {
  const hdr = o.hdr ?? 40;
  const stride = Math.floor((w * bpp + 31) / 32) * 4;
  const pal = (o.palette ?? []).flatMap((c) =>
    hdr === 12 ? [c[2]!, c[1]!, c[0]!] : [c[2]!, c[1]!, c[0]!, 0],
  );
  const extraMasks = hdr === 40 && o.masks ? o.masks.flatMap(u32le) : [];
  const off = 14 + hdr + extraMasks.length + pal.length;
  const pixels = new Uint8Array(stride * h);
  for (let y = 0; y < h; y++) {
    const row = (o.topDown ? y : h - 1 - y) * stride;
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (bpp <= 8) setBits(pixels, row * 8 + x * bpp, bpp, o.idx![i]!);
      else {
        const [r, g, b, a] = o.rgba![i]!;
        if (bpp === 24) pixels.set([b!, g!, r!], row + x * 3);
        else if (bpp === 32) pixels.set([b!, g!, r!, a!], row + x * 4);
        else {
          const v = ((r! >> 3) << 11) | ((g! >> 2) << 5) | (b! >> 3); // 565
          pixels.set(u16le(v), row + x * 2);
        }
      }
    }
  }
  let header: number[];
  if (hdr === 12) header = [...u32le(12), ...u16le(w), ...u16le(h), ...u16le(1), ...u16le(bpp)];
  else {
    header = [
      ...u32le(hdr),
      ...u32le(w),
      ...u32le(o.topDown ? -h : h),
      ...u16le(1),
      ...u16le(bpp),
      ...u32le(o.comp ?? 0),
      ...u32le(pixels.length),
      ...u32le(2835),
      ...u32le(2835),
      ...u32le(o.palette?.length ?? 0),
      ...u32le(0),
    ];
    if (hdr === 108) {
      header.push(...(o.masks ?? [0, 0, 0, 0]).flatMap(u32le));
      while (header.length < 108) header.push(0);
    }
  }
  const total = off + pixels.length;
  return Uint8Array.from([
    0x42,
    0x4d,
    ...u32le(total),
    0,
    0,
    0,
    0,
    ...u32le(off),
    ...header,
    ...extraMasks,
    ...pal,
    ...pixels,
  ]);
}

/** Classic TIFF: data at offset 8, then the IFD; entries are [tag, type, values]. */
function tiff(
  le: boolean,
  entries: [number, number, number[]][],
  strips: Uint8Array[],
): Uint8Array {
  const out: number[] = [];
  const w16 = (v: number): number[] => (le ? u16le(v) : [v >>> 8, v & 255]);
  const w32 = (v: number): number[] => (le ? u32le(v) : u32be(v));
  out.push(...(le ? [0x49, 0x49] : [0x4d, 0x4d]), ...w16(42), 0, 0, 0, 0);
  const stripOffs: number[] = [];
  for (const s of strips) {
    stripOffs.push(out.length);
    out.push(...s);
    if (out.length & 1) out.push(0);
  }
  type Entry = [number, number, number[]];
  const all: Entry[] = [
    ...entries,
    [273, 4, stripOffs] as Entry,
    [279, 4, strips.map((s) => s.length)] as Entry,
  ].sort((a, b) => a[0] - b[0]);
  const ifd = out.length;
  out.splice(4, 4, ...w32(ifd));
  const extra: number[] = [];
  const extraBase = ifd + 2 + all.length * 12 + 4;
  out.push(...w16(all.length));
  for (const [tag, type, vals] of all) {
    const size = type === 3 ? 2 : 4;
    const enc = vals.flatMap((v) => (size === 2 ? w16(v) : w32(v)));
    out.push(...w16(tag), ...w16(type), ...w32(vals.length));
    if (enc.length <= 4) out.push(...enc, ...new Array(4 - enc.length).fill(0));
    else {
      out.push(...w32(extraBase + extra.length));
      extra.push(...enc);
    }
  }
  out.push(0, 0, 0, 0, ...extra);
  return Uint8Array.from(out);
}

function packbits(b: Uint8Array): Uint8Array {
  const out: number[] = [];
  let i = 0;
  while (i < b.length) {
    let run = 1;
    while (i + run < b.length && run < 128 && b[i + run] === b[i]) run++;
    if (run >= 2) {
      out.push(257 - run, b[i]!);
      i += run;
    } else {
      out.push(0, b[i]!);
      i++;
    }
  }
  return Uint8Array.from(out);
}

const b64 = (s: string): Uint8Array => Uint8Array.from(Buffer.from(s, 'base64'));
/** 16×16 JPEG (macOS sips), quadrants red / green / blue / white. */
const JPEG = b64(
  '/9j/4AAQSkZJRgABAQAASABIAAD/4QBMRXhpZgAATU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKAC' +
    'AAQAAAABAAAAEKADAAQAAAABAAAAEAAAAAD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZ' +
    'jwCyBOmACZjs+EJ+/8AAEQgAEAAQAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIB' +
    'AwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNE' +
    'RUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfI' +
    'ycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIB' +
    'AgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpD' +
    'REVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXG' +
    'x8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/bAEMAAQEBAQEBAgEBAgMCAgIDBAMDAwMEBQQEBAQEBQYFBQUFBQUG' +
    'BgYGBgYGBgcHBwcHBwgICAgICQkJCQkJCQkJCf/bAEMBAQEBAgICBAICBAkGBQYJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJ' +
    'CQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCf/dAAQAAf/aAAwDAQACEQMRAD8A/F+v6UK/g3r/AHEKP2mP7DP/AIhz/Yn/ABlH' +
    '1j6x9Y/5g/Z8vs/Yf9RU735/K1ut9J/bBeJX/E6P+rv7n+x/7H+tfa+t+2+t/V/7uG9n7P6t/f5+f7PL73//2Q==',
);
/** 16×16 GIF87a (macOS sips), same quadrants. */
const GIF = b64(
  'R0lGODdhEAAQAKIAAAAAAAAA//8AAAD/AP///wAAAAAAAAAAACH5BAQAAAAALAAAAAAQABAAAAMsKLo6/o4xCOWizzY8tOCdBnqjyAVoihJsy6qq68Kp3NKrTeCBvuM+XhCoSwAAOw==',
);

// Quadrant test pattern (2×2): red, green / blue, white.
const RED = [255, 0, 0];
const GREEN = [0, 255, 0];
const BLUE = [0, 0, 255];
const WHITE = [255, 255, 255];
const QUAD = [RED, GREEN, BLUE, WHITE];

// ---- decoder tests ---------------------------------------------------------------------------

describe('PNG decoder', () => {
  const W = 13;
  const H = 11;
  for (const interlace of [false, true]) {
    for (const [ctype, depths] of [
      [0, [1, 2, 4, 8, 16]],
      [2, [8, 16]],
      [3, [1, 2, 4, 8]],
      [4, [8, 16]],
      [6, [8, 16]],
    ] as const) {
      for (const depth of depths) {
        it(`colour type ${ctype}, ${depth}-bit${interlace ? ', Adam7' : ''}`, () => {
          const ch = CHANNELS[ctype]!;
          const rb = Math.ceil((W * ch * depth) / 8);
          let rows = bytes(rb * H, ctype * 100 + depth);
          if (ctype === 3) {
            // keep indices inside a 3-entry palette... except for 1-bit which has 2
            rows = rows.map((v) => (depth === 8 ? v % 3 : v));
          }
          // Clear the padding bits at the end of each row (encoders write zeros there).
          const used = (W * ch * depth) % 8;
          if (used)
            for (let y = 0; y < H; y++) rows[y * rb + rb - 1]! &= (0xff << (8 - used)) & 0xff;
          const palette = ctype === 3 ? Array.from(bytes(3 * (1 << depth), 7)) : undefined;
          const img = decodeImage(
            png(W, H, ctype, depth, rows, {
              interlace,
              ...(palette ? { palette } : {}),
              split: true,
            }),
          );
          expect([img.width, img.height, img.bpc]).toEqual([W, H, depth]);
          if (ctype === 4 || ctype === 6) {
            const sb = depth / 8;
            const cb = (ch - 1) * sb;
            const color: number[] = [];
            const alpha: number[] = [];
            for (let i = 0; i < W * H; i++) {
              color.push(...rows.subarray(i * ch * sb, i * ch * sb + cb));
              alpha.push(...rows.subarray(i * ch * sb + cb, (i + 1) * ch * sb));
            }
            expect(Array.from(img.data)).toEqual(color);
            expect(Array.from(img.alpha!)).toEqual(alpha);
            expect(img.alphaBpc).toBe(depth);
          } else {
            expect(Array.from(img.data)).toEqual(Array.from(rows));
            expect(img.alpha).toBeUndefined();
          }
          expect(img.color).toBe(ctype === 0 || ctype === 4 ? 'gray' : 'rgb');
          if (ctype === 3) expect(Array.from(img.palette!)).toEqual(palette);
        });
      }
    }
  }

  it('palette transparency becomes an alpha plane', () => {
    const rows = Uint8Array.from([0b00011011]); // 2-bit indices 0,1,2,3
    const img = decodeImage(
      png(4, 1, 3, 2, rows, { palette: [...RED, ...GREEN, ...BLUE, ...WHITE], trns: [0, 128] }),
    );
    expect(Array.from(img.alpha!)).toEqual([0, 128, 255, 255]);
  });

  it('grey and RGB tRNS become colour-key masks', () => {
    expect(decodeImage(png(1, 1, 0, 8, Uint8Array.of(7), { trns: [0, 7] })).colorKey).toEqual([
      7, 7,
    ]);
    expect(
      decodeImage(png(1, 1, 2, 8, Uint8Array.of(1, 2, 3), { trns: [0, 1, 0, 2, 0, 3] })).colorKey,
    ).toEqual([1, 1, 2, 2, 3, 3]);
  });

  it('drops a fully opaque alpha channel', () => {
    const img = decodeImage(png(2, 1, 6, 8, Uint8Array.of(1, 2, 3, 255, 4, 5, 6, 255)));
    expect(img.alpha).toBeUndefined();
    expect(Array.from(img.data)).toEqual([1, 2, 3, 4, 5, 6]);
  });
});

describe('GIF decoder', () => {
  const pal = [...RED, ...GREEN, ...BLUE, ...WHITE, 9, 9, 9];
  const idx = Uint8Array.from({ length: 7 * 9 }, (_, i) => (i * 7) % 5);
  for (const variant of [{}, { interlace: true }, { local: true }, { transparent: 4 }]) {
    it(`decodes ${JSON.stringify(variant)}`, () => {
      const img = decodeImage(gif(7, 9, pal, idx, variant));
      expect([img.width, img.height, img.bpc]).toEqual([7, 9, 8]);
      expect(Array.from(img.data)).toEqual(Array.from(idx));
      expect(Array.from(img.palette!).slice(0, 15)).toEqual(pal);
      if ('transparent' in variant) expect(img.colorKey).toEqual([4, 4]);
    });
  }

  it('decodes a real encoder’s LZW stream', () => {
    const img = decodeImage(GIF);
    expect([img.width, img.height]).toEqual([16, 16]);
    const at = (x: number, y: number): number[] => {
      const i = img.data[y * 16 + x]!;
      return Array.from(img.palette!.subarray(i * 3, i * 3 + 3));
    };
    expect([at(2, 2), at(12, 2), at(2, 12), at(12, 12)]).toEqual(QUAD);
  });
});

describe('BMP decoder', () => {
  const rgba = [
    [255, 0, 0, 255],
    [0, 255, 0, 128],
    [0, 0, 255, 0],
    [255, 255, 255, 255],
  ];
  const rgb = [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255];

  it('24-bit bottom-up and top-down, core header', () => {
    for (const o of [{}, { topDown: true }, { hdr: 12 as const }]) {
      const img = decodeImage(bmp(2, 2, 24, { rgba, ...o }));
      expect(Array.from(img.data)).toEqual(rgb);
      expect(img.alpha).toBeUndefined();
    }
  });

  it('32-bit: padding byte ignored when all zero, alpha via BITFIELDS', () => {
    const zeroA = rgba.map((c) => [c[0]!, c[1]!, c[2]!, 0]);
    expect(decodeImage(bmp(2, 2, 32, { rgba: zeroA })).alpha).toBeUndefined();
    const masks = [0xff0000, 0xff00, 0xff, 0xff000000];
    const img = decodeImage(bmp(2, 2, 32, { rgba, hdr: 108, comp: 3, masks }));
    expect(Array.from(img.data)).toEqual(rgb);
    expect(Array.from(img.alpha!)).toEqual([255, 128, 0, 255]);
  });

  it('16-bit 565 bitfields', () => {
    const img = decodeImage(bmp(2, 2, 16, { rgba, comp: 3, masks: [0xf800, 0x7e0, 0x1f] }));
    expect(Array.from(img.data)).toEqual(rgb);
  });

  it('1/4/8-bit palettes stay indexed', () => {
    for (const bpp of [1, 4, 8]) {
      const img = decodeImage(bmp(3, 2, bpp, { idx: [0, 1, 1, 1, 0, 1], palette: [RED, BLUE] }));
      expect(img.bpc).toBe(bpp);
      expect(Array.from(img.palette!)).toEqual([...RED, ...BLUE]);
      const rb = Math.ceil((3 * bpp) / 8);
      const read = (y: number, x: number): number => getBits(img.data, y * rb * 8 + x * bpp, bpp);
      expect([read(0, 0), read(0, 1), read(0, 2), read(1, 0), read(1, 1), read(1, 2)]).toEqual([
        0, 1, 1, 1, 0, 1,
      ]);
    }
  });

  it('rejects RLE', () => {
    expect(() => decodeImage(bmp(2, 1, 8, { idx: [0, 0], palette: [RED], comp: 1 }))).toThrow(
      /compression/,
    );
  });
});

describe('TIFF decoder', () => {
  const base = (
    w: number,
    h: number,
    bps: number[],
    photo: number,
    spp = bps.length,
  ): [number, number, number[]][] => [
    [256, 3, [w]],
    [257, 3, [h]],
    [258, 3, bps],
    [259, 3, [1]],
    [262, 3, [photo]],
    [277, 3, [spp]],
    [278, 3, [h]],
  ];

  it('uncompressed RGB, little- and big-endian, several strips', () => {
    const data = Uint8Array.from(rgb2x2());
    for (const le of [true, false]) {
      const entries = base(2, 2, [8, 8, 8], 2).map((e): [number, number, number[]] =>
        e[0] === 278 ? [278, 3, [1]] : e,
      );
      const img = decodeImage(tiff(le, entries, [data.subarray(0, 6), data.subarray(6)]));
      expect(img.color).toBe('rgb');
      expect(Array.from(img.data)).toEqual(Array.from(data));
    }
  });

  it('PackBits grey, WhiteIsZero bilevel, 16-bit byte order', () => {
    const grey = Uint8Array.from([5, 5, 5, 5, 9, 1, 2, 3]);
    const e = base(4, 2, [8], 1).map((x): [number, number, number[]] =>
      x[0] === 259 ? [259, 3, [32773]] : x,
    );
    expect(Array.from(decodeImage(tiff(false, e, [packbits(grey)])).data)).toEqual(
      Array.from(grey),
    );
    const bi = decodeImage(tiff(true, base(8, 1, [1], 0), [Uint8Array.of(0b10100000)]));
    expect(bi.decode).toEqual([1, 0]);
    expect(bi.bpc).toBe(1);
    const g16 = decodeImage(
      tiff(true, base(2, 1, [16], 1), [Uint8Array.of(0x34, 0x12, 0x78, 0x56)]),
    );
    expect(Array.from(g16.data)).toEqual([0x12, 0x34, 0x56, 0x78]);
  });

  it('palette, CMYK and RGBA (extra sample alpha)', () => {
    const map = [65535, 0, 0, 0, 0, 65535, 0, 0, 0, 0, 0, 0].slice(0, 12);
    // 2-bit palette: R = [ffff,0,0,0], G = [0,ffff,0,0], B = [0,0,0,0]
    const pal = decodeImage(
      tiff(true, [...base(2, 1, [2], 3), [320, 3, map]], [Uint8Array.of(0b00010000)]),
    );
    expect(Array.from(pal.palette!.subarray(0, 6))).toEqual([255, 0, 0, 0, 255, 0]);
    expect(
      decodeImage(tiff(true, base(1, 1, [8, 8, 8, 8], 5), [Uint8Array.of(1, 2, 3, 4)])).color,
    ).toBe('cmyk');
    const rgbaImg = decodeImage(
      tiff(
        true,
        [...base(2, 1, [8, 8, 8, 8], 2), [338, 3, [2]]],
        [Uint8Array.of(1, 2, 3, 40, 4, 5, 6, 255)],
      ),
    );
    expect(Array.from(rgbaImg.data)).toEqual([1, 2, 3, 4, 5, 6]);
    expect(Array.from(rgbaImg.alpha!)).toEqual([40, 255]);
  });

  it('rejects LZW and tiles', () => {
    const lzw = base(1, 1, [8], 1).map((x): [number, number, number[]] =>
      x[0] === 259 ? [259, 3, [5]] : x,
    );
    expect(() => decodeImage(tiff(true, lzw, [Uint8Array.of(0)]))).toThrow(/compression 5/);
    expect(() =>
      decodeImage(tiff(true, [...base(1, 1, [8], 1), [322, 3, [16]]], [Uint8Array.of(0)])),
    ).toThrow(/tiled/);
  });
});

function rgb2x2(): number[] {
  return QUAD.flat();
}

describe('JPEG passthrough', () => {
  it('reads the frame header and keeps the bytes', () => {
    const img = decodeImage(JPEG);
    expect(img).toMatchObject({ width: 16, height: 16, bpc: 8, color: 'rgb', dct: true });
    expect(img.data).toBe(JPEG);
  });

  it('inverts Adobe CMYK and handles grey', () => {
    const seg = (m: number, body: number[]): number[] => [
      0xff,
      m,
      (body.length + 2) >> 8,
      (body.length + 2) & 255,
      ...body,
    ];
    const sof = (comps: number): number[] =>
      seg(0xc0, [8, 0, 3, 0, 5, comps, ...new Array(comps * 3).fill(1)]);
    const adobe = seg(0xee, [...[...'Adobe'].map((c) => c.charCodeAt(0)), 0, 100, 0, 0, 0, 0, 2]);
    const cmyk = decodeImage(Uint8Array.from([0xff, 0xd8, ...adobe, ...sof(4), 0xff, 0xd9]));
    expect(cmyk).toMatchObject({
      width: 5,
      height: 3,
      color: 'cmyk',
      decode: [1, 0, 1, 0, 1, 0, 1, 0],
    });
    const plain = decodeImage(Uint8Array.from([0xff, 0xd8, ...sof(4), 0xff, 0xd9]));
    expect(plain.decode).toBeUndefined();
    expect(decodeImage(Uint8Array.from([0xff, 0xd8, ...sof(1), 0xff, 0xd9])).color).toBe('gray');
  });
});

// ---- rendering -------------------------------------------------------------------------------

/** The 2×2 quadrant pattern in every format. */
function quadImages(): ImageResource[] {
  const res = (id: string, contentType: string, b: Uint8Array): ImageResource => ({
    id,
    contentType,
    bytes: b,
  });
  const rgba = QUAD.map((c, i) => [...c, i === 3 ? 254 : 255]); // a real (if faint) alpha plane
  return [
    res('png-rgb', 'image/png', png(2, 2, 2, 8, Uint8Array.from(rgb2x2()))),
    res(
      'png-rgb16-i',
      'image/png',
      png(2, 2, 2, 16, Uint8Array.from(rgb2x2().flatMap((v) => [v, v])), { interlace: true }),
    ),
    res(
      'png-pal',
      'image/png',
      png(2, 2, 3, 2, Uint8Array.of(0b00010000, 0b10110000), { palette: rgb2x2() }),
    ),
    res('png-rgba', 'image/png', png(2, 2, 6, 8, Uint8Array.from(rgba.flat()))),
    res('jpeg', 'image/jpeg', JPEG),
    res('gif', 'image/gif', gif(2, 2, rgb2x2(), Uint8Array.of(0, 1, 2, 3))),
    res('gif-sips', 'image/gif', GIF),
    res('bmp24', 'image/bmp', bmp(2, 2, 24, { rgba })),
    res('bmp8', 'image/bmp', bmp(2, 2, 8, { idx: [0, 1, 2, 3], palette: QUAD })),
    res(
      'tiff',
      'image/tiff',
      tiff(
        true,
        [
          [256, 3, [2]],
          [257, 3, [2]],
          [258, 3, [8, 8, 8]],
          [262, 3, [2]],
          [277, 3, [3]],
        ],
        [Uint8Array.from(rgb2x2())],
      ),
    ),
  ];
}

const close = (a: number[], b: number[], tol = 40): boolean =>
  a.every((v, i) => Math.abs(v - b[i]!) <= tol);

describe('image rendering', () => {
  const images = quadImages();
  const page = (item: Partial<ImageItem> & { image: string }) => ({
    width: 72,
    height: 72,
    items: [{ kind: 'image' as const, x: 0, y: 0, width: 72, height: 72, ...item }],
  });

  it.skipIf(!HAS_QPDF)('every format embeds into a valid PDF', () => {
    const warnings: string[] = [];
    const pdf = renderPdf(
      layoutOf(
        images.map((im) => page({ image: im.id })),
        [],
        images,
      ),
      {
        onWarning: (w) => warnings.push(w.message),
      },
    );
    expect(warnings).toEqual([]);
    dump('images.pdf', pdf);
    expect(qpdfCheck(pdf).status).toBe(0);
    if (HAS_POPPLER) {
      const list = runOn(pdf, 'pdfimages', ['-list']).stdout;
      expect(list.match(/^\s+\d+\s+\d+\s+image/gm)?.length).toBe(images.length);
      expect(list).toMatch(/smask/); // png-rgba
      expect(list).toMatch(/jpeg/);
      expect(list).toMatch(/index/);
    }
  });

  for (const im of images) {
    it.skipIf(!HAS_POPPLER)(`${im.id} renders upright with the right colours`, () => {
      const r = rasterize(renderPdf(layoutOf([page({ image: im.id })], [], images)), 72);
      expect(r.stderr).toBe('');
      const got = [r.px(18, 18), r.px(54, 18), r.px(18, 54), r.px(54, 54)];
      got.forEach((c, i) => {
        expect(close(c, QUAD[i]!), `${im.id} quadrant ${i}: ${c}`).toBe(true);
      });
    });
  }

  it.skipIf(!HAS_POPPLER)('alpha composites over the page', () => {
    const half = png(
      2,
      2,
      6,
      8,
      Uint8Array.from([255, 0, 0, 0, 0, 255, 0, 255, 0, 0, 255, 255, 0, 0, 0, 128]),
    );
    const imgs = [{ id: 'a', contentType: 'image/png', bytes: half }];
    const r = rasterize(renderPdf(layoutOf([page({ image: 'a' })], [], imgs)), 72);
    expect(r.px(18, 18)).toEqual([255, 255, 255]); // transparent
    expect(close(r.px(54, 54), [127, 127, 127], 8)).toBe(true); // 50% black
  });

  it.skipIf(!HAS_POPPLER)('crops and rotates', () => {
    // Crop away the left half: green over white fills the box.
    const crop = rasterize(
      renderPdf(
        layoutOf(
          [page({ image: 'png-rgb', crop: { left: 0.5, top: 0, right: 0, bottom: 0 } })],
          [],
          images,
        ),
      ),
    );
    expect([crop.px(18, 18), crop.px(54, 18), crop.px(18, 54), crop.px(54, 54)]).toEqual([
      GREEN,
      GREEN,
      WHITE,
      WHITE,
    ]);
    // 90° clockwise: red moves to the top right.
    const rot = rasterize(
      renderPdf(layoutOf([page({ image: 'png-rgb', rotation: 90 })], [], images)),
    );
    expect([rot.px(18, 18), rot.px(54, 18), rot.px(18, 54), rot.px(54, 54)]).toEqual([
      BLUE,
      RED,
      WHITE,
      GREEN,
    ]);
  });

  it('unsupported formats draw a placeholder and warn once', () => {
    const emf = {
      id: 'emf',
      contentType: 'image/x-emf',
      bytes: Uint8Array.of(1, 0, 0, 0, 0x6c, 0, 0, 0),
    };
    const codes: string[] = [];
    const pdf = renderPdf(
      layoutOf(
        [
          [
            ...[1, 2].map(() => page({ image: 'emf' }).items[0]!),
            page({ image: 'nope' }).items[0]!,
          ],
        ],
        [],
        [emf],
      ),
      { onWarning: (w) => codes.push(w.code), compress: false },
    );
    expect(codes).toEqual(['PDF_IMAGE_UNSUPPORTED', 'PDF_IMAGE_MISSING']);
    const s = Buffer.from(pdf).toString('latin1');
    expect(s).toContain('0.85 g 0 720 72 72 re f');
    expect(s).not.toContain('/XObject');
    if (HAS_QPDF) expect(qpdfCheck(pdf).status).toBe(0);
  });

  it('shares one XObject per image across pages', () => {
    const pdf = renderPdf(layoutOf([page({ image: 'jpeg' }), page({ image: 'jpeg' })], [], images));
    const s = Buffer.from(pdf).toString('latin1');
    expect(s.match(/\/Subtype \/Image/g)?.length).toBe(1);
  });
});
