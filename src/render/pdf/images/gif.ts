// GIF decoder: the first frame only (an animated GIF prints as its first frame, like Word does).
// Variable-width LZW, interlaced row order, global/local palettes and the graphics-control
// transparent index (a /Mask colour key on the indexed image).

import { checkSize, type RasterImage } from './raster.js';

export function decodeGif(b: Uint8Array): RasterImage {
  if (String.fromCharCode(...b.subarray(0, 4)) !== 'GIF8') throw new Error('not a GIF');
  const u16 = (p: number): number => (b[p] ?? 0) | ((b[p + 1] ?? 0) << 8);
  const flags = b[10] ?? 0;
  let p = 13;
  let global: Uint8Array | undefined;
  if (flags & 0x80) {
    const n = 3 << ((flags & 7) + 1);
    global = b.subarray(p, p + n);
    p += n;
  }
  let transparent = -1;
  const skipBlocks = (): void => {
    while (p < b.length && b[p] !== 0) p += b[p]! + 1;
    p++;
  };
  while (p < b.length) {
    const kind = b[p]!;
    if (kind === 0x21) {
      if (b[p + 1] === 0xf9 && b[p + 2]! >= 4 && b[p + 3]! & 1) transparent = b[p + 6]!;
      p += 2;
      skipBlocks();
    } else if (kind === 0x2c) {
      const width = u16(p + 5);
      const height = u16(p + 7);
      const f = b[p + 9] ?? 0;
      p += 10;
      checkSize(width, height);
      let palette = global;
      if (f & 0x80) {
        const n = 3 << ((f & 7) + 1);
        palette = b.subarray(p, p + n);
        p += n;
      }
      const minCode = b[p++] ?? 8;
      const parts: number[] = [];
      while (p < b.length && b[p] !== 0) {
        const n = b[p]!;
        for (let i = 1; i <= n; i++) parts.push(b[p + i] ?? 0);
        p += n + 1;
      }
      let px = lzw(Uint8Array.from(parts), minCode, width * height);
      if (f & 0x40) px = deinterlace(px, width, height);
      if (!palette) {
        palette = new Uint8Array(768);
        for (let i = 0; i < 768; i++) palette[i] = Math.floor(i / 3);
      }
      const img: RasterImage = { width, height, bpc: 8, color: 'rgb', palette, data: px };
      if (transparent >= 0) img.colorKey = [transparent, transparent];
      return img;
    } else break;
  }
  throw new Error('GIF without an image');
}

/** GIF LZW (LSB-first codes, clear/end codes, code width growing to 12 bits). */
function lzw(data: Uint8Array, minSize: number, count: number): Uint8Array {
  const out = new Uint8Array(count);
  const clear = 1 << minSize;
  const end = clear + 1;
  const prefix = new Int16Array(4096);
  const suffix = new Uint8Array(4096);
  const stack = new Uint8Array(4097);
  for (let i = 0; i < clear; i++) suffix[i] = i;
  let size = minSize + 1;
  let next = end + 1;
  let old = -1;
  let first = 0;
  let acc = 0;
  let bits = 0;
  let pos = 0;
  let op = 0;
  while (op < count) {
    while (bits < size && pos < data.length) {
      acc |= data[pos++]! << bits;
      bits += 8;
    }
    if (bits < size) break;
    let code = acc & ((1 << size) - 1);
    acc >>>= size;
    bits -= size;
    if (code === clear) {
      size = minSize + 1;
      next = end + 1;
      old = -1;
      continue;
    }
    if (code === end) break;
    if (old < 0) {
      if (code >= clear) break;
      out[op++] = code;
      old = first = code;
      continue;
    }
    const cur = code;
    let sp = 0;
    if (code >= next) {
      if (code > next) break; // corrupt
      stack[sp++] = first;
      code = old;
    }
    while (code > end) {
      stack[sp++] = suffix[code]!;
      code = prefix[code]!;
    }
    first = suffix[code]!;
    stack[sp++] = first;
    while (sp > 0 && op < count) out[op++] = stack[--sp]!;
    if (next < 4096) {
      prefix[next] = old;
      suffix[next] = first;
      next++;
      if (next === 1 << size && size < 12) size++;
    }
    old = cur;
  }
  return out;
}

/** Interlaced GIF rows come in four passes (every 8th from 0, 8th from 4, 4th from 2, 2nd from 1). */
function deinterlace(px: Uint8Array, w: number, h: number): Uint8Array {
  const out = new Uint8Array(px.length);
  let src = 0;
  for (const [start, step] of [
    [0, 8],
    [4, 8],
    [2, 4],
    [1, 2],
  ] as const) {
    for (let y = start; y < h; y += step) {
      out.set(px.subarray(src * w, src * w + w), y * w);
      src++;
    }
  }
  return out;
}
