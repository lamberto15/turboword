// PNG decoder: chunks (IHDR, PLTE, tRNS, IDAT), zlib inflate (fflate), the five scanline filters,
// Adam7 de-interlacing, bit depths 1-16 and every colour type. Palettes stay palettes (/Indexed);
// alpha channels and palette transparency become an alpha plane (/SMask); a tRNS colour key on
// grey/RGB images becomes a /Mask range. Gamma/ICC chunks are ignored.

import { unzlibSync } from 'fflate';
import { checkSize, opaque, type RasterImage, rowBytes, sample } from './raster.js';

const CHANNELS: Readonly<Record<number, number>> = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };
// Adam7 passes: x0, y0, dx, dy.
const ADAM7 = [
  [0, 0, 8, 8],
  [4, 0, 8, 8],
  [0, 4, 4, 8],
  [2, 0, 4, 4],
  [0, 2, 2, 4],
  [1, 0, 2, 2],
  [0, 1, 1, 2],
] as const;

export function decodePng(b: Uint8Array): RasterImage {
  const sig = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  if (!sig.every((v, i) => b[i] === v)) throw new Error('not a PNG');
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let width = 0;
  let height = 0;
  let depth = 8;
  let ctype = 0;
  let interlace = 0;
  let palette: Uint8Array | undefined;
  let trns: Uint8Array | undefined;
  const idat: Uint8Array[] = [];
  let p = 8;
  while (p + 8 <= b.length) {
    const len = dv.getUint32(p);
    const type = String.fromCharCode(b[p + 4]!, b[p + 5]!, b[p + 6]!, b[p + 7]!);
    const data = b.subarray(p + 8, Math.min(b.length, p + 8 + len));
    if (type === 'IHDR') {
      width = dv.getUint32(p + 8);
      height = dv.getUint32(p + 12);
      depth = data[8]!;
      ctype = data[9]!;
      interlace = data[12]!;
    } else if (type === 'PLTE') palette = data;
    else if (type === 'tRNS') trns = data;
    else if (type === 'IDAT') idat.push(data);
    else if (type === 'IEND') break;
    p += 12 + len;
  }
  const channels = CHANNELS[ctype];
  if (!channels || ![1, 2, 4, 8, 16].includes(depth)) throw new Error('unsupported PNG format');
  checkSize(width, height);
  let total = 0;
  for (const c of idat) total += c.length;
  const z = new Uint8Array(total);
  total = 0;
  for (const c of idat) {
    z.set(c, total);
    total += c.length;
  }
  const raw = unzlibSync(z);
  const bitsPP = channels * depth;
  const bpp = Math.max(1, bitsPP >> 3);
  const rb = rowBytes(width, bitsPP);
  let pixels = new Uint8Array(rb * height);
  if (!interlace) unfilter(raw, 0, rb, height, bpp, pixels);
  else {
    let off = 0;
    for (const [x0, y0, dx, dy] of ADAM7) {
      const pw = Math.ceil((width - x0) / dx);
      const ph = Math.ceil((height - y0) / dy);
      if (pw <= 0 || ph <= 0) continue;
      const prb = rowBytes(pw, bitsPP);
      const pass = new Uint8Array(prb * ph);
      off = unfilter(raw, off, prb, ph, bpp, pass);
      for (let j = 0; j < ph; j++) {
        for (let i = 0; i < pw; i++) {
          const x = x0 + i * dx;
          const y = y0 + j * dy;
          if (bitsPP >= 8) {
            pixels.set(pass.subarray(j * prb + i * bpp, j * prb + i * bpp + bpp), y * rb + x * bpp);
          } else {
            const v = sample(pass, j * prb, i, depth);
            const bit = x * depth;
            pixels[y * rb + (bit >> 3)]! |= v << (8 - depth - (bit & 7));
          }
        }
      }
    }
  }

  if (ctype === 3) {
    if (!palette) throw new Error('PNG palette missing');
    const img: RasterImage = { width, height, bpc: depth, color: 'rgb', palette, data: pixels };
    if (trns?.some((a) => a !== 0xff)) {
      const alpha = new Uint8Array(width * height);
      for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
          alpha[y * width + x] = trns[sample(pixels, y * rb, x, depth)] ?? 0xff;
        }
      }
      img.alpha = alpha;
      img.alphaBpc = 8;
    }
    return img;
  }
  if (ctype === 0 || ctype === 2) {
    const img: RasterImage = {
      width,
      height,
      bpc: depth,
      color: ctype ? 'rgb' : 'gray',
      data: pixels,
    };
    if (trns && trns.length >= channels * 2) {
      const key: number[] = [];
      for (let c = 0; c < channels; c++) {
        const v = ((trns[c * 2]! << 8) | trns[c * 2 + 1]!) & ((1 << depth) - 1);
        key.push(v, v);
      }
      img.colorKey = key;
    }
    return img;
  }
  // Grey+alpha / RGBA: split the alpha samples out.
  const sb = depth >> 3;
  const cb = (channels - 1) * sb;
  const color = new Uint8Array(width * height * cb);
  const alpha = new Uint8Array(width * height * sb);
  for (let i = 0, n = width * height; i < n; i++) {
    const s = i * bpp;
    color.set(pixels.subarray(s, s + cb), i * cb);
    alpha.set(pixels.subarray(s + cb, s + bpp), i * sb);
  }
  pixels = color;
  const img: RasterImage = {
    width,
    height,
    bpc: depth,
    color: ctype === 4 ? 'gray' : 'rgb',
    data: pixels,
  };
  if (!opaque(alpha)) {
    img.alpha = alpha;
    img.alphaBpc = depth;
  }
  return img;
}

/** Undoes the per-row filters of `h` rows of `rb` bytes starting at `off`; returns the end offset. */
function unfilter(
  src: Uint8Array,
  off: number,
  rb: number,
  h: number,
  bpp: number,
  out: Uint8Array,
): number {
  let p = off;
  for (let y = 0; y < h; y++) {
    const ft = src[p++] ?? 0;
    const row = y * rb;
    const prev = row - rb;
    for (let x = 0; x < rb; x++) {
      const raw = src[p + x] ?? 0;
      const a = x >= bpp ? out[row + x - bpp]! : 0;
      const up = y > 0 ? out[prev + x]! : 0;
      let v = raw;
      if (ft === 1) v = raw + a;
      else if (ft === 2) v = raw + up;
      else if (ft === 3) v = raw + ((a + up) >> 1);
      else if (ft === 4) {
        const c = y > 0 && x >= bpp ? out[prev + x - bpp]! : 0;
        const pa = Math.abs(up - c);
        const pb = Math.abs(a - c);
        const pc = Math.abs(a + up - 2 * c);
        v = raw + (pa <= pb && pa <= pc ? a : pb <= pc ? up : c);
      }
      out[row + x] = v & 0xff;
    }
    p += rb;
  }
  return p;
}
