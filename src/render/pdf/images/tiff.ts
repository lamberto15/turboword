// TIFF decoder for the simple cases Word documents carry: the first IFD, strips (not tiles),
// chunky planar layout, no compression or PackBits, no predictor. Bilevel/grey (both photometric
// senses), RGB, palette and CMYK, 1-16 bits per sample; an extra alpha sample becomes an /SMask.
// LZW/Deflate/JPEG/CCITT-compressed TIFFs throw (the renderer then draws a placeholder box).

import { checkSize, opaque, type RasterImage, rowBytes } from './raster.js';

const TYPE_SIZE: Readonly<Record<number, number>> = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 16: 8 };

export function decodeTiff(b: Uint8Array): RasterImage {
  const le = b[0] === 0x49 && b[1] === 0x49;
  if (!le && !(b[0] === 0x4d && b[1] === 0x4d)) throw new Error('not a TIFF');
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const u16 = (p: number): number => dv.getUint16(p, le);
  const u32 = (p: number): number => dv.getUint32(p, le);
  if (u16(2) !== 42) throw new Error('not a classic TIFF');
  const ifd = u32(4);
  const tags = new Map<number, number[]>();
  for (let i = 0, n = u16(ifd); i < n; i++) {
    const e = ifd + 2 + i * 12;
    const type = u16(e + 2);
    const count = u32(e + 4);
    const size = TYPE_SIZE[type] ?? 1;
    if (count > 1 << 20) continue;
    const at = count * size <= 4 ? e + 8 : u32(e + 8);
    const vals: number[] = [];
    for (let k = 0; k < count && at + k * size + size <= b.length; k++) {
      const p = at + k * size;
      vals.push(size === 1 ? b[p]! : size === 2 ? u16(p) : u32(p));
    }
    tags.set(u16(e), vals);
  }
  const one = (tag: number, def: number): number => tags.get(tag)?.[0] ?? def;
  const width = one(256, 0);
  const height = one(257, 0);
  checkSize(width, height);
  const bps = one(258, 1);
  const compression = one(259, 1);
  const photo = one(262, 1);
  const spp = one(277, 1);
  if (tags.has(322)) throw new Error('tiled TIFF');
  if (compression !== 1 && compression !== 32773)
    throw new Error(`TIFF compression ${compression}`);
  if (one(284, 1) !== 1 || one(317, 1) !== 1) throw new Error('unsupported TIFF layout');
  if (![1, 2, 4, 8, 16].includes(bps)) throw new Error(`TIFF depth ${bps}`);
  const base = photo === 2 ? 3 : photo === 5 ? 4 : 1;
  if (spp < base || (bps < 8 && spp !== base)) throw new Error('unsupported TIFF samples');

  const rb = rowBytes(width, spp * bps);
  const offsets = tags.get(273) ?? [];
  const counts = tags.get(279) ?? [];
  const raw = new Uint8Array(rb * height);
  let w = 0;
  offsets.forEach((off, i) => {
    const strip = b.subarray(off, off + (counts[i] ?? 0));
    const data = compression === 32773 ? packbits(strip) : strip;
    const n = Math.min(data.length, raw.length - w);
    raw.set(data.subarray(0, n), w);
    w += n;
  });
  if (bps === 16 && le) {
    for (let i = 0; i + 1 < raw.length; i += 2) {
      const t = raw[i]!;
      raw[i] = raw[i + 1]!;
      raw[i + 1] = t;
    }
  }

  let data = raw;
  let alpha: Uint8Array | undefined;
  if (spp > base) {
    // Keep the colour samples; the first extra sample is alpha when ExtraSamples says so.
    const sb = bps >> 3;
    const isAlpha = [1, 2].includes(tags.get(338)?.[0] ?? 0);
    data = new Uint8Array(width * height * base * sb);
    if (isAlpha) alpha = new Uint8Array(width * height * sb);
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const s = y * rb + x * spp * sb;
        const o = y * width + x;
        data.set(raw.subarray(s, s + base * sb), o * base * sb);
        alpha?.set(raw.subarray(s + base * sb, s + base * sb + sb), o * sb);
      }
    }
  }
  const img: RasterImage = {
    width,
    height,
    bpc: bps,
    color: photo === 2 ? 'rgb' : photo === 5 ? 'cmyk' : 'gray',
    data,
  };
  if (photo === 0) img.decode = [1, 0];
  else if (photo === 3) {
    const map = tags.get(320) ?? [];
    const n = Math.min(1 << bps, 256, Math.floor(map.length / 3));
    const palette = new Uint8Array(n * 3);
    for (let i = 0; i < n; i++) {
      for (let c = 0; c < 3; c++) palette[i * 3 + c] = (map[c * (map.length / 3) + i] ?? 0) >> 8;
    }
    img.color = 'rgb';
    img.palette = palette;
  } else if (photo !== 1 && photo !== 2 && photo !== 5) {
    throw new Error(`TIFF photometric ${photo}`);
  }
  if (alpha && !opaque(alpha)) {
    img.alpha = alpha;
    img.alphaBpc = bps;
  }
  return img;
}

/** PackBits run-length decoding. */
function packbits(src: Uint8Array): Uint8Array {
  const out: number[] = [];
  let p = 0;
  while (p < src.length) {
    const n = (src[p++]! << 24) >> 24;
    if (n >= 0) {
      for (let i = 0; i <= n && p < src.length; i++) out.push(src[p++]!);
    } else if (n !== -128) {
      const v = src[p++] ?? 0;
      for (let i = 0; i < 1 - n; i++) out.push(v);
    }
  }
  return Uint8Array.from(out);
}
