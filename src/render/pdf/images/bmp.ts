// BMP decoder: OS/2 core and Windows info headers (v3-v5), bottom-up or top-down rows, 1/4/8-bit
// palettes (kept as /Indexed), 24-bit BGR, 16/32-bit with default or BI_BITFIELDS masks (an alpha
// mask becomes an /SMask; 32-bit BI_RGB uses its fourth byte as alpha only when it is not all zero).
// RLE-compressed bitmaps are not supported.

import { checkSize, opaque, type RasterImage, rowBytes } from './raster.js';

export function decodeBmp(b: Uint8Array): RasterImage {
  if (b[0] !== 0x42 || b[1] !== 0x4d) throw new Error('not a BMP');
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const dataOff = dv.getUint32(10, true);
  const hdr = dv.getUint32(14, true);
  const core = hdr === 12;
  const width = core ? dv.getUint16(18, true) : dv.getInt32(18, true);
  const rawH = core ? dv.getInt16(20, true) : dv.getInt32(22, true);
  const bpp = dv.getUint16(core ? 24 : 28, true);
  const comp = core ? 0 : dv.getUint32(30, true);
  const height = Math.abs(rawH);
  const topDown = rawH < 0;
  checkSize(width, height);
  if (comp !== 0 && comp !== 3 && comp !== 6) throw new Error(`BMP compression ${comp}`);
  let masks: number[] | undefined;
  let palOff = 14 + hdr;
  if (comp === 3 || comp === 6) {
    const at = hdr >= 52 ? 54 : 14 + hdr;
    const n = comp === 6 || hdr >= 56 ? 4 : 3;
    masks = [];
    for (let i = 0; i < n; i++) masks.push(dv.getUint32(at + i * 4, true));
    if (hdr < 52) palOff += n * 4;
  }
  const stride = Math.floor((width * bpp + 31) / 32) * 4;
  const srcRow = (y: number): number => dataOff + (topDown ? y : height - 1 - y) * stride;
  if (stride * height + dataOff > b.length) throw new Error('truncated BMP');

  if (bpp <= 8) {
    if (![1, 4, 8].includes(bpp)) throw new Error(`BMP depth ${bpp}`);
    const entry = core ? 3 : 4;
    const used = core ? 0 : dv.getUint32(46, true);
    const n = Math.min(used || 1 << bpp, 256);
    const palette = new Uint8Array(n * 3);
    for (let i = 0; i < n; i++) {
      const e = palOff + i * entry;
      palette[i * 3] = b[e + 2] ?? 0;
      palette[i * 3 + 1] = b[e + 1] ?? 0;
      palette[i * 3 + 2] = b[e] ?? 0;
    }
    const rb = rowBytes(width, bpp);
    const data = new Uint8Array(rb * height);
    for (let y = 0; y < height; y++) data.set(b.subarray(srcRow(y), srcRow(y) + rb), y * rb);
    return { width, height, bpc: bpp, color: 'rgb', palette, data };
  }

  if (bpp !== 16 && bpp !== 24 && bpp !== 32) throw new Error(`BMP depth ${bpp}`);
  const rgb = new Uint8Array(width * height * 3);
  const alpha = new Uint8Array(width * height);
  const m = masks ?? (bpp === 16 ? [0x7c00, 0x3e0, 0x1f] : [0xff0000, 0xff00, 0xff, 0xff000000]);
  const chans = m.map((mask) => {
    const shift = mask ? 31 - Math.clz32(mask & -mask) : 0;
    return { mask: mask >>> 0, shift, max: mask ? (mask >>> 0) >>> shift : 0 };
  });
  for (let y = 0; y < height; y++) {
    const row = srcRow(y);
    for (let x = 0; x < width; x++) {
      const o = y * width + x;
      if (bpp === 24) {
        const s = row + x * 3;
        rgb[o * 3] = b[s + 2]!;
        rgb[o * 3 + 1] = b[s + 1]!;
        rgb[o * 3 + 2] = b[s]!;
        continue;
      }
      const v = bpp === 16 ? dv.getUint16(row + x * 2, true) : dv.getUint32(row + x * 4, true);
      for (let c = 0; c < chans.length; c++) {
        const ch = chans[c]!;
        const s = ch.max ? Math.round((((v & ch.mask) >>> ch.shift) * 255) / ch.max) : 255;
        if (c < 3) rgb[o * 3 + c] = s;
        else alpha[o] = s;
      }
    }
  }
  const img: RasterImage = { width, height, bpc: 8, color: 'rgb', data: rgb };
  const hasAlpha = bpp === 32 && chans.length > 3 && chans[3]!.max > 0;
  // 32-bit BI_RGB: the fourth byte is usually padding (all zero), sometimes real alpha.
  if (hasAlpha && (comp !== 0 || alpha.some((v) => v !== 0)) && !opaque(alpha)) {
    img.alpha = alpha;
    img.alphaBpc = 8;
  }
  return img;
}
