// Minimal PNG encoder for bitmaps decoded from metafiles (EMF/WMF DIBs), so every renderer only
// has to understand PNG/JPEG. 8-bit RGBA, or RGB when every pixel is opaque; filter 0 on every
// row (the zlib stream from fflate does the compression). Deterministic for the same pixels.

import { zlibSync } from 'fflate';

let crcTable: Uint32Array | undefined;

function crc32(data: Uint8Array, start: number, end: number): number {
  if (crcTable === undefined) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (let i = start; i < end; i++) {
    c = (crcTable[(c ^ (data[i] ?? 0)) & 0xff] ?? 0) ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  dv.setUint32(8 + data.length, crc32(out, 4, 8 + data.length));
  return out;
}

/** Encodes `rgba` (width × height × 4 bytes, top row first) as a PNG file. */
export function encodePng(width: number, height: number, rgba: Uint8Array): Uint8Array {
  const w = Math.max(0, Math.floor(width));
  const h = Math.max(0, Math.floor(height));
  let opaque = true;
  for (let i = 3; i < w * h * 4; i += 4) {
    if ((rgba[i] ?? 255) !== 255) {
      opaque = false;
      break;
    }
  }
  const bpp = opaque ? 3 : 4;
  const stride = w * bpp + 1;
  const raw = new Uint8Array(stride * h);
  for (let y = 0; y < h; y++) {
    let o = y * stride + 1; // filter byte 0
    let s = y * w * 4;
    for (let x = 0; x < w; x++, s += 4) {
      raw[o++] = rgba[s] ?? 0;
      raw[o++] = rgba[s + 1] ?? 0;
      raw[o++] = rgba[s + 2] ?? 0;
      if (!opaque) raw[o++] = rgba[s + 3] ?? 0;
    }
  }
  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, w);
  dv.setUint32(4, h);
  ihdr[8] = 8; // bit depth
  ihdr[9] = opaque ? 2 : 6; // colour type: RGB / RGBA
  const parts = [
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlibSync(raw, { level: 6 })),
    chunk('IEND', new Uint8Array(0)),
  ];
  let len = 0;
  for (const p of parts) len += p.length;
  const out = new Uint8Array(len);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}
