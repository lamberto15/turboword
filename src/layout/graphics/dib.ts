// Device-independent bitmaps (DIBs) as found in EMF/WMF bitmap records and pattern brushes:
// BITMAPCOREHEADER and BITMAPINFOHEADER (+ V4/V5), 1/4/8 bpp with a palette, 16/24/32 bpp,
// BI_RGB and BI_BITFIELDS, bottom-up and top-down rows. BI_JPEG / BI_PNG payloads are passed
// through as they are. RLE-compressed bitmaps are not decoded (undefined). Every read is
// bounds-checked; missing pixel data decodes as transparent.

/** A decoded bitmap: RGBA pixels (top row first), or a pass-through JPEG/PNG payload. */
export interface DecodedDib {
  width: number;
  height: number;
  /** width × height × 4 bytes, top row first. */
  rgba?: Uint8Array;
  /** Compressed payload (BI_JPEG / BI_PNG). */
  bytes?: Uint8Array;
  contentType?: string;
}

/** Largest bitmap decoded (pixels). */
const MAX_PIXELS = 1 << 24;

const BI_RGB = 0;
const BI_BITFIELDS = 3;
const BI_JPEG = 4;
const BI_PNG = 5;
const BI_ALPHABITFIELDS = 6;

interface Header {
  size: number;
  width: number;
  height: number;
  bpp: number;
  compression: number;
  colors: number;
  core: boolean;
  masks?: [number, number, number, number];
}

function readHeader(b: Uint8Array): Header | undefined {
  if (b.length < 12) return undefined;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const size = dv.getUint32(0, true);
  if (size === 12) {
    return {
      size,
      width: dv.getUint16(4, true),
      height: dv.getInt16(6, true),
      bpp: dv.getUint16(10, true),
      compression: BI_RGB,
      colors: 0,
      core: true,
    };
  }
  if (size < 40 || b.length < 40) return undefined;
  const h: Header = {
    size,
    width: dv.getInt32(4, true),
    height: dv.getInt32(8, true),
    bpp: dv.getUint16(14, true),
    compression: dv.getUint32(16, true),
    colors: dv.getUint32(32, true),
    core: false,
  };
  if (h.compression === BI_BITFIELDS || h.compression === BI_ALPHABITFIELDS) {
    const n = h.compression === BI_ALPHABITFIELDS ? 4 : 3;
    if (b.length >= 40 + n * 4) {
      const m = (i: number): number => (i < n ? dv.getUint32(40 + i * 4, true) : 0);
      h.masks = [m(0), m(1), m(2), size >= 56 && n === 3 ? dv.getUint32(52, true) : m(3)];
    }
  }
  return h;
}

/** Bytes taken by the header, colour masks and palette of a packed DIB. */
function infoSize(h: Header): number {
  let n = h.size;
  if (!h.core && h.size === 40) {
    if (h.compression === BI_BITFIELDS) n += 12;
    else if (h.compression === BI_ALPHABITFIELDS) n += 16;
  }
  const entries = h.colors > 0 ? Math.min(h.colors, 256) : h.bpp <= 8 ? 1 << h.bpp : 0;
  return n + entries * (h.core ? 3 : 4);
}

/** Shift and width of a contiguous bit mask. */
function maskInfo(mask: number): [number, number] {
  if (mask === 0) return [0, 0];
  let shift = 0;
  while (((mask >>> shift) & 1) === 0 && shift < 32) shift++;
  let bits = 0;
  while (((mask >>> (shift + bits)) & 1) === 1 && shift + bits < 32) bits++;
  return [shift, bits];
}

/**
 * Decodes a DIB from its BITMAPINFO (`bmi`: header, masks, palette) and pixel bits. `usage` 1
 * (DIB_PAL_COLORS) palettes hold palette indices, which are shown as greys.
 */
export function decodeDib(bmi: Uint8Array, bits: Uint8Array, usage = 0): DecodedDib | undefined {
  const h = readHeader(bmi);
  if (h === undefined) return undefined;
  const width = Math.abs(h.width);
  const height = Math.abs(h.height);
  if (width === 0 || height === 0 || width * height > MAX_PIXELS) return undefined;
  if (h.compression === BI_JPEG || h.compression === BI_PNG) {
    return {
      width,
      height,
      bytes: bits.slice(),
      contentType: h.compression === BI_JPEG ? 'image/jpeg' : 'image/png',
    };
  }
  if (h.compression !== BI_RGB && h.compression !== BI_BITFIELDS) {
    if (h.compression !== BI_ALPHABITFIELDS) return undefined;
  }
  const bpp = h.bpp;
  if (![1, 4, 8, 16, 24, 32].includes(bpp)) return undefined;
  // Palette.
  const palette: number[] = [];
  if (bpp <= 8) {
    let at = h.size;
    if (!h.core && h.size === 40 && h.compression === BI_BITFIELDS) at += 12;
    const entry = h.core ? 3 : 4;
    const count = h.colors > 0 ? Math.min(h.colors, 256) : 1 << bpp;
    for (let i = 0; i < count; i++) {
      const o = at + i * entry;
      if (usage === 1) {
        const v = Math.round((i * 255) / Math.max(1, count - 1));
        palette.push((v << 16) | (v << 8) | v);
      } else if (o + 2 < bmi.length) {
        palette.push(((bmi[o + 2] ?? 0) << 16) | ((bmi[o + 1] ?? 0) << 8) | (bmi[o] ?? 0));
      } else palette.push(0);
    }
  }
  let masks = h.masks;
  if (masks === undefined && bpp === 16) masks = [0x7c00, 0x03e0, 0x001f, 0];
  if (masks === undefined && bpp === 32 && h.compression !== BI_RGB) {
    masks = [0xff0000, 0xff00, 0xff, 0];
  }
  const mi = masks?.map(maskInfo) as [number, number][] | undefined;
  const chan = (v: number, k: number): number => {
    const m = mi?.[k];
    if (m === undefined || m[1] === 0) return k === 3 ? 255 : 0;
    const raw = (v >>> m[0]) & ((2 ** m[1] - 1) >>> 0);
    return m[1] >= 8 ? raw >>> (m[1] - 8) : Math.round((raw * 255) / (2 ** m[1] - 1));
  };
  const stride = (((width * bpp + 31) >>> 5) << 2) >>> 0;
  const rgba = new Uint8Array(width * height * 4);
  const bottomUp = h.height > 0;
  let alphaSeen = false;
  for (let row = 0; row < height; row++) {
    const src = row * stride;
    if (src + stride > bits.length) break;
    const y = bottomUp ? height - 1 - row : row;
    let o = y * width * 4;
    for (let x = 0; x < width; x++, o += 4) {
      let r = 0;
      let g = 0;
      let b = 0;
      let a = 255;
      if (bpp <= 8) {
        const bit = x * bpp;
        const byte = bits[src + (bit >>> 3)] ?? 0;
        const idx = (byte >>> (8 - bpp - (bit & 7))) & ((1 << bpp) - 1);
        const c = palette[idx] ?? 0;
        r = (c >>> 16) & 255;
        g = (c >>> 8) & 255;
        b = c & 255;
      } else if (bpp === 24) {
        const s = src + x * 3;
        b = bits[s] ?? 0;
        g = bits[s + 1] ?? 0;
        r = bits[s + 2] ?? 0;
      } else if (bpp === 16) {
        const s = src + x * 2;
        const v = (bits[s] ?? 0) | ((bits[s + 1] ?? 0) << 8);
        r = chan(v, 0);
        g = chan(v, 1);
        b = chan(v, 2);
        if (mi?.[3]?.[1]) a = chan(v, 3);
      } else {
        const s = src + x * 4;
        if (mi === undefined) {
          b = bits[s] ?? 0;
          g = bits[s + 1] ?? 0;
          r = bits[s + 2] ?? 0;
          a = bits[s + 3] ?? 0;
          if (a !== 0) alphaSeen = true;
        } else {
          const v =
            ((bits[s] ?? 0) |
              ((bits[s + 1] ?? 0) << 8) |
              ((bits[s + 2] ?? 0) << 16) |
              ((bits[s + 3] ?? 0) << 24)) >>>
            0;
          r = chan(v, 0);
          g = chan(v, 1);
          b = chan(v, 2);
          a = mi[3]?.[1] ? chan(v, 3) : 255;
        }
      }
      rgba[o] = r;
      rgba[o + 1] = g;
      rgba[o + 2] = b;
      rgba[o + 3] = a;
    }
  }
  // 32 bpp BI_RGB: the fourth byte is usually unused (all zero) → opaque.
  if (bpp === 32 && mi === undefined && !alphaSeen) {
    for (let i = 3; i < rgba.length; i += 4) rgba[i] = 255;
  }
  return { width, height, rgba };
}

/** Decodes a packed DIB (BITMAPINFO immediately followed by the bits, as in WMF records). */
export function decodePackedDib(dib: Uint8Array, usage = 0): DecodedDib | undefined {
  const h = readHeader(dib);
  if (h === undefined) return undefined;
  const at = infoSize(h);
  if (at > dib.length) return undefined;
  return decodeDib(dib.subarray(0, at), dib.subarray(at), usage);
}

/** Average colour of a decoded bitmap as RRGGBB (pattern brushes are approximated by it). */
export function averageDibColor(d: DecodedDib): string | undefined {
  const px = d.rgba;
  if (px === undefined || px.length === 0) return undefined;
  let r = 0;
  let g = 0;
  let b = 0;
  const n = px.length / 4;
  for (let i = 0; i < px.length; i += 4) {
    const a = (px[i + 3] ?? 255) / 255;
    r += (px[i] ?? 0) * a + 255 * (1 - a);
    g += (px[i + 1] ?? 0) * a + 255 * (1 - a);
    b += (px[i + 2] ?? 0) * a + 255 * (1 - a);
  }
  const h = (v: number): string =>
    Math.round(v / n)
      .toString(16)
      .padStart(2, '0');
  return (h(r) + h(g) + h(b)).toUpperCase();
}
