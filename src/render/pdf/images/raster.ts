// The decoded-image shape every decoder returns, ready to become a PDF image XObject: packed
// samples (each row byte-aligned, as both PDF and most formats store them), the colour space or an
// RGB palette for /Indexed, and optional transparency as an 8/16-bit alpha plane (/SMask) or a
// colour-key range (/Mask).

export interface RasterImage {
  width: number;
  height: number;
  /** Bits per sample (1, 2, 4, 8 or 16). */
  bpc: number;
  /** Base colour space (palette entries are always RGB). */
  color: 'gray' | 'rgb' | 'cmyk';
  /** RGB triples: the samples are indices into this palette (/Indexed). */
  palette?: Uint8Array;
  /** Packed rows, or the whole JPEG file when `dct`. */
  data: Uint8Array;
  /** JPEG passthrough (/DCTDecode). */
  dct?: boolean;
  /** /Decode array (inverted CMYK, WhiteIsZero). */
  decode?: number[];
  /** One sample per pixel, `alphaBpc` bits (8 or 16). */
  alpha?: Uint8Array;
  alphaBpc?: number;
  /** Colour-key mask ranges, [min max] per component. */
  colorKey?: number[];
}

/** Bytes in one packed row. */
export const rowBytes = (width: number, bitsPerPixel: number): number =>
  Math.ceil((width * bitsPerPixel) / 8);

/** Reads an `n`-bit sample (n ≤ 8, MSB first) at pixel index `i` of a packed row starting at `off`. */
export function sample(b: Uint8Array, off: number, i: number, n: number): number {
  if (n === 8) return b[off + i] ?? 0;
  const bit = i * n;
  return ((b[off + (bit >> 3)] ?? 0) >> (8 - n - (bit & 7))) & ((1 << n) - 1);
}

/** Rejects absurd dimensions before allocating. */
export function checkSize(width: number, height: number): void {
  if (!(width > 0 && height > 0) || width * height > 0x10000000) {
    throw new Error(`unsupported image size ${width}x${height}`);
  }
}

/** True when every alpha sample is fully opaque (so the mask can be dropped). */
export const opaque = (alpha: Uint8Array): boolean => alpha.every((v) => v === 0xff);
