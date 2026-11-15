// JPEG passthrough: the file is embedded as-is with /DCTDecode. Only the frame header is parsed
// (size, components). Adobe APP14 CMYK JPEGs store inverted samples, hence the /Decode inversion.

import type { RasterImage } from './raster.js';

export function decodeJpeg(b: Uint8Array): RasterImage {
  if (b[0] !== 0xff || b[1] !== 0xd8) throw new Error('not a JPEG');
  let p = 2;
  let adobe = false;
  while (p + 4 <= b.length) {
    if (b[p] !== 0xff) {
      p++;
      continue;
    }
    const marker = b[p + 1]!;
    if (marker === 0xff) {
      p++; // fill byte
      continue;
    }
    p += 2;
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    const len = (b[p]! << 8) | b[p + 1]!;
    if (marker === 0xee && String.fromCharCode(...b.subarray(p + 2, p + 7)) === 'Adobe') {
      adobe = true;
    }
    const sof =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (sof && p + 8 <= b.length) {
      const height = (b[p + 3]! << 8) | b[p + 4]!;
      const width = (b[p + 5]! << 8) | b[p + 6]!;
      const comps = b[p + 7]!;
      // Keep scanning for a later APP14? It precedes SOF in practice; stop here.
      if (comps !== 1 && comps !== 3 && comps !== 4)
        throw new Error(`JPEG with ${comps} components`);
      const img: RasterImage = {
        width,
        height,
        bpc: 8,
        color: comps === 1 ? 'gray' : comps === 3 ? 'rgb' : 'cmyk',
        data: b,
        dct: true,
      };
      if (comps === 4 && adobe) img.decode = [1, 0, 1, 0, 1, 0, 1, 0];
      if (!width || !height) throw new Error('JPEG without dimensions');
      return img;
    }
    if (marker === 0xda) break;
    p += len;
  }
  throw new Error('JPEG frame header not found');
}
