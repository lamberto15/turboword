// Image XObjects. Each ImageResource is decoded once, on first use, by sniffing its bytes (the
// content type is only a fallback hint) and written as one shared XObject, plus an /SMask for
// alpha. Formats that cannot be decoded (EMF/WMF/SVG, LZW TIFF, corrupt files) yield undefined so
// the page draws a grey placeholder box, and a warning is reported once per image.

import type { ImageResource } from '../../../layout/types.js';
import type { Warn } from '../fonts.js';
import { hex, type PdfWriter, ref } from '../writer.js';
import { decodeBmp } from './bmp.js';
import { decodeGif } from './gif.js';
import { decodeJpeg } from './jpeg.js';
import { decodePng } from './png.js';
import type { RasterImage } from './raster.js';
import { rowBytes } from './raster.js';
import { decodeTiff } from './tiff.js';

export type { RasterImage } from './raster.js';

/** Decodes any supported raster format (JPEG, PNG, GIF, BMP, TIFF). Throws when unsupported. */
export function decodeImage(bytes: Uint8Array, contentType = ''): RasterImage {
  const b0 = bytes[0];
  const b1 = bytes[1];
  if (b0 === 0xff && b1 === 0xd8) return decodeJpeg(bytes);
  if (b0 === 0x89 && b1 === 0x50) return decodePng(bytes);
  if (b0 === 0x47 && b1 === 0x49) return decodeGif(bytes);
  if (b0 === 0x42 && b1 === 0x4d) return decodeBmp(bytes);
  if ((b0 === 0x49 && b1 === 0x49) || (b0 === 0x4d && b1 === 0x4d)) return decodeTiff(bytes);
  throw new Error(`unsupported image format ${contentType || 'unknown'}`);
}

/** Collects the images the pages draw and writes them as XObjects. */
export class ImageRegistry {
  private readonly names = new Map<string, string | null>();
  private readonly entries: string[] = [];

  constructor(
    private readonly images: ReadonlyMap<string, ImageResource>,
    private readonly w: PdfWriter,
    private readonly warn: Warn,
  ) {}

  /** The XObject resource name for image `id`, or undefined when it cannot be drawn. */
  use(id: string): string | undefined {
    const known = this.names.get(id);
    if (known !== undefined) return known ?? undefined;
    const res = this.images.get(id);
    let name: string | null = null;
    if (!res) this.warn('PDF_IMAGE_MISSING', `image ${id} is not in the layout`);
    else {
      try {
        const obj = this.write(decodeImage(res.bytes, res.contentType));
        name = `Im${this.entries.length}`;
        this.entries.push(`/${name} ${ref(obj)}`);
      } catch (e) {
        this.warn(
          'PDF_IMAGE_UNSUPPORTED',
          `image ${id} (${res.contentType}) drawn as a placeholder: ${(e as Error).message}`,
        );
      }
    }
    this.names.set(id, name);
    return name ?? undefined;
  }

  /** The /XObject resource entries. */
  resources(): string {
    return this.entries.join(' ');
  }

  private write(img: RasterImage): number {
    const { width, height, bpc } = img;
    let cs =
      img.color === 'gray' ? '/DeviceGray' : img.color === 'cmyk' ? '/DeviceCMYK' : '/DeviceRGB';
    const comps = img.color === 'gray' ? 1 : img.color === 'cmyk' ? 4 : 3;
    if (img.palette && img.palette.length >= 3) {
      const pal = img.palette.subarray(0, Math.min(256, Math.floor(img.palette.length / 3)) * 3);
      cs = `[/Indexed /DeviceRGB ${pal.length / 3 - 1} <${hex(pal)}>]`;
    }
    let dict =
      `/Type /XObject /Subtype /Image /Width ${width} /Height ${height} ` +
      `/ColorSpace ${cs} /BitsPerComponent ${bpc}`;
    if (img.decode) dict += ` /Decode [${img.decode.join(' ')}]`;
    if (img.colorKey) dict += ` /Mask [${img.colorKey.join(' ')}]`;
    if (img.alpha) {
      const abpc = img.alphaBpc ?? 8;
      const mask = this.w.addStream(
        `/Type /XObject /Subtype /Image /Width ${width} /Height ${height} ` +
          `/ColorSpace /DeviceGray /BitsPerComponent ${abpc}`,
        exact(img.alpha, rowBytes(width, abpc) * height),
        true,
      );
      dict += ` /SMask ${ref(mask)}`;
    }
    if (img.dct) return this.w.addStream(`${dict} /Filter /DCTDecode`, img.data, false);
    const bits = img.palette ? bpc : bpc * comps;
    return this.w.addStream(dict, exact(img.data, rowBytes(width, bits) * height), true);
  }
}

/** Pads (or trims) sample data to the exact size the dictionary promises. */
function exact(data: Uint8Array, size: number): Uint8Array {
  if (data.length === size) return data;
  const out = new Uint8Array(size);
  out.set(data.subarray(0, size));
  return out;
}
