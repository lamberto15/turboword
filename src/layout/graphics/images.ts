// Image resources found by graphics code (pictures inside groups, blips of metafiles, bitmaps
// embedded in EMF/WMF records). Images are keyed by their part name (`word/media/image1.png`),
// derived ids for metafile bitmaps (`word/media/image3.emf#dib0`), so the same picture is
// registered once however many times it is drawn.

import type { Part } from '../../opc/part.js';
import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import type { LayoutEnv } from '../engine.js';
import type { ImageResource } from '../types.js';

/** Registers an image (first registration wins) and returns its id. */
export function addImage(
  env: Pick<LayoutEnv, 'images'>,
  id: string,
  contentType: string,
  bytes: Uint8Array,
  pixelWidth?: number,
  pixelHeight?: number,
): string {
  if (!env.images.has(id)) {
    const res: ImageResource = { id, contentType, bytes };
    if (pixelWidth !== undefined) res.pixelWidth = pixelWidth;
    if (pixelHeight !== undefined) res.pixelHeight = pixelHeight;
    env.images.set(id, res);
  }
  return id;
}

/** Content type of an image part, from the package or sniffed from its bytes. */
export function imageContentType(part: Part): string {
  const ct = part.contentType;
  if (ct !== undefined && ct !== 'application/octet-stream') return ct;
  const b = part.bytes();
  if (b[0] === 0x89 && b[1] === 0x50) return 'image/png';
  if (b[0] === 0xff && b[1] === 0xd8) return 'image/jpeg';
  if (b[0] === 0x47 && b[1] === 0x49) return 'image/gif';
  if (b[0] === 0x42 && b[1] === 0x4d) return 'image/bmp';
  if (b[0] === 0x01 && b[1] === 0 && b[2] === 0 && b[3] === 0) return 'image/x-emf';
  if (
    (b[0] === 0xd7 && b[1] === 0xcd) ||
    (b[0] === 1 && b[1] === 0) ||
    (b[0] === 2 && b[1] === 0)
  ) {
    return 'image/x-wmf';
  }
  const name = part.name.toLowerCase();
  if (name.endsWith('.emf')) return 'image/x-emf';
  if (name.endsWith('.wmf')) return 'image/x-wmf';
  return ct ?? 'application/octet-stream';
}

/** The image part an a:blip (r:embed) or v:imagedata (r:id) points at. */
export function blipPart(el: XElement, env: Pick<LayoutEnv, 'partOf'>): Part | undefined {
  const id = el.attrNS(NS.r, 'embed') ?? el.attrNS(NS.r, 'id') ?? el.attrNS(NS.r, 'link');
  if (id === undefined) return undefined;
  return env.partOf(el)?.related(id);
}
