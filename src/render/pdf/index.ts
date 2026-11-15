// PDF renderer (subpath turboword/pdf, experimental). Turns a PageModel (DocumentLayout) into a
// PDF 1.7 file; never reads the document or the layout IR.
//
// Module map: render.ts (renderPdf: document structure, links, outline, metadata, tagging),
// content.ts (page operators), path.ts (SVG path data), fonts.ts + sfnt.ts (embedding, GID-
// preserving TrueType subsetting, ToUnicode), images/ (JPEG passthrough; PNG, GIF, BMP, TIFF
// decoders), writer.ts (objects, xref, trailer /ID), md5.ts (identifiers).

import { paginate } from '../../layout/index.js';
import type { LayoutOptions } from '../../layout/types.js';
import type { WordDocument } from '../../wml/document.js';
import { type PdfOptions, renderPdf } from './render.js';

export { decodeImage, type RasterImage } from './images/index.js';
export { svgPathToPdf } from './path.js';
export { type PdfOptions, renderPdf } from './render.js';

/** paginate + renderPdf. */
export async function toPdf(
  doc: WordDocument,
  o: LayoutOptions & PdfOptions = {},
): Promise<Uint8Array> {
  return renderPdf(await paginate(doc, o), o);
}
