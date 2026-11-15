// Images: format sniffing, image parts (deduplicated by content), and Word's inline drawing.
//
// An image in a .docx is three things: a binary part under `word/media/`, a relationship from
// the part that shows it (document, header, footer, notes, comments) and DrawingML markup in
// that part pointing at the relationship (`a:blip r:embed`). This module makes each of them the
// way Word does, so the result opens in Word without a repair prompt:
// - `imageInfo` reads the pixel size and resolution from the file header (PNG, JPEG, GIF, BMP,
//   TIFF, WebP, SVG, EMF, WMF): enough to size the picture without decoding it;
// - `addImagePart` stores identical bytes once per package (CRC and length first, then a byte
//   compare), names new parts `imageN.<ext>` and keeps [Content_Types].xml complete;
// - `inlineDrawing` writes the `wp:inline` block Word writes for "Insert > Picture", with a
//   document-unique `wp:docPr/@id` (from DocumentIds) and namespace declarations on the
//   DrawingML elements themselves, like Word, so the markup is valid wherever it is pasted.
// `imagesOf` lists every image reference across the stories (for WordDocument.images()).

import { UsageError } from '../core/errors.js';
import type { Part } from '../opc/part.js';
import { REL, relTypeIs } from '../opc/rel-types.js';
import { partDir, partKey, relativeTarget } from '../opc/uri.js';
import { XElement } from '../xml/dom.js';
import { isStrictNs, NS } from '../xml/ns.js';
import { crc32 } from '../zip/index.js';
import type { WordDocument } from './document.js';
import { DocumentIds, type NumericIdKind } from './ids.js';
import { setWAttr, wPrefix } from './order.js';
import type { Story } from './story.js';

export interface ImageInfo {
  /** image/png, image/jpeg, image/gif, image/bmp, image/tiff, image/svg+xml, image/x-emf, image/x-wmf, image/webp. */
  contentType: string;
  /**
   * Pixels. Vector formats (SVG, EMF, WMF) report their natural size in CSS pixels (1/96 inch),
   * which may be fractional.
   */
  width: number;
  height: number;
  /** Dots per inch when the file says (default 96). */
  dpiX: number;
  dpiY: number;
}

export interface AddedImage {
  /** Relationship id from `owner` to the image part. */
  relId: string;
  part: Part;
  info: ImageInfo;
}

// ---- sniffing -----------------------------------------------------------------------------

const DEFAULT_DPI = 96;

/** A resolution as Word uses it (whole dpi), or the 96 dpi default when absent or absurd. */
function dpi(v: number | undefined): number {
  if (v === undefined || !Number.isFinite(v)) return DEFAULT_DPI;
  const r = Math.round(v);
  return r >= 1 && r <= 100_000 ? r : DEFAULT_DPI;
}

function ascii(b: Uint8Array, at: number, n: number): string {
  let s = '';
  for (let i = at; i < at + n && i < b.length; i++) s += String.fromCharCode(b[i]!);
  return s;
}

function info(
  contentType: string,
  width: number,
  height: number,
  dpiX = DEFAULT_DPI,
  dpiY = dpiX,
): ImageInfo | undefined {
  if (!(width > 0 && height > 0 && Number.isFinite(width) && Number.isFinite(height))) {
    return undefined;
  }
  return { contentType, width, height, dpiX, dpiY };
}

/** PNG: IHDR size; pHYs pixels per metre (unit 1) as dpi. */
function png(b: Uint8Array, dv: DataView): ImageInfo | undefined {
  const width = dv.getUint32(16);
  const height = dv.getUint32(20);
  let dx = DEFAULT_DPI;
  let dy = DEFAULT_DPI;
  for (let p = 8, n = 0; p + 8 <= b.length && n < 1000; n++) {
    const len = dv.getUint32(p);
    const type = ascii(b, p + 4, 4);
    if (type === 'pHYs' && len >= 9 && b[p + 16] === 1) {
      dx = dpi(dv.getUint32(p + 8) * 0.0254);
      dy = dpi(dv.getUint32(p + 12) * 0.0254);
      break;
    }
    if (type === 'IDAT' || type === 'IEND') break;
    p += 12 + len;
  }
  return info('image/png', width, height, dx, dy);
}

/** Tags of the first IFD of a TIFF structure at `base` (SHORT, LONG and RATIONAL values only). */
function tiffTags(dv: DataView, base: number): Map<number, number> | undefined {
  const order = dv.getUint16(base);
  const le = order === 0x4949;
  if (!le && order !== 0x4d4d) return undefined;
  if (dv.getUint16(base + 2, le) !== 42) return undefined;
  const ifd = base + dv.getUint32(base + 4, le);
  const n = Math.min(dv.getUint16(ifd, le), 1000);
  const out = new Map<number, number>();
  for (let k = 0; k < n; k++) {
    const e = ifd + 2 + 12 * k;
    const tag = dv.getUint16(e, le);
    const type = dv.getUint16(e + 2, le);
    if (type === 3) out.set(tag, dv.getUint16(e + 8, le));
    else if (type === 4) out.set(tag, dv.getUint32(e + 8, le));
    else if (type === 5) {
      const off = base + dv.getUint32(e + 8, le);
      const den = dv.getUint32(off + 4, le);
      if (den !== 0) out.set(tag, dv.getUint32(off, le) / den);
    }
  }
  return out;
}

/** XResolution/YResolution (tags 282/283) in dpi per ResolutionUnit (296: 2 inch, 3 cm). */
function tiffDpi(tags: Map<number, number>, tag: number): number | undefined {
  const r = tags.get(tag);
  const unit = tags.get(296) ?? 2;
  if (r === undefined || r <= 0) return undefined;
  return unit === 3 ? r * 2.54 : unit === 2 ? r : undefined;
}

/** JPEG: SOFn size; JFIF density, else Exif resolution. */
function jpeg(b: Uint8Array, dv: DataView): ImageInfo | undefined {
  let jfif: [number, number] | undefined;
  let exif: [number | undefined, number | undefined] | undefined;
  let i = 2;
  while (i + 4 <= b.length) {
    if (b[i] !== 0xff) return undefined;
    const marker = b[i + 1]!;
    if (marker === 0xff) {
      i++;
      continue;
    }
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      i += 2;
      continue;
    }
    if (marker === 0xd9 || marker === 0xda) return undefined; // no frame header before the scan
    const len = dv.getUint16(i + 2);
    const d = i + 4;
    if (marker === 0xe0 && ascii(b, d, 5) === 'JFIF\0') {
      const units = b[d + 7];
      const x = dv.getUint16(d + 8);
      const y = dv.getUint16(d + 10);
      if (units === 1) jfif = [x, y];
      else if (units === 2) jfif = [x * 2.54, y * 2.54];
    } else if (marker === 0xe1 && ascii(b, d, 6) === 'Exif\0\0') {
      const tags = tiffTags(dv, d + 6);
      if (tags) exif = [tiffDpi(tags, 282), tiffDpi(tags, 283)];
    } else if (
      marker >= 0xc0 &&
      marker <= 0xcf &&
      marker !== 0xc4 &&
      marker !== 0xc8 &&
      marker !== 0xcc
    ) {
      const height = dv.getUint16(d + 1);
      const width = dv.getUint16(d + 3);
      const [dx, dy] = jfif ?? exif ?? [undefined, undefined];
      return info('image/jpeg', width, height, dpi(dx), dpi(dy ?? dx));
    }
    i += 2 + len;
  }
  return undefined;
}

/** BMP: BITMAPCOREHEADER or BITMAPINFOHEADER (and later) size and pixels per metre. */
function bmp(dv: DataView): ImageInfo | undefined {
  const header = dv.getUint32(14, true);
  if (header === 12) return info('image/bmp', dv.getUint16(18, true), dv.getUint16(20, true));
  if (header < 40) return undefined;
  const width = dv.getInt32(18, true);
  const height = Math.abs(dv.getInt32(22, true)); // negative = top-down rows
  const px = dv.getInt32(38, true);
  const py = dv.getInt32(42, true);
  return info(
    'image/bmp',
    width,
    height,
    px > 0 ? dpi(px * 0.0254) : DEFAULT_DPI,
    py > 0 ? dpi(py * 0.0254) : DEFAULT_DPI,
  );
}

function tiff(dv: DataView): ImageInfo | undefined {
  const tags = tiffTags(dv, 0);
  if (tags === undefined) return undefined;
  const x = tiffDpi(tags, 282);
  return info(
    'image/tiff',
    tags.get(256) ?? 0,
    tags.get(257) ?? 0,
    dpi(x),
    dpi(tiffDpi(tags, 283) ?? x),
  );
}

/** WebP: lossy (VP8), lossless (VP8L) or extended (VP8X) header. */
function webp(b: Uint8Array, dv: DataView): ImageInfo | undefined {
  const kind = ascii(b, 12, 4);
  const d = 20;
  if (kind === 'VP8 ') {
    if (b[d + 3] !== 0x9d || b[d + 4] !== 0x01 || b[d + 5] !== 0x2a) return undefined;
    return info(
      'image/webp',
      dv.getUint16(d + 6, true) & 0x3fff,
      dv.getUint16(d + 8, true) & 0x3fff,
    );
  }
  if (kind === 'VP8L') {
    if (b[d] !== 0x2f) return undefined;
    const bits = dv.getUint32(d + 1, true);
    return info('image/webp', (bits & 0x3fff) + 1, ((bits >>> 14) & 0x3fff) + 1);
  }
  if (kind === 'VP8X') {
    const u24 = (at: number): number => b[at]! | (b[at + 1]! << 8) | (b[at + 2]! << 16);
    return info('image/webp', u24(d + 4) + 1, u24(d + 7) + 1);
  }
  return undefined;
}

/** EMF: the header's frame rectangle (0.01 mm) as CSS pixels. */
function emf(dv: DataView): ImageInfo | undefined {
  const w = dv.getInt32(32, true) - dv.getInt32(24, true);
  const h = dv.getInt32(36, true) - dv.getInt32(28, true);
  return info('image/x-emf', (w * 96) / 2540, (h * 96) / 2540);
}

/** Placeable WMF: the bounding box in its own units per inch, as CSS pixels. */
function wmf(dv: DataView): ImageInfo | undefined {
  const w = dv.getInt16(10, true) - dv.getInt16(6, true);
  const h = dv.getInt16(12, true) - dv.getInt16(8, true);
  const perInch = dv.getUint16(14, true);
  if (perInch === 0) return undefined;
  return info('image/x-wmf', (Math.abs(w) * 96) / perInch, (Math.abs(h) * 96) / perInch);
}

/** CSS length → CSS pixels (undefined for percentages and junk). */
function cssLength(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const m = /^\s*\+?(\d*\.?\d+(?:e[+-]?\d+)?)\s*(px|pt|pc|in|cm|mm|em|ex|%)?\s*$/i.exec(v);
  if (m === null) return undefined;
  const n = Number(m[1]);
  switch ((m[2] ?? 'px').toLowerCase()) {
    case 'px':
      return n;
    case 'pt':
      return (n * 96) / 72;
    case 'pc':
      return n * 16;
    case 'in':
      return n * 96;
    case 'cm':
      return (n * 96) / 2.54;
    case 'mm':
      return (n * 96) / 25.4;
    case 'em':
      return n * 16;
    case 'ex':
      return n * 8;
    default:
      return undefined;
  }
}

/** SVG: the root's width/height, completed (or replaced, for percentages) from its viewBox. */
function svg(b: Uint8Array): ImageInfo | undefined {
  const text = new TextDecoder().decode(b.subarray(0, Math.min(b.length, 1 << 20)));
  if (!/^\s*</.test(text.replace(/^﻿/, ''))) return undefined;
  const tag = /<(?:[A-Za-z_][\w.-]*:)?svg(\s[^>]*)?>/.exec(text);
  if (tag === null) return undefined;
  const attrs = tag[1] ?? '';
  const attr = (name: string): string | undefined =>
    new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(attrs)?.slice(1).join('');
  let width = cssLength(attr('width'));
  let height = cssLength(attr('height'));
  const vb = attr('viewBox')
    ?.trim()
    .split(/[\s,]+/)
    .map(Number);
  const vw = vb?.[2];
  const vh = vb?.[3];
  if (vb?.length === 4 && vw !== undefined && vh !== undefined && vw > 0 && vh > 0) {
    if (width === undefined && height === undefined) {
      width = vw;
      height = vh;
    } else if (width === undefined) width = ((height ?? 0) * vw) / vh;
    else if (height === undefined) height = (width * vh) / vw;
  }
  return info('image/svg+xml', width ?? 300, height ?? 150);
}

/**
 * Detects format and pixel size from the bytes; undefined when unrecognized (or truncated, or a
 * WMF without the placeable header, which has no size).
 */
export function imageInfo(bytes: Uint8Array): ImageInfo | undefined {
  const b = bytes;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  try {
    if (b.length >= 24 && dv.getUint32(0) === 0x89504e47 && dv.getUint32(4) === 0x0d0a1a0a) {
      return png(b, dv);
    }
    if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return jpeg(b, dv);
    const sig = ascii(b, 0, 6);
    if (sig === 'GIF87a' || sig === 'GIF89a') {
      return info('image/gif', dv.getUint16(6, true), dv.getUint16(8, true));
    }
    if (sig.startsWith('BM')) return bmp(dv);
    if (sig.startsWith('II*\0') || sig.startsWith('MM\0*')) return tiff(dv);
    if (sig.startsWith('RIFF') && ascii(b, 8, 4) === 'WEBP') return webp(b, dv);
    if (b.length >= 88 && dv.getUint32(0, true) === 1 && dv.getUint32(40, true) === 0x464d4520) {
      return emf(dv);
    }
    if (b.length >= 22 && dv.getUint32(0, true) === 0x9ac6cdd7) return wmf(dv);
    return svg(b);
  } catch (e) {
    if (e instanceof RangeError) return undefined; // truncated header
    throw e;
  }
}

// ---- ids ----------------------------------------------------------------------------------

/** `w:` elements whose `w:id` is a revision id. */
const REVISION: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'moveFromRangeStart',
  'moveToRangeStart',
  'rPrChange',
  'pPrChange',
  'sectPrChange',
  'tblPrChange',
  'tblGridChange',
  'trPrChange',
  'tcPrChange',
  'numberingChange',
  'cellIns',
  'cellDel',
  'cellMerge',
]);

/** @internal A fresh document-wide numeric id of `kind`, reserved (see DocumentIds). */
export function freshId(doc: WordDocument, kind: NumericIdKind): number {
  return DocumentIds.of(doc).next(kind);
}

/**
 * @internal Gives copied markup fresh document-unique ids: drawings (`wp:docPr/@id`), content
 * controls (`w:sdtPr/w:id`) and revisions (`w:ins`, `w:del`, `*Change`, ... `@w:id`). With
 * `deep` false only `root` itself (and its own `w:sdtPr`) is renewed: its content was moved,
 * not copied.
 */
export function renewIds(doc: WordDocument, root: XElement, deep: boolean): void {
  const ids = DocumentIds.of(doc);
  const els = deep
    ? [root, ...root.descendants()]
    : [root, ...(root.element(NS.w, 'sdtPr')?.elements(NS.w, 'id') ?? [])];
  for (const el of els) {
    if (el.is(NS.wp, 'docPr')) {
      if (el.attr('id') !== undefined) el.setAttr('id', String(ids.next('docPr')));
    } else if (el.uri === NS.w) {
      if (el.local === 'id') {
        const parent = el.parent;
        if (
          parent?.nodeType === 1 &&
          parent.is(NS.w, 'sdtPr') &&
          el.attrNS(NS.w, 'val') !== undefined
        ) {
          setWAttr(el, 'val', String(ids.next('sdt')));
        }
      } else if (REVISION.has(el.local) && el.attrNS(NS.w, 'id') !== undefined) {
        setWAttr(el, 'id', String(ids.next('revision')));
      }
    }
  }
}

// ---- parts --------------------------------------------------------------------------------

const STRICT_REL_BASE = 'http://purl.oclc.org/ooxml/officeDocument/relationships/';

/** @internal `type` (a transitional relationship type) in the flavour the package uses. */
export function relTypeFor(doc: WordDocument, type: string): string {
  const main = doc.package.rels.all().find((r) => relTypeIs(r.type, 'officeDocument'));
  if (main?.type.startsWith(STRICT_REL_BASE) !== true) return type;
  return STRICT_REL_BASE + type.slice(type.lastIndexOf('/') + 1);
}

/** Usual extension per image content type (Word's own spellings). */
const EXTENSIONS: Readonly<Record<string, string>> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/tiff': 'tiff',
  'image/svg+xml': 'svg',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
  'image/webp': 'webp',
};

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/**
 * An image part holding exactly `bytes` among the document's own parts (under the main part's
 * directory: never `docProps/thumbnail.jpeg`), if any.
 */
function findSameImage(doc: WordDocument, bytes: Uint8Array): Part | undefined {
  let crc: number | undefined;
  const dir = partKey(partDir(doc.mainPart.name));
  for (const part of doc.package.parts()) {
    if (part.contentType?.startsWith('image/') !== true) continue;
    if (!partKey(part.name).startsWith(dir)) continue;
    const entry = part.dirty ? undefined : part._entry;
    if (entry !== undefined) {
      // A clean source part: compare size and CRC from the zip directory before inflating.
      if (entry.size !== bytes.length) continue;
      crc ??= crc32(bytes);
      if (entry.crc32 !== crc) continue;
    }
    if (sameBytes(part.bytes(), bytes)) return part;
  }
  return undefined;
}

/** `word/media/imageN.<ext>` (next to the main part) with N one past the largest in use. */
function newMediaName(doc: WordDocument, ext: string): string {
  const dir = `${partDir(doc.mainPart.name)}media/`;
  const prefix = partKey(dir);
  let max = 0;
  for (const part of doc.package.parts()) {
    const key = partKey(part.name);
    if (!key.startsWith(prefix)) continue;
    const m = /^image(\d{1,9})\.[^/]*$/.exec(key.slice(prefix.length));
    if (m !== null) max = Math.max(max, Number(m[1]));
  }
  return `${dir}image${max + 1}.${ext}`;
}

/**
 * Adds (or reuses, when identical bytes exist) an image part under word/media/ and a relationship
 * from `owner` (document, header, footer, notes, comments part) to it.
 */
export function addImagePart(
  doc: WordDocument,
  owner: Part,
  bytes: Uint8Array,
  o?: { contentType?: string },
): AddedImage {
  const sniffed = imageInfo(bytes);
  const contentType = o?.contentType ?? sniffed?.contentType;
  if (contentType === undefined) {
    throw new UsageError('addImagePart(): unrecognized image format (pass contentType)');
  }
  const meta: ImageInfo = sniffed ?? {
    contentType,
    width: 0,
    height: 0,
    dpiX: DEFAULT_DPI,
    dpiY: DEFAULT_DPI,
  };
  if (sniffed !== undefined && o?.contentType !== undefined) meta.contentType = o.contentType;

  let part = findSameImage(doc, bytes);
  if (part === undefined) {
    const ct = doc.package.contentTypes;
    let ext: string | undefined;
    for (const [e, t] of ct.defaults()) {
      if (t.toLowerCase() === contentType.toLowerCase() && /^[a-z0-9]+$/.test(e)) {
        ext = e;
        break;
      }
    }
    if (ext === undefined) {
      ext =
        EXTENSIONS[contentType.toLowerCase()] ??
        (contentType
          .replace(/^.*\//, '')
          .replace(/^x-|\+.*$/g, '')
          .replace(/[^a-z0-9]/gi, '')
          .toLowerCase() ||
          'bin');
      if (ct.getDefault(ext) === undefined) ct.setDefault(ext, contentType);
    }
    part = doc.package.addPart(newMediaName(doc, ext), contentType, bytes);
  }

  const key = partKey(part.name);
  const existing = owner.rels
    .byType('image')
    .find((r) => !r.external && r.partName !== undefined && partKey(r.partName) === key);
  const relId =
    existing?.id ??
    owner.rels.add(relTypeFor(doc, REL.image), relativeTarget(owner.name, part.name)).id;
  return { relId, part, info: meta };
}

// ---- markup -------------------------------------------------------------------------------

const STRICT_A = 'http://purl.oclc.org/ooxml/drawingml/main';
const STRICT_PIC = 'http://purl.oclc.org/ooxml/drawingml/picture';
const SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
const SVG_EXT = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';

/** A prefix bound to `uri` in `scope` usable on attributes, else `fallback` (declared on save). */
function prefixFor(scope: XElement, uri: string, fallback: string): string {
  const p = scope.lookupPrefix(uri);
  return p === undefined || p === '' ? fallback : p;
}

type Attr = readonly [qname: string, uri: string, value: string];

function make(name: string, uri: string, attrs: readonly Attr[], ...kids: XElement[]): XElement {
  const flat: string[] = [];
  for (const [q, u, v] of attrs) flat.push(q, u, v);
  const el = new XElement(name, uri, flat);
  if (kids.length > 0) el.append(...kids);
  return el;
}

const plain = (name: string, value: string | number): Attr => [name, '', String(value)];

/**
 * A `w:drawing` with `wp:inline` for an image relationship, sized in EMUs, with a fresh
 * document-unique wp:docPr id. Created in `owner`'s XML namespace scope.
 *
 * `svgRelId` adds Word's SVG extension (`asvg:svgBlip`): `relId` is then the raster fallback
 * that older consumers show (it may be the SVG itself when there is no fallback).
 */
export function inlineDrawing(
  doc: WordDocument,
  owner: Part,
  relId: string,
  widthEmu: number,
  heightEmu: number,
  o?: { name?: string; description?: string; svgRelId?: string },
): XElement {
  const root = owner.xml().root;
  const w = wPrefix(root);
  const strict = isStrictNs(root.lookupNamespace(w) ?? '');
  const aUri = strict ? STRICT_A : NS.a;
  const picUri = strict ? STRICT_PIC : NS.pic;
  const wp = prefixFor(root, NS.wp, 'wp');
  const r = prefixFor(root, NS.r, 'r');
  const q = (p: string, local: string): string => (p === '' ? local : `${p}:${local}`);
  const cx = Math.max(1, Math.round(widthEmu));
  const cy = Math.max(1, Math.round(heightEmu));
  const id = DocumentIds.of(doc).next('docPr');
  const file = owner.related(relId)?.name.replace(/^.*\//, '');
  const name = o?.name ?? `Picture ${id}`;
  const docPrAttrs: Attr[] = [plain('id', id), plain('name', name)];
  if (o?.description !== undefined) docPrAttrs.push(plain('descr', o.description));
  const xmlnsA: Attr = ['xmlns:a', NS.xmlns, aUri];

  const blip = make('a:blip', NS.a, [[`${r}:embed`, NS.r, relId]]);
  if (o?.svgRelId !== undefined) {
    blip.append(
      make(
        'a:extLst',
        NS.a,
        [],
        make(
          'a:ext',
          NS.a,
          [plain('uri', SVG_EXT)],
          make('asvg:svgBlip', SVG_NS, [
            ['xmlns:asvg', NS.xmlns, SVG_NS],
            [`${r}:embed`, NS.r, o.svgRelId],
          ]),
        ),
      ),
    );
  }
  const pic = make(
    'pic:pic',
    NS.pic,
    [['xmlns:pic', NS.xmlns, picUri]],
    make(
      'pic:nvPicPr',
      NS.pic,
      [],
      make('pic:cNvPr', NS.pic, [plain('id', id), plain('name', o?.name ?? file ?? name)]),
      make('pic:cNvPicPr', NS.pic, []),
    ),
    make(
      'pic:blipFill',
      NS.pic,
      [],
      blip,
      make('a:stretch', NS.a, [], make('a:fillRect', NS.a, [])),
    ),
    make(
      'pic:spPr',
      NS.pic,
      [],
      make(
        'a:xfrm',
        NS.a,
        [],
        make('a:off', NS.a, [plain('x', 0), plain('y', 0)]),
        make('a:ext', NS.a, [plain('cx', cx), plain('cy', cy)]),
      ),
      make('a:prstGeom', NS.a, [plain('prst', 'rect')], make('a:avLst', NS.a, [])),
    ),
  );
  const inline = make(
    q(wp, 'inline'),
    NS.wp,
    [plain('distT', 0), plain('distB', 0), plain('distL', 0), plain('distR', 0)],
    make(q(wp, 'extent'), NS.wp, [plain('cx', cx), plain('cy', cy)]),
    make(q(wp, 'effectExtent'), NS.wp, [
      plain('l', 0),
      plain('t', 0),
      plain('r', 0),
      plain('b', 0),
    ]),
    make(q(wp, 'docPr'), NS.wp, docPrAttrs),
    make(
      q(wp, 'cNvGraphicFramePr'),
      NS.wp,
      [],
      make('a:graphicFrameLocks', NS.a, [xmlnsA, plain('noChangeAspect', 1)]),
    ),
    make('a:graphic', NS.a, [xmlnsA], make('a:graphicData', NS.a, [plain('uri', picUri)], pic)),
  );
  return make(q(w, 'drawing'), NS.w, [], inline);
}

// ---- listing ------------------------------------------------------------------------------

/** One reference to an image from a story (see WordDocument.images()). */
export interface ImageRef {
  /** The story whose markup references the image. */
  readonly story: Story;
  /** The referencing element: `a:blip` (DrawingML) or `v:imagedata` (VML). */
  readonly element: XElement;
  /** Relationship id (from the story's part). */
  readonly relId: string;
  /** The image part; undefined for a linked (external) image or a dangling relationship. */
  readonly part: Part | undefined;
  /** Relationship target as written (a URL for linked images). */
  readonly target: string | undefined;
  /** True for a linked image (`r:link` / external relationship). */
  readonly external: boolean;
  /** Displayed size in points, from the drawing extent or the VML shape style, when known. */
  readonly width: number | undefined;
  readonly height: number | undefined;
}

const EMU_PER_POINT = 12700;

/** A CSS length in points (VML `style` values). */
function cssPoints(v: string | undefined): number | undefined {
  const px = cssLength(v);
  return px === undefined ? undefined : (px * 72) / 96;
}

function blipSize(el: XElement): [number | undefined, number | undefined] {
  const frame = el.closest(NS.wp, 'inline') ?? el.closest(NS.wp, 'anchor');
  const ext = frame?.element(NS.wp, 'extent');
  const cx = Number(ext?.attr('cx'));
  const cy = Number(ext?.attr('cy'));
  return [
    Number.isFinite(cx) && ext ? cx / EMU_PER_POINT : undefined,
    Number.isFinite(cy) && ext ? cy / EMU_PER_POINT : undefined,
  ];
}

function vmlSize(el: XElement): [number | undefined, number | undefined] {
  const style = el.parent?.nodeType === 1 ? (el.parent.attr('style') ?? '') : '';
  const prop = (name: string): string | undefined =>
    new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([^;]*)`, 'i').exec(style)?.[1];
  return [cssPoints(prop('width')), cssPoints(prop('height'))];
}

/** True inside an `mc:Fallback` whose AlternateContent also has a Choice (a duplicate). */
function inDuplicateFallback(el: XElement): boolean {
  for (let p = el.parent; p !== null && p.nodeType === 1; p = p.parent) {
    if (p.is(NS.mc, 'Fallback')) {
      const ac = p.parent;
      if (ac !== null && ac.nodeType === 1 && ac.element(NS.mc, 'Choice') !== undefined)
        return true;
    }
  }
  return false;
}

/**
 * @internal Every image reference in `stories`, in document order per story: `a:blip`
 * (`r:embed`, else `r:link`) and VML `v:imagedata` (`r:id`). Fallback branches of
 * `mc:AlternateContent` that duplicate a Choice are skipped.
 */
export function imagesOf(stories: readonly Story[]): ImageRef[] {
  const out: ImageRef[] = [];
  for (const story of stories) {
    for (const el of story.element.descendants()) {
      let relId: string | undefined;
      let size: [number | undefined, number | undefined];
      if (el.is(NS.a, 'blip')) {
        relId = el.attrNS(NS.r, 'embed') ?? el.attrNS(NS.r, 'link');
        if (relId === undefined) continue;
        size = blipSize(el);
      } else if (el.is(NS.v, 'imagedata')) {
        relId = el.attrNS(NS.r, 'id') ?? el.attrNS(NS.r, 'href');
        if (relId === undefined) continue;
        size = vmlSize(el);
      } else continue;
      if (inDuplicateFallback(el)) continue;
      const rel = story.part.rels.get(relId);
      const part = rel?.partName === undefined ? undefined : story.part.package.part(rel.partName);
      out.push({
        story,
        element: el,
        relId,
        part,
        target: rel?.target,
        external: rel?.external === true,
        width: size[0],
        height: size[1],
      });
    }
  }
  return out;
}
