/// <reference types="node" />
// Editing content: images (format sniffing, parts, Word's inline drawing), hyperlinks,
// bookmarks, paragraph/run formatting, structural paragraph edits (split, insert, replace,
// remove, clone) and story-level inserts. Every written document is validated and its untouched
// parts are checked to keep their raw zip bytes. Outputs go to out/edit-content/ so they can be
// opened in Word or converted with `node scripts/libreoffice-check.mjs out/edit-content`.
import { mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';
import { beforeAll, describe, expect, it } from 'vitest';
import {
  createDocx,
  NS,
  Paragraph,
  parseXml,
  REL,
  readDocx,
  serializeXml,
  UsageError,
  validatePackage,
  type WordDocument,
  writeDocx,
  type XElement,
} from '../src/index.js';
import { Bookmark } from '../src/wml/bookmarks.js';
import { Hyperlink } from '../src/wml/hyperlinks.js';
import { DocumentIds } from '../src/wml/ids.js';
import { addImagePart, imageInfo, inlineDrawing } from '../src/wml/media.js';
import { crc32, unzipLazy } from '../src/zip/index.js';
import { fixture, fixtures, strictFixtures } from './helpers/fixtures.js';

const OUT = fileURLToPath(new URL('../out/edit-content/', import.meta.url));
const W = NS.w;
const HWI = 'docxcompose/header_with_image.docx';
const TOC2 = 'open-xml-sdk/TOC2.docx';
const CRP = 'python-docx/comments-rich-para.docx';
const EMU = 12700;

beforeAll(() => {
  mkdirSync(`${OUT}fixtures/`, { recursive: true });
});

// ---- image builders ---------------------------------------------------------------------

function concat(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

function bytesOf(n: number, fill: (dv: DataView, b: Uint8Array) => void): Uint8Array {
  const b = new Uint8Array(n);
  fill(new DataView(b.buffer), b);
  return b;
}

const ascii = (s: string): Uint8Array => Uint8Array.from(s, (c) => c.charCodeAt(0));

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const body = concat(ascii(type), data);
  return concat(
    bytesOf(4, (dv) => dv.setUint32(0, data.length)),
    body,
    bytesOf(4, (dv) => dv.setUint32(0, crc32(body))),
  );
}

/** A valid RGB PNG of `w` x `h` pixels in one colour, with a pHYs chunk when `dpi` is given. */
function png(w: number, h: number, o: { dpi?: number; rgb?: [number, number, number] } = {}) {
  const ihdr = bytesOf(13, (dv, b) => {
    dv.setUint32(0, w);
    dv.setUint32(4, h);
    b[8] = 8;
    b[9] = 2;
  });
  const [r, g, bl] = o.rgb ?? [200, 40, 40];
  const row = new Uint8Array(1 + w * 3);
  for (let x = 0; x < w; x++) row.set([r, g, bl], 1 + x * 3);
  const raw = new Uint8Array(row.length * h);
  for (let y = 0; y < h; y++) raw.set(row, y * row.length);
  const chunks = [pngChunk('IHDR', ihdr)];
  if (o.dpi !== undefined) {
    const ppm = Math.round(o.dpi / 0.0254);
    chunks.push(
      pngChunk(
        'pHYs',
        bytesOf(9, (dv, b) => {
          dv.setUint32(0, ppm);
          dv.setUint32(4, ppm);
          b[8] = 1;
        }),
      ),
    );
  }
  chunks.push(
    pngChunk('IDAT', new Uint8Array(deflateSync(raw))),
    pngChunk('IEND', new Uint8Array()),
  );
  return concat(Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), ...chunks);
}

/** A valid 1x1 GIF. */
const GIF = Uint8Array.from(
  atob('R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw=='),
  (c) => c.charCodeAt(0),
);

/** A valid 24-bit BMP (`w` x `h`, bottom-up unless `topDown`). */
function bmp(w: number, h: number, ppm: number, topDown = false): Uint8Array {
  const stride = Math.ceil((w * 3) / 4) * 4;
  return bytesOf(54 + stride * h, (dv, b) => {
    b.set(ascii('BM'));
    dv.setUint32(2, b.length, true);
    dv.setUint32(10, 54, true);
    dv.setUint32(14, 40, true);
    dv.setInt32(18, w, true);
    dv.setInt32(22, topDown ? -h : h, true);
    dv.setUint16(26, 1, true);
    dv.setUint16(28, 24, true);
    dv.setUint32(34, stride * h, true);
    dv.setInt32(38, ppm, true);
    dv.setInt32(42, ppm, true);
  });
}

/** A TIFF header + IFD: size, X/Y resolution and resolution unit. */
function tiff(le: boolean, w: number, h: number, rx: number, ry: number, unit: number) {
  return bytesOf(90, (dv, b) => {
    b.set(ascii(le ? 'II' : 'MM'));
    dv.setUint16(2, 42, le);
    dv.setUint32(4, 8, le);
    dv.setUint16(8, 5, le);
    const entry = (k: number, tag: number, type: number, value: number): void => {
      const e = 10 + 12 * k;
      dv.setUint16(e, tag, le);
      dv.setUint16(e + 2, type, le);
      dv.setUint32(e + 4, 1, le);
      if (type === 3) dv.setUint16(e + 8, value, le);
      else dv.setUint32(e + 8, value, le);
    };
    entry(0, 256, 3, w);
    entry(1, 257, 4, h);
    entry(2, 282, 5, 74);
    entry(3, 283, 5, 82);
    entry(4, 296, 3, unit);
    dv.setUint32(74, rx, le);
    dv.setUint32(78, 1, le);
    dv.setUint32(82, ry, le);
    dv.setUint32(86, 1, le);
  });
}

/** JPEG markers only (enough for sniffing): JFIF or Exif density, then a SOF0 frame header. */
function jpegHeader(w: number, h: number, density: 'jfif' | 'exif' | 'none'): Uint8Array {
  const parts: Uint8Array[] = [Uint8Array.from([0xff, 0xd8])];
  if (density === 'jfif') {
    parts.push(
      bytesOf(18, (dv, b) => {
        b.set([0xff, 0xe0]);
        dv.setUint16(2, 16);
        b.set(ascii('JFIF\0'), 4);
        b.set([1, 1, 1], 9);
        dv.setUint16(12, 300);
        dv.setUint16(14, 150);
      }),
    );
  } else if (density === 'exif') {
    const t = tiff(false, 1, 1, 200, 200, 2);
    parts.push(
      bytesOf(4, (dv, b) => {
        b.set([0xff, 0xe1]);
        dv.setUint16(2, 2 + 6 + t.length);
      }),
      ascii('Exif\0\0'),
      t,
    );
  }
  parts.push(
    bytesOf(19, (dv, b) => {
      b.set([0xff, 0xc0]);
      dv.setUint16(2, 17);
      b[4] = 8;
      dv.setUint16(5, h);
      dv.setUint16(7, w);
      b[9] = 3;
    }),
    Uint8Array.from([0xff, 0xd9]),
  );
  return concat(...parts);
}

function webp(kind: 'VP8X' | 'VP8L' | 'VP8 ', w: number, h: number): Uint8Array {
  return bytesOf(40, (dv, b) => {
    b.set(ascii('RIFF'));
    dv.setUint32(4, 32, true);
    b.set(ascii('WEBP'), 8);
    b.set(ascii(kind), 12);
    dv.setUint32(16, 20, true);
    if (kind === 'VP8X') {
      const put = (at: number, v: number): void => b.set([v & 255, (v >> 8) & 255, v >> 16], at);
      put(24, w - 1);
      put(27, h - 1);
    } else if (kind === 'VP8L') {
      b[20] = 0x2f;
      dv.setUint32(21, (w - 1) | ((h - 1) << 14), true);
    } else {
      b.set([0x9d, 0x01, 0x2a], 23);
      dv.setUint16(26, w, true);
      dv.setUint16(28, h, true);
    }
  });
}

/** An EMF header whose frame is `wmm` x `hmm` hundredths of a millimetre. */
function emf(w100mm: number, h100mm: number): Uint8Array {
  return bytesOf(88, (dv) => {
    dv.setUint32(0, 1, true);
    dv.setUint32(4, 88, true);
    dv.setInt32(32, w100mm, true);
    dv.setInt32(36, h100mm, true);
    dv.setUint32(40, 0x464d4520, true);
  });
}

function wmf(w: number, h: number, perInch: number): Uint8Array {
  return bytesOf(22, (dv) => {
    dv.setUint32(0, 0x9ac6cdd7, true);
    dv.setInt16(10, w, true);
    dv.setInt16(12, h, true);
    dv.setUint16(14, perInch, true);
  });
}

const svg = (attrs: string): Uint8Array =>
  new TextEncoder().encode(
    `<?xml version="1.0"?>\n<!-- logo --><svg xmlns="http://www.w3.org/2000/svg" ${attrs}><rect stroke-width="9" width="5" height="5"/></svg>`,
  );

/** The JPEG of the header_with_image fixture. */
function nasaJpeg(): Uint8Array {
  const doc = readDocx(fixture(HWI).bytes);
  return doc.images()[0]!.part!.bytes();
}

// ---- helpers ----------------------------------------------------------------------------

function errors(doc: WordDocument | Uint8Array): string[] {
  const d = doc instanceof Uint8Array ? readDocx(doc) : doc;
  return validatePackage(d)
    .filter((i) => i.severity === 'error')
    .map((i) => `${i.code} ${i.part ?? ''}: ${i.message}`);
}

/** Saves (validating), writes out/edit-content/<name>, reads back (validated again). */
function save(doc: WordDocument, name: string): WordDocument {
  expect(errors(doc), name).toEqual([]);
  const bytes = writeDocx(doc, { validate: true });
  writeFileSync(`${OUT}${name}`, bytes);
  const back = readDocx(bytes);
  expect(errors(back), name).toEqual([]);
  return back;
}

function xmlOf(doc: WordDocument, name: string): string {
  return serializeXml(doc.package.part(name)!.xml());
}

/** Entries of `output` that differ from `input` in raw payload or CRC (and new ones). */
function changedEntries(input: Uint8Array, output: Uint8Array): string[] {
  const a = unzipLazy(input);
  const b = unzipLazy(output);
  const out: string[] = [];
  for (const [name, e] of b) {
    const old = a.get(name);
    if (old === undefined || old.crc32 !== e.crc32 || !sameBytes(old.raw(), e.raw()))
      out.push(name);
  }
  for (const name of a.keys()) if (!b.has(name)) out.push(`-${name}`);
  return out.sort();
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function kids(el: XElement | undefined): string[] {
  return (el?.elements() ?? []).map((k) => k.local);
}

function pOf(doc: WordDocument, xml: string): Paragraph {
  const el = parseXml(
    `<w:p xmlns:w="${W}" xmlns:r="${NS.r}" xmlns:w14="${NS.w14}">${xml}</w:p>`,
  ).root;
  const body = doc.body.element;
  body.insertBefore(el, body.lastElement() ?? null);
  return new Paragraph(el, doc.body);
}

function extent(run: { element: XElement }): [number, number] {
  const e = run.element.descendants(NS.wp, 'extent')[0]!;
  return [Number(e.attr('cx')), Number(e.attr('cy'))];
}

// ---- imageInfo --------------------------------------------------------------------------

describe('imageInfo', () => {
  it('reads PNG size and pHYs resolution', () => {
    expect(imageInfo(png(40, 20))).toEqual({
      contentType: 'image/png',
      width: 40,
      height: 20,
      dpiX: 96,
      dpiY: 96,
    });
    expect(imageInfo(png(3, 2, { dpi: 300 }))).toMatchObject({ width: 3, dpiX: 300, dpiY: 300 });
    expect(imageInfo(png(3, 2, { dpi: 72 }))).toMatchObject({ dpiX: 72 });
  });

  it('reads JPEG frame size with JFIF or Exif density', () => {
    expect(imageInfo(jpegHeader(640, 480, 'jfif'))).toEqual({
      contentType: 'image/jpeg',
      width: 640,
      height: 480,
      dpiX: 300,
      dpiY: 150,
    });
    expect(imageInfo(jpegHeader(10, 20, 'exif'))).toMatchObject({ width: 10, dpiX: 200 });
    expect(imageInfo(jpegHeader(10, 20, 'none'))).toMatchObject({ height: 20, dpiX: 96 });
    const real = imageInfo(nasaJpeg());
    expect(real?.contentType).toBe('image/jpeg');
    expect(real!.width).toBeGreaterThan(0);
  });

  it('reads GIF, BMP, TIFF and WebP', () => {
    expect(imageInfo(GIF)).toEqual({
      contentType: 'image/gif',
      width: 1,
      height: 1,
      dpiX: 96,
      dpiY: 96,
    });
    expect(imageInfo(bmp(5, 3, 3780))).toMatchObject({
      contentType: 'image/bmp',
      width: 5,
      height: 3,
      dpiX: 96,
    });
    expect(imageInfo(bmp(5, 3, 11811, true))).toMatchObject({ height: 3, dpiX: 300 });
    expect(imageInfo(bmp(5, 3, 0))).toMatchObject({ dpiX: 96 });
    expect(imageInfo(tiff(true, 100, 50, 300, 150, 2))).toEqual({
      contentType: 'image/tiff',
      width: 100,
      height: 50,
      dpiX: 300,
      dpiY: 150,
    });
    expect(imageInfo(tiff(false, 7, 9, 118, 118, 3))).toMatchObject({ width: 7, dpiX: 300 });
    expect(imageInfo(webp('VP8X', 1000, 600))).toMatchObject({
      contentType: 'image/webp',
      width: 1000,
      height: 600,
    });
    expect(imageInfo(webp('VP8L', 33, 44))).toMatchObject({ width: 33, height: 44 });
    expect(imageInfo(webp('VP8 ', 320, 200))).toMatchObject({ width: 320, height: 200 });
  });

  it('reads SVG, EMF and placeable WMF natural sizes in CSS pixels', () => {
    expect(imageInfo(svg('width="2in" height="1in"'))).toMatchObject({
      contentType: 'image/svg+xml',
      width: 192,
      height: 96,
    });
    expect(imageInfo(svg('viewBox="0 0 300 120"'))).toMatchObject({ width: 300, height: 120 });
    expect(imageInfo(svg('width="100%" viewBox="0,0,40,20"'))).toMatchObject({
      width: 40,
      height: 20,
    });
    expect(imageInfo(svg('width="80" viewBox="0 0 40 20"'))).toMatchObject({
      width: 80,
      height: 40,
    });
    expect(imageInfo(svg('width="72pt" height="10mm"'))!.height).toBeCloseTo(37.795, 2);
    expect(imageInfo(svg(''))).toMatchObject({ width: 300, height: 150 });
    expect(imageInfo(emf(5080, 2540))).toMatchObject({
      contentType: 'image/x-emf',
      width: 192,
      height: 96,
    });
    expect(imageInfo(wmf(2880, 1440, 1440))).toMatchObject({
      contentType: 'image/x-wmf',
      width: 192,
      height: 96,
    });
  });

  it('returns undefined for unknown, truncated or unsized data', () => {
    expect(imageInfo(new Uint8Array())).toBeUndefined();
    expect(imageInfo(ascii('hello world'))).toBeUndefined();
    expect(imageInfo(png(4, 4).subarray(0, 12))).toBeUndefined();
    expect(imageInfo(jpegHeader(4, 4, 'jfif').subarray(0, 10))).toBeUndefined();
    expect(imageInfo(bmp(0, 3, 0))).toBeUndefined();
    expect(imageInfo(Uint8Array.from([1, 0, 9, 0, 0, 3]))).toBeUndefined(); // non-placeable WMF
    expect(imageInfo(ascii('<html><body/></html>'))).toBeUndefined();
  });
});

// ---- image parts and drawings -----------------------------------------------------------

describe('addImagePart / inlineDrawing', () => {
  it('adds word/media/imageN.ext with a content-type Default and an image relationship', () => {
    const doc = createDocx();
    const a = addImagePart(doc, doc.mainPart, png(10, 10));
    expect(a.part.name).toBe('word/media/image1.png');
    expect(a.info).toMatchObject({ contentType: 'image/png', width: 10 });
    const ct = doc.package.contentTypes;
    expect(ct.getDefault('png')).toBe('image/png');
    expect(ct.overrides().has('word/media/image1.png')).toBe(false);
    const rel = doc.mainPart.rels.get(a.relId)!;
    expect(rel).toMatchObject({ type: REL.image, target: 'media/image1.png', external: false });
    const b = addImagePart(doc, doc.mainPart, nasaJpeg());
    expect(b.part.name).toBe('word/media/image2.jpeg');
    expect(ct.getDefault('jpeg')).toBe('image/jpeg');
  });

  it('stores identical bytes once and reuses the owner relationship', () => {
    const doc = readDocx(fixture(HWI).bytes);
    const header = doc.headersFooters()[0]!;
    const before = doc.package.parts().length;
    const a = addImagePart(doc, doc.mainPart, png(8, 8));
    const again = addImagePart(doc, doc.mainPart, png(8, 8));
    expect(again.part).toBe(a.part);
    expect(again.relId).toBe(a.relId);
    const fromHeader = addImagePart(doc, header.part, png(8, 8));
    expect(fromHeader.part).toBe(a.part);
    expect(header.part.rels.get(fromHeader.relId)?.target).toBe('media/image2.png');
    const other = addImagePart(doc, doc.mainPart, png(8, 8, { rgb: [0, 0, 255] }));
    expect(other.part.name).toBe('word/media/image3.png');
    // The fixture's JPEG is found without touching it; new JPEGs reuse its `jpg` extension.
    const existing = doc.images()[0]!;
    const same = addImagePart(doc, header.part, existing.part!.bytes());
    expect(same.part).toBe(existing.part);
    expect(same.relId).toBe(existing.relId);
    expect(same.part.dirty).toBe(false);
    expect(addImagePart(doc, doc.mainPart, jpegHeader(2, 2, 'none')).part.name).toBe(
      'word/media/image4.jpg',
    );
    expect(doc.package.parts().length).toBe(before + 3); // 3 media
  });

  it('throws for unrecognized bytes unless a content type is given', () => {
    const doc = createDocx();
    expect(() => addImagePart(doc, doc.mainPart, ascii('nope'))).toThrow(UsageError);
    const a = addImagePart(doc, doc.mainPart, ascii('nope'), { contentType: 'image/x-icon' });
    expect(a.part.name).toBe('word/media/image1.icon');
    expect(doc.package.contentTypes.getDefault('icon')).toBe('image/x-icon');
  });

  it("writes Word's wp:inline markup with declared namespaces and a fresh docPr id", () => {
    const doc = createDocx();
    const { relId } = addImagePart(doc, doc.mainPart, png(4, 2));
    const d1 = inlineDrawing(doc, doc.mainPart, relId, 2 * 914400, 914400, {
      name: 'Logo',
      description: 'Company logo',
    });
    const d2 = inlineDrawing(doc, doc.mainPart, relId, 10, 10);
    expect(d1.name).toBe('w:drawing');
    const inline = d1.element(NS.wp, 'inline')!;
    expect(kids(inline)).toEqual([
      'extent',
      'effectExtent',
      'docPr',
      'cNvGraphicFramePr',
      'graphic',
    ]);
    const docPr = inline.element(NS.wp, 'docPr')!;
    expect(docPr.attr('name')).toBe('Logo');
    expect(docPr.attr('descr')).toBe('Company logo');
    const id2 = d2.descendants(NS.wp, 'docPr')[0]!.attr('id');
    expect(id2).not.toBe(docPr.attr('id'));
    const graphic = inline.element(NS.a, 'graphic')!;
    expect(graphic.attr('xmlns:a')).toBe(NS.a);
    expect(graphic.element(NS.a, 'graphicData')?.attr('uri')).toBe(NS.pic);
    const pic = graphic.descendants(NS.pic, 'pic')[0]!;
    expect(pic.attr('xmlns:pic')).toBe(NS.pic);
    expect(kids(pic)).toEqual(['nvPicPr', 'blipFill', 'spPr']);
    expect(pic.descendants(NS.a, 'blip')[0]!.attrNS(NS.r, 'embed')).toBe(relId);
    expect(kids(pic.element(NS.pic, 'blipFill'))).toEqual(['blip', 'stretch']);
    expect(pic.descendants(NS.a, 'ext')[0]!.attr('cx')).toBe('1828800');
    expect(pic.descendants(NS.a, 'prstGeom')[0]!.attr('prst')).toBe('rect');
    expect(inline.element(NS.wp, 'extent')!.attr('cy')).toBe('914400');
    const r = doc.body.paragraphs()[0]!.appendRun('');
    r.element.append(d1);
    const back = save(doc, 'inline-drawing.docx');
    expect(back.images()).toHaveLength(1);
    expect(back.images()[0]).toMatchObject({ relId, width: 144, height: 72, external: false });
  });

  it('uses Strict relationship types and namespaces in Strict documents', () => {
    const [strict] = strictFixtures([HWI]);
    const doc = readDocx(strict!.bytes);
    const p = doc.body.appendParagraph('Strict picture: ');
    const run = p.appendImage(png(6, 3));
    const rel = doc.mainPart.rels.byType('image').at(-1)!;
    expect(rel.type).toBe('http://purl.oclc.org/ooxml/officeDocument/relationships/image');
    expect(run.element.descendants(NS.a, 'graphic')[0]!.attr('xmlns:a')).toBe(
      'http://purl.oclc.org/ooxml/drawingml/main',
    );
    const link = p.appendHyperlink(' link', 'https://example.com/');
    expect(doc.mainPart.rels.get(link.relId!)?.type).toBe(
      'http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink',
    );
    p.indentLeft = 36;
    p.alignment = 'right';
    expect(p.element.element(W, 'pPr')!.element(W, 'ind')!.attrNS(W, 'start')).toBe('720');
    expect(p.alignment).toBe('end');
    const back = save(doc, 'strict-edits.docx');
    expect(back.body.paragraphs().at(-1)!.text).toBe('Strict picture:  link');
  });
});

// ---- Paragraph.appendImage --------------------------------------------------------------

describe('Paragraph.appendImage', () => {
  it('sizes from pixels and dpi and keeps the aspect ratio', () => {
    const doc = createDocx();
    const p = doc.body.paragraphs()[0]!;
    expect(extent(p.appendImage(png(96, 48)))).toEqual([72 * EMU, 36 * EMU]);
    expect(extent(p.appendImage(png(96, 48, { dpi: 192 })))).toEqual([36 * EMU, 18 * EMU]);
    expect(extent(p.appendImage(png(96, 48), { width: 144 }))).toEqual([144 * EMU, 72 * EMU]);
    expect(extent(p.appendImage(png(96, 48), { height: 72 }))).toEqual([144 * EMU, 72 * EMU]);
    expect(extent(p.appendImage(png(96, 48), { width: 10, height: 30 }))).toEqual([
      10 * EMU,
      30 * EMU,
    ]);
    const r = p.appendImage(png(96, 48), { name: 'Chart', description: 'Q3' });
    expect(r.element.descendants(NS.wp, 'docPr')[0]!.attr('descr')).toBe('Q3');
    expect(r.paragraph).toBe(p);
    const ids = p.element.descendants(NS.wp, 'docPr').map((d) => d.attr('id'));
    expect(new Set(ids).size).toBe(6);
    expect(doc.package.parts().filter((x) => x.name.startsWith('word/media/'))).toHaveLength(2);
    expect(() => p.appendImage(ascii('xx'))).toThrow(UsageError);
    expect(() => p.appendImage(png(2, 2), { width: -1 })).toThrow(UsageError);
    save(doc, 'image-sizes.docx');
  });

  it('works in the body, a header and a table cell; untouched parts keep their bytes', () => {
    const input = fixture(HWI).bytes;
    const doc = readDocx(input);
    const before = doc.images().length;
    // body
    doc.body.appendParagraph('Body picture: ').appendImage(png(120, 60), { width: 90 });
    // header: through addImagePart with the header part as owner, and through appendImage
    const header = doc.headersFooters()[0]!;
    const hp = header.story.paragraphs()[0]!;
    const added = addImagePart(doc, header.part, png(30, 30, { rgb: [20, 160, 20] }));
    const drawing = inlineDrawing(doc, header.part, added.relId, 20 * EMU, 20 * EMU);
    hp.appendRun(' ').element.append(drawing);
    hp.appendImage(GIF, { width: 12 });
    // table cell
    const table = doc.body.appendTable(2, 2);
    table
      .cell(1, 1)!
      .paragraphs()[0]!
      .appendImage(bmp(4, 4, 3780), { width: 24 });
    table.cell(0, 0)!.paragraphs()[0]!.appendRun('cell');

    const back = save(doc, 'images-everywhere.docx');
    const imgs = back.images();
    expect(imgs).toHaveLength(before + 4);
    expect(imgs.every((i) => i.part !== undefined && !i.external)).toBe(true);
    expect(imgs.map((i) => i.story.kind).sort()).toEqual([
      'body',
      'body',
      'header',
      'header',
      'header',
    ]);
    expect(imgs.find((i) => i.part!.name.endsWith('.gif'))).toMatchObject({
      width: 12,
      height: 12,
    });
    const changed = changedEntries(input, writeDocx(back));
    expect(changed).toEqual(
      [
        '[Content_Types].xml',
        'word/_rels/document.xml.rels',
        'word/_rels/header1.xml.rels',
        'word/document.xml',
        'word/header1.xml',
        'word/media/image2.png',
        'word/media/image3.png',
        'word/media/image4.gif',
        'word/media/image5.bmp',
      ].sort(),
    );
  });

  it('references SVG through the svgBlip extension, with an optional raster fallback', () => {
    const doc = createDocx();
    const p = doc.body.paragraphs()[0]!;
    const plain = p.appendImage(svg('width="96" height="48"'));
    expect(extent(plain)).toEqual([72 * EMU, 36 * EMU]);
    const blip = plain.element.descendants(NS.a, 'blip')[0]!;
    const svgBlip = blip.descendants().find((e) => e.local === 'svgBlip')!;
    expect(svgBlip.attrNS(NS.r, 'embed')).toBe(blip.attrNS(NS.r, 'embed'));
    const withFallback = p.appendImage(svg('viewBox="0 0 10 10"'), { fallback: png(10, 10) });
    const b2 = withFallback.element.descendants(NS.a, 'blip')[0]!;
    const s2 = b2.descendants().find((e) => e.local === 'svgBlip')!;
    expect(doc.mainPart.related(b2.attrNS(NS.r, 'embed')!)?.name).toMatch(/\.png$/);
    expect(doc.mainPart.related(s2.attrNS(NS.r, 'embed')!)?.name).toMatch(/\.svg$/);
    expect(doc.package.contentTypes.getDefault('svg')).toBe('image/svg+xml');
    save(doc, 'svg.docx');
  });
});

// ---- hyperlinks -------------------------------------------------------------------------

describe('Paragraph.appendHyperlink and doc.hyperlinks()', () => {
  it('external links get a TargetMode External relationship and the Hyperlink style', () => {
    const doc = createDocx();
    const p = doc.body.paragraphs()[0]!;
    p.appendRun('See ');
    const link = p.appendHyperlink('the site', 'https://example.com/a?b=1&c=2', {
      tooltip: 'Open',
      bold: true,
    });
    expect(link).toBeInstanceOf(Hyperlink);
    const rel = doc.mainPart.rels.get(link.relId!)!;
    expect(rel).toMatchObject({
      type: REL.hyperlink,
      external: true,
      target: 'https://example.com/a?b=1&c=2',
    });
    expect(link.url).toBe('https://example.com/a?b=1&c=2');
    expect(link.text).toBe('the site');
    expect(link.tooltip).toBe('Open');
    const run = link.element.element(W, 'r')!;
    expect(kids(run.element(W, 'rPr'))).toEqual(['rStyle', 'b']);
    expect(run.element(W, 'rPr')!.element(W, 'rStyle')!.attrNS(W, 'val')).toBe('Hyperlink');
    p.appendHyperlink(' again', 'https://example.org/');
    const styles = xmlOf(doc, 'word/styles.xml');
    expect(styles.match(/w:styleId="Hyperlink"/g)).toHaveLength(1);
    expect(styles).toContain(
      '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style>',
    );
    expect(p.text).toBe('See the site again');
    const back = save(doc, 'hyperlinks.docx');
    expect(back.hyperlinks().map((h) => [h.url, h.text])).toEqual([
      ['https://example.com/a?b=1&c=2', 'the site'],
      ['https://example.org/', ' again'],
    ]);
  });

  it('internal links name a bookmark; a localized Hyperlink style is found by name', () => {
    const doc = createDocx();
    const styles = doc.mainPart.relatedByType('styles')[0]!.xml().root;
    styles.append(
      parseXml(
        `<w:style xmlns:w="${W}" w:type="character" w:styleId="Hipervnculo"><w:name w:val="Hyperlink"/></w:style>`,
      ).root,
    );
    const target = doc.body.paragraphs()[0]!;
    target.appendRun('Target');
    doc.bookmarks.add('Intro', target);
    const link = doc.body.appendParagraph().appendHyperlink('Go to intro', { anchor: 'Intro' });
    expect(link.anchor).toBe('Intro');
    expect(link.url).toBeUndefined();
    expect(link.relId).toBeUndefined();
    expect(link.element.attrNS(W, 'history')).toBe('1');
    expect(
      link.element.element(W, 'r')!.element(W, 'rPr')!.element(W, 'rStyle')!.attrNS(W, 'val'),
    ).toBe('Hipervnculo');
    expect(xmlOf(doc, 'word/styles.xml')).not.toContain('w:styleId="Hyperlink"');
    save(doc, 'hyperlink-anchor.docx');
  });

  it('creates a styles part when the document has none', () => {
    const doc = createDocx();
    const rel = doc.mainPart.rels.byType('styles')[0]!;
    doc.mainPart.rels.remove(rel.id);
    doc.package.removePart('word/styles.xml');
    doc.body.paragraphs()[0]!.appendHyperlink('x', 'https://example.com/');
    const part = doc.mainPart.relatedByType('styles')[0]!;
    expect(part.name).toBe('word/styles1.xml');
    expect(serializeXml(part.xml())).toContain('w:styleId="Hyperlink"');
    save(doc, 'hyperlink-new-styles.docx');
  });

  it('lists hyperlinks across stories and retargets them', () => {
    const doc = readDocx(fixture(CRP).bytes);
    const links = doc.hyperlinks();
    expect(links.map((h) => [h.story.kind, h.url, h.text])).toEqual([
      ['comment', 'https://google.com', expect.any(String)],
    ]);
    const link = links[0]!;
    const id = link.relId!;
    link.url = 'https://example.com/new';
    expect(link.relId).toBe(id); // rewritten in place
    expect(link.story.part.rels.get(id)?.target).toBe('https://example.com/new');
    const toc = readDocx(fixture(TOC2).bytes);
    const anchors = toc.hyperlinks();
    expect(anchors.length).toBeGreaterThan(3);
    expect(anchors[0]!.anchor).toMatch(/^_Toc/);
    // From anchor to URL and back.
    anchors[0]!.url = 'https://example.com/#x';
    expect(anchors[0]!.anchor).toBeUndefined();
    expect(anchors[0]!.url).toBe('https://example.com/#x');
    const rid = anchors[0]!.relId!;
    anchors[0]!.anchor = '_Toc167001276';
    expect(anchors[0]!.relId).toBeUndefined();
    expect(toc.mainPart.rels.get(rid)).toBeUndefined(); // unused relationship removed
    save(doc, 'hyperlink-retarget.docx');
    save(toc, 'hyperlink-toc.docx');
  });

  it('gives a link its own relationship when the old one is shared', () => {
    const doc = createDocx();
    const p = doc.body.paragraphs()[0]!;
    const a = p.appendHyperlink('a', 'https://one.example/');
    const b = p.appendHyperlink('b', 'https://two.example/');
    b.element.setAttrNS(NS.r, 'id', a.relId!); // both now share a's relationship
    doc.mainPart.rels.remove(doc.mainPart.rels.byType('hyperlink')[1]!.id);
    b.url = 'https://three.example/';
    expect(b.relId).not.toBe(a.relId);
    expect(a.url).toBe('https://one.example/');
    expect(b.url).toBe('https://three.example/');
    expect(doc.hyperlinks().map((h) => h.url)).toEqual([
      'https://one.example/',
      'https://three.example/',
    ]);
    save(doc, 'hyperlink-shared.docx');
  });
});

// ---- split ------------------------------------------------------------------------------

describe('Paragraph.split', () => {
  it('splits inside a run; both halves keep the paragraph properties', () => {
    const doc = createDocx();
    const p = pOf(
      doc,
      '<w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Hello world</w:t></w:r><w:r><w:t xml:space="preserve"> again</w:t></w:r>',
    );
    const second = p.split(5);
    expect(p.text).toBe('Hello');
    expect(second.text).toBe(' world again');
    expect(second.style).toBe('Heading1');
    expect(second.alignment).toBe('center');
    expect(p.style).toBe('Heading1');
    expect(p.element.nextElement()).toBe(second.element);
    expect(second.runs()[0]!.bold).toBe(true);
    expect(second.runs()[0]!.element.element(W, 't')!.attr('xml:space')).toBe('preserve');
    save(doc, 'split.docx');
  });

  it('splits at run boundaries without leaving empty runs, and at both ends', () => {
    const doc = createDocx();
    const p = pOf(doc, '<w:r><w:t>ab</w:t></w:r><w:r><w:t>cd</w:t></w:r>');
    const q = p.split(2);
    expect([p.runs().length, q.runs().length]).toEqual([1, 1]);
    const end = q.split(2);
    expect(q.text).toBe('cd');
    expect(end.text).toBe('');
    const start = p.split(0);
    expect(p.text).toBe('');
    expect(start.text).toBe('ab');
    expect(() => start.split(3)).toThrow(RangeError);
    expect(() => start.split(-1)).toThrow(RangeError);
    expect(() => start.split(0.5)).toThrow(RangeError);
  });

  it('splits hyperlinks and content controls, renewing the copied ids', () => {
    const doc = createDocx();
    const p = doc.body.paragraphs()[0]!;
    const link = p.appendHyperlink('click here', 'https://example.com/');
    const sdt = parseXml(
      `<w:sdt xmlns:w="${W}"><w:sdtPr><w:id w:val="77"/></w:sdtPr><w:sdtContent><w:r><w:t>xy</w:t></w:r><w:r><w:t>zw</w:t></w:r></w:sdtContent></w:sdt>`,
    ).root;
    p.element.append(sdt);
    const second = p.split(6);
    expect(p.text).toBe('click ');
    expect(second.text).toBe('herexyzw');
    const links = doc.hyperlinks();
    expect(links.map((h) => h.url)).toEqual(['https://example.com/', 'https://example.com/']);
    expect(links[1]!.relId).toBe(link.relId);
    const third = second.split(6);
    expect(third.text).toBe('zw');
    const ids = doc.body.element
      .descendants(W, 'sdtPr')
      .map((s) => s.element(W, 'id')!.attrNS(W, 'val'));
    expect(ids).toHaveLength(2);
    expect(new Set(ids).size).toBe(2);
    save(doc, 'split-inline.docx');
  });

  it('keeps simple fields whole and splits tracked insertions with a fresh revision id', () => {
    const doc = createDocx();
    const p = pOf(
      doc,
      '<w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>12</w:t></w:r></w:fldSimple><w:ins w:id="5" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:t>new text</w:t></w:r></w:ins>',
    );
    const q = p.split(6); // inside the field result
    expect(p.text).toBe('Page ');
    expect(q.text).toBe('12new text');
    const r = q.split(6);
    expect(q.text).toBe('12new ');
    expect(r.text).toBe('text');
    const ins = doc.body.element.descendants(W, 'ins');
    expect(ins.map((i) => i.attrNS(W, 'author'))).toEqual(['A', 'A']);
    expect(ins[0]!.attrNS(W, 'id')).toBe('5');
    expect(ins[1]!.attrNS(W, 'id')).not.toBe('5');
    save(doc, 'split-fields.docx');
  });

  it('moves a section break to the second half and drops w14 ids from the copy', () => {
    const doc = createDocx();
    const p = pOf(
      doc,
      '<w:pPr><w:spacing w:after="0"/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>one two</w:t></w:r>',
    );
    p.element.setAttr('w14:paraId', '1A2B3C4D');
    const q = p.split(3);
    expect(p.element.element(W, 'pPr')!.element(W, 'sectPr')).toBeUndefined();
    expect(q.element.element(W, 'pPr')!.element(W, 'sectPr')).toBeDefined();
    expect(q.spaceAfter).toBe(0);
    expect(p.spaceAfter).toBe(0);
    expect(q.element.attr('w14:paraId')).toBeUndefined();
    expect(p.element.attr('w14:paraId')).toBe('1A2B3C4D');
    expect(doc.sections()).toHaveLength(2);
    save(doc, 'split-section.docx');
  });
});

// ---- insert / replace / remove / clone --------------------------------------------------

describe('Paragraph structure edits', () => {
  it('insertBefore / insertAfter with text or a moved Paragraph', () => {
    const doc = createDocx();
    const mid = doc.body.paragraphs()[0]!;
    mid.appendRun('mid');
    const before = mid.insertBefore('before', { style: 'Heading1' });
    const after = mid.insertAfter('after');
    expect(doc.body.paragraphs().map((p) => p.text)).toEqual(['before', 'mid', 'after']);
    expect(before.style).toBe('Heading1');
    before.insertAfter(after); // moves it
    expect(doc.body.paragraphs().map((p) => p.text)).toEqual(['before', 'after', 'mid']);
    expect(() => mid.insertBefore(mid)).toThrow(UsageError);
    expect(doc.body.element.lastElement()!.local).toBe('sectPr');
    const empty = mid.insertAfter('');
    expect(empty.runs()).toHaveLength(0);
  });

  it('replaceWith keeps a section break', () => {
    const doc = createDocx();
    const p = pOf(
      doc,
      '<w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>old</w:t></w:r>',
    );
    const n = p.replaceWith('new');
    expect(n.text).toBe('new');
    expect(n.element.element(W, 'pPr')!.element(W, 'sectPr')).toBeDefined();
    expect(p.element.parent).toBeNull();
    expect(doc.sections()).toHaveLength(2);
    save(doc, 'replace.docx');
  });

  it('remove() keeps a paragraph in table cells, headers and notes', () => {
    const doc = readDocx(fixture('docx-mailmerge/test_footnote_header_footer.docx').bytes);
    const table = doc.body.appendTable(1, 2);
    const cell = table.cell(0, 0)!;
    const only = cell.paragraphs()[0]!;
    only.appendRun('x');
    only.remove();
    expect(cell.paragraphs()).toHaveLength(1);
    expect(cell.element.lastElement()!.local).toBe('p');
    // a cell that ends with a nested table gets a paragraph after it
    const c2 = table.cell(0, 1)!;
    c2.paragraphs()[0]!.insertBefore('first');
    c2.element.insertBefore(
      parseXml(
        `<w:tbl xmlns:w="${W}"><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
      ).root,
      c2.element.lastElement()!,
    );
    c2.element.lastElement()!.remove();
    expect(c2.element.lastElement()!.local).toBe('tbl');
    new Paragraph(c2.element.elements(W, 'p')[0]!, doc.body).remove();
    expect(c2.element.lastElement()!.local).toBe('p');
    // header
    const hf = doc.headersFooters()[0]!;
    for (const p of hf.story.paragraphs()) {
      if (p.element.closest(W, 'tc') === undefined) p.remove();
    }
    expect(hf.story.element.elements(W, 'p').length).toBeGreaterThanOrEqual(1);
    // footnote
    const note = doc.footnotes()[0]!;
    for (const p of note.paragraphs()) p.remove();
    expect(note.paragraphs()).toHaveLength(1);
    save(doc, 'remove.docx');
  });

  it('remove() moves a section break to the previous paragraph, else drops it', () => {
    const doc = createDocx();
    const first = doc.body.paragraphs()[0]!;
    first.appendRun('first');
    const brk = pOf(
      doc,
      '<w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240"/></w:sectPr></w:pPr><w:r><w:t>end of 1</w:t></w:r>',
    );
    brk.remove();
    expect(
      first.element.element(W, 'pPr')!.element(W, 'sectPr')!.element(W, 'pgSz')!.attrNS(W, 'w'),
    ).toBe('15840');
    expect(doc.sections()).toHaveLength(2);
    first.remove(); // nothing before it: the break is dropped
    expect(doc.sections()).toHaveLength(1);
    expect(doc.body.paragraphs()).toHaveLength(0);
    doc.body.appendParagraph('again');
    save(doc, 'remove-section.docx');
  });

  it('clone() renews ids and leaves out bookmarks and comment anchors', () => {
    const doc = createDocx();
    const p = pOf(
      doc,
      '<w:bookmarkStart w:id="0" w:name="bm"/><w:commentRangeStart w:id="0"/><w:r><w:t>text</w:t></w:r><w:ins w:id="9" w:author="A"><w:r><w:t>!</w:t></w:r></w:ins><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:bookmarkEnd w:id="0"/>',
    );
    p.element.setAttr('w14:paraId', '00000001');
    p.appendImage(png(4, 4));
    const c = p.clone();
    expect(c.element.parent).toBeNull();
    expect(c.text).toBe('text!');
    expect(c.element.descendants(W, 'bookmarkStart')).toHaveLength(0);
    expect(c.element.descendants(W, 'commentReference')).toHaveLength(0);
    expect(c.element.descendants(W, 'commentRangeStart')).toHaveLength(0);
    expect(c.element.attr('w14:paraId')).toBeUndefined();
    expect(c.element.descendants(W, 'ins')[0]!.attrNS(W, 'id')).not.toBe('9');
    const ids = [p, c].map((x) => x.element.descendants(NS.wp, 'docPr')[0]!.attr('id'));
    expect(ids[0]).not.toBe(ids[1]);
    p.insertAfter(c);
    p.element.descendants(W, 'commentRangeStart')[0]!.remove(); // no comments part here
    p.element.descendants(W, 'commentRangeEnd')[0]!.remove();
    (p.element.descendants(W, 'commentReference')[0]!.parent as XElement).remove();
    expect(doc.body.paragraphs().map((x) => x.text)).toEqual(['', 'text!', 'text!']);
    save(doc, 'clone.docx');
  });
});

// ---- formatting -------------------------------------------------------------------------

describe('paragraph and run properties', () => {
  it('paragraph setters write schema order and read back', () => {
    const doc = createDocx();
    const p = doc.body.paragraphs()[0]!;
    p.appendRun('Formatted paragraph');
    p.alignment = 'center';
    p.lineSpacing = 1.5;
    p.firstLineIndent = -18;
    p.indentRight = 18;
    p.indentLeft = 36;
    p.spaceAfter = 6;
    p.spaceBefore = 12;
    p.pageBreakBefore = true;
    p.keepNext = true;
    p.keepLines = false;
    p.style = 'Heading2';
    const pPr = p.element.element(W, 'pPr')!;
    expect(kids(pPr)).toEqual([
      'pStyle',
      'keepNext',
      'keepLines',
      'pageBreakBefore',
      'spacing',
      'ind',
      'jc',
    ]);
    expect(serializeXml(parseXml(`<x xmlns:w="${W}"/>`))).toBeTruthy();
    const ind = pPr.element(W, 'ind')!;
    expect([ind.attrNS(W, 'hanging'), ind.attrNS(W, 'right'), ind.attrNS(W, 'left')]).toEqual([
      '360',
      '360',
      '720',
    ]);
    const sp = pPr.element(W, 'spacing')!;
    expect([
      sp.attrNS(W, 'line'),
      sp.attrNS(W, 'lineRule'),
      sp.attrNS(W, 'after'),
      sp.attrNS(W, 'before'),
    ]).toEqual(['360', 'auto', '120', '240']);
    expect([p.alignment, p.indentLeft, p.indentRight, p.firstLineIndent]).toEqual([
      'center',
      36,
      18,
      -18,
    ]);
    expect([p.spaceBefore, p.spaceAfter, p.lineSpacing]).toEqual([
      12,
      6,
      { value: 1.5, rule: 'multiple' },
    ]);
    expect([p.keepNext, p.keepLines, p.pageBreakBefore]).toEqual([true, false, true]);
    p.firstLineIndent = 9;
    expect(ind.attrNS(W, 'hanging')).toBeUndefined();
    expect(p.firstLineIndent).toBe(9);
    p.lineSpacing = { value: 14, rule: 'exact' };
    expect([sp.attrNS(W, 'line'), sp.attrNS(W, 'lineRule')]).toEqual(['280', 'exact']);
    expect(p.lineSpacing).toEqual({ value: 14, rule: 'exact' });
    expect(() => {
      p.lineSpacing = 0;
    }).toThrow(UsageError);
    save(doc, 'paragraph-props.docx');
    for (const k of [
      'alignment',
      'lineSpacing',
      'firstLineIndent',
      'indentRight',
      'indentLeft',
      'spaceAfter',
      'spaceBefore',
      'pageBreakBefore',
      'keepNext',
      'keepLines',
      'style',
    ] as const) {
      (p as unknown as Record<string, undefined>)[k] = undefined;
    }
    expect(p.element.element(W, 'pPr')).toBeUndefined();
  });

  it('reads universal measures and start/end indents', () => {
    const doc = createDocx();
    const p = pOf(
      doc,
      '<w:pPr><w:spacing w:before="12pt"/><w:ind w:start="1in" w:end="360"/></w:pPr>',
    );
    expect([p.spaceBefore, p.indentLeft, p.indentRight]).toEqual([12, 72, 18]);
    p.indentLeft = 10;
    const ind = p.element.element(W, 'pPr')!.element(W, 'ind')!;
    expect([ind.attrNS(W, 'start'), ind.attrNS(W, 'left')]).toEqual(['200', undefined]);
  });

  it('run setters write schema order and read back', () => {
    const doc = createDocx();
    const p = doc.body.paragraphs()[0]!;
    const r = p.appendRun('styled', {
      underline: 'double',
      highlight: 'yellow',
      size: 10.5,
      color: '#ff8000',
      bold: true,
      font: 'Arial',
      style: 'Strong',
      italic: false,
    });
    expect(kids(r.element.element(W, 'rPr'))).toEqual([
      'rStyle',
      'rFonts',
      'b',
      'i',
      'color',
      'sz',
      'szCs',
      'highlight',
      'u',
    ]);
    expect([r.underline, r.highlight, r.size, r.color, r.font, r.style]).toEqual([
      'double',
      'yellow',
      10.5,
      'FF8000',
      'Arial',
      'Strong',
    ]);
    const rPr = r.element.element(W, 'rPr')!;
    expect(rPr.element(W, 'sz')!.attrNS(W, 'val')).toBe('21');
    expect(rPr.element(W, 'szCs')!.attrNS(W, 'val')).toBe('21');
    rPr.element(W, 'color')!.setAttrNS(W, 'themeColor', 'accent1');
    rPr.element(W, 'rFonts')!.setAttrNS(W, 'asciiTheme', 'minorHAnsi');
    r.color = 'auto';
    r.font = 'Courier New';
    expect(rPr.element(W, 'color')!.attrNS(W, 'themeColor')).toBeUndefined();
    expect(rPr.element(W, 'rFonts')!.attrNS(W, 'asciiTheme')).toBeUndefined();
    expect([r.color, r.font]).toEqual(['auto', 'Courier New']);
    expect(() => {
      r.color = 'red';
    }).toThrow(UsageError);
    expect(() => {
      r.size = 0;
    }).toThrow(UsageError);
    save(doc, 'run-props.docx');
    for (const k of [
      'underline',
      'highlight',
      'size',
      'color',
      'font',
      'style',
      'bold',
      'italic',
    ] as const) {
      (r as unknown as Record<string, undefined>)[k] = undefined;
    }
    expect(r.element.element(W, 'rPr')).toBeUndefined();
  });
});

// ---- stories ----------------------------------------------------------------------------

describe('Story edits', () => {
  it('insertParagraph inserts at an index of blocks()', () => {
    const doc = createDocx();
    doc.body.paragraphs()[0]!.appendRun('b');
    doc.body.insertParagraph(0, 'a', { style: 'Title' });
    doc.body.insertParagraph(2, 'c');
    doc.body.insertParagraph(1, 'a2');
    expect(doc.body.paragraphs().map((p) => p.text)).toEqual(['a', 'a2', 'b', 'c']);
    expect(doc.body.paragraphs()[0]!.style).toBe('Title');
    expect(doc.body.element.lastElement()!.local).toBe('sectPr');
    expect(() => doc.body.insertParagraph(9)).toThrow(RangeError);
    expect(() => doc.body.insertParagraph(-1)).toThrow(RangeError);
    const hf = readDocx(fixture(HWI).bytes).headersFooters()[0]!;
    expect(hf.story.insertParagraph(0, 'top').element).toBe(hf.story.element.firstElement());
  });

  it('appendTable, appendPageBreak', () => {
    const doc = createDocx();
    const t = doc.body.appendTable(2, 3);
    expect(t.rows()).toHaveLength(2);
    expect(t.rows()[0]!.cells()).toHaveLength(3);
    expect(doc.body.element.lastElement()!.local).toBe('sectPr');
    const pb = doc.body.appendPageBreak();
    const br = pb.element.descendants(W, 'br')[0]!;
    expect(br.attrNS(W, 'type')).toBe('page');
    doc.body.appendParagraph('page 2');
    // empty first paragraph, 6 empty cells, the break (LF), then "page 2"
    expect(doc.text()).toBe(`${'\n'.repeat(9)}page 2`);
    save(doc, 'table-pagebreak.docx');
  });

  it('clear() keeps the body sectPr; headers and notes keep one empty paragraph', () => {
    const doc = readDocx(fixture('docx-mailmerge/test_footnote_header_footer.docx').bytes);
    const sectPr = doc.body.element.lastElement()!;
    doc.body.clear();
    expect(doc.body.element.children).toEqual([sectPr]);
    expect(doc.body.blocks()).toHaveLength(0);
    doc.body.appendParagraph('fresh');
    const hf = doc.headersFooters()[0]!;
    hf.story.clear();
    expect(kids(hf.story.element)).toEqual(['p']);
    const note = doc.footnotes()[0]!;
    note.clear();
    expect(kids(note.element)).toEqual(['p']);
    expect(doc.text()).toBe('fresh');
    save(doc, 'clear.docx');
  });
});

// ---- bookmarks --------------------------------------------------------------------------

describe('doc.bookmarks', () => {
  it('lists bookmarks across stories', () => {
    const toc = readDocx(fixture(TOC2).bytes);
    const list = toc.bookmarks.list();
    expect(list.length).toBeGreaterThan(3);
    const first = list[0]!;
    expect(first).toBeInstanceOf(Bookmark);
    expect(first.name).toBe('_Toc167001275');
    expect(first.hidden).toBe(true);
    expect(first.end?.attrNS(W, 'id')).toBe(first.id);
    expect(first.paragraph?.text.length).toBeGreaterThan(0);
    expect(toc.bookmarks.get('_TOC167001275')?.start).toBe(first.start);
    expect([...toc.bookmarks].length).toBe(list.length);
    const mm = readDocx(fixture('docx-mailmerge/test_footnote_header_footer.docx').bytes);
    const kinds = new Set(mm.bookmarks.list().map((b) => b.story.kind));
    expect(kinds.size).toBeGreaterThan(0);
  });

  it('adds bookmarks around a paragraph or a range, with unique ids', () => {
    const doc = createDocx();
    const [a] = doc.body.paragraphs();
    a!.appendRun('alpha');
    const b = doc.body.appendParagraph('beta', { style: 'Heading1' });
    const c = doc.body.appendParagraph('gamma');
    const one = doc.bookmarks.add('Alpha', b);
    expect(one.start.previousElement()?.local).toBe('pPr');
    expect(b.element.lastElement()?.local).toBe('bookmarkEnd');
    const range = doc.bookmarks.add('_Range', { start: a!, end: c });
    expect(range.start.parent).toBe(a!.element);
    expect(range.end?.parent).toBe(c.element);
    expect(one.id).not.toBe(range.id);
    expect(doc.bookmarks.list().map((x) => x.name)).toEqual(['_Range', 'Alpha']);
    const hf = readDocx(fixture(HWI).bytes);
    const head = hf.headersFooters()[0]!.story.paragraphs()[0]!;
    expect(hf.bookmarks.add('InHeader', head).story.kind).toBe('header');
    expect(DocumentIds.of(doc).has('bookmark', Number(range.id))).toBe(true);
    save(doc, 'bookmarks.docx');
  });

  it('rejects invalid names, duplicates and bad ranges', () => {
    const doc = createDocx();
    const p = doc.body.paragraphs()[0]!;
    const q = doc.body.appendParagraph('q');
    doc.bookmarks.add('ok', p);
    expect(() => doc.bookmarks.add('OK', q)).toThrow(/already exists/);
    expect(() => doc.bookmarks.add('has space', q)).toThrow(UsageError);
    expect(() => doc.bookmarks.add('', q)).toThrow(UsageError);
    expect(() => doc.bookmarks.add('x'.repeat(41), q)).toThrow(/40/);
    expect(doc.bookmarks.add('x'.repeat(40), q).name).toHaveLength(40);
    expect(() => doc.bookmarks.add('back', { start: q, end: p })).toThrow(/before/);
    const other = readDocx(fixture(HWI).bytes);
    const hp = other.headersFooters()[0]!.story.paragraphs()[0]!;
    expect(() =>
      other.bookmarks.add('mixed', { start: other.body.paragraphs()[0]!, end: hp }),
    ).toThrow(/same story/);
  });

  it('removes bookmarks', () => {
    const toc = readDocx(fixture(TOC2).bytes);
    const n = toc.bookmarks.list().length;
    const id = toc.bookmarks.get('_Toc167001277')!.id;
    expect(toc.bookmarks.remove('_Toc167001277')).toBe(true);
    expect(toc.bookmarks.remove('_Toc167001277')).toBe(false);
    expect(toc.bookmarks.list()).toHaveLength(n - 1);
    const ends = toc.body.element
      .descendants(W, 'bookmarkEnd')
      .filter((e) => e.attrNS(W, 'id') === id);
    expect(ends).toHaveLength(0);
  });

  it('renames bookmarks and the hyperlinks and fields that point at them', () => {
    const input = fixture(TOC2).bytes;
    const toc = readDocx(input);
    toc.bookmarks.rename('_Toc167001276', 'Chapter2');
    expect(toc.bookmarks.get('_Toc167001276')).toBeUndefined();
    expect(toc.bookmarks.get('chapter2')).toBeDefined();
    const xml = xmlOf(toc, 'word/document.xml');
    expect(xml).not.toContain('_Toc167001276');
    expect(xml).toContain('w:anchor="Chapter2"');
    expect(xml).toContain('PAGEREF Chapter2 \\h');
    const doc = createDocx();
    const p = doc.body.paragraphs()[0]!;
    doc.bookmarks.add('Old', p);
    const field = pOf(
      doc,
      '<w:fldSimple w:instr=" REF Old \\h "><w:r><w:t>x</w:t></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "Old" </w:instrText></w:r><w:r><w:instrText> REF Older </w:instrText></w:r>',
    );
    doc.bookmarks.rename('Old', 'New');
    expect(field.element.element(W, 'fldSimple')!.attrNS(W, 'instr')).toBe(' REF New \\h ');
    const instr = field.element.descendants(W, 'instrText').map((e) => e.text);
    expect(instr).toEqual([' HYPERLINK \\l "New" ', ' REF Older ']);
    doc.bookmarks.rename('New', 'NEW'); // case-only change of the same bookmark
    expect(() => doc.bookmarks.rename('missing', 'x')).toThrow(UsageError);
    doc.bookmarks.add('Taken', field);
    expect(() => doc.bookmarks.rename('NEW', 'taken')).toThrow(/already exists/);
    const noRefs = createDocx();
    noRefs.bookmarks.add('A', noRefs.body.paragraphs()[0]!);
    pOf(noRefs, '<w:fldSimple w:instr=" REF A "/>');
    noRefs.bookmarks.rename('A', 'B', { updateReferences: false });
    expect(xmlOf(noRefs, 'word/document.xml')).toContain('w:instr=" REF A "');
    const back = save(toc, 'bookmark-rename.docx');
    expect(changedEntries(input, writeDocx(back))).toEqual(['word/document.xml']);
  });
});

// ---- document-level image API -----------------------------------------------------------

describe('doc.addImage / doc.images()', () => {
  it('addImage adds a part related from the main document', () => {
    const doc = createDocx();
    const a = doc.addImage(GIF);
    expect(doc.mainPart.related(a.relId)).toBe(a.part);
    expect(doc.images()).toHaveLength(0); // not shown until referenced
  });

  it('lists images in headers, comments and VML with part and size', () => {
    const hwi = readDocx(fixture(HWI).bytes).images();
    expect(hwi).toHaveLength(1);
    expect(hwi[0]).toMatchObject({ external: false, target: expect.stringMatching(/media\//) });
    expect(hwi[0]!.story.kind).toBe('header');
    expect(hwi[0]!.width).toBeCloseTo(607101 / EMU, 3);
    const crp = readDocx(fixture(CRP).bytes).images();
    expect(crp.map((i) => [i.story.kind, i.width, i.height])).toEqual([['comment', 204, 204]]);
    const doc = createDocx();
    doc.body
      .paragraphs()[0]!
      .element.append(
        parseXml(
          `<w:r xmlns:w="${W}" xmlns:v="${NS.v}" xmlns:r="${NS.r}"><w:pict><v:shape style="width:1in;height:36pt"><v:imagedata r:id="rIdX"/></v:shape></w:pict></w:r>`,
        ).root,
      );
    expect(doc.images()).toEqual([
      expect.objectContaining({ relId: 'rIdX', part: undefined, width: 72, height: 36 }),
    ]);
    for (const f of fixtures()) {
      const d = readDocx(f.bytes);
      for (const img of d.images()) {
        if (!img.external) expect(img.part, `${f.name} ${img.relId}`).toBeDefined();
      }
    }
  });
});

// ---- corpus: edits keep packages valid and untouched parts byte-identical -----------------

describe('edits across the fixture corpus', () => {
  const corpus = [...fixtures(), ...strictFixtures([HWI, CRP])];
  it.each(corpus.map((f) => [f.name, f] as const))('%s', (name, f) => {
    const before = new Set(errors(f.bytes));
    const doc = readDocx(f.bytes);
    const p = doc.body.appendParagraph('Edited: ');
    p.appendHyperlink('turboword', 'https://example.com/turboword');
    p.appendImage(png(64, 32, { dpi: 192 }));
    p.split(3);
    doc.bookmarks.add('turboword_edit', p);
    const cell = doc.body.tables()[0]?.cell(0, 0);
    cell?.paragraphs()[0]?.appendImage(GIF, { width: 9 });
    const out = writeDocx(doc);
    const added = errors(out).filter((e) => !before.has(e));
    expect(added).toEqual([]);
    writeFileSync(`${OUT}fixtures/${name.replace(/[/\\]/g, '__')}`, out);
    const allowed = new Set([
      '[Content_Types].xml',
      'word/document.xml',
      doc.mainPart.name,
      `${doc.mainPart.name.replace(/[^/]*$/, '')}_rels/${doc.mainPart.name.replace(/^.*\//, '')}.rels`,
    ]);
    for (const c of changedEntries(f.bytes, out)) {
      if (allowed.has(c) || /\/media\/image\d+\.(png|gif)$/.test(c) || /styles\d*\.xml$/.test(c))
        continue;
      throw new Error(`${name}: unexpected change to ${c}`);
    }
  });
});
