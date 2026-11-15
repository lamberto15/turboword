// sfnt helpers for font embedding: the few header fields a PDF FontDescriptor needs (bbox, fixed
// pitch, weight, serif class) and a GID-preserving TrueType subsetter.
//
// The subsetter keeps numGlyphs and every glyph id, empties the outlines of glyphs that are not
// used (directly or as a component of a used composite glyph), rebuilds glyf/loca (short or long
// offsets as needed), keeps hinting (cvt, fpgm, prep) and the metric tables (zeroing the hmtx
// entries of dropped glyphs so they compress away), drops everything else and recomputes table
// checksums and head.checkSumAdjustment. Because ids are preserved, the PDF
// can use CIDToGIDMap /Identity and the glyph ids the layout produced, unchanged.

interface Table {
  offset: number;
  length: number;
}

/** The table directory of a single sfnt face. */
export function tables(bytes: Uint8Array): Map<string, Table> {
  const dv = view(bytes);
  const out = new Map<string, Table>();
  const n = dv.getUint16(4);
  for (let i = 0; i < n; i++) {
    const e = 12 + i * 16;
    if (e + 16 > bytes.length) break;
    const tag = String.fromCharCode(bytes[e]!, bytes[e + 1]!, bytes[e + 2]!, bytes[e + 3]!);
    const offset = dv.getUint32(e + 8);
    const length = dv.getUint32(e + 12);
    if (offset + length <= bytes.length) out.set(tag, { offset, length });
  }
  return out;
}

const view = (b: Uint8Array): DataView => new DataView(b.buffer, b.byteOffset, b.byteLength);

/** Header facts for the FontDescriptor. */
export interface FaceInfo {
  unitsPerEm: number;
  bbox: [number, number, number, number];
  fixedPitch: boolean;
  serif: boolean;
  weight: number;
}

/** Reads head/post/OS/2 facts; tolerant of missing tables. */
export function faceInfo(bytes: Uint8Array): FaceInfo {
  const info: FaceInfo = {
    unitsPerEm: 1000,
    bbox: [0, 0, 0, 0],
    fixedPitch: false,
    serif: false,
    weight: 400,
  };
  if (bytes.length < 12) return info;
  const dv = view(bytes);
  const t = tables(bytes);
  const head = t.get('head');
  if (head && head.length >= 54) {
    info.unitsPerEm = dv.getUint16(head.offset + 18) || 1000;
    const o = head.offset + 36;
    info.bbox = [dv.getInt16(o), dv.getInt16(o + 2), dv.getInt16(o + 4), dv.getInt16(o + 6)];
  }
  const post = t.get('post');
  if (post && post.length >= 16) info.fixedPitch = dv.getUint32(post.offset + 12) !== 0;
  const os2 = t.get('OS/2');
  if (os2 && os2.length >= 32) {
    info.weight = dv.getUint16(os2.offset + 4) || 400;
    const cls = dv.getUint8(os2.offset + 30);
    info.serif = cls >= 1 && cls <= 7;
  }
  return info;
}

/** Tables copied into a subset (besides the rebuilt glyf/loca and the patched head). */
const KEEP = ['cvt ', 'fpgm', 'hhea', 'hmtx', 'maxp', 'prep'];

/**
 * A GID-preserving subset of a TrueType face containing the outlines of `used` glyph ids (plus
 * .notdef and composite components). Returns the input unchanged when it is not a glyf font.
 */
export function subsetTrueType(bytes: Uint8Array, used: Iterable<number>): Uint8Array {
  const t = tables(bytes);
  const head = t.get('head');
  const loca = t.get('loca');
  const glyf = t.get('glyf');
  const maxp = t.get('maxp');
  if (!head || !loca || !glyf || !maxp || head.length < 54) return bytes;
  const dv = view(bytes);
  const numGlyphs = dv.getUint16(maxp.offset + 4);
  const longLoca = dv.getInt16(head.offset + 50) === 1;
  const locaAt = (g: number): number => {
    const o = loca.offset + g * (longLoca ? 4 : 2);
    if (o + (longLoca ? 4 : 2) > loca.offset + loca.length) return 0;
    return longLoca ? dv.getUint32(o) : dv.getUint16(o) * 2;
  };
  const span = (g: number): [number, number] => {
    const s = Math.min(locaAt(g), glyf.length);
    const e = Math.min(Math.max(locaAt(g + 1), s), glyf.length);
    return [glyf.offset + s, glyf.offset + e];
  };

  // Closure over composite components (explicit stack; the keep array guards against cycles).
  const keep = new Uint8Array(numGlyphs);
  const stack: number[] = [0];
  keep[0] = 1;
  for (const g of used) {
    if (g >= 0 && g < numGlyphs && !keep[g]) {
      keep[g] = 1;
      stack.push(g);
    }
  }
  while (stack.length) {
    const [s, e] = span(stack.pop()!);
    if (e - s < 10 || dv.getInt16(s) >= 0) continue;
    let p = s + 10;
    for (;;) {
      if (p + 4 > e) break;
      const flags = dv.getUint16(p);
      const gi = dv.getUint16(p + 2);
      if (gi < numGlyphs && !keep[gi]) {
        keep[gi] = 1;
        stack.push(gi);
      }
      p += 4 + (flags & 1 ? 4 : 2);
      p += flags & 8 ? 2 : flags & 0x40 ? 4 : flags & 0x80 ? 8 : 0;
      if (!(flags & 0x20)) break;
    }
  }

  // Rebuild glyf and loca.
  let size = 0;
  for (let g = 0; g < numGlyphs; g++) {
    if (!keep[g]) continue;
    const [s, e] = span(g);
    size += (e - s + 3) & ~3;
  }
  const newGlyf = new Uint8Array(size);
  const offsets = new Uint32Array(numGlyphs + 1);
  let p = 0;
  for (let g = 0; g < numGlyphs; g++) {
    offsets[g] = p;
    if (!keep[g]) continue;
    const [s, e] = span(g);
    newGlyf.set(bytes.subarray(s, e), p);
    p += (e - s + 3) & ~3;
  }
  offsets[numGlyphs] = p;
  const long = p > 0x1fffe;
  const newLoca = new Uint8Array((numGlyphs + 1) * (long ? 4 : 2));
  const lv = view(newLoca);
  for (let g = 0; g <= numGlyphs; g++) {
    if (long) lv.setUint32(g * 4, offsets[g]!);
    else lv.setUint16(g * 2, offsets[g]! / 2);
  }
  const newHead = bytes.slice(head.offset, head.offset + head.length);
  const hv = view(newHead);
  hv.setUint32(8, 0);
  hv.setInt16(50, long ? 1 : 0);

  const out = new Map<string, Uint8Array>([
    ['glyf', newGlyf],
    ['head', newHead],
    ['loca', newLoca],
  ]);
  for (const tag of KEEP) {
    const tb = t.get(tag);
    if (tb) out.set(tag, bytes.subarray(tb.offset, tb.offset + tb.length));
  }
  // Zero the metrics of dropped glyphs: same layout and glyph count, but the table now compresses
  // to almost nothing (PDF viewers take advances from /W, and those glyphs are empty anyway).
  const hmtx = out.get('hmtx');
  const hhea = t.get('hhea');
  if (hmtx && hhea && hhea.length >= 36) {
    const m = hmtx.slice();
    const nhm = dv.getUint16(hhea.offset + 34);
    for (let g = 0; g < numGlyphs; g++) {
      if (keep[g]) continue;
      const o = g < nhm ? g * 4 : nhm * 4 + (g - nhm) * 2;
      m.fill(0, Math.min(o, m.length), Math.min(o + (g < nhm ? 4 : 2), m.length));
    }
    out.set('hmtx', m);
  }
  return build(dv.getUint32(0), out);
}

/** Assembles an sfnt from tables (sorted by tag) and fixes head.checkSumAdjustment. */
function build(version: number, parts: Map<string, Uint8Array>): Uint8Array {
  const tags = [...parts.keys()].sort();
  const n = tags.length;
  let size = 12 + n * 16;
  for (const tag of tags) size += (parts.get(tag)!.length + 3) & ~3;
  const out = new Uint8Array(size);
  const dv = view(out);
  const log2 = Math.floor(Math.log2(n));
  dv.setUint32(0, version);
  dv.setUint16(4, n);
  dv.setUint16(6, 16 << log2);
  dv.setUint16(8, log2);
  dv.setUint16(10, n * 16 - (16 << log2));
  let p = 12 + n * 16;
  let headAt = -1;
  tags.forEach((tag, i) => {
    const data = parts.get(tag)!;
    out.set(data, p);
    const e = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[e + k] = tag.charCodeAt(k);
    dv.setUint32(e + 4, checksum(out, p, data.length));
    dv.setUint32(e + 8, p);
    dv.setUint32(e + 12, data.length);
    if (tag === 'head') headAt = p;
    p += (data.length + 3) & ~3;
  });
  if (headAt >= 0) dv.setUint32(headAt + 8, (0xb1b0afba - checksum(out, 0, size)) >>> 0);
  return out;
}

/** Sum of big-endian uint32 words, zero-padded to 4 bytes. */
function checksum(b: Uint8Array, off: number, len: number): number {
  let sum = 0;
  const end = off + len;
  for (let i = off; i < end; i += 4) {
    sum =
      (sum +
        (((b[i] ?? 0) << 24) |
          ((i + 1 < end ? b[i + 1]! : 0) << 16) |
          ((i + 2 < end ? b[i + 2]! : 0) << 8) |
          (i + 3 < end ? b[i + 3]! : 0))) >>>
      0;
  }
  return sum;
}
