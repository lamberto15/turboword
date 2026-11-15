// sfnt (TrueType / OpenType-CFF) parser: loadFont.
//
// Reads just what layout and PDF output need: the table directory (plain sfnt or one face of a
// TrueType Collection), head/hhea/maxp/hmtx for advances, OS/2 + post for vertical metrics and
// style, name for the family/PostScript names, cmap for code point → glyph, and GPOS/kern for
// pair kerning (lazily). Required tables that are missing or malformed throw
// TurbowordError('FONT_FORMAT'); optional ones (OS/2, post, name, GPOS, kern) degrade to
// fallbacks. All reads are bounds-checked by Reader, so arbitrary bytes never crash the parser.

import type { Font, FontMetrics } from '../types.js';
import { parseCmap } from './cmap.js';
import { buildKerner, type Kerner } from './kerning.js';
import { decodeNames, type FontNames } from './names.js';
import { fontError, Reader } from './reader.js';

/** Extra facts about a face beyond the Font contract (PDF FontDescriptor, matching). */
export interface FontDetails {
  /** Font bounding box [xMin, yMin, xMax, yMax], font units. */
  readonly bbox: readonly [number, number, number, number];
  /** OS/2 usWeightClass (400 regular, 700 bold). */
  readonly weight: number;
  /** Monospaced (post.isFixedPitch or OS/2 panose proportion). */
  readonly fixedPitch: boolean;
  /** A symbol font (cmap 3,0 or a standard-14 symbolic face): not a text font. */
  readonly symbolic: boolean;
  /** OS/2 fsType embedding permissions (0 = installable). */
  readonly fsType: number;
  /** Typographic family (name ID 16) when it differs from the legacy RIBBI family. */
  readonly typographicFamily: string;
  readonly fullName: string;
  /** Face index within the collection it was loaded from (0 for single fonts). */
  readonly faceIndex: number;
}

const SFNT_TRUETYPE = 0x00010000;
const TAG_TRUE = 0x74727565; // 'true'
const TAG_OTTO = 0x4f54544f; // 'OTTO'
const TAG_TTCF = 0x74746366; // 'ttcf'
const TAG_WOFF = 0x774f4646; // 'wOFF'
const TAG_WOF2 = 0x774f4632; // 'wOF2'

interface TableRecord {
  tag: string;
  checksum: number;
  offset: number;
  length: number;
}

/**
 * Parses a TrueType/OpenType font (or face `index` of a .ttc/.otc collection). Throws
 * TurbowordError('FONT_FORMAT') when the data is not a usable font.
 */
export function loadFont(bytes: Uint8Array, index = 0): Font {
  if (!(bytes instanceof Uint8Array)) fontError('loadFont: expected a Uint8Array');
  if (bytes.length < 12) fontError(`not a font: only ${bytes.length} bytes`);
  const file = new Reader(bytes, 0, bytes.length, 'font');
  const magic = file.u32(0);
  let faceOffset = 0;
  let collection = false;
  if (magic === TAG_TTCF) {
    const numFonts = file.u32(8);
    file.array(12, numFonts, 4);
    if (!(Number.isInteger(index) && index >= 0 && index < numFonts)) {
      fontError(`collection has ${numFonts} faces; face ${index} requested`);
    }
    faceOffset = file.u32(12 + index * 4);
    collection = true;
  } else if (magic === TAG_WOFF || magic === TAG_WOF2) {
    fontError('WOFF/WOFF2 fonts are not supported: decompress them to TTF/OTF first');
  } else if (index !== 0) {
    fontError(`face ${index} requested but the font is not a collection`);
  }
  const tables = readDirectory(file, faceOffset);
  const face = collection ? extractFace(bytes, file, faceOffset, tables) : bytes;
  return new SfntFont(face, tables, collection ? index : 0, collection ? undefined : file);
}

function readDirectory(file: Reader, at: number): Map<string, TableRecord> {
  const version = file.u32(at);
  if (version !== SFNT_TRUETYPE && version !== TAG_OTTO && version !== TAG_TRUE) {
    fontError(`not a TrueType/OpenType font (signature 0x${version.toString(16)})`);
  }
  const numTables = file.u16(at + 4);
  file.array(at + 12, numTables, 16);
  const tables = new Map<string, TableRecord>();
  for (let i = 0; i < numTables; i++) {
    const p = at + 12 + i * 16;
    const tag = file.tag(p);
    const offset = file.u32(p + 8);
    let length = file.u32(p + 12);
    if (offset > file.length) continue; // dangling entry: treat the table as missing
    if (offset + length > file.length) length = file.length - offset; // truncated last table
    if (!tables.has(tag)) tables.set(tag, { tag, checksum: file.u32(p + 4), offset, length });
  }
  return tables;
}

/** Copies one face of a collection into a standalone sfnt (tables 4-byte aligned). */
function extractFace(
  bytes: Uint8Array,
  file: Reader,
  at: number,
  tables: Map<string, TableRecord>,
): Uint8Array {
  const list = [...tables.values()].sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0));
  const headerSize = 12 + list.length * 16;
  let size = headerSize;
  for (const t of list) size += (t.length + 3) & ~3;
  const out = new Uint8Array(size);
  const view = new DataView(out.buffer);
  view.setUint32(0, file.u32(at));
  const n = list.length;
  let pow = 1;
  let log = 0;
  while (pow * 2 <= n) {
    pow *= 2;
    log++;
  }
  view.setUint16(4, n);
  view.setUint16(6, pow * 16);
  view.setUint16(8, log);
  view.setUint16(10, n * 16 - pow * 16);
  let offset = headerSize;
  const relocated = new Map<string, TableRecord>();
  list.forEach((t, i) => {
    const p = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[p + k] = t.tag.charCodeAt(k);
    view.setUint32(p + 4, t.checksum);
    view.setUint32(p + 8, offset);
    view.setUint32(p + 12, t.length);
    out.set(bytes.subarray(t.offset, t.offset + t.length), offset);
    relocated.set(t.tag, { ...t, offset });
    offset += (t.length + 3) & ~3;
  });
  tables.clear();
  for (const [k, v] of relocated) tables.set(k, v);
  return out;
}

const EMPTY_NAMES: FontNames = {
  family: '',
  subfamily: '',
  postscriptName: '',
  typographicFamily: '',
  fullName: '',
};

/** A parsed sfnt face. */
export class SfntFont implements Font, FontDetails {
  readonly family: string;
  readonly subfamily: string;
  readonly postscriptName: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly metrics: FontMetrics;
  readonly outlines: 'truetype' | 'cff';
  readonly bytes: Uint8Array;
  readonly numGlyphs: number;
  readonly bbox: readonly [number, number, number, number];
  readonly weight: number;
  readonly fixedPitch: boolean;
  readonly symbolic: boolean;
  readonly fsType: number;
  readonly typographicFamily: string;
  readonly fullName: string;
  readonly faceIndex: number;
  readonly #advances: Uint16Array;
  readonly #cmap: (cp: number) => number;
  readonly #tables: Map<string, TableRecord>;
  #kerner: Kerner | undefined;

  /** @internal */
  constructor(
    bytes: Uint8Array,
    tables: Map<string, TableRecord>,
    faceIndex: number,
    whole?: Reader,
  ) {
    this.bytes = bytes;
    this.faceIndex = faceIndex;
    this.#tables = tables;
    const file = whole ?? new Reader(bytes, 0, bytes.length, 'font');
    const table = (tag: string): Reader | undefined => {
      const t = tables.get(tag);
      return t === undefined ? undefined : file.sub(t.offset, t.length, tag);
    };
    const need = (tag: string): Reader =>
      table(tag) ?? fontError(`missing required table '${tag}'`);

    if (tables.has('glyf')) this.outlines = 'truetype';
    else if (tables.has('CFF ') || tables.has('CFF2')) this.outlines = 'cff';
    else fontError('no glyph outlines (neither glyf nor CFF)');

    // head
    const head = need('head');
    const unitsPerEm = head.u16(18);
    if (unitsPerEm < 16 || unitsPerEm > 16384) fontError(`head: bad unitsPerEm ${unitsPerEm}`);
    const macStyle = head.u16(44);
    this.bbox = [head.i16(36), head.i16(38), head.i16(40), head.i16(42)];
    const locFormat = head.i16(50);

    // maxp
    const numGlyphs = need('maxp').u16(4);
    if (numGlyphs === 0) fontError('maxp: font has no glyphs');
    this.numGlyphs = numGlyphs;

    // hhea + hmtx
    const hhea = need('hhea');
    const hheaAsc = hhea.i16(4);
    const hheaDesc = hhea.i16(6);
    const hheaGap = hhea.i16(8);
    const hmtx = need('hmtx');
    let numH = Math.min(hhea.u16(34), numGlyphs, hmtx.length >>> 2);
    if (numH === 0) {
      if (hmtx.length < 2) fontError('hmtx: no metrics');
      numH = 1;
    }
    const advances = new Uint16Array(numH);
    for (let i = 0; i < numH; i++) advances[i] = hmtx.u16(i * 4);
    this.#advances = advances;

    // post (optional)
    let italicAngle = 0;
    let underlinePosition = Math.round(-unitsPerEm * 0.1);
    let underlineThickness = Math.round(unitsPerEm * 0.05);
    let fixedPitch = false;
    const post = optional(() => table('post'));
    if (post?.has(0, 16)) {
      italicAngle = post.fixed(4);
      underlinePosition = post.i16(8);
      underlineThickness = post.i16(10) || underlineThickness;
      fixedPitch = post.u32(12) !== 0;
    }

    // OS/2 (optional; very old Apple fonts lack it)
    const os2 = optional(() => table('OS/2'));
    let weight = macStyle & 1 ? 700 : 400;
    let fsSelection = 0;
    let fsType = 0;
    let typo: [number, number, number] | undefined;
    let win: [number, number] | undefined;
    let strikeoutSize = 0;
    let strikeoutPosition = 0;
    let capHeight = 0;
    let xHeight = 0;
    if (os2?.has(0, 78)) {
      const version = os2.u16(0);
      weight = os2.u16(4) || weight;
      fsType = os2.u16(8);
      strikeoutSize = os2.i16(26);
      strikeoutPosition = os2.i16(28);
      if (os2.u8(32 + 3) === 9) fixedPitch = true; // panose bProportion = monospaced
      fsSelection = os2.u16(62);
      typo = [os2.i16(68), os2.i16(70), os2.i16(72)];
      win = [os2.u16(74), os2.u16(76)];
      if (version >= 2 && os2.has(86, 4)) {
        xHeight = os2.i16(86);
        capHeight = os2.i16(88);
      }
    }
    this.weight = weight;
    this.fsType = fsType;
    this.fixedPitch = fixedPitch;
    this.bold = (fsSelection & 0x20) !== 0 || (macStyle & 1) !== 0;
    this.italic = (fsSelection & 0x201) !== 0 || (macStyle & 2) !== 0;

    let ascender: number;
    let descender: number;
    let lineGap: number;
    if (typo !== undefined && (fsSelection & 0x80) !== 0) {
      [ascender, descender, lineGap] = typo;
    } else if (hheaAsc !== 0 || hheaDesc !== 0) {
      [ascender, descender, lineGap] = [hheaAsc, hheaDesc, hheaGap];
    } else if (typo !== undefined && (typo[0] !== 0 || typo[1] !== 0)) {
      [ascender, descender, lineGap] = typo;
    } else if (win !== undefined) {
      [ascender, descender, lineGap] = [win[0], -win[1], 0];
    } else {
      [ascender, descender, lineGap] = [this.bbox[3], this.bbox[1], 0];
    }
    if (win === undefined || (win[0] === 0 && win[1] === 0)) {
      win = [Math.max(ascender, 0), Math.max(-descender, 0)];
    }

    // cmap (required)
    const cmap = parseCmap(need('cmap'), numGlyphs);
    this.#cmap = cmap.lookup;
    this.symbolic = cmap.symbol;

    // Cap/x height fallbacks: the glyph bounds of 'H' and 'x' (TrueType), else typical ratios.
    if (capHeight <= 0) {
      capHeight = optional(() => glyphTop(table, locFormat, this.#cmap(0x48), numGlyphs)) ?? 0;
      if (capHeight <= 0) capHeight = Math.round(unitsPerEm * 0.7);
    }
    if (xHeight <= 0) {
      xHeight = optional(() => glyphTop(table, locFormat, this.#cmap(0x78), numGlyphs)) ?? 0;
      if (xHeight <= 0) xHeight = Math.round(unitsPerEm * 0.5);
    }
    if (strikeoutSize <= 0) strikeoutSize = underlineThickness;
    if (strikeoutPosition <= 0) strikeoutPosition = Math.round(xHeight / 2);

    this.metrics = {
      unitsPerEm,
      ascender,
      descender,
      lineGap,
      winAscent: win[0],
      winDescent: win[1],
      capHeight,
      xHeight,
      underlinePosition,
      underlineThickness,
      strikeoutPosition,
      strikeoutSize,
      italicAngle,
    };

    // name (optional)
    const nameTable = optional(() => table('name'));
    const names =
      (nameTable !== undefined && optional(() => decodeNames(nameTable))) || EMPTY_NAMES;
    this.postscriptName = names.postscriptName;
    this.family = names.family || names.typographicFamily || names.postscriptName;
    this.subfamily =
      names.subfamily ||
      (this.bold ? (this.italic ? 'Bold Italic' : 'Bold') : this.italic ? 'Italic' : 'Regular');
    this.typographicFamily = names.typographicFamily || this.family;
    this.fullName = names.fullName || `${this.family} ${this.subfamily}`.trim();
  }

  glyphId(codePoint: number): number {
    if (!(codePoint >= 0 && codePoint <= 0x10ffff)) return 0;
    return this.#cmap(codePoint);
  }

  advance(glyphId: number): number {
    if (!(glyphId >= 0 && glyphId < this.numGlyphs)) return 0;
    const a = this.#advances;
    const g = Math.floor(glyphId);
    return g < a.length ? a[g]! : a[a.length - 1]!;
  }

  kerning(left: number, right: number): number {
    if (this.#kerner === undefined) {
      const file = new Reader(this.bytes, 0, this.bytes.length, 'font');
      const table = (tag: string): Reader | undefined => {
        const t = this.#tables.get(tag);
        return t === undefined ? undefined : optional(() => file.sub(t.offset, t.length, tag));
      };
      this.#kerner = buildKerner(table('GPOS'), table('kern'));
    }
    return this.#kerner(left, right);
  }

  /** True when the face has an sfnt table `tag` (e.g. 'GPOS', 'glyf'). */
  hasTable(tag: string): boolean {
    return this.#tables.has(tag);
  }

  /** Window [offset, length] of table `tag` within `bytes`, for subsetters. */
  tableRange(tag: string): { offset: number; length: number } | undefined {
    const t = this.#tables.get(tag);
    return t === undefined ? undefined : { offset: t.offset, length: t.length };
  }
}

/** Runs `f`, turning a malformed optional table into undefined. */
function optional<T>(f: () => T): T | undefined {
  try {
    return f();
  } catch {
    return undefined;
  }
}

/** yMax of a TrueType glyph (from its glyf header), 0 when empty or unknown. */
function glyphTop(
  table: (tag: string) => Reader | undefined,
  locFormat: number,
  gid: number,
  numGlyphs: number,
): number {
  if (gid <= 0 || gid >= numGlyphs) return 0;
  const loca = table('loca');
  const glyf = table('glyf');
  if (loca === undefined || glyf === undefined) return 0;
  const start = locFormat === 0 ? loca.u16(gid * 2) * 2 : loca.u32(gid * 4);
  const end = locFormat === 0 ? loca.u16(gid * 2 + 2) * 2 : loca.u32(gid * 4 + 4);
  if (end <= start || !glyf.has(start, 10)) return 0;
  return glyf.i16(start + 8);
}
