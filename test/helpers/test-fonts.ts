// Test fonts: the OFL oracle set in test/fixtures/fonts/ (Carlito, Caladea, Liberation Sans /
// Serif / Mono, regular + bold) and a tiny sfnt writer for synthetic fonts.
//
// The fonts are stored gzip-compressed (byte-identical after gunzip; see
// test/fixtures/fonts/SOURCES.md) to keep the repository small. testFontResolver() serves them
// to layout as a FontResolver: it answers for the real family names only (Calibri reaches
// Carlito through FONT_SUBSTITUTIONS, exactly as with a user's resolver), returns the bold face
// for bold requests and nothing for italic ones (no italic faces in the set, so layout falls back
// to the regular/bold face with a synthetic slant).
//
// Test-only code: uses node:fs and node:zlib.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';
import { loadFont } from '../../src/layout/fonts/index.js';
import type { Font, FontResolver } from '../../src/layout/types.js';

/** Absolute path of test/fixtures/fonts/. */
export const TEST_FONTS_DIR: string = fileURLToPath(new URL('../fixtures/fonts/', import.meta.url));

/** The oracle font set: family → regular/bold file names (gzip) in TEST_FONTS_DIR. */
export const TEST_FONTS: Readonly<Record<string, { regular: string; bold: string }>> = {
  Carlito: { regular: 'Carlito-Regular.ttf.gz', bold: 'Carlito-Bold.ttf.gz' },
  Caladea: { regular: 'Caladea-Regular.ttf.gz', bold: 'Caladea-Bold.ttf.gz' },
  'Liberation Sans': {
    regular: 'LiberationSans-Regular.ttf.gz',
    bold: 'LiberationSans-Bold.ttf.gz',
  },
  'Liberation Serif': {
    regular: 'LiberationSerif-Regular.ttf.gz',
    bold: 'LiberationSerif-Bold.ttf.gz',
  },
  'Liberation Mono': {
    regular: 'LiberationMono-Regular.ttf.gz',
    bold: 'LiberationMono-Bold.ttf.gz',
  },
};

const bytesCache = new Map<string, Uint8Array>();
const fontCache = new Map<string, Font>();

/** The decompressed TTF bytes of a file in TEST_FONTS_DIR (cached; do not mutate). */
export function testFontBytes(file: string): Uint8Array {
  let b = bytesCache.get(file);
  if (b === undefined) {
    b = new Uint8Array(gunzipSync(readFileSync(join(TEST_FONTS_DIR, file))));
    bytesCache.set(file, b);
  }
  return b;
}

/** A parsed face of the oracle set (one shared Font instance per face). */
export function testFont(family: keyof typeof TEST_FONTS | string, bold = false): Font {
  const entry = TEST_FONTS[family];
  if (entry === undefined) throw new Error(`No test font ${family}`);
  const file = bold ? entry.bold : entry.regular;
  let f = fontCache.get(file);
  if (f === undefined) {
    f = loadFont(testFontBytes(file));
    fontCache.set(file, f);
  }
  return f;
}

const LOWER_FAMILIES = new Map(Object.keys(TEST_FONTS).map((k) => [k.toLowerCase(), k]));

/** A FontResolver over the oracle set (see the file header). */
export function testFontResolver(): FontResolver {
  return (req) => {
    if (req.italic) return undefined;
    const family = LOWER_FAMILIES.get(req.family.trim().toLowerCase());
    return family === undefined ? undefined : testFont(family, req.bold);
  };
}

// ---- synthetic fonts -------------------------------------------------------------------------

/** Big-endian byte writer for building font tables in tests. */
export class ByteWriter {
  readonly #bytes: number[] = [];
  get length(): number {
    return this.#bytes.length;
  }
  u8(v: number): this {
    this.#bytes.push(v & 0xff);
    return this;
  }
  u16(v: number): this {
    return this.u8(v >>> 8).u8(v);
  }
  i16(v: number): this {
    return this.u16(v & 0xffff);
  }
  u32(v: number): this {
    return this.u16(v >>> 16).u16(v & 0xffff);
  }
  tag(t: string): this {
    for (let i = 0; i < 4; i++) this.u8(t.charCodeAt(i));
    return this;
  }
  bytes(b: Uint8Array | readonly number[]): this {
    for (const x of b) this.u8(x);
    return this;
  }
  /** Overwrites a u16 at `at` (for offsets patched after the fact). */
  setU16(at: number, v: number): this {
    this.#bytes[at] = (v >>> 8) & 0xff;
    this.#bytes[at + 1] = v & 0xff;
    return this;
  }
  done(): Uint8Array {
    return Uint8Array.from(this.#bytes);
  }
}

/** Assembles an sfnt from raw tables (sorted by tag, 4-byte aligned). */
export function buildSfnt(
  tables: Readonly<Record<string, Uint8Array>>,
  signature = 0x00010000,
): Uint8Array {
  const tags = Object.keys(tables).sort();
  const w = new ByteWriter();
  w.u32(signature).u16(tags.length).u16(0).u16(0).u16(0);
  let offset = 12 + tags.length * 16;
  for (const tag of tags) {
    const t = tables[tag]!;
    w.tag(tag).u32(0).u32(offset).u32(t.length);
    offset += (t.length + 3) & ~3;
  }
  for (const tag of tags) {
    const t = tables[tag]!;
    w.bytes(t);
    while (w.length % 4 !== 0) w.u8(0);
  }
  return w.done();
}
