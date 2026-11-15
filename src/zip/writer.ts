// Zip writer: entries arrive already compressed (RawEntry), so a clean entry read from the source
// archive is copied with its original payload and CRC (`rawEntry`) and only dirty parts pay for
// deflate (`makeEntry`). Ported from turbomerge (src/zip.ts, same author, MIT) and made
// runtime-neutral: deflate is `node:zlib` when asked for and available, else fflate; CRC-32 is
// native when available, else a table.
//
// Output is deterministic: entries are written in the given order, with a fixed date
// (1980-01-01 00:00) unless one is supplied, clean local headers (no data descriptors, no extra
// fields) and the UTF-8 name flag always set. ZIP64 is not written: an archive that would need
// it (> 65534 entries, or 4 GiB) throws ZipLimitError.

import { deflateSync } from 'fflate';
import { UsageError } from '../core/errors.js';
import { crc32 } from './crc32.js';
import { assertSafePath, ZipLimitError, ZipPathError } from './limits.js';
import { zlibDeflateSync } from './node.js';
import type { CompressOptions, LazyEntry, RawEntry } from './types.js';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
/** Version 2.0 (deflate), made by MS-DOS: what Word writes. */
const VERSION = 20;
const FLAG_UTF8 = 0x800;
/** Largest value of a 32-bit field that does not mean "see ZIP64". */
const MAX32 = 0xfffffffe;

/** DOS time field of the fixed zip date (00:00:00). */
export const FIXED_DOS_TIME: number = 0;
/** DOS date field of the fixed zip date (1980-01-01). */
export const FIXED_DOS_DATE: number = (0 << 9) | (1 << 5) | 1;

const encoder = /* @__PURE__ */ new TextEncoder();

/**
 * Extensions of formats that are already compressed, so deflating them again costs time and
 * saves (almost) nothing: raster images (JPEG XR / HD Photo `wdp`/`jxr`/`hdp` included, which
 * Office uses for image effects), gzip-wrapped metafiles (`emz`/`wmz`/`svgz`), audio/video,
 * compressed fonts and zip-based containers (embedded Office documents).
 *
 * Deliberately NOT here, because deflate shrinks them a lot: `odttf` (an obfuscated but
 * uncompressed TrueType font, typically -40-60%), `ttf`/`otf`, `emf`/`wmf`, `bmp`, `tif(f)`,
 * `svg`, and `bin` (OLE objects, vbaProject).
 */
const STORED_EXTENSIONS: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'png',
  'jpg',
  'jpeg',
  'jpe',
  'jfif',
  'gif',
  'webp',
  'avif',
  'heic',
  'wdp',
  'jxr',
  'hdp',
  'emz',
  'wmz',
  'svgz',
  'mp3',
  'm4a',
  'aac',
  'ogg',
  'oga',
  'opus',
  'wma',
  'mp4',
  'm4v',
  'mov',
  'webm',
  'wmv',
  'avi',
  'mpg',
  'mpeg',
  'woff',
  'woff2',
  'zip',
  'gz',
  '7z',
  'docx',
  'docm',
  'dotx',
  'dotm',
  'xlsx',
  'xlsm',
  'xltx',
  'xltm',
  'pptx',
  'pptm',
  'potx',
  'potm',
  'ppsx',
  'ppsm',
  'odt',
  'ods',
  'odp',
]);

/**
 * True when the entry `name` holds an already-compressed format (by extension, case-insensitive)
 * and should be stored rather than deflated; pass it as `makeEntry`'s `store` option.
 */
export function shouldStore(name: string): boolean {
  const dot = name.lastIndexOf('.');
  if (dot < 0 || dot < name.lastIndexOf('/')) return false;
  return STORED_EXTENSIONS.has(name.slice(dot + 1).toLowerCase());
}

/**
 * DOS date/time fields for `d`, from its UTC components (so output does not depend on the
 * machine's time zone). Clamped to the DOS range 1980-01-01 00:00:00 .. 2107-12-31 23:59:58;
 * seconds have 2-second resolution. Throws UsageError for an invalid Date.
 */
export function dosDateTime(d: Date): { date: number; time: number } {
  if (Number.isNaN(d.getTime())) throw new UsageError('Invalid Date for a zip entry');
  const year = d.getUTCFullYear();
  if (year < 1980) return { date: FIXED_DOS_DATE, time: FIXED_DOS_TIME };
  if (year > 2107) return { date: (127 << 9) | (12 << 5) | 31, time: (23 << 11) | (59 << 5) | 29 };
  return {
    date: ((year - 1980) << 9) | ((d.getUTCMonth() + 1) << 5) | d.getUTCDate(),
    time: (d.getUTCHours() << 11) | (d.getUTCMinutes() << 5) | (d.getUTCSeconds() >> 1),
  };
}

type Level = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

/** Raw DEFLATE with the requested backend ('native' falls back to fflate without node:zlib). */
function deflateRaw(data: Uint8Array, level: number, backend: 'native' | 'fflate'): Uint8Array {
  const native = backend === 'native' ? zlibDeflateSync(data, level) : undefined;
  return native ?? deflateSync(data, { level: level as Level });
}

/**
 * Compresses `data` into a RawEntry named `name`. The entry is stored instead when `store` is
 * set, `level` is 0, or deflate would not make it smaller. A stored entry keeps a reference to
 * `data` (no copy): do not mutate it afterwards. Throws UsageError for a level outside 0-9 or
 * an invalid `date`.
 */
export function makeEntry(name: string, data: Uint8Array, o: CompressOptions = {}): RawEntry {
  const level = o.level ?? 6;
  if (!Number.isInteger(level) || level < 0 || level > 9) {
    throw new UsageError(`Deflate level must be an integer 0-9, got ${level}`);
  }
  const { date, time } =
    o.date === undefined ? { date: FIXED_DOS_DATE, time: FIXED_DOS_TIME } : dosDateTime(o.date);
  const crc = crc32(data);
  const size = data.length;
  if (o.store !== true && level > 0 && size > 0) {
    const packed = deflateRaw(data, level, o.deflate ?? 'native');
    if (packed.length < size) {
      return { name, method: 8, crc32: crc, size, compressed: packed, time, date };
    }
  }
  return { name, method: 0, crc32: crc, size, compressed: data, time, date };
}

/**
 * A RawEntry that copies `e` without recompressing: same payload bytes (a view, no copy), method,
 * CRC, sizes and time/date. Throws ZipLimitError for entries that cannot be copied (encrypted,
 * methods other than 0/8, truncated).
 */
export function rawEntry(e: LazyEntry): RawEntry {
  return {
    name: e.path,
    method: e.compression,
    crc32: e.crc32,
    size: e.size,
    compressed: e.raw(),
    time: e.time,
    date: e.date,
  };
}

/** The 26 bytes shared by local and central headers, from "version needed" to "extra length". */
function common(dv: DataView, at: number, e: RawEntry, nameLength: number): void {
  dv.setUint16(at, VERSION, true);
  dv.setUint16(at + 2, FLAG_UTF8, true);
  dv.setUint16(at + 4, e.method, true);
  dv.setUint16(at + 6, e.time ?? FIXED_DOS_TIME, true);
  dv.setUint16(at + 8, e.date ?? FIXED_DOS_DATE, true);
  dv.setUint32(at + 10, e.crc32, true);
  dv.setUint32(at + 14, e.compressed.length, true);
  dv.setUint32(at + 18, e.size, true);
  dv.setUint16(at + 22, nameLength, true);
  dv.setUint16(at + 24, 0, true);
}

/**
 * Writes a zip archive from `entries`, in order: UTF-8 names, no data descriptors, no extra
 * fields, no comments. Throws ZipPathError for unsafe or duplicate names, and ZipLimitError for
 * an invalid entry (method other than 0/8, stored entry whose size differs from its payload) or
 * when ZIP64 would be needed (more than 65534 entries, or 4 GiB).
 */
export function writeZip(entries: readonly RawEntry[]): Uint8Array {
  const n = entries.length;
  if (n > 0xfffe) throw new ZipLimitError(`Too many zip entries without ZIP64: ${n} (max 65534)`);
  const names: Uint8Array[] = [];
  const seen = new Set<string>();
  let localSize = 0;
  let centralSize = 0;
  for (const e of entries) {
    assertSafePath(e.name);
    if (seen.has(e.name)) throw new ZipPathError(`Duplicate zip entry: ${e.name}`);
    seen.add(e.name);
    if (e.method !== 0 && e.method !== 8) {
      throw new ZipLimitError(`Unsupported zip compression method ${e.method}: ${e.name}`);
    }
    if (e.method === 0 && e.size !== e.compressed.length) {
      throw new ZipLimitError(`Stored zip entry size does not match its data: ${e.name}`);
    }
    if (e.size > MAX32 || e.compressed.length > MAX32) {
      throw new ZipLimitError(`Zip entry needs ZIP64 (4 GiB or more): ${e.name}`);
    }
    const name = encoder.encode(e.name);
    if (name.length > 0xffff) throw new ZipLimitError(`Zip entry name too long: ${e.name}`);
    names.push(name);
    localSize += 30 + name.length + e.compressed.length;
    centralSize += 46 + name.length;
  }
  if (localSize + centralSize > MAX32) {
    throw new ZipLimitError('Zip archive needs ZIP64 (4 GiB or more)');
  }

  const out = new Uint8Array(localSize + centralSize + 22);
  const dv = new DataView(out.buffer);
  const offsets: number[] = [];
  let pos = 0;
  for (let i = 0; i < n; i++) {
    const e = entries[i]!;
    const name = names[i]!;
    offsets.push(pos);
    dv.setUint32(pos, SIG_LOCAL, true);
    common(dv, pos + 4, e, name.length);
    out.set(name, pos + 30);
    out.set(e.compressed, pos + 30 + name.length);
    pos += 30 + name.length + e.compressed.length;
  }
  const centralStart = pos;
  for (let i = 0; i < n; i++) {
    const name = names[i]!;
    dv.setUint32(pos, SIG_CENTRAL, true);
    dv.setUint16(pos + 4, VERSION, true); // version made by
    common(dv, pos + 6, entries[i]!, name.length);
    // comment length, disk number, internal and external attributes stay 0
    dv.setUint32(pos + 42, offsets[i]!, true);
    out.set(name, pos + 46);
    pos += 46 + name.length;
  }
  dv.setUint32(pos, SIG_EOCD, true);
  dv.setUint16(pos + 8, n, true);
  dv.setUint16(pos + 10, n, true);
  dv.setUint32(pos + 12, pos - centralStart, true);
  dv.setUint32(pos + 16, centralStart, true);
  return out;
}
