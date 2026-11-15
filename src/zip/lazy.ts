// Lazy synchronous unzip: parses the central directory only; an entry is inflated when its
// `bytes()` is called, and `raw()` hands out the stored payload untouched so the writer can copy
// clean entries without recompressing them. Ported from turbosheet (src/zip/lazy.ts).
//
// - Every value comes from the central directory. Entries written with a data descriptor
//   (flag bit 3) may have zeros in their local header; only its name/extra lengths are read.
// - Limits (see limits.ts) are enforced on the ACTUAL inflated output of every access: native
//   zlib is capped with `maxOutputLength`; the fflate fallback is fed 16 KiB slices and checked
//   after each one, so a bomb never materializes in full.
// - ZIP64 archives are read (the writer does not produce them).
// - Names use UTF-8 when flag bit 11 says so, and also when they decode as valid UTF-8 without
//   it (macOS and many tools omit the flag); otherwise latin1.

import type { Limits } from '../core/limits.js';
import { Budget, RATIO_MIN_BYTES, ZipFormatError, ZipLimitError, ZipPathError } from './limits.js';
import { zlibInflateSync } from './node.js';
import { inflateStream } from './stream.js';
import type { LazyEntry } from './types.js';

/** Max compressed bytes fed to the fflate inflater at once; bounds the memory burst of a bomb. */
const SLICE = 16 * 1024;
/** Largest output asked of native zlib in one piece (above: fflate streaming). */
const NATIVE_MAX = 2 ** 31 - 1;
const EMPTY = /* @__PURE__ */ new Uint8Array(0);

const SIG_EOCD = 0x06054b50;
const SIG_EOCD64 = 0x06064b50;
const SIG_LOCATOR64 = 0x07064b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_LOCAL = 0x04034b50;
const MAX32 = 0xffffffff;

function malformed(what: string): never {
  throw new ZipFormatError(`Malformed zip archive: ${what}`);
}

const utf8 = /* @__PURE__ */ new TextDecoder();
const utf8Strict = /* @__PURE__ */ new TextDecoder('utf-8', { fatal: true });

function decodeName(b: Uint8Array, flags: number): string {
  if (flags & 0x800) return utf8.decode(b);
  try {
    return utf8Strict.decode(b);
  } catch {
    let s = '';
    for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]!);
    return s;
  }
}

/** Locates the end-of-central-directory record; returns [entries, cdOffset, cdSize]. */
function findDirectory(data: Uint8Array, dv: DataView): [number, number, number] {
  const n = data.length;
  const stop = Math.max(0, n - 22 - 0xffff);
  for (let i = n - 22; i >= stop; i--) {
    if (dv.getUint32(i, true) !== SIG_EOCD) continue;
    let count = dv.getUint16(i + 10, true);
    let size = dv.getUint32(i + 12, true);
    let off = dv.getUint32(i + 16, true);
    if (count === 0xffff || size === MAX32 || off === MAX32) {
      // ZIP64: locator (20 bytes) right before the EOCD
      const loc = i - 20;
      if (loc >= 0 && dv.getUint32(loc, true) === SIG_LOCATOR64) {
        const e64 = Number(dv.getBigUint64(loc + 8, true));
        if (e64 + 56 > n || dv.getUint32(e64, true) !== SIG_EOCD64) malformed('bad ZIP64 record');
        count = Number(dv.getBigUint64(e64 + 32, true));
        size = Number(dv.getBigUint64(e64 + 40, true));
        off = Number(dv.getBigUint64(e64 + 48, true));
      }
    }
    if (off + size > n) malformed('central directory out of bounds');
    return [count, off, size];
  }
  return malformed('end of central directory not found (not a zip file?)');
}

/** One central directory record. */
interface DirEntry {
  readonly name: string;
  readonly flags: number;
  readonly method: number;
  readonly time: number;
  readonly date: number;
  readonly crc: number;
  readonly csize: number;
  readonly usize: number;
  /** Offset of the local file header. */
  readonly lho: number;
}

/**
 * Parses `count` central directory records from `cd` (a buffer starting at the directory).
 * Validates entry count, names, duplicates and declared sizes; directory entries are skipped.
 */
function parseDirectory(cd: Uint8Array, count: number, budget: Budget): DirEntry[] {
  if (count > budget.limits.maxEntries) {
    throw new ZipLimitError(`Too many zip entries (max ${budget.limits.maxEntries})`);
  }
  const dv = new DataView(cd.buffer, cd.byteOffset, cd.byteLength);
  const out: DirEntry[] = [];
  const seen = new Set<string>();
  let p = 0;
  for (let k = 0; k < count; k++) {
    if (p + 46 > cd.length || dv.getUint32(p, true) !== SIG_CENTRAL) {
      malformed('bad central directory entry');
    }
    const flags = dv.getUint16(p + 8, true);
    const method = dv.getUint16(p + 10, true);
    const time = dv.getUint16(p + 12, true);
    const date = dv.getUint16(p + 14, true);
    const crc = dv.getUint32(p + 16, true);
    let csize = dv.getUint32(p + 20, true);
    let usize = dv.getUint32(p + 24, true);
    const nl = dv.getUint16(p + 28, true);
    const xl = dv.getUint16(p + 30, true);
    const cl = dv.getUint16(p + 32, true);
    let lho = dv.getUint32(p + 42, true);
    if (p + 46 + nl + xl + cl > cd.length) malformed('central directory entry out of bounds');
    const name = decodeName(cd.subarray(p + 46, p + 46 + nl), flags);
    if (usize === MAX32 || csize === MAX32 || lho === MAX32) {
      // ZIP64 extended information extra field: only the fields set to 0xffffffff, in order
      let x = p + 46 + nl;
      const xe = x + xl;
      while (x + 4 <= xe) {
        const id = dv.getUint16(x, true);
        const len = dv.getUint16(x + 2, true);
        if (x + 4 + len > xe) break;
        if (id === 1) {
          let q = x + 4;
          const end = q + len;
          const next = (): number => {
            if (q + 8 > end) malformed(`bad ZIP64 extra field for ${name}`);
            const v = Number(dv.getBigUint64(q, true));
            q += 8;
            return v;
          };
          if (usize === MAX32) usize = next();
          if (csize === MAX32) csize = next();
          if (lho === MAX32) lho = next();
          break;
        }
        x += 4 + len;
      }
    }
    p += 46 + nl + xl + cl;
    if (!budget.enter(name, usize)) continue;
    if (seen.has(name)) throw new ZipPathError(`Duplicate zip entry: ${name}`);
    seen.add(name);
    out.push({ name, flags, method, time, date, crc, csize, usize, lho });
  }
  return out;
}

/**
 * Indexes a zip archive without decompressing anything; Map order = central directory order.
 * Entry count, names (see `assertSafePath`), duplicates and declared sizes are validated up
 * front; byte and ratio limits while inflating. Directory entries are skipped. Encrypted entries
 * and methods other than stored/deflate throw on access. Throws `ZipLimitError` for anything
 * that is not a readable zip (including empty input).
 */
export function unzipLazy(data: Uint8Array, limits?: Partial<Limits>): Map<string, LazyEntry> {
  const budget = new Budget(limits);
  const out = new Map<string, LazyEntry>();
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const [count, cdOff] = findDirectory(data, dv);
  for (const e of parseDirectory(data.subarray(cdOff), count, budget)) {
    out.set(e.name, new Entry(data, dv, budget, e));
  }
  return out;
}

class Entry implements LazyEntry {
  readonly path: string;
  readonly size: number;
  readonly compressedSize: number;
  readonly compression: number;
  readonly crc32: number;
  readonly flags: number;
  readonly time: number;
  readonly date: number;
  private readonly lho: number;
  /** Output bytes of this entry already counted against the total (no double counting). */
  private seen = 0;
  /** Payload offset, resolved from the local header on first access. */
  private start = -1;

  constructor(
    private readonly data: Uint8Array,
    private readonly dv: DataView,
    private readonly budget: Budget,
    e: DirEntry,
  ) {
    this.path = e.name;
    this.size = e.usize;
    this.compressedSize = e.csize;
    this.compression = e.method;
    this.crc32 = e.crc;
    this.flags = e.flags;
    this.time = e.time;
    this.date = e.date;
    this.lho = e.lho;
  }

  raw(): Uint8Array {
    if (this.flags & 1) throw new ZipLimitError(`Encrypted zip entry not supported: ${this.path}`);
    if (this.compression !== 0 && this.compression !== 8) {
      throw new ZipLimitError(
        `Unsupported zip compression method ${this.compression}: ${this.path}`,
      );
    }
    if (this.start < 0) {
      const h = this.lho;
      if (h + 30 > this.data.length || this.dv.getUint32(h, true) !== SIG_LOCAL) {
        malformed(`bad local header for ${this.path}`);
      }
      this.start = h + 30 + this.dv.getUint16(h + 26, true) + this.dv.getUint16(h + 28, true);
    }
    const end = this.start + this.compressedSize;
    if (end > this.data.length) malformed(`truncated data for ${this.path}`);
    return this.data.subarray(this.start, end);
  }

  bytes(): Uint8Array {
    const raw = this.raw();
    if (this.compression === 0) {
      this.account(raw.length, raw.length, false);
      return raw;
    }
    return this.native(raw) ?? this.inflate(raw);
  }

  private account(out: number, compressed: number, ratio: boolean): void {
    const fresh = out > this.seen ? out - this.seen : 0;
    if (fresh) this.seen = out;
    this.budget.output(this.path, out, fresh, compressed, ratio);
  }

  /**
   * Inflates the whole entry with native zlib; undefined without `node:zlib`, for corrupt data
   * (the fflate path then reports it) or past NATIVE_MAX. The limits are enforced through
   * zlib's `maxOutputLength`, so a bomb never allocates more than one byte past a limit.
   */
  private native(raw: Uint8Array): Uint8Array | undefined {
    const L = this.budget.limits;
    const ratioMax = Math.max(RATIO_MIN_BYTES, raw.length * L.maxCompressionRatio);
    const limit = Math.min(L.maxEntryBytes, this.budget.remaining() + this.seen, ratioMax);
    const max = Math.min(limit, NATIVE_MAX);
    let out: Uint8Array | 'over' | undefined;
    try {
      out = zlibInflateSync(raw, max + 1);
    } catch (e) {
      if ((e as { code?: string }).code?.startsWith('Z_')) return undefined;
      throw e;
    }
    if (out === 'over') {
      if (max < limit) return undefined;
      this.account(max + 1, raw.length, true); // throws the matching ZipLimitError
      return undefined;
    }
    if (out !== undefined) this.account(out.length, raw.length, true);
    return out;
  }

  /** Portable path: fflate streaming inflate, limits checked after every slice. */
  private inflate(raw: Uint8Array): Uint8Array {
    const parts: Uint8Array[] = [];
    let out = 0;
    let fed = 0;
    let done = false;
    const inf = inflateStream((d, final) => {
      if (final) done = true;
      if (d.length === 0) return;
      out += d.length;
      this.account(out, fed, true);
      parts.push(d);
    });
    const len = raw.length;
    try {
      if (len === 0) inf.push(EMPTY, true);
      for (let i = 0; i < len; i += SLICE) {
        fed = Math.min(i + SLICE, len);
        inf.push(raw.subarray(i, fed), fed === len);
      }
    } catch (e) {
      if (e instanceof ZipLimitError) throw e;
      malformed(`corrupt deflate data in ${this.path} (${(e as Error).message})`);
    }
    if (!done) malformed(`truncated data for ${this.path}`);
    if (parts.length === 1) return parts[0]!;
    const buf = new Uint8Array(out);
    let o = 0;
    for (const c of parts) {
      buf.set(c, o);
      o += c.length;
    }
    return buf;
  }
}
