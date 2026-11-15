// Safety limits for reading untrusted archives: entry count, unsafe entry names, per-entry and
// total decompressed bytes, and compression ratio. Bytes and ratio are always measured on the
// ACTUAL decompressed output, never trusted from headers alone, so lying headers cannot bypass
// them. Ported from turbosheet (src/zip/limits.ts, same author, MIT).

import { TurbowordError } from '../core/errors.js';
import type { Limits } from '../core/limits.js';
import { resolveLimits } from '../core/limits.js';

/**
 * Thrown when a configured limit (entries, bytes, ratio) is exceeded, the archive is malformed,
 * an entry cannot be read (encrypted, unsupported method) or an archive cannot be written
 * without ZIP64. `code` is `'ZIP_LIMIT'`.
 */
export class ZipLimitError extends TurbowordError {
  constructor(message: string) {
    super('ZIP_LIMIT', message);
    this.name = 'ZipLimitError';
  }
}

/**
 * Thrown for unsafe or duplicate entry names (path traversal, absolute paths, ...). It is also a
 * `ZipLimitError`, so one `catch` covers every zip failure. `code` is `'ZIP_PATH'`.
 */
export class ZipPathError extends ZipLimitError {
  override readonly code: string = 'ZIP_PATH';

  constructor(message: string) {
    super(message);
    this.name = 'ZipPathError';
  }
}

/**
 * Thrown when the input is not a readable zip archive (no central directory, bad headers,
 * truncated or corrupt data). Also a `ZipLimitError`. `code` is `'ZIP_MALFORMED'`.
 */
export class ZipFormatError extends ZipLimitError {
  override readonly code: string = 'ZIP_MALFORMED';

  constructor(message: string) {
    super(message);
    this.name = 'ZipFormatError';
  }
}

/** Ratio checks only start once an entry has produced this many bytes. */
export const RATIO_MIN_BYTES: number = 1024 * 1024;

/**
 * Throws `ZipPathError` for an entry name that could escape an extraction directory or is not a
 * valid zip path: empty, NUL, backslash, leading `/`, a drive letter (`C:`), or a `..` segment.
 */
export function assertSafePath(path: string): void {
  if (
    path.length === 0 ||
    path.includes('\0') ||
    path.includes('\\') ||
    path.startsWith('/') ||
    /^[A-Za-z]:/.test(path) ||
    path.split('/').includes('..')
  ) {
    throw new ZipPathError(`Unsafe zip entry path: ${JSON.stringify(path)}`);
  }
}

/** Limit accounting for one archive. */
export class Budget {
  readonly limits: Limits;
  private entries = 0;
  /** Decompressed bytes produced so far (all entries). */
  private total = 0;

  constructor(limits?: Partial<Limits>) {
    this.limits = resolveLimits(limits);
  }

  /**
   * Validates an entry header: count, name, declared size. Returns false for directory entries
   * (nothing to extract).
   */
  enter(name: string, declared: number): boolean {
    const L = this.limits;
    if (++this.entries > L.maxEntries) {
      throw new ZipLimitError(`Too many zip entries (max ${L.maxEntries})`);
    }
    assertSafePath(name);
    if (name.endsWith('/')) return false;
    if (declared > L.maxEntryBytes) throw new ZipLimitError(`Zip entry too large: ${name}`);
    return true;
  }

  /** Decompressed bytes still allowed by `maxTotalBytes`. */
  remaining(): number {
    return this.limits.maxTotalBytes - this.total;
  }

  /**
   * Accounts output of entry `name`: it has produced `out` bytes in total from `compressed` input
   * bytes, `fresh` of which were not accounted before. `ratio` = false for stored entries.
   */
  output(name: string, out: number, fresh: number, compressed: number, ratio: boolean): void {
    const L = this.limits;
    this.total += fresh;
    if (out > L.maxEntryBytes) {
      throw new ZipLimitError(`Zip entry too large: ${name} (max ${L.maxEntryBytes} bytes)`);
    }
    if (this.total > L.maxTotalBytes) {
      throw new ZipLimitError(`Zip total uncompressed size exceeds ${L.maxTotalBytes} bytes`);
    }
    if (ratio && out > RATIO_MIN_BYTES && out / Math.max(compressed, 1) > L.maxCompressionRatio) {
      throw new ZipLimitError(
        `Zip entry ${name} exceeds max compression ratio ${L.maxCompressionRatio}`,
      );
    }
  }
}
