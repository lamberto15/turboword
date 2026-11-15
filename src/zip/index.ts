// The zip layer (L1): a lazy reader that indexes an archive without inflating it and hands out
// stored payloads untouched, and a writer that copies those payloads back raw, so a save only
// pays for compressing the parts that actually changed. Runtime-neutral: native zlib through
// node.ts when available, fflate otherwise. See lazy.ts, writer.ts and limits.ts.

export { crc32 } from './crc32.js';
export { unzipLazy } from './lazy.js';
export { assertSafePath, ZipFormatError, ZipLimitError, ZipPathError } from './limits.js';
export type { CompressOptions, LazyEntry, RawEntry } from './types.js';
export { dosDateTime, makeEntry, rawEntry, shouldStore, writeZip } from './writer.js';
