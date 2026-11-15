// Public shapes of the zip layer: entries being read (LazyEntry), entries ready to be written
// (RawEntry) and compression options. Kept apart from the implementations so the reader and the
// writer can share them without importing each other.

/** One entry of an archive being read (central directory data + lazy access). */
export interface LazyEntry {
  /** Entry name (zip path, no leading slash). */
  readonly path: string;
  /** Uncompressed size as declared by the central directory (not trusted for limits). */
  readonly size: number;
  readonly compressedSize: number;
  /** 0 = stored, 8 = deflate. */
  readonly compression: number;
  /** CRC-32 from the central directory. */
  readonly crc32: number;
  /** General purpose bit flags (bit 0 = encrypted, bit 3 = data descriptor, bit 11 = UTF-8). */
  readonly flags: number;
  /** DOS last-modified time and date fields, as stored. */
  readonly time: number;
  readonly date: number;
  /**
   * The whole decompressed entry. Inflates on every call (cache it if needed); stored entries
   * return a view into the archive (no copy). Entry, total and ratio limits are enforced on the
   * actual output. Throws `ZipLimitError` for encrypted entries, methods other than 0/8, a
   * truncated archive, corrupt data or an exceeded limit.
   */
  bytes(): Uint8Array;
  /**
   * The entry's compressed payload exactly as stored (a view into the archive; no copy). Throws
   * `ZipLimitError` for encrypted entries, methods other than 0/8 or a truncated archive.
   */
  raw(): Uint8Array;
}

/** An entry ready to be written: already compressed (or stored). */
export interface RawEntry {
  readonly name: string;
  /** 0 = stored, 8 = deflate. */
  readonly method: number;
  readonly crc32: number;
  /** Uncompressed size. */
  readonly size: number;
  readonly compressed: Uint8Array;
  /** DOS time/date; default = the fixed zip date (1980-01-01 00:00). */
  readonly time?: number;
  readonly date?: number;
}

export interface CompressOptions {
  /** Deflate level 0-9 (0 = store). Default 6. */
  level?: number;
  /**
   * 'native' (node:zlib when available, else fflate) or 'fflate'. Default 'native'. The two
   * produce different (equally valid) deflate streams; pick 'fflate' for output that is
   * byte-identical across runtimes.
   */
  deflate?: 'native' | 'fflate';
  /** Timestamp written for this entry (UTC fields). Default 1980-01-01 00:00. */
  date?: Date;
  /** Force stored (no compression), e.g. for already-compressed media (see `shouldStore`). */
  store?: boolean;
}
