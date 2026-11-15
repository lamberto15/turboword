// Safety limits applied while reading untrusted input. Zip limits are measured on actual
// decompressed output (see zip/limits.ts); XML limits while parsing (see xml/parse.ts).

export interface Limits {
  /** Max zip entries. Default 10 000. */
  maxEntries: number;
  /** Max decompressed bytes of one entry. Default 1 GiB. */
  maxEntryBytes: number;
  /** Max decompressed bytes of all entries read. Default 2 GiB. */
  maxTotalBytes: number;
  /** Max decompressed/compressed ratio of an entry (checked past 1 MiB of output). Default 1000. */
  maxCompressionRatio: number;
  /** Max element nesting depth of an XML part. Default 1024. */
  maxXmlDepth: number;
}

export const DEFAULT_LIMITS: Readonly<Limits> = Object.freeze({
  maxEntries: 10_000,
  maxEntryBytes: 1024 * 1024 * 1024,
  maxTotalBytes: 2 * 1024 * 1024 * 1024,
  maxCompressionRatio: 1000,
  maxXmlDepth: 1024,
});

/** `DEFAULT_LIMITS` overridden by `limits`. */
export function resolveLimits(limits?: Partial<Limits>): Limits {
  return { ...DEFAULT_LIMITS, ...limits };
}
