// Minimal push-style raw DEFLATE decompressor (a thin fflate wrapper). It is the portable
// fallback of the lazy reader when `node:zlib` is unavailable: input is fed in small slices, so
// limits are checked while inflating and a bomb never materializes in full.
// Ported from turbosheet (src/zip/stream.ts); the deflate side is not needed in v0.1.

import { Inflate } from 'fflate';

export interface ByteStream {
  /** Pushes input; `final = true` on the last push flushes and ends the stream. */
  push(data: Uint8Array, final?: boolean): void;
}

export type ByteSink = (data: Uint8Array, final: boolean) => void;

/** Raw DEFLATE decompressor. Output chunks are fresh arrays (safe to keep). */
export function inflateStream(onData: ByteSink): ByteStream {
  const i = new Inflate(onData);
  return { push: (data, final = false) => i.push(data, final) };
}
