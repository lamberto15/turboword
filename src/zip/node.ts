// Node / Bun / Deno `node:zlib`, looked up at runtime with `process.getBuiltinModule` (Node >= 22.3,
// Bun, Deno 2). This is the ONE place in src/ that touches a Node built-in: there is no static
// `node:*` import, so browser bundles stay free of Node code, and every caller falls back to
// fflate (deflate/inflate) or a table CRC when the lookup fails. Ported from turbosheet
// (src/zip/node.ts), trimmed to the synchronous calls v0.1 needs.

/** The subset of `node:zlib` used here. */
export interface NodeZlib {
  deflateRawSync(data: Uint8Array, opts: { level: number }): Uint8Array;
  inflateRawSync(data: Uint8Array, opts?: { maxOutputLength?: number }): Uint8Array;
  /** Node >= 22.2. */
  crc32?: (data: Uint8Array, value?: number) => number;
}

let zlib: NodeZlib | null | undefined;

/** `node:zlib`, or undefined outside Node-compatible runtimes. Looked up once. */
export function nodeZlib(): NodeZlib | undefined {
  if (zlib === undefined) {
    zlib = null;
    try {
      const p = (globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } })
        .process;
      const z = p?.getBuiltinModule?.('node:zlib') as NodeZlib | undefined;
      if (typeof z?.inflateRawSync === 'function' && typeof z.deflateRawSync === 'function') {
        zlib = z;
      }
    } catch {
      // not available
    }
  }
  return zlib ?? undefined;
}

/** Plain Uint8Array view of a Buffer (same memory, no copy). */
const view = (b: Uint8Array): Uint8Array => new Uint8Array(b.buffer, b.byteOffset, b.byteLength);

/** Raw DEFLATE of `data` with `zlib.deflateRawSync`; undefined without `node:zlib`. */
export function zlibDeflateSync(data: Uint8Array, level: number): Uint8Array | undefined {
  const z = nodeZlib();
  return z ? view(z.deflateRawSync(data, { level })) : undefined;
}

/**
 * Inflates a whole raw-deflate buffer with `zlib.inflateRawSync`, producing at most `max` bytes.
 * Returns undefined without `node:zlib`, and 'over' when the output would exceed `max`.
 * Throws the zlib error for corrupt or truncated data.
 */
export function zlibInflateSync(data: Uint8Array, max: number): Uint8Array | 'over' | undefined {
  const z = nodeZlib();
  if (!z) return undefined;
  try {
    return view(z.inflateRawSync(data, { maxOutputLength: Math.max(1, max) }));
  } catch (e) {
    if ((e as { code?: string }).code === 'ERR_BUFFER_TOO_LARGE') return 'over';
    throw e;
  }
}
