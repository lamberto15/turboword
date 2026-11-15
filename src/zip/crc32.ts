// CRC-32 (IEEE 802.3, reflected polynomial 0xEDB88320), as used by zip. Native `zlib.crc32`
// (Node >= 22.2, several GB/s) when the runtime has it; otherwise a byte-at-a-time table
// implementation. The table is built on first use so importing this module costs nothing.

import { nodeZlib } from './node.js';

let table: Int32Array | undefined;

function crcTable(): Int32Array {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
}

/**
 * Portable table-driven CRC-32 of `data`, continuing from `seed` (a previous result; 0 to
 * start). The fallback used by `crc32` when `zlib.crc32` is unavailable.
 */
export function crc32Table(data: Uint8Array, seed = 0): number {
  table ??= crcTable();
  const t = table;
  let c = ~seed;
  for (let i = 0; i < data.length; i++) c = t[(c ^ data[i]!) & 0xff]! ^ (c >>> 8);
  return ~c >>> 0;
}

/**
 * CRC-32 (IEEE) of `data`, continuing from `seed` (a previous result; 0 to start), so
 * `crc32(b, crc32(a))` equals the CRC of `a` followed by `b`. Native `zlib.crc32` when
 * available. Always returns an unsigned 32-bit integer.
 */
export function crc32(data: Uint8Array, seed = 0): number {
  const f = nodeZlib()?.crc32;
  return typeof f === 'function' ? f(data, seed >>> 0) : crc32Table(data, seed);
}
