// Hash algorithms of the agile EncryptionInfo descriptor: the synchronous SHA family from core
// (for the 100,000-round password spin loop) paired with the WebCrypto name (for the HMAC over
// the whole package) and the name written in the XML descriptor.

import { type HashFn, shaByName } from '../core/sha512.js';

export interface Hash {
  fn: HashFn;
  size: number;
  /** WebCrypto name (`SHA-512`). */
  web: string;
  /** Name written in the agile XML descriptor (`SHA512`). */
  xml: string;
}

/** Look up a hash by name; accepts `SHA512`, `SHA-512`, `sha512`, ... */
export function hashByName(name: string): Hash | undefined {
  const h = shaByName(name);
  return h && { fn: h.fn, size: h.size, web: h.name, xml: h.name.replace('-', '') };
}
