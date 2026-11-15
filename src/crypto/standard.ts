// ECMA-376 standard encryption (MS-OFFCRYPTO 2.3.4.5–2.3.4.9), read only: Office 2007-era files
// with a binary EncryptionInfo (version 3.2/4.2), AES-128/192/256-ECB and SHA-1 with 50,000
// iterations. Ported from turbosheet (MIT, same author).

import { sha1 } from '../core/sha512.js';
import { decryptPackage } from './agile.js';
import { aesKey, badPassword, cbcDecrypt, concat, equal, fail, le32, utf16 } from './util.js';

/** AES-ECB decrypt via CBC: CBC with a zero IV, then undo the chaining XOR block by block. */
async function ecbDecrypt(key: CryptoKey, data: Uint8Array): Promise<Uint8Array> {
  const out = await cbcDecrypt(key, new Uint8Array(16), data);
  for (let i = 16; i < out.length; i++) out[i]! ^= data[i - 16]!;
  return out;
}

/** Key derivation (2.3.4.7). */
export function standardKey(salt: Uint8Array, password: string, keyBytes: number): Uint8Array {
  let h = sha1(concat(salt, utf16(password)));
  const buf = new Uint8Array(24);
  for (let i = 0; i < 50000; i++) {
    buf.set(le32(i));
    buf.set(h, 4);
    h = sha1(buf);
  }
  h = sha1(concat(h, le32(0)));
  const x = (c: number) => {
    const b = new Uint8Array(64).fill(c);
    for (let i = 0; i < 20; i++) b[i]! ^= h[i]!;
    return sha1(b);
  };
  return concat(x(0x36), x(0x5c)).subarray(0, keyBytes);
}

/** Decrypts a standard-encrypted package; `info` is the whole EncryptionInfo stream. */
export async function standardDecrypt(
  info: Uint8Array,
  pkg: Uint8Array,
  password: string,
  maxBytes: number,
): Promise<Uint8Array> {
  if (info.length < 12) fail('EncryptionInfo is too short');
  const dv = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const flags = dv.getUint32(4, true);
  const hdrSize = dv.getUint32(8, true);
  const v = 12 + hdrSize;
  if (hdrSize < 32 || v + 72 > info.length) fail('Bad standard EncryptionInfo header');
  const algId = dv.getUint32(20, true);
  const keyBits = dv.getUint32(28, true) || 128;
  // fCryptoAPI without fAES (or the RC4 AlgID) is RC4 CryptoAPI; fExternal is extensible.
  if (flags & 0x10 || !(flags & 0x20) || algId === 0x6801)
    fail('RC4 CryptoAPI and extensible encryption are not supported');
  if (![128, 192, 256].includes(keyBits) || dv.getUint32(24, true) !== 0x8004)
    fail(`Unsupported standard encryption (${keyBits}-bit)`);
  const saltSize = dv.getUint32(v, true);
  if (saltSize !== 16) fail('Bad standard encryption verifier salt size');
  const salt = info.subarray(v + 4, v + 20);
  const key = await aesKey(standardKey(salt, password, keyBits / 8));
  const verifier = await ecbDecrypt(key, info.subarray(v + 20, v + 36));
  const vhash = await ecbDecrypt(key, info.subarray(v + 40, v + 72));
  if (!equal(vhash.subarray(0, 20), sha1(verifier))) throw badPassword();
  return decryptPackage(pkg, maxBytes, (data) => ecbDecrypt(key, data));
}
