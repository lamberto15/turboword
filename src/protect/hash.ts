// Word's document-protection password hash (w:documentProtection hash / hashValue).
//
// Word never hashes the password itself. It first reduces it to the 32-bit key of the legacy
// Word 97 "editing restrictions" scheme (ECMA-376 Part 4 §14.8.2.4 / MS-OE376 2.1.410 note a;
// the high/low-word XOR algorithm below), because a document converted from .doc only carries
// that key. The key, byte-reversed and written as 8 uppercase hex digits ("Example" → key
// 0x64CEED7E → "7EEDCE64"), is what gets hashed: H0 = H(salt + UTF-16LE(hex)), then
// Hn = H(Hn-1 + LE32(n-1)) for spinCount rounds (the initial hash is not counted; the iterator
// is appended, unlike agile encryption where it is prepended).
//
// Verified against Word-written files (Apache POI's bug56076.docx: "Example", SHA-1, 100,000
// rounds) and Aspose.Words' published test vectors (MIT). Aspose applies the same pre-processing
// to the ISO attribute form (algorithmName/hashValue), and falls back to hashing the raw
// password; verifyProtectionPassword does the same.

import { type HashFn, shaByName } from '../core/sha512.js';

/** Initial high-order word by password length (1..15). */
const INITIAL_CODE = [
  0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9,
  0x280c, 0xa96a, 0x4ec3,
];

/** Encryption matrix: 15 rows (character position from the end) x 7 bits. */
const MATRIX = [
  [0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09],
  [0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf],
  [0x4563, 0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68, 0x5ad0],
  [0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40],
  [0xd849, 0xa0b3, 0x5147, 0xa28e, 0x553d, 0xaa7a, 0x44d5],
  [0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496, 0x390d, 0x721a],
  [0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe, 0xa7fc, 0x5fd9],
  [0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0],
  [0xb861, 0x60e3, 0xc1c6, 0x93ad, 0x377b, 0x6ef6, 0xddec],
  [0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10],
  [0xaa51, 0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168],
  [0x76b4, 0xed68, 0xcaf1, 0x85c3, 0x1ba7, 0x374e, 0x6e9c],
  [0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6, 0x1dad],
  [0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373, 0x06e6, 0x0dcc],
  [0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48c4],
];

/**
 * The legacy 32-bit Word password key (high word: matrix XOR; low word: rotate-XOR checksum).
 * Only the first 15 characters count, each reduced to one byte (low byte, or the high byte when
 * the low one is zero). The empty password has key 0.
 */
export function legacyPasswordKey(password: string): number {
  const n = Math.min(password.length, 15);
  if (n === 0) return 0;
  const b: number[] = [];
  for (let i = 0; i < n; i++) {
    const c = password.charCodeAt(i);
    b.push(c & 0xff || c >> 8);
  }
  let hi = INITIAL_CODE[n - 1]!;
  for (let i = 0; i < n; i++) {
    const row = MATRIX[15 - n + i]!;
    for (let bit = 0; bit < 7; bit++) if (b[i]! & (1 << bit)) hi ^= row[bit]!;
  }
  const step = (lo: number, x: number): number => (((lo >> 14) & 1) | ((lo << 1) & 0x7fff)) ^ x;
  let lo = 0;
  for (let i = n - 1; i >= 0; i--) lo = step(lo, b[i]!);
  lo = step(lo, n) ^ 0xce4b;
  return ((hi << 16) | lo) >>> 0;
}

/** The legacy key as Word hashes it: bytes reversed, 8 uppercase hex digits ("7EEDCE64"). */
export function legacyKeyString(password: string): string {
  const k = legacyPasswordKey(password);
  let s = '';
  for (let i = 0; i < 4; i++) s += ((k >>> (i * 8)) & 0xff).toString(16).padStart(2, '0');
  return s.toUpperCase();
}

/** `cryptAlgorithmSid` → hash name (MS-OE376 2.1.410 note h; 1-3 are MD2/MD4/MD5). */
const SID_NAMES: Readonly<Record<string, string>> = {
  '4': 'SHA-1',
  '12': 'SHA-256',
  '13': 'SHA-384',
  '14': 'SHA-512',
};

/** Hash name → `cryptAlgorithmSid`. */
export const SIDS: Readonly<Record<string, number>> = {
  'SHA-1': 4,
  'SHA-256': 12,
  'SHA-384': 13,
  'SHA-512': 14,
};

/** The hash function for a `cryptAlgorithmSid` or an `algorithmName`, if implemented. */
export function protectionHashFn(sidOrName: string): HashFn | undefined {
  return shaByName(SID_NAMES[sidOrName.trim()] ?? sidOrName.trim())?.fn;
}

/** Iterated salted hash: H0 = H(salt + UTF-16LE(input)), Hn = H(Hn-1 + LE32(n-1)). */
export function spunHash(
  fn: HashFn,
  salt: Uint8Array,
  input: string,
  spinCount: number,
): Uint8Array {
  const first = new Uint8Array(salt.length + input.length * 2);
  first.set(salt);
  for (let i = 0; i < input.length; i++) {
    const c = input.charCodeAt(i);
    first[salt.length + i * 2] = c & 0xff;
    first[salt.length + i * 2 + 1] = c >> 8;
  }
  let h = fn(first);
  let buf = new Uint8Array(0);
  for (let i = 0; i < spinCount; i++) {
    if (buf.length !== h.length + 4) buf = new Uint8Array(h.length + 4);
    buf.set(h);
    const o = h.length;
    buf[o] = i & 0xff;
    buf[o + 1] = (i >>> 8) & 0xff;
    buf[o + 2] = (i >>> 16) & 0xff;
    buf[o + 3] = i >>> 24;
    h = fn(buf);
  }
  return h;
}

/** Word's document-protection hash of `password` (legacy key pre-processing, then spunHash). */
export function protectionHash(
  fn: HashFn,
  salt: Uint8Array,
  password: string,
  spinCount: number,
): Uint8Array {
  return spunHash(fn, salt, legacyKeyString(password), spinCount);
}

export function toBase64(b: Uint8Array): string {
  let s = '';
  for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]!);
  return btoa(s);
}

/** Decodes base64, or undefined when the text is not valid base64. */
export function fromBase64(s: string): Uint8Array | undefined {
  try {
    return Uint8Array.from(atob(s.replace(/\s+/g, '')), (c) => c.charCodeAt(0));
  } catch {
    return undefined;
  }
}
