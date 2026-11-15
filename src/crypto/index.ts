// `turboword/crypto`: password-encrypted ("Encrypt with Password") Word documents. Office wraps
// the encrypted .docx zip in an OLE compound file (CFB) holding `EncryptionInfo` and
// `EncryptedPackage` streams (ECMA-376 Part 2 / MS-OFFCRYPTO). Reading supports Agile (Word
// 2010+) and Standard (Word 2007) encryption; writing produces Agile AES-256 / SHA-512, exactly
// what Word writes. Web Crypto only (`crypto.subtle`), so it runs in every runtime; kept out of
// the main entry so it costs nothing unless imported. Ported from turbosheet (MIT, same author).
//
//   const doc = readDocx(await decryptDocx(bytes, 'secret'));
//   const out = await encryptDocx(writeDocx(doc), 'secret');

import { type Limits, resolveLimits } from '../core/limits.js';
import { agileDecrypt, agileEncrypt, type EncryptOptions } from './agile.js';
import { isCfb, readCfb, writeCfb } from './cfb.js';
import { standardDecrypt } from './standard.js';
import { concat, fail, utf16 } from './util.js';

export type { EncryptOptions } from './agile.js';

/** Options of {@link decryptDocx}. */
export interface DecryptOptions {
  /**
   * Verify the agile `dataIntegrity` HMAC and reject tampered files (default true; files without
   * a dataIntegrity element are accepted).
   */
  verifyIntegrity?: boolean;
  /** Size limits: `maxEntryBytes` caps every container stream and the decrypted package. */
  limits?: Partial<Limits>;
  /** Reject files asking for more password-hash iterations than this (default 10,000,000). */
  maxSpinCount?: number;
}

const bytesOf = (data: Uint8Array | ArrayBuffer): Uint8Array =>
  data instanceof Uint8Array ? data : new Uint8Array(data);

/** True when `data` is an encrypted Office document (CFB with an `EncryptionInfo` stream). */
export function isEncrypted(data: Uint8Array | ArrayBuffer): boolean {
  const bytes = bytesOf(data);
  if (!isCfb(bytes)) return false;
  try {
    return readCfb(bytes).entries.some((e) => e.type === 2 && e.path === 'EncryptionInfo');
  } catch {
    return false;
  }
}

/**
 * Decrypts a password-encrypted .docx (or any encrypted OOXML file) to the plain zip bytes.
 * Throws `TurbowordError` with code `'BAD_PASSWORD'` on a wrong password, and `DocxFormatError`
 * when the file is not an encrypted document, is malformed or tampered with, or uses a scheme
 * this library does not implement (RC4 CryptoAPI, extensible encryption, IRM).
 */
export async function decryptDocx(
  data: Uint8Array | ArrayBuffer,
  password: string,
  o: DecryptOptions = {},
): Promise<Uint8Array> {
  const bytes = bytesOf(data);
  if (!isCfb(bytes)) fail('Not an encrypted Office document (no compound file signature)');
  const maxBytes = resolveLimits(o.limits).maxEntryBytes;
  const cfb = readCfb(bytes, maxBytes);
  const info = cfb.read('EncryptionInfo');
  const pkg = cfb.read('EncryptedPackage');
  if (!info || !pkg) return fail('Not an encrypted Office document (no EncryptionInfo stream)');
  if (info.length < 8) fail('EncryptionInfo is too short');
  const major = info[0]! | (info[1]! << 8);
  const minor = info[2]! | (info[3]! << 8);
  if (major === 4 && minor === 4) {
    const xml = new TextDecoder().decode(info.subarray(8));
    const ao: { verifyIntegrity?: boolean; maxSpinCount?: number; maxBytes: number } = { maxBytes };
    if (o.verifyIntegrity !== undefined) ao.verifyIntegrity = o.verifyIntegrity;
    if (o.maxSpinCount !== undefined) ao.maxSpinCount = o.maxSpinCount;
    return agileDecrypt(xml, pkg, password, ao);
  }
  if (minor === 2 && major >= 2 && major <= 4) {
    return standardDecrypt(info, pkg, password, maxBytes);
  }
  return fail(`Unsupported EncryptionInfo version ${major}.${minor}`);
}

// \x06DataSpaces storage (MS-OFFCRYPTO 2.1, 2.3.4.15 "strong encryption" data space). Constant
// content; Word writes it alongside agile-encrypted packages, so we do too.

/** UNICODE-LP-P4: u32 byte length, UTF-16LE, padded to 4 bytes. */
function lp(s: string): Uint8Array {
  const b = utf16(s);
  const out = new Uint8Array(4 + Math.ceil(b.length / 4) * 4);
  new DataView(out.buffer).setUint32(0, b.length, true);
  out.set(b, 4);
  return out;
}

function u32s(...v: number[]): Uint8Array {
  const out = new Uint8Array(v.length * 4);
  const dv = new DataView(out.buffer);
  v.forEach((x, i) => {
    dv.setUint32(i * 4, x, true);
  });
  return out;
}

function dataSpaces(): [string, Uint8Array][] {
  const D = '\x06DataSpaces/';
  const versions = u32s(1, 1, 1); // reader / updater / writer 1.0
  const map = concat(lp('EncryptedPackage'), lp('StrongEncryptionDataSpace'));
  const id = lp('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
  return [
    [`${D}Version`, concat(lp('Microsoft.Container.DataSpaces'), versions)],
    [`${D}DataSpaceMap`, concat(u32s(8, 1, 12 + map.length, 1, 0), map)],
    [
      `${D}DataSpaceInfo/StrongEncryptionDataSpace`,
      concat(u32s(8, 1), lp('StrongEncryptionTransform')),
    ],
    [
      `${D}TransformInfo/StrongEncryptionTransform/\x06Primary`,
      concat(
        u32s(8 + id.length, 1),
        id,
        lp('Microsoft.Container.EncryptionTransform'),
        versions,
        u32s(0, 0, 0, 4), // EncryptionName (empty), block size, cipher mode, reserved
      ),
    ],
  ];
}

/**
 * Encrypts .docx bytes (e.g. from `writeDocx`) with a password using Agile encryption — what Word
 * 2010+ writes: AES-256-CBC, SHA-512, 100,000 password-hash iterations, HMAC data integrity.
 * The keys and salts come from `crypto.getRandomValues`, so every call gives different bytes.
 */
export async function encryptDocx(
  data: Uint8Array | ArrayBuffer,
  password: string,
  o?: EncryptOptions,
): Promise<Uint8Array> {
  const [info, pkg] = await agileEncrypt(bytesOf(data), password, o);
  return writeCfb([...dataSpaces(), ['EncryptionInfo', info], ['EncryptedPackage', pkg]]);
}
