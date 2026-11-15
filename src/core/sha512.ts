// Small synchronous SHA-1 / SHA-256 / SHA-384 / SHA-512 (FIPS 180-4). Two users, both of which
// hash tiny inputs many times and cannot use WebCrypto (async; an awaited digest per iteration
// costs seconds over 100,000 rounds):
// - `protect`: document-protection password hashes (w:documentProtection), computed by sync APIs;
// - `crypto`: the agile/standard encryption password spin loop.
// They live in core because protect and crypto are same-rank siblings that may not import each
// other. Bulk hashing (the HMAC over a whole package) goes through WebCrypto instead.
//
// SHA-512 works on 32-bit halves. Round constants are derived on first use (fractional parts of
// cube / square roots of the first primes, via BigInt) instead of shipping a 1.3 KB table;
// SHA-256 reuses the high words of SHA-512's constants. Ported from turbosheet (MIT, same author).

let K: Int32Array | undefined;
let IV: Int32Array | undefined;
let IV384: Int32Array | undefined;

/** Integer k-th root of `n` (Newton iteration). */
function iroot(n: bigint, k: bigint): bigint {
  let x = 1n << (BigInt(n.toString(2).length) / k + 1n);
  for (;;) {
    const y = ((k - 1n) * x + n / x ** (k - 1n)) / k;
    if (y >= x) return x;
    x = y;
  }
}

const MASK64 = (1n << 64n) - 1n;

/** Writes the 64-bit value `v` as hi/lo 32-bit words at `out[i * 2]`. */
function put64(out: Int32Array, i: number, v: bigint): void {
  out[i * 2] = Number(v >> 32n) | 0;
  out[i * 2 + 1] = Number(v & 0xffffffffn) | 0;
}

function init(): void {
  const k = new Int32Array(160);
  const iv = new Int32Array(16);
  let found = 0;
  for (let p = 2; found < 80; p++) {
    let prime = true;
    for (let d = 2; d * d <= p; d++) {
      if (p % d === 0) {
        prime = false;
        break;
      }
    }
    if (!prime) continue;
    put64(k, found, iroot(BigInt(p) << 192n, 3n) & MASK64);
    if (found < 8) put64(iv, found, iroot(BigInt(p) << 128n, 2n) & MASK64);
    found++;
  }
  K = k;
  IV = iv;
}

const W = /* @__PURE__ */ new Int32Array(160);

/** SHA-512 digest (64 bytes). A custom initial value `iv` (8 hi/lo word pairs) gives SHA-384. */
export function sha512(msg: Uint8Array, iv?: Int32Array): Uint8Array {
  if (K === undefined) init();
  const k = K!;
  const H = (iv ?? IV!).slice();
  const len = msg.length;
  const blocks = ((len + 16) >> 7) + 1;
  const buf = new Uint8Array(blocks * 128);
  buf.set(msg);
  buf[len] = 0x80;
  const dv = new DataView(buf.buffer);
  dv.setUint32(buf.length - 8, Math.floor(len / 0x20000000), false);
  dv.setUint32(buf.length - 4, (len << 3) >>> 0, false);
  const w = W;
  for (let b = 0; b < blocks; b++) {
    const off = b * 128;
    for (let t = 0; t < 32; t++) w[t] = dv.getInt32(off + t * 4, false);
    for (let t = 32; t < 160; t += 2) {
      // s0 = rotr1 ^ rotr8 ^ shr7 of w[t-15]; s1 = rotr19 ^ rotr61 ^ shr6 of w[t-2]
      let xh = w[t - 30]!;
      let xl = w[t - 29]!;
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      xh = w[t - 4]!;
      xl = w[t - 3]!;
      const s1h = ((xh >>> 19) | (xl << 13)) ^ ((xl >>> 29) | (xh << 3)) ^ (xh >>> 6);
      const s1l =
        ((xl >>> 19) | (xh << 13)) ^ ((xh >>> 29) | (xl << 3)) ^ ((xl >>> 6) | (xh << 26));
      const lo = (s0l >>> 0) + (s1l >>> 0) + (w[t - 13]! >>> 0) + (w[t - 31]! >>> 0);
      w[t] = s0h + s1h + w[t - 14]! + w[t - 32]! + Math.floor(lo / 0x100000000);
      w[t + 1] = lo;
    }
    let ah = H[0]!;
    let al = H[1]!;
    let bh = H[2]!;
    let bl = H[3]!;
    let ch = H[4]!;
    let cl = H[5]!;
    let dh = H[6]!;
    let dl = H[7]!;
    let eh = H[8]!;
    let el = H[9]!;
    let fh = H[10]!;
    let fl = H[11]!;
    let gh = H[12]!;
    let gl = H[13]!;
    let hh = H[14]!;
    let hl = H[15]!;
    for (let t = 0; t < 160; t += 2) {
      // S1 = rotr14 ^ rotr18 ^ rotr41 (e); ch = (e & f) ^ (~e & g)
      const S1h =
        ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l =
        ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh);
      const chl = (el & fl) ^ (~el & gl);
      const t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (k[t + 1]! >>> 0) + (w[t + 1]! >>> 0);
      const t1h = hh + S1h + chh + k[t]! + w[t]! + Math.floor(t1l / 0x100000000);
      // S0 = rotr28 ^ rotr34 ^ rotr39 (a); maj = (a & b) ^ (a & c) ^ (b & c)
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch);
      const mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
      hh = gh;
      hl = gl;
      gh = fh;
      gl = fl;
      fh = eh;
      fl = el;
      let lo = (dl >>> 0) + (t1l >>> 0);
      eh = (dh + t1h + Math.floor(lo / 0x100000000)) | 0;
      el = lo | 0;
      dh = ch;
      dl = cl;
      ch = bh;
      cl = bl;
      bh = ah;
      bl = al;
      lo = (t1l >>> 0) + (S0l >>> 0) + (mjl >>> 0);
      ah = (t1h + S0h + mjh + Math.floor(lo / 0x100000000)) | 0;
      al = lo | 0;
    }
    const v = [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl];
    for (let i = 0; i < 16; i += 2) {
      const lo = (H[i + 1]! >>> 0) + (v[i + 1]! >>> 0);
      H[i] = H[i]! + v[i]! + Math.floor(lo / 0x100000000);
      H[i + 1] = lo;
    }
  }
  return bytes(H);
}

/** SHA-384 digest (48 bytes): the SHA-512 core with its own initial value. */
export function sha384(msg: Uint8Array): Uint8Array {
  if (IV384 === undefined) {
    // Fractional square roots of the 9th..16th primes.
    const iv = new Int32Array(16);
    [23, 29, 31, 37, 41, 43, 47, 53].forEach((p, i) => {
      put64(iv, i, iroot(BigInt(p) << 128n, 2n) & MASK64);
    });
    IV384 = iv;
  }
  return sha512(msg, IV384).slice(0, 48);
}

// ---- SHA-1 / SHA-256 (64-byte blocks, 32-bit words) -----------------------------------------

const rotl = (x: number, n: number): number => (x << n) | (x >>> (32 - n));
const rotr = (x: number, n: number): number => (x >>> n) | (x << (32 - n));

/** Merkle–Damgård padding into big-endian 32-bit words (64-byte blocks). */
function words(msg: Uint8Array): Int32Array {
  const n = ((msg.length + 8) >> 6) + 1;
  const w = new Int32Array(n * 16);
  for (let i = 0; i < msg.length; i++) w[i >> 2]! |= msg[i]! << (24 - (i & 3) * 8);
  w[msg.length >> 2]! |= 0x80 << (24 - (msg.length & 3) * 8);
  w[n * 16 - 2] = Math.floor(msg.length / 0x20000000);
  w[n * 16 - 1] = msg.length << 3;
  return w;
}

/** Big-endian bytes of 32-bit words. */
function bytes(h: ArrayLike<number>): Uint8Array {
  const out = new Uint8Array(h.length * 4);
  const dv = new DataView(out.buffer);
  for (let i = 0; i < h.length; i++) dv.setInt32(i * 4, h[i]!, false);
  return out;
}

/** SHA-1 digest (20 bytes). */
export function sha1(msg: Uint8Array): Uint8Array {
  const m = words(msg);
  const h = [0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0];
  const w = new Int32Array(80);
  for (let b = 0; b < m.length; b += 16) {
    for (let t = 0; t < 80; t++)
      w[t] = t < 16 ? m[b + t]! : rotl(w[t - 3]! ^ w[t - 8]! ^ w[t - 14]! ^ w[t - 16]!, 1);
    let [a, bb, c, d, e] = h as [number, number, number, number, number];
    for (let t = 0; t < 80; t++) {
      const f =
        t < 20
          ? ((bb & c) | (~bb & d)) + 0x5a827999
          : t < 40
            ? (bb ^ c ^ d) + 0x6ed9eba1
            : t < 60
              ? ((bb & c) | (bb & d) | (c & d)) + 0x8f1bbcdc
              : (bb ^ c ^ d) + 0xca62c1d6;
      const tmp = (rotl(a, 5) + f + e + w[t]!) | 0;
      e = d;
      d = c;
      c = rotl(bb, 30);
      bb = a;
      a = tmp;
    }
    h[0] = (h[0]! + a) | 0;
    h[1] = (h[1]! + bb) | 0;
    h[2] = (h[2]! + c) | 0;
    h[3] = (h[3]! + d) | 0;
    h[4] = (h[4]! + e) | 0;
  }
  return bytes(h);
}

/** SHA-256 digest (32 bytes). */
export function sha256(msg: Uint8Array): Uint8Array {
  if (K === undefined) init();
  const k = K!;
  const m = words(msg);
  const h = Int32Array.from({ length: 8 }, (_, i) => IV![i * 2]!);
  const w = new Int32Array(64);
  for (let b = 0; b < m.length; b += 16) {
    for (let t = 0; t < 64; t++) {
      if (t < 16) w[t] = m[b + t]!;
      else {
        const x = w[t - 15]!;
        const y = w[t - 2]!;
        w[t] =
          (rotr(x, 7) ^ rotr(x, 18) ^ (x >>> 3)) +
          (rotr(y, 17) ^ rotr(y, 19) ^ (y >>> 10)) +
          w[t - 7]! +
          w[t - 16]!;
      }
    }
    let a = h[0]!;
    let bb = h[1]!;
    let c = h[2]!;
    let d = h[3]!;
    let e = h[4]!;
    let f = h[5]!;
    let g = h[6]!;
    let hh = h[7]!;
    for (let t = 0; t < 64; t++) {
      const t1 =
        hh + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + ((e & f) ^ (~e & g)) + k[t * 2]! + w[t]!;
      const t2 = (rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + ((a & bb) ^ (a & c) ^ (bb & c));
      hh = g;
      g = f;
      f = e;
      e = (d + t1) | 0;
      d = c;
      c = bb;
      bb = a;
      a = (t1 + t2) | 0;
    }
    const v = [a, bb, c, d, e, f, g, hh];
    for (let i = 0; i < 8; i++) h[i] = h[i]! + v[i]!;
  }
  return bytes(h);
}

/** A synchronous hash function. */
export type HashFn = (msg: Uint8Array) => Uint8Array;

/**
 * The SHA family by name: accepts `SHA512`, `SHA-512`, `sha512`, `SHA1`, `SHA-1`, ...
 * Returns undefined for anything else (MD2/MD4/MD5, RIPEMD, WHIRLPOOL are not implemented).
 */
export function shaByName(name: string): { fn: HashFn; size: number; name: string } | undefined {
  switch (name.toUpperCase().replace('-', '')) {
    case 'SHA1':
      return { fn: sha1, size: 20, name: 'SHA-1' };
    case 'SHA256':
      return { fn: sha256, size: 32, name: 'SHA-256' };
    case 'SHA384':
      return { fn: sha384, size: 48, name: 'SHA-384' };
    case 'SHA512':
      return { fn: sha512, size: 64, name: 'SHA-512' };
    default:
      return undefined;
  }
}
