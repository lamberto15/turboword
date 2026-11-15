// Zip layer: lazy reader, raw-copy writer, CRC-32, deflate backends, limits and malformed input.
// Archives are built in code (with our writer or with fflate as an independent implementation).

import { inflateSync, strToU8, unzipSync, Zip, ZipDeflate, ZipPassThrough, zipSync } from 'fflate';
import { describe, expect, it, vi } from 'vitest';
import { TurbowordError, UsageError } from '../src/core/errors.js';
import { crc32Table } from '../src/zip/crc32.js';
import {
  assertSafePath,
  crc32,
  dosDateTime,
  makeEntry,
  type RawEntry,
  rawEntry,
  shouldStore,
  unzipLazy,
  writeZip,
  ZipFormatError,
  ZipLimitError,
  ZipPathError,
} from '../src/zip/index.js';
import { nodeZlib } from '../src/zip/node.js';

// ---------------------------------------------------------------------------------------------
// helpers

/** mulberry32: small seeded PRNG so random inputs are reproducible. */
function prng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomBytes(n: number, seed: number): Uint8Array {
  const rnd = prng(seed);
  const out = new Uint8Array(n);
  for (let i = 0; i < n; i++) out[i] = Math.floor(rnd() * 256);
  return out;
}

/** Compressible XML-ish text. */
function xmlText(rows: number): Uint8Array {
  const s: string[] = [];
  for (let i = 0; i < rows; i++) s.push(`<w:p><w:r><w:t>Paragraph ${i} é ü 日本</w:t></w:r></w:p>`);
  return strToU8(s.join(''));
}

/** fflate's CRC-32 (through ZipPassThrough, which computes it while streaming). */
function fflateCrc(data: Uint8Array): number {
  const p = new ZipPassThrough('x');
  p.ondata = () => {};
  p.push(data, true);
  return p.crc >>> 0;
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

/** Offset of the central directory of an archive without a comment. */
function cdOffset(zip: Uint8Array): number {
  return new DataView(zip.buffer, zip.byteOffset).getUint32(zip.length - 22 + 16, true);
}

function dv(zip: Uint8Array): DataView {
  return new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
}

const sample: [string, Uint8Array][] = [
  ['[Content_Types].xml', strToU8('<Types/>')],
  ['word/document.xml', xmlText(500)],
  ['word/日本語/ü.xml', strToU8('<x>unicode name</x>')],
  ['empty.txt', new Uint8Array(0)],
  ['word/media/image1.png', randomBytes(3000, 1)],
  ['bin.dat', Uint8Array.from({ length: 5000 }, (_, i) => (i * 7) & 255)],
];

function sampleZip(o?: Parameters<typeof makeEntry>[2]): Uint8Array {
  return writeZip(
    sample.map(([n, d]) => makeEntry(n, d, { ...o, store: o?.store ?? shouldStore(n) })),
  );
}

let bombCache: Uint8Array | undefined;
/** 20 MB of zeros at level 9: ratio ~1030, above the default limit of 1000. */
function bomb(): Uint8Array {
  bombCache ??= zipSync({ 'zeros.bin': new Uint8Array(20_000_000) }, { level: 9 });
  return bombCache;
}

// ---------------------------------------------------------------------------------------------

describe('write -> read round-trip', () => {
  it('reads back every entry with its metadata', () => {
    const entries = sample.map(([n, d]) => makeEntry(n, d, { store: shouldStore(n) }));
    const zip = writeZip(entries);
    const m = unzipLazy(zip);
    expect([...m.keys()]).toEqual(sample.map(([n]) => n));
    for (const [i, [name, data]] of sample.entries()) {
      const e = m.get(name)!;
      const w = entries[i]!;
      expect(e.bytes()).toEqual(data);
      expect(e.path).toBe(name);
      expect(e.size).toBe(data.length);
      expect(e.compressedSize).toBe(w.compressed.length);
      expect(e.compression).toBe(w.method);
      expect(e.crc32).toBe(fflateCrc(data));
      expect(e.flags).toBe(0x800); // UTF-8, no data descriptor
      expect(e.time).toBe(0);
      expect(e.date).toBe(0x21); // 1980-01-01
      expect(e.raw()).toEqual(w.compressed);
    }
    expect(m.get('word/document.xml')!.compression).toBe(8);
    expect(m.get('word/media/image1.png')!.compression).toBe(0);
    expect(m.get('empty.txt')!.compression).toBe(0);
  });

  it('is readable by fflate', () => {
    const files = unzipSync(sampleZip());
    expect(Object.keys(files)).toEqual(sample.map(([n]) => n));
    for (const [name, data] of sample) expect(files[name]).toEqual(data);
  });

  it('raw() and stored bytes() are views into the archive (no copy)', () => {
    const zip = sampleZip();
    const m = unzipLazy(zip);
    expect(m.get('word/document.xml')!.raw().buffer).toBe(zip.buffer);
    expect(m.get('word/media/image1.png')!.bytes().buffer).toBe(zip.buffer);
  });

  it('writes an empty archive and reads it back', () => {
    const zip = writeZip([]);
    expect(zip.length).toBe(22);
    expect(unzipLazy(zip).size).toBe(0);
  });

  it('skips directory entries', () => {
    const zip = zipSync({ 'dir/': new Uint8Array(0), 'dir/a.txt': strToU8('a') });
    expect([...unzipLazy(zip).keys()]).toEqual(['dir/a.txt']);
  });
});

describe('rawEntry copy', () => {
  it('preserves compressed bytes, CRC, sizes, method and time/date exactly', () => {
    const mtime = new Date(2021, 5, 15, 10, 30, 20);
    const src = zipSync(
      {
        'a.xml': [xmlText(300), { level: 9, mtime }],
        'b.bin': [randomBytes(2000, 2), { level: 0, mtime }],
        'c.txt': [strToU8('hello'), { level: 1, mtime }],
      },
      {},
    );
    const a = unzipLazy(src);
    const copy = writeZip([...a.values()].map(rawEntry));
    const b = unzipLazy(copy);
    expect([...b.keys()]).toEqual([...a.keys()]);
    for (const [name, x] of a) {
      const y = b.get(name)!;
      expect(y.raw()).toEqual(x.raw());
      expect(y.crc32).toBe(x.crc32);
      expect(y.size).toBe(x.size);
      expect(y.compressedSize).toBe(x.compressedSize);
      expect(y.compression).toBe(x.compression);
      expect(y.time).toBe(x.time);
      expect(y.date).toBe(x.date);
      expect(y.bytes()).toEqual(x.bytes());
    }
    expect(a.get('a.xml')!.date).not.toBe(0x21);
  });

  it('re-writing our own archive raw is byte-identical', () => {
    const zip = sampleZip();
    expect(writeZip([...unzipLazy(zip).values()].map(rawEntry))).toEqual(zip);
  });
});

describe('crc32', () => {
  it('matches fflate on random data (native and table)', () => {
    const lengths = [0, 1, 3, 7, 8, 9, 63, 64, 255, 1000, 4096, 65_537, 300_001];
    for (const [i, n] of lengths.entries()) {
      const data = randomBytes(n, 100 + i);
      const want = fflateCrc(data);
      expect(crc32(data)).toBe(want);
      expect(crc32Table(data)).toBe(want);
    }
  });

  it('known vector and continuation from a seed', () => {
    const v = strToU8('123456789');
    expect(crc32(v)).toBe(0xcbf43926);
    expect(crc32Table(v)).toBe(0xcbf43926);
    const a = randomBytes(1234, 7);
    const b = randomBytes(4321, 8);
    const whole = fflateCrc(concat([a, b]));
    expect(crc32(b, crc32(a))).toBe(whole);
    expect(crc32Table(b, crc32Table(a))).toBe(whole);
  });

  it('uses native zlib.crc32 in this runtime', () => {
    expect(typeof nodeZlib()?.crc32).toBe('function');
  });
});

describe('deflate backends', () => {
  const data = xmlText(4000);

  it('native and fflate both inflate back correctly', () => {
    expect(nodeZlib()).toBeDefined();
    const native = makeEntry('a.xml', data, { deflate: 'native' });
    const fflate = makeEntry('a.xml', data, { deflate: 'fflate' });
    // the native entry really came from node:zlib
    expect(native.compressed).toEqual(
      new Uint8Array(nodeZlib()!.deflateRawSync(data, { level: 6 })),
    );
    for (const e of [native, fflate]) {
      expect(e.method).toBe(8);
      expect(e.compressed.length).toBeLessThan(data.length / 4);
      expect(inflateSync(e.compressed)).toEqual(data);
      expect(
        unzipLazy(writeZip([e]))
          .get('a.xml')!
          .bytes(),
      ).toEqual(data);
    }
  });

  it('honors the level for both backends', () => {
    for (const deflate of ['native', 'fflate'] as const) {
      const fast = makeEntry('a', data, { level: 1, deflate });
      const best = makeEntry('a', data, { level: 9, deflate });
      expect(best.compressed.length).toBeLessThanOrEqual(fast.compressed.length);
      expect(inflateSync(fast.compressed)).toEqual(data);
      expect(inflateSync(best.compressed)).toEqual(data);
    }
  });

  it('stores when asked, at level 0, for empty data and when deflate does not help', () => {
    const s = makeEntry('a', data, { store: true });
    expect(s.method).toBe(0);
    expect(s.compressed).toBe(data);
    expect(makeEntry('a', data, { level: 0 }).method).toBe(0);
    const empty = makeEntry('e', new Uint8Array(0));
    expect([empty.method, empty.size, empty.crc32]).toEqual([0, 0, 0]);
    const noise = randomBytes(1000, 3);
    expect(makeEntry('n', noise).method).toBe(0);
    expect(makeEntry('n', noise, { deflate: 'fflate' }).method).toBe(0);
  });

  it('rejects an invalid level', () => {
    for (const level of [-1, 10, 1.5, Number.NaN]) {
      expect(() => makeEntry('a', data, { level })).toThrow(UsageError);
    }
  });
});

describe('data descriptors', () => {
  function streamZip(files: [string, Uint8Array, boolean][]): Uint8Array {
    const parts: Uint8Array[] = [];
    const z = new Zip((err, c) => {
      if (err) throw err;
      parts.push(c);
    });
    for (const [name, data, deflate] of files) {
      const f = deflate ? new ZipDeflate(name, { level: 6 }) : new ZipPassThrough(name);
      z.add(f);
      f.push(data, true);
    }
    z.end();
    return concat(parts);
  }

  const files: [string, Uint8Array, boolean][] = [
    ['s0.xml', strToU8('data-0-'.repeat(500)), true],
    ['s1.bin', randomBytes(777, 4), false],
    ['s2.xml', xmlText(200), true],
  ];

  it('reads entries whose local headers carry zeros (central directory values used)', () => {
    const zip = streamZip(files);
    const h = dv(zip);
    // fflate streaming: local header CRC and sizes are 0, real values follow the data
    expect(h.getUint16(6, true) & 8).toBe(8);
    expect([h.getUint32(14, true), h.getUint32(18, true), h.getUint32(22, true)]).toEqual([
      0, 0, 0,
    ]);
    const m = unzipLazy(zip);
    expect([...m.keys()]).toEqual(files.map(([n]) => n));
    for (const [name, data, deflate] of files) {
      const e = m.get(name)!;
      expect(e.flags & 8).toBe(8);
      expect(e.compression).toBe(deflate ? 8 : 0);
      expect(e.size).toBe(data.length);
      expect(e.crc32).toBe(fflateCrc(data));
      expect(e.bytes()).toEqual(data);
    }
  });

  it('raw copies get clean local headers without the descriptor flag', () => {
    const src = unzipLazy(streamZip(files));
    const out = writeZip([...src.values()].map(rawEntry));
    const h = dv(out);
    expect(h.getUint16(6, true)).toBe(0x800);
    expect(h.getUint32(14, true)).toBe(fflateCrc(files[0]![1]));
    expect(h.getUint32(22, true)).toBe(files[0]![1].length);
    const back = unzipSync(out);
    for (const [name, data] of files) expect(back[name]).toEqual(data);
    for (const e of unzipLazy(out).values()) expect(e.flags).toBe(0x800);
  });
});

describe('limits', () => {
  it('entry count', () => {
    const zip = sampleZip();
    expect(() => unzipLazy(zip, { maxEntries: sample.length - 1 })).toThrow(ZipLimitError);
    expect(() => unzipLazy(zip, { maxEntries: sample.length - 1 })).toThrow(/Too many/);
    expect(unzipLazy(zip, { maxEntries: sample.length }).size).toBe(sample.length);
  });

  it('ratio bomb', () => {
    const zip = bomb();
    const e = unzipLazy(zip).get('zeros.bin')!;
    expect(e.size / e.compressedSize).toBeGreaterThan(1000);
    expect(() => e.bytes()).toThrow(ZipLimitError);
    expect(() => e.bytes()).toThrow(/ratio/);
    // passes when the ratio limit is relaxed
    const ok = unzipLazy(zip, { maxCompressionRatio: 1e9 }).get('zeros.bin')!;
    expect(ok.bytes().length).toBe(20_000_000);
  });

  it('does not flag small highly compressible entries', () => {
    const zip = writeZip([makeEntry('a.xml', strToU8(' '.repeat(500_000)), { level: 9 })]);
    expect(unzipLazy(zip).get('a.xml')!.bytes().length).toBe(500_000);
  });

  it('entry size is checked on actual output, even when headers lie', () => {
    const zip = writeZip([makeEntry('a', strToU8('x'.repeat(100_000)))]);
    expect(() => unzipLazy(zip, { maxEntryBytes: 1000 })).toThrow(/too large/); // declared
    dv(zip).setUint32(cdOffset(zip) + 24, 1, true); // central directory now claims 1 byte
    const e = unzipLazy(zip, { maxEntryBytes: 1000 }).get('a')!;
    expect(() => e.bytes()).toThrow(/too large/);
  });

  it('total bytes, without double counting re-reads', () => {
    const zip = writeZip([
      makeEntry('a', strToU8('x'.repeat(6000))),
      makeEntry('b', strToU8('y'.repeat(6000))),
    ]);
    const m = unzipLazy(zip, { maxTotalBytes: 10_000 });
    m.get('a')!.bytes();
    m.get('a')!.bytes();
    expect(() => m.get('b')!.bytes()).toThrow(/total/i);
  });

  it.each([
    '../x',
    'a/../../x',
    '..',
    '/abs',
    'C:x',
    'C:/x',
    'c:\\x',
    'a\\b',
    'word\\document.xml',
    'a\0b',
  ])('rejects unsafe path %j', (p) => {
    const zip = zipSync({ [p]: strToU8('x') });
    expect(() => unzipLazy(zip)).toThrow(ZipPathError);
    expect(() => assertSafePath(p)).toThrow(ZipPathError);
    try {
      unzipLazy(zip);
    } catch (e) {
      expect(e).toBeInstanceOf(ZipLimitError);
      expect(e).toBeInstanceOf(TurbowordError);
      expect((e as TurbowordError).code).toBe('ZIP_PATH');
    }
  });

  it('accepts ordinary OPC names', () => {
    for (const p of ['[Content_Types].xml', '_rels/.rels', 'word/a..b.xml', 'customXml/', 'x']) {
      expect(() => assertSafePath(p)).not.toThrow();
    }
    expect(() => assertSafePath('')).toThrow(ZipPathError);
  });

  it('rejects duplicate entry names', () => {
    const zip = writeZip([makeEntry('a1', strToU8('one')), makeEntry('a2', strToU8('two'))]);
    // rename the second entry to 'a1' in its local header and central directory record
    const local2 = 30 + 2 + 3;
    zip[local2 + 30 + 1] = 0x31;
    zip[cdOffset(zip) + 46 + 2 + 46 + 1] = 0x31;
    expect(() => unzipLazy(zip)).toThrow(ZipPathError);
    expect(() => unzipLazy(zip)).toThrow(/Duplicate/);
  });

  it('encrypted entries and unsupported methods throw on access', () => {
    const make = (): Uint8Array => writeZip([makeEntry('a.xml', xmlText(50))]);
    const enc = make();
    dv(enc).setUint16(cdOffset(enc) + 8, 0x801, true);
    const e1 = unzipLazy(enc).get('a.xml')!;
    expect(e1.flags & 1).toBe(1);
    expect(() => e1.raw()).toThrow(/Encrypted/);
    expect(() => e1.bytes()).toThrow(ZipLimitError);
    expect(() => rawEntry(e1)).toThrow(ZipLimitError);

    const bz = make();
    dv(bz).setUint16(cdOffset(bz) + 10, 12, true); // bzip2
    const e2 = unzipLazy(bz).get('a.xml')!;
    expect(e2.compression).toBe(12);
    expect(() => e2.raw()).toThrow(/method 12/);
    expect(() => e2.bytes()).toThrow(ZipLimitError);
  });
});

describe('malformed input', () => {
  // Malformed input is a ZipFormatError: still a ZipLimitError, with its own code.
  const expectZipLimit = (fn: () => unknown): void => {
    let err: unknown;
    try {
      fn();
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(ZipLimitError);
    expect(err).toBeInstanceOf(ZipFormatError);
    expect((err as TurbowordError).code).toBe('ZIP_MALFORMED');
  };

  it('not a zip', () => {
    expectZipLimit(() => unzipLazy(new Uint8Array(0)));
    expectZipLimit(() => unzipLazy(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])));
    expectZipLimit(() => unzipLazy(strToU8('<?xml version="1.0"?><w:document/>')));
    expectZipLimit(() => unzipLazy(randomBytes(5000, 9)));
  });

  it('truncated archives', () => {
    const zip = sampleZip();
    expectZipLimit(() => unzipLazy(zip.subarray(0, zip.length - 30))); // EOCD gone
    expectZipLimit(() => unzipLazy(zip.subarray(0, zip.length >> 1)));
    expectZipLimit(() => unzipLazy(zip.subarray(100))); // offsets now point past the end
    // central directory claims more data than the archive holds
    const big = sampleZip();
    dv(big).setUint32(cdOffset(big) + 20, 0x7fffffff, true);
    const e = unzipLazy(big).get('[Content_Types].xml')!;
    expectZipLimit(() => e.raw());
    expectZipLimit(() => e.bytes());
  });

  it('bad central directory and local header signatures', () => {
    const cd = sampleZip();
    cd[cdOffset(cd)] = 0;
    expectZipLimit(() => unzipLazy(cd));
    const lh = sampleZip();
    lh[0] = 0;
    expectZipLimit(() => unzipLazy(lh).get('[Content_Types].xml')!.bytes());
  });

  it('corrupt deflate data', () => {
    const zip = writeZip([makeEntry('a', xmlText(3000))]);
    for (let i = 130; i < 170; i++) zip[30 + 1 + i] = 0xff;
    expectZipLimit(() => unzipLazy(zip).get('a')!.bytes());
  });
});

describe('ZIP64', () => {
  /** A hand-built ZIP64 archive (all 32-bit fields set to 0xffffffff) with one stored entry. */
  function zip64(data: Uint8Array): Uint8Array {
    const name = strToU8('a.txt');
    const crc = fflateCrc(data);
    const localLen = 30 + name.length + data.length;
    const cdLen = 46 + name.length + 28;
    const out = new Uint8Array(localLen + cdLen + 56 + 20 + 22);
    const v = dv(out);
    v.setUint32(0, 0x04034b50, true);
    v.setUint16(4, 45, true);
    v.setUint32(14, crc, true);
    v.setUint32(18, data.length, true);
    v.setUint32(22, data.length, true);
    v.setUint16(26, name.length, true);
    out.set(name, 30);
    out.set(data, 30 + name.length);
    let p = localLen;
    v.setUint32(p, 0x02014b50, true);
    v.setUint16(p + 4, 45, true);
    v.setUint16(p + 6, 45, true);
    v.setUint32(p + 16, crc, true);
    v.setUint32(p + 20, 0xffffffff, true);
    v.setUint32(p + 24, 0xffffffff, true);
    v.setUint16(p + 28, name.length, true);
    v.setUint16(p + 30, 28, true);
    v.setUint32(p + 42, 0xffffffff, true);
    out.set(name, p + 46);
    const x = p + 46 + name.length;
    v.setUint16(x, 1, true);
    v.setUint16(x + 2, 24, true);
    v.setBigUint64(x + 4, BigInt(data.length), true); // uncompressed
    v.setBigUint64(x + 12, BigInt(data.length), true); // compressed
    v.setBigUint64(x + 20, 0n, true); // local header offset
    p += cdLen;
    const e64 = p;
    v.setUint32(p, 0x06064b50, true);
    v.setBigUint64(p + 4, 44n, true);
    v.setUint16(p + 12, 45, true);
    v.setUint16(p + 14, 45, true);
    v.setBigUint64(p + 24, 1n, true);
    v.setBigUint64(p + 32, 1n, true);
    v.setBigUint64(p + 40, BigInt(cdLen), true);
    v.setBigUint64(p + 48, BigInt(localLen), true);
    p += 56;
    v.setUint32(p, 0x07064b50, true);
    v.setBigUint64(p + 8, BigInt(e64), true);
    v.setUint32(p + 16, 1, true);
    p += 20;
    v.setUint32(p, 0x06054b50, true);
    v.setUint16(p + 8, 0xffff, true);
    v.setUint16(p + 10, 0xffff, true);
    v.setUint32(p + 12, 0xffffffff, true);
    v.setUint32(p + 16, 0xffffffff, true);
    return out;
  }

  it('reads ZIP64 archives', () => {
    const data = strToU8('hello zip64');
    const e = unzipLazy(zip64(data)).get('a.txt')!;
    expect(e.size).toBe(data.length);
    expect(e.compressedSize).toBe(data.length);
    expect(e.bytes()).toEqual(data);
  });

  it('refuses to write archives that would need ZIP64', () => {
    const empty = new Uint8Array(0);
    const many: RawEntry[] = Array.from({ length: 65_535 }, (_, i) => ({
      name: `f${i}`,
      method: 0,
      crc32: 0,
      size: 0,
      compressed: empty,
    }));
    expect(() => writeZip(many)).toThrow(ZipLimitError);
    const max = writeZip(many.slice(1));
    expect(dv(max).getUint16(max.length - 22 + 10, true)).toBe(65_534);
    const huge: RawEntry = {
      name: 'big',
      method: 8,
      crc32: 0,
      size: 2 ** 32,
      compressed: new Uint8Array(1),
    };
    expect(() => writeZip([huge])).toThrow(/ZIP64/);
  });
});

describe('writer validation', () => {
  it('rejects unsafe or duplicate names and invalid entries', () => {
    const e = makeEntry('a', strToU8('a'));
    expect(() => writeZip([{ ...e, name: '../a' }])).toThrow(ZipPathError);
    expect(() => writeZip([e, e])).toThrow(/Duplicate/);
    expect(() => writeZip([{ ...e, method: 12 }])).toThrow(ZipLimitError);
    expect(() => writeZip([{ ...e, size: 5 }])).toThrow(ZipLimitError);
  });
});

describe('determinism', () => {
  it('identical inputs give identical bytes', () => {
    for (const deflate of ['native', 'fflate'] as const) {
      const a = sampleZip({ deflate });
      const b = sampleZip({ deflate });
      expect(a).toEqual(b);
    }
    const date = new Date(Date.UTC(2024, 0, 2, 3, 4, 6));
    expect(sampleZip({ date })).toEqual(sampleZip({ date }));
  });

  it('uses the fixed 1980-01-01 date in local and central headers', () => {
    const zip = writeZip([makeEntry('a', strToU8('abc'))]);
    const v = dv(zip);
    expect([v.getUint16(10, true), v.getUint16(12, true)]).toEqual([0, 0x21]);
    const cd = cdOffset(zip);
    expect([v.getUint16(cd + 12, true), v.getUint16(cd + 14, true)]).toEqual([0, 0x21]);
  });
});

describe('dosDateTime', () => {
  it('encodes UTC components with 2-second resolution', () => {
    const d = new Date(Date.UTC(2024, 2, 5, 13, 45, 59));
    expect(dosDateTime(d)).toEqual({
      date: (44 << 9) | (3 << 5) | 5,
      time: (13 << 11) | (45 << 5) | 29,
    });
    expect(dosDateTime(new Date(Date.UTC(1980, 0, 1)))).toEqual({ date: 0x21, time: 0 });
  });

  it('clamps to 1980-2107', () => {
    expect(dosDateTime(new Date(0))).toEqual({ date: 0x21, time: 0 });
    expect(dosDateTime(new Date(Date.UTC(1979, 11, 31, 23, 59, 59)))).toEqual({
      date: 0x21,
      time: 0,
    });
    const max = { date: (127 << 9) | (12 << 5) | 31, time: (23 << 11) | (59 << 5) | 29 };
    expect(dosDateTime(new Date(Date.UTC(2200, 5, 1)))).toEqual(max);
    expect(dosDateTime(new Date(Date.UTC(2107, 11, 31, 23, 59, 59)))).toEqual(max);
  });

  it('rejects an invalid Date', () => {
    expect(() => dosDateTime(new Date(Number.NaN))).toThrow(UsageError);
  });

  it('is what makeEntry writes for a date option', () => {
    const date = new Date(Date.UTC(2023, 6, 14, 8, 9, 10));
    const e = unzipLazy(writeZip([makeEntry('a', strToU8('x'), { date })])).get('a')!;
    expect({ date: e.date, time: e.time }).toEqual(dosDateTime(date));
  });
});

describe('shouldStore', () => {
  it('stores already-compressed media and containers', () => {
    for (const n of [
      'word/media/image1.png',
      'word/media/image2.JPEG',
      'word/media/hdphoto1.wdp',
      'word/media/a.jxr',
      'word/media/image3.emz',
      'word/media/media1.mp4',
      'word/embeddings/Microsoft_Excel_Worksheet.xlsx',
      'word/embeddings/doc.docx',
    ]) {
      expect(shouldStore(n)).toBe(true);
    }
  });

  it('deflates XML, fonts, metafiles and extensionless names', () => {
    for (const n of [
      'word/document.xml',
      '_rels/.rels',
      'word/fonts/font1.odttf',
      'word/media/image1.emf',
      'word/media/image1.bmp',
      'word/embeddings/oleObject1.bin',
      'noext',
      'dir.png/file',
    ]) {
      expect(shouldStore(n)).toBe(false);
    }
  });
});

describe('without node:zlib (browser-like runtime)', () => {
  it('falls back to fflate and the table CRC; limits still hold', async () => {
    const orig = process.getBuiltinModule.bind(process);
    const spy = vi
      .spyOn(process, 'getBuiltinModule')
      .mockImplementation(((id: string) =>
        id === 'node:zlib' || id === 'zlib' ? undefined : orig(id)) as never);
    vi.resetModules();
    try {
      const z = await import('../src/zip/index.js');
      const node = await import('../src/zip/node.js');
      expect(node.nodeZlib()).toBeUndefined();

      const data = xmlText(2000);
      const a = z.makeEntry('a.xml', data, { deflate: 'native' });
      const b = makeEntry('a.xml', data, { deflate: 'fflate' });
      expect(a.compressed).toEqual(b.compressed);
      expect(a.crc32).toBe(fflateCrc(data));
      expect(z.crc32(randomBytes(9999, 5))).toBe(fflateCrc(randomBytes(9999, 5)));
      expect(
        z
          .unzipLazy(z.writeZip([a]))
          .get('a.xml')!
          .bytes(),
      ).toEqual(data);

      const e = z.unzipLazy(bomb()).get('zeros.bin')!;
      expect(() => e.bytes()).toThrow(z.ZipLimitError);
      expect(() => e.bytes()).toThrow(/ratio/);
      const big = z.unzipLazy(bomb(), { maxCompressionRatio: 1e9 }).get('zeros.bin')!;
      expect(big.bytes().length).toBe(20_000_000);

      const corrupt = z.writeZip([a]);
      for (let i = 130; i < 170; i++) corrupt[30 + 5 + i] = 0xff;
      expect(() => z.unzipLazy(corrupt).get('a.xml')!.bytes()).toThrow(z.ZipLimitError);
    } finally {
      spy.mockRestore();
      vi.resetModules();
    }
  });
});
