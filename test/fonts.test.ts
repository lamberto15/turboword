// sfnt parsing (loadFont), standard-14 faces, kerning, collections, obfuscation and robustness.
//
// Expected metrics, glyph ids and kerning values were read from the fixture fonts with fontTools
// 4.68 and HarfBuzz (uharfbuzz; kerning = advance with `kern` on minus advance with it off).

import { describe, expect, it } from 'vitest';
import { TurbowordError } from '../src/core/errors.js';
import {
  deobfuscateFont,
  fontKeyBytes,
  loadFont,
  obfuscateFont,
  type SfntFont,
  STANDARD14_NAMES,
  standardFont,
  winAnsiCode,
} from '../src/layout/fonts/index.js';
import type { Font } from '../src/layout/types.js';
import { ByteWriter, buildSfnt, testFont, testFontBytes } from './helpers/test-fonts.js';

const cp = (s: string): number => s.codePointAt(0)!;

function expectFontError(f: () => unknown): void {
  try {
    f();
  } catch (err) {
    expect(err).toBeInstanceOf(TurbowordError);
    expect((err as TurbowordError).code).toBe('FONT_FORMAT');
    return;
  }
  throw new Error('expected a FONT_FORMAT error');
}

// ---- synthetic font builder ---------------------------------------------------------------------

interface SynthOptions {
  numGlyphs?: number;
  /** cmap subtables: [platform, encoding, subtable bytes]. */
  cmaps?: [number, number, Uint8Array][];
  extra?: Record<string, Uint8Array>;
  advances?: number[];
  family?: string;
  noOs2?: boolean;
}

function headTable(unitsPerEm = 1000, macStyle = 0): Uint8Array {
  const w = new ByteWriter();
  w.u32(0x10000).u32(0x10000).u32(0).u32(0x5f0f3cf5).u16(0).u16(unitsPerEm);
  w.u32(0).u32(0).u32(0).u32(0); // created, modified
  w.i16(-50).i16(-200).i16(1000).i16(900); // bbox
  w.u16(macStyle).u16(8).i16(2).i16(0).i16(0);
  return w.done();
}

function hheaTable(numH: number): Uint8Array {
  const w = new ByteWriter();
  w.u32(0x10000).i16(800).i16(-200).i16(90).u16(1000);
  for (let i = 0; i < 11; i++) w.i16(0); // min bearings, extent, caret, reserved, data format
  w.u16(numH);
  return w.done();
}

function maxpTable(numGlyphs: number): Uint8Array {
  return new ByteWriter().u32(0x5000).u16(numGlyphs).done();
}

function hmtxTable(advances: number[]): Uint8Array {
  const w = new ByteWriter();
  for (const a of advances) w.u16(a).i16(0);
  return w.done();
}

function nameTable(family: string, subfamily: string, ps: string): Uint8Array {
  const recs: [number, string][] = [
    [1, family],
    [2, subfamily],
    [6, ps],
  ];
  const w = new ByteWriter();
  w.u16(0)
    .u16(recs.length)
    .u16(6 + recs.length * 12);
  let off = 0;
  for (const [id, s] of recs) {
    w.u16(3)
      .u16(1)
      .u16(0x409)
      .u16(id)
      .u16(s.length * 2)
      .u16(off);
    off += s.length * 2;
  }
  for (const [, s] of recs) for (const ch of s) w.u16(ch.charCodeAt(0));
  return w.done();
}

function cmapFormat4(map: [number, number][]): Uint8Array {
  // One segment per mapping (idDelta), plus the 0xFFFF terminator.
  const segs = [...map].sort((a, b) => a[0] - b[0]).map(([c, g]) => [c, c, (g - c) & 0xffff]);
  segs.push([0xffff, 0xffff, 1]);
  const n = segs.length;
  const w = new ByteWriter();
  w.u16(4)
    .u16(16 + n * 8)
    .u16(0)
    .u16(n * 2)
    .u16(0)
    .u16(0)
    .u16(0);
  for (const s of segs) w.u16(s[1]!);
  w.u16(0);
  for (const s of segs) w.u16(s[0]!);
  for (const s of segs) w.u16(s[2]!);
  for (let i = 0; i < n; i++) w.u16(0);
  return w.done();
}

function cmapFormat12(groups: [number, number, number][]): Uint8Array {
  const w = new ByteWriter();
  w.u16(12)
    .u16(0)
    .u32(16 + groups.length * 12)
    .u32(0)
    .u32(groups.length);
  for (const [s, e, g] of groups) w.u32(s).u32(e).u32(g);
  return w.done();
}

function cmapTable(subtables: [number, number, Uint8Array][]): Uint8Array {
  const w = new ByteWriter();
  w.u16(0).u16(subtables.length);
  let off = 4 + subtables.length * 8;
  for (const [p, e, t] of subtables) {
    w.u16(p).u16(e).u32(off);
    off += t.length;
  }
  for (const [, , t] of subtables) w.bytes(t);
  return w.done();
}

/** A minimal CFF-flavoured font (no glyf/loca needed) with the given cmap subtables. */
function synthFont(o: SynthOptions = {}): Uint8Array {
  const numGlyphs = o.numGlyphs ?? 10;
  const advances = o.advances ?? Array.from({ length: numGlyphs }, (_, i) => 100 + i * 10);
  const tables: Record<string, Uint8Array> = {
    head: headTable(),
    hhea: hheaTable(advances.length),
    maxp: maxpTable(numGlyphs),
    hmtx: hmtxTable(advances),
    cmap: cmapTable(o.cmaps ?? [[3, 1, cmapFormat4([[0x41, 1]])]]),
    'CFF ': new Uint8Array([1, 0, 4, 1]),
    name: nameTable(o.family ?? 'Synth', 'Regular', 'Synth-Regular'),
    ...o.extra,
  };
  return buildSfnt(tables, 0x4f54544f);
}

// ---- real fonts: metrics, names, cmap, advances ---------------------------------------------------

interface Expected {
  file: [string, boolean];
  upm: number;
  numGlyphs: number;
  ascender: number;
  descender: number;
  lineGap: number;
  win: [number, number];
  cap: number;
  xh: number;
  names: [string, string, string];
  bold: boolean;
  gA: number;
  advA: number;
  advSpace: number;
  gEuro: number;
  advEuro: number;
  fixed: boolean;
  underline: [number, number];
  strike: [number, number];
}

const EXPECTED: Expected[] = [
  {
    file: ['Liberation Sans', false],
    upm: 2048,
    numGlyphs: 2620,
    ascender: 1854,
    descender: -434,
    lineGap: 67,
    win: [1854, 434],
    cap: 1409,
    xh: 1082,
    names: ['Liberation Sans', 'Regular', 'LiberationSans'],
    bold: false,
    gA: 36,
    advA: 1366,
    advSpace: 569,
    gEuro: 2088,
    advEuro: 1139,
    fixed: false,
    underline: [-67, 150],
    strike: [530, 102],
  },
  {
    file: ['Liberation Sans', true],
    upm: 2048,
    numGlyphs: 2620,
    ascender: 1854,
    descender: -434,
    lineGap: 67,
    win: [1854, 434],
    cap: 1409,
    xh: 1082,
    names: ['Liberation Sans', 'Bold', 'LiberationSans-Bold'],
    bold: true,
    gA: 36,
    advA: 1479,
    advSpace: 569,
    gEuro: 2088,
    advEuro: 1139,
    fixed: false,
    underline: [-2, 215],
    strike: [530, 102],
  },
  {
    file: ['Liberation Serif', false],
    upm: 2048,
    numGlyphs: 2602,
    ascender: 1825,
    descender: -443,
    lineGap: 87,
    win: [1825, 443],
    cap: 1341,
    xh: 940,
    names: ['Liberation Serif', 'Regular', 'LiberationSerif'],
    bold: false,
    gA: 36,
    advA: 1479,
    advSpace: 512,
    gEuro: 2088,
    advEuro: 1024,
    fixed: false,
    underline: [-123, 100],
    strike: [420, 100],
  },
  {
    file: ['Liberation Mono', true],
    upm: 2048,
    numGlyphs: 2423,
    ascender: 1705,
    descender: -615,
    lineGap: 0,
    win: [1705, 615],
    cap: 1349,
    xh: 1082,
    names: ['Liberation Mono', 'Bold', 'LiberationMono-Bold'],
    bold: true,
    gA: 36,
    advA: 1229,
    advSpace: 1229,
    gEuro: 2074,
    advEuro: 1229,
    fixed: true,
    underline: [-272, 205],
    strike: [530, 102],
  },
  {
    file: ['Carlito', false],
    upm: 2048,
    numGlyphs: 2783,
    ascender: 1950,
    descender: -550,
    lineGap: 0,
    win: [1950, 550],
    cap: 1314,
    xh: 978,
    names: ['Carlito', 'Regular', 'Carlito-Regular'],
    bold: false,
    gA: 3,
    advA: 1185,
    advSpace: 463,
    gEuro: 361,
    advEuro: 1038,
    fixed: false,
    underline: [-103, 194],
    strike: [512, 134],
  },
  {
    file: ['Caladea', true],
    upm: 1000,
    numGlyphs: 464,
    // USE_TYPO_METRICS (fsSelection 0x80) is set: typo metrics win, identical to hhea here.
    ascender: 900,
    descender: -250,
    lineGap: 0,
    win: [1050, 250],
    cap: 667,
    xh: 484,
    names: ['Caladea', 'Bold', 'Caladea-Bold'],
    bold: true,
    gA: 5,
    advA: 653,
    advSpace: 220,
    gEuro: 376,
    advEuro: 640,
    fixed: false,
    underline: [-75, 50],
    strike: [290, 50],
  },
];

describe('loadFont on the oracle fonts', () => {
  for (const e of EXPECTED) {
    it(`${e.file[0]} ${e.file[1] ? 'Bold' : 'Regular'}: metrics, names, cmap, advances`, () => {
      const f = testFont(e.file[0], e.file[1]);
      expect(f.outlines).toBe('truetype');
      expect(f.numGlyphs).toBe(e.numGlyphs);
      expect(f.metrics).toMatchObject({
        unitsPerEm: e.upm,
        ascender: e.ascender,
        descender: e.descender,
        lineGap: e.lineGap,
        winAscent: e.win[0],
        winDescent: e.win[1],
        capHeight: e.cap,
        xHeight: e.xh,
        underlinePosition: e.underline[0],
        underlineThickness: e.underline[1],
        strikeoutPosition: e.strike[0],
        strikeoutSize: e.strike[1],
        italicAngle: 0,
      });
      expect([f.family, f.subfamily, f.postscriptName]).toEqual(e.names);
      expect(f.bold).toBe(e.bold);
      expect(f.italic).toBe(false);
      expect(f.glyphId(cp('A'))).toBe(e.gA);
      expect(f.advance(e.gA)).toBe(e.advA);
      expect(f.advance(f.glyphId(0x20))).toBe(e.advSpace);
      expect(f.glyphId(0x20ac)).toBe(e.gEuro);
      expect(f.advance(e.gEuro)).toBe(e.advEuro);
      expect((f as SfntFont).fixedPitch).toBe(e.fixed);
      expect((f as SfntFont).weight).toBe(e.bold ? 700 : 400);
    });
  }

  it('returns 0 for missing characters and out-of-range input', () => {
    const f = testFont('Carlito');
    expect(f.glyphId(0x4e2d)).toBe(0); // 中
    expect(f.glyphId(0x2032)).toBe(0); // ′ (Carlito lacks it, Liberation has it)
    expect(f.glyphId(-1)).toBe(0);
    expect(f.glyphId(0x110000)).toBe(0);
    expect(f.glyphId(Number.NaN)).toBe(0);
    expect(f.advance(-1)).toBe(0);
    expect(f.advance(f.numGlyphs)).toBe(0);
    expect(f.advance(1.5)).toBe(f.advance(1));
    expect(testFont('Liberation Sans').glyphId(0x2032)).not.toBe(0);
  });

  it('keeps the input bytes as the face bytes and exposes details', () => {
    const bytes = testFontBytes('Caladea-Regular.ttf.gz');
    const f = loadFont(bytes) as SfntFont;
    expect(f.bytes).toBe(bytes);
    expect(f.bbox).toEqual([-313, -222, 1199, 936]);
    expect(f.fsType).toBe(0);
    expect(f.symbolic).toBe(false);
    expect(f.hasTable('GPOS')).toBe(true);
    expect(f.tableRange('glyf')!.length).toBeGreaterThan(0);
    expect(f.faceIndex).toBe(0);
  });
});

// ---- kerning ------------------------------------------------------------------------------------

/** HarfBuzz kern-on minus kern-off advance of the first glyph of each pair. */
const KERN_REFERENCE: [string, boolean, string, number][] = [
  ['Liberation Sans', false, 'AV', -152],
  ['Liberation Sans', false, 'To', -227],
  ['Liberation Sans', false, 'Ty', -113],
  ['Liberation Sans', true, 'To', -152],
  ['Liberation Serif', false, 'AV', -264],
  ['Liberation Serif', true, 'LT', -188],
  ['Liberation Mono', false, 'AV', 0],
  ['Carlito', false, 'AV', -89],
  ['Carlito', false, 'To', -182],
  ['Carlito', true, 'Ty', -64],
  ['Caladea', false, 'AV', -117],
  ['Caladea', false, 'VA', -119],
  ['Caladea', true, 'To', -88],
  ['Caladea', true, 'LT', -80],
];

describe('kerning', () => {
  for (const [family, bold, pair, value] of KERN_REFERENCE) {
    it(`${family}${bold ? ' Bold' : ''} "${pair}" = ${value}`, () => {
      const f = testFont(family, bold);
      const l = f.glyphId(cp(pair[0]!));
      const r = f.glyphId(cp(pair[1]!));
      expect(f.kerning(l, r)).toBe(value);
    });
  }

  it('is 0 for unkerned pairs and invalid glyph ids', () => {
    const f = testFont('Liberation Sans');
    expect(f.kerning(f.glyphId(cp('o')), f.glyphId(cp('o')))).toBe(0);
    expect(f.kerning(-1, 5)).toBe(0);
    expect(f.kerning(5, 0x10000)).toBe(0);
  });

  it('falls back to the legacy kern table when there is no GPOS', () => {
    // Rename the GPOS table tag so only the kern table (format 0) remains.
    const bytes = testFontBytes('LiberationSans-Regular.ttf.gz').slice();
    const n = (bytes[4]! << 8) | bytes[5]!;
    for (let i = 0; i < n; i++) {
      const p = 12 + i * 16;
      if (String.fromCharCode(...bytes.subarray(p, p + 4)) === 'GPOS') bytes[p] = 0x58; // 'XPOS'
    }
    const f = loadFont(bytes) as SfntFont;
    expect(f.hasTable('GPOS')).toBe(false);
    expect(f.hasTable('kern')).toBe(true);
    const k = (a: string, b: string): number => f.kerning(f.glyphId(cp(a)), f.glyphId(cp(b)));
    expect(k('A', 'V')).toBe(-152);
    expect(k('T', 'o')).toBe(-227);
  });

  it('reads synthetic kern tables (Microsoft and Apple headers) and GPOS extension lookups', () => {
    const cmap: [number, number, Uint8Array][] = [
      [
        3,
        1,
        cmapFormat4([
          [0x41, 1],
          [0x56, 2],
        ]),
      ],
    ];
    const kernMs = new ByteWriter()
      .u16(0)
      .u16(1) // version 0, 1 subtable
      .u16(0)
      .u16(6 + 8 + 6)
      .u16(0x0001) // format 0, horizontal
      .u16(1)
      .u16(6)
      .u16(0)
      .u16(0)
      .u16(1)
      .u16(2)
      .i16(-80)
      .done();
    expect(kernOf(synthFont({ cmaps: cmap, extra: { kern: kernMs } }))).toBe(-80);

    const kernApple = new ByteWriter()
      .u32(0x10000)
      .u32(1)
      .u32(8 + 8 + 6)
      .u16(0x0000) // format 0, horizontal
      .u16(0)
      .u16(1)
      .u16(6)
      .u16(0)
      .u16(0)
      .u16(1)
      .u16(2)
      .i16(-33)
      .done();
    expect(kernOf(synthFont({ cmaps: cmap, extra: { kern: kernApple } }))).toBe(-33);

    // Vertical / cross-stream subtables are ignored.
    const kernVertical = kernMs.slice();
    kernVertical[9] = 0x04; // cross-stream bit
    expect(kernOf(synthFont({ cmaps: cmap, extra: { kern: kernVertical } }))).toBe(0);

    expect(kernOf(synthFont({ cmaps: cmap, extra: { GPOS: gposPairFormat1(true) } }))).toBe(-55);
    expect(kernOf(synthFont({ cmaps: cmap, extra: { GPOS: gposPairFormat1(false) } }))).toBe(-55);
    // A malformed GPOS disables GPOS kerning and falls back to the kern table.
    const broken = gposPairFormat1(false).slice(0, 20);
    expect(kernOf(synthFont({ cmaps: cmap, extra: { GPOS: broken, kern: kernMs } }))).toBe(-80);
  });
});

function kernOf(bytes: Uint8Array): number {
  const f = loadFont(bytes);
  return f.kerning(f.glyphId(0x41), f.glyphId(0x56));
}

/** GPOS with one kern lookup: PairPos format 1, glyph 1 + glyph 2 → XAdvance -55. */
function gposPairFormat1(extension: boolean): Uint8Array {
  // PairPos subtable: format 1, coverage [1], valueFormat1 = XAdvance, valueFormat2 = 0, 1 set.
  const pair = new ByteWriter();
  pair.u16(1).u16(12).u16(0x0004).u16(0).u16(1).u16(18); // header (12 bytes) + set offset
  pair.u16(1).u16(1).u16(1); // coverage format 1, 1 glyph: 1   (at 12)
  pair.u16(1).u16(2).i16(-55); // PairSet: 1 record: second glyph 2, XAdvance -55 (at 18)
  const pairBytes = pair.done();
  const sub = extension ? new ByteWriter().u16(1).u16(2).u32(8).bytes(pairBytes).done() : pairBytes;
  const lookup = new ByteWriter()
    .u16(extension ? 9 : 2)
    .u16(0)
    .u16(1)
    .u16(8)
    .bytes(sub)
    .done();
  const w = new ByteWriter();
  // header: version 1.0, ScriptList @10, FeatureList @30, LookupList @44
  w.u16(1).u16(0).u16(10).u16(30).u16(44);
  // ScriptList (10): 1 record DFLT → script @8
  w.u16(1).tag('DFLT').u16(8);
  // Script (18): default LangSys @4, 0 langs
  w.u16(4).u16(0);
  // LangSys (22): no order, no required feature, 1 feature: 0
  w.u16(0).u16(0xffff).u16(1).u16(0);
  // FeatureList (30): 1 record 'kern' → feature @8
  w.u16(1).tag('kern').u16(8);
  // Feature (38): no params, 1 lookup: 0
  w.u16(0).u16(1).u16(0);
  // LookupList (44): 1 lookup @4
  w.u16(1).u16(4);
  w.bytes(lookup);
  return w.done();
}

// ---- cmap formats ---------------------------------------------------------------------------------

describe('cmap', () => {
  it('format 12 maps supplementary characters and wins over format 4', () => {
    const font = loadFont(
      synthFont({
        cmaps: [
          [3, 1, cmapFormat4([[0x41, 1]])],
          [
            3,
            10,
            cmapFormat12([
              [0x41, 0x43, 3], // A-C → 3..5
              [0x1f600, 0x1f601, 7], // 😀😁 → 7, 8
              [0x20000, 0x20000, 9],
            ]),
          ],
        ],
      }),
    );
    expect(font.glyphId(0x41)).toBe(3);
    expect(font.glyphId(0x43)).toBe(5);
    expect(font.glyphId(0x1f600)).toBe(7);
    expect(font.glyphId(0x1f601)).toBe(8);
    expect(font.glyphId(cp('𠀀'))).toBe(9);
    expect(font.glyphId(0x1f602)).toBe(0);
    expect(font.advance(7)).toBe(170);
    expect(font.outlines).toBe('cff');
  });

  it('falls through the subtable chain for characters missing from the best one', () => {
    const font = loadFont(
      synthFont({
        cmaps: [
          [3, 10, cmapFormat12([[0x1f600, 0x1f600, 7]])],
          [3, 1, cmapFormat4([[0x42, 2]])],
        ],
      }),
    );
    expect(font.glyphId(0x1f600)).toBe(7);
    expect(font.glyphId(0x42)).toBe(2);
  });

  it('decodes format 0 (Mac Roman) and format 6, and skips format 14', () => {
    const f0 = new ByteWriter().u16(0).u16(262).u16(0);
    for (let c = 0; c < 256; c++) f0.u8(c === 0x41 ? 4 : c === 0x80 ? 5 : 0);
    const f6 = new ByteWriter().u16(6).u16(16).u16(0).u16(0x2010).u16(3).u16(6).u16(0).u16(8);
    const f14 = new ByteWriter().u16(14).u32(10).u32(0);
    const font = loadFont(
      synthFont({
        cmaps: [
          [0, 5, f14.done()],
          [1, 0, f0.done()],
          [3, 1, f6.done()],
        ],
      }),
    );
    expect(font.glyphId(0x41)).toBe(4);
    expect(font.glyphId(0xc4)).toBe(5); // Ä is Mac Roman 0x80
    expect(font.glyphId(0x2010)).toBe(6);
    expect(font.glyphId(0x2011)).toBe(0);
    expect(font.glyphId(0x2012)).toBe(8);
  });

  it('maps Latin code points through U+F0xx in symbol fonts', () => {
    const font = loadFont(
      synthFont({
        cmaps: [
          [
            3,
            0,
            cmapFormat4([
              [0xf041, 3],
              [0xf0b7, 4],
            ]),
          ],
        ],
      }),
    );
    expect(font.glyphId(0x41)).toBe(3);
    expect(font.glyphId(0xf041)).toBe(3);
    expect(font.glyphId(0xb7)).toBe(4);
    expect((font as SfntFont).symbolic).toBe(true);
  });

  it('ignores glyph ids beyond numGlyphs', () => {
    const font = loadFont(synthFont({ numGlyphs: 4, cmaps: [[3, 1, cmapFormat4([[0x41, 9]])]] }));
    expect(font.glyphId(0x41)).toBe(0);
  });

  it('reads a font without OS/2, post or name', () => {
    const tables: Record<string, Uint8Array> = {
      head: headTable(1000, 3),
      hhea: hheaTable(2),
      maxp: maxpTable(3),
      hmtx: hmtxTable([500, 600]),
      cmap: cmapTable([[3, 1, cmapFormat4([[0x41, 2]])]]),
      'CFF ': new Uint8Array(4),
    };
    const font = loadFont(buildSfnt(tables, 0x4f54544f));
    expect(font.bold).toBe(true); // head.macStyle
    expect(font.italic).toBe(true);
    expect(font.advance(2)).toBe(600); // beyond numberOfHMetrics: the last advance
    expect(font.metrics).toMatchObject({
      ascender: 800,
      descender: -200,
      winAscent: 800,
      winDescent: 200,
    });
    expect(font.family).toBe('');
    expect(font.subfamily).toBe('Bold Italic');
  });
});

// ---- collections --------------------------------------------------------------------------------

/** Builds a TTC from whole sfnt files (each face keeps its own tables). */
function buildTtc(faces: Uint8Array[]): Uint8Array {
  const headerSize = 12 + faces.length * 4;
  const dirs: { tag: string; data: Uint8Array; checksum: number }[][] = faces.map((b) => {
    const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
    const n = v.getUint16(4);
    const out: { tag: string; data: Uint8Array; checksum: number }[] = [];
    for (let i = 0; i < n; i++) {
      const p = 12 + i * 16;
      const tag = String.fromCharCode(...b.subarray(p, p + 4));
      const off = v.getUint32(p + 8);
      const len = v.getUint32(p + 12);
      out.push({ tag, data: b.subarray(off, off + len), checksum: v.getUint32(p + 4) });
    }
    return out;
  });
  const dirSizes = dirs.map((d) => 12 + d.length * 16);
  let offset = headerSize + dirSizes.reduce((a, b) => a + b, 0);
  const w = new ByteWriter();
  w.tag('ttcf').u32(0x10000).u32(faces.length);
  let dirAt = headerSize;
  for (const size of dirSizes) {
    w.u32(dirAt);
    dirAt += size;
  }
  const tableData: Uint8Array[] = [];
  for (const d of dirs) {
    w.u32(0x10000).u16(d.length).u16(0).u16(0).u16(0);
    for (const t of d) {
      w.tag(t.tag).u32(t.checksum).u32(offset).u32(t.data.length);
      tableData.push(t.data);
      offset += (t.data.length + 3) & ~3;
    }
  }
  for (const t of tableData) {
    w.bytes(t);
    while (w.length % 4 !== 0) w.u8(0);
  }
  return w.done();
}

describe('TrueType collections', () => {
  const regular = testFontBytes('Caladea-Regular.ttf.gz');
  const bold = testFontBytes('Caladea-Bold.ttf.gz');
  const ttc = buildTtc([regular, bold]);

  it('loads each face by index and extracts it as a standalone sfnt', () => {
    const f0 = loadFont(ttc, 0);
    const f1 = loadFont(ttc, 1);
    expect([f0.family, f0.subfamily]).toEqual(['Caladea', 'Regular']);
    expect([f1.family, f1.subfamily]).toEqual(['Caladea', 'Bold']);
    expect(loadFont(ttc).subfamily).toBe('Regular');
    expect((f1 as SfntFont).faceIndex).toBe(1);
    expect(f1.advance(f1.glyphId(cp('A')))).toBe(653);
    expect(f1.kerning(f1.glyphId(cp('T')), f1.glyphId(cp('o')))).toBe(-88);
    // The face bytes are a single-face font that parses on its own to the same face.
    const again = loadFont(f1.bytes);
    expect(again.postscriptName).toBe('Caladea-Bold');
    expect(again.metrics).toEqual(f1.metrics);
    expect(String.fromCharCode(...f1.bytes.subarray(0, 4))).not.toBe('ttcf');
  });

  it('rejects bad face indices', () => {
    expectFontError(() => loadFont(ttc, 2));
    expectFontError(() => loadFont(ttc, -1));
    expectFontError(() => loadFont(ttc, 0.5));
    expectFontError(() => loadFont(regular, 1));
  });
});

// ---- standard 14 --------------------------------------------------------------------------------

describe('standardFont', () => {
  it('provides the 14 faces with AFM widths and WinAnsi codes', () => {
    expect(STANDARD14_NAMES).toHaveLength(14);
    for (const name of STANDARD14_NAMES) {
      const f = standardFont(name);
      expect(f.postscriptName).toBe(name);
      expect(f.outlines).toBe('standard14');
      expect(f.bytes.length).toBe(0);
      expect(f.metrics.unitsPerEm).toBe(1000);
      expect(f.numGlyphs).toBe(256);
    }
    const h = standardFont('Helvetica');
    expect(h.glyphId(cp('A'))).toBe(65);
    expect(h.advance(65)).toBe(667);
    expect(h.advance(h.glyphId(0x20))).toBe(278);
    expect(h.glyphId(0x20ac)).toBe(0x80); // € in WinAnsi
    expect(h.advance(0x80)).toBe(556);
    expect(h.glyphId(0x201c)).toBe(0x93);
    expect(h.glyphId(0xe9)).toBe(0xe9);
    expect(h.glyphId(0x4e2d)).toBe(0);
    expect(h.glyphId(0x09)).toBe(0);
    expect(h.kerning(65, 86)).toBe(0);
    expect(standardFont('Helvetica-Bold').advance(65)).toBe(722);
    expect(standardFont('Times-Roman').advance(0x20)).toBe(250);
    expect(standardFont('Times-Italic').metrics.italicAngle).toBe(-15.5);
    expect(standardFont('Courier-BoldOblique').advance(cp('W'))).toBe(600);
  });

  it('encodes Symbol and ZapfDingbats with their built-in encodings and Word PUA codes', () => {
    const s = standardFont('Symbol');
    expect(s.glyphId(0x3b1)).toBe(0x61); // α
    expect(s.advance(0x61)).toBe(631);
    expect(s.glyphId(0xf061)).toBe(0x61);
    expect(s.glyphId(0x2200)).toBe(0x22); // ∀
    expect(s.glyphId(0x41)).toBe(0); // no Latin 'A' in Symbol (Α is U+0391)
    const z = standardFont('ZapfDingbats');
    expect(z.glyphId(0x2701)).toBe(0x21);
    expect(z.advance(0x21)).toBe(974);
    expect(z.glyphId(0xf021)).toBe(0x21);
  });

  it('matches the vertical metrics of the Word fonts it stands in for', () => {
    const sans = testFont('Liberation Sans').metrics;
    const h = standardFont('Helvetica').metrics;
    expect(h.winAscent).toBe(Math.round((sans.winAscent * 1000) / sans.unitsPerEm));
    expect(h.winDescent).toBe(Math.round((sans.winDescent * 1000) / sans.unitsPerEm));
    const serif = testFont('Liberation Serif').metrics;
    expect(standardFont('Times-Bold').metrics.winAscent).toBe(
      Math.round((serif.winAscent * 1000) / serif.unitsPerEm),
    );
    const mono = testFont('Liberation Mono').metrics;
    expect(standardFont('Courier').metrics.winDescent).toBe(
      Math.round((mono.winDescent * 1000) / mono.unitsPerEm),
    );
  });

  it('accepts case-insensitive names and family aliases, rejects others', () => {
    expect(standardFont('helvetica-boldoblique').postscriptName).toBe('Helvetica-BoldOblique');
    expect(standardFont('Arial').postscriptName).toBe('Helvetica');
    expect(standardFont('Arial,Bold').postscriptName).toBe('Helvetica-Bold');
    expect(standardFont('Times New Roman').postscriptName).toBe('Times-Roman');
    expect(standardFont('Times-Italic').italic).toBe(true);
    expect(standardFont('Courier New').fixedPitch).toBe(true);
    expect(standardFont('Helvetica')).toBe(standardFont('Helvetica')); // cached
    expect(() => standardFont('Comic Sans')).toThrow(TurbowordError);
  });

  it('winAnsiCode covers cp1252', () => {
    expect(winAnsiCode(0x41)).toBe(0x41);
    expect(winAnsiCode(0x152)).toBe(0x8c); // Œ
    expect(winAnsiCode(0x178)).toBe(0x9f); // Ÿ
    expect(winAnsiCode(0x2011)).toBe(0x2d); // non-breaking hyphen → hyphen
    expect(winAnsiCode(0x81)).toBe(0);
    expect(winAnsiCode(0x7f)).toBe(0);
  });
});

// ---- odttf obfuscation -----------------------------------------------------------------------------

describe('embedded font obfuscation', () => {
  const GUID = '{0C1E5E4E-0EC1-4E8B-9B57-7F3AD3D2C9A1}';

  it('derives the key from the GUID bytes in reverse order', () => {
    const key = fontKeyBytes(GUID)!;
    expect(Array.from(key.subarray(0, 4))).toEqual([0xa1, 0xc9, 0xd2, 0xd3]);
    expect(key[15]).toBe(0x0c);
    expect(fontKeyBytes('not-a-guid')).toBeUndefined();
  });

  it('round-trips and only touches the first 32 bytes', () => {
    const bytes = testFontBytes('Caladea-Regular.ttf.gz');
    const obf = obfuscateFont(bytes, GUID);
    expect(obf).not.toBe(bytes);
    expect(obf.subarray(32)).toEqual(bytes.subarray(32));
    expect(obf.subarray(0, 32)).not.toEqual(bytes.subarray(0, 32));
    expect(obf[0]).toBe(bytes[0]! ^ 0xa1);
    expect(obf[16]).toBe(bytes[16]! ^ 0xa1);
    expectFontError(() => loadFont(obf));
    const back = deobfuscateFont(obf, GUID);
    expect(back).toEqual(bytes);
    expect(loadFont(back).family).toBe('Caladea');
    expect(() => deobfuscateFont(bytes, 'xyz')).toThrow(TypeError);
  });
});

// ---- robustness ---------------------------------------------------------------------------------

/** Deterministic PRNG (mulberry32). */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Exercises every Font method; must never throw. */
function exercise(f: Font): void {
  for (const c of [0x20, 0x41, 0x56, 0xe9, 0x20ac, 0x3b1, 0x4e2d, 0x1f600, 0xf041]) {
    const g = f.glyphId(c);
    expect(Number.isInteger(g)).toBe(true);
    expect(Number.isFinite(f.advance(g))).toBe(true);
  }
  for (let l = 0; l < 40; l += 3) for (let r = 0; r < 40; r += 7) f.kerning(l, r);
  expect(typeof f.family).toBe('string');
  expect(Number.isFinite(f.metrics.ascender)).toBe(true);
}

function loadOrFontError(bytes: Uint8Array, index?: number): void {
  let font: Font;
  try {
    font = loadFont(bytes, index);
  } catch (err) {
    if (!(err instanceof TurbowordError) || err.code !== 'FONT_FORMAT') throw err;
    return;
  }
  exercise(font);
}

describe('malformed fonts', () => {
  it('rejects non-fonts with FONT_FORMAT', () => {
    expectFontError(() => loadFont(new Uint8Array(0)));
    expectFontError(() => loadFont(new Uint8Array(100)));
    expectFontError(() => loadFont(new TextEncoder().encode('wOFF and some more bytes')));
    expectFontError(() => loadFont(new TextEncoder().encode('wOF2 and some more bytes')));
    expectFontError(() => loadFont(new TextEncoder().encode('<html><body>not a font</body>')));
    expectFontError(() => loadFont('abc' as unknown as Uint8Array));
    // ttcf header claiming a billion faces in 16 bytes
    expectFontError(() =>
      loadFont(new ByteWriter().tag('ttcf').u32(0x10000).u32(1e9).u32(0).done()),
    );
    // sfnt header claiming 65535 tables
    expectFontError(() =>
      loadFont(new ByteWriter().u32(0x10000).u16(0xffff).u16(0).u16(0).u16(0).done()),
    );
  });

  it('rejects fonts missing required tables', () => {
    const base: Record<string, Uint8Array> = {
      head: headTable(),
      hhea: hheaTable(2),
      maxp: maxpTable(3),
      hmtx: hmtxTable([500, 600]),
      cmap: cmapTable([[3, 1, cmapFormat4([[0x41, 2]])]]),
      'CFF ': new Uint8Array(4),
    };
    expect(loadFont(buildSfnt(base, 0x4f54544f)).glyphId(0x41)).toBe(2);
    for (const tag of Object.keys(base)) {
      const { [tag]: _, ...rest } = base;
      expectFontError(() => loadFont(buildSfnt(rest, 0x4f54544f)));
    }
    expectFontError(() => loadFont(buildSfnt({ ...base, head: headTable(0) }, 0x4f54544f)));
    expectFontError(() => loadFont(buildSfnt({ ...base, maxp: maxpTable(0) }, 0x4f54544f)));
    expectFontError(() => loadFont(buildSfnt({ ...base, cmap: cmapTable([]) }, 0x4f54544f)));
    expectFontError(() => loadFont(buildSfnt({ ...base, head: new Uint8Array(20) }, 0x4f54544f)));
  });

  it('survives truncation at any length (real font)', () => {
    const bytes = testFontBytes('Caladea-Regular.ttf.gz');
    const next = rng(1);
    const cuts = [12, 13, 100, 200, 300, 1000, bytes.length - 1];
    for (let i = 0; i < 60; i++) cuts.push(Math.floor(next() * bytes.length));
    for (const n of cuts) loadOrFontError(bytes.subarray(0, n));
  });

  it('survives random byte corruption (real fonts, structural regions)', () => {
    const files = ['Caladea-Regular.ttf.gz', 'Carlito-Regular.ttf.gz'];
    const next = rng(42);
    for (const file of files) {
      const bytes = testFontBytes(file);
      const n = (bytes[4]! << 8) | bytes[5]!;
      // Table offsets of the small structural tables, where corruption is most dangerous.
      const regions: [number, number][] = [[0, 12 + n * 16]];
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      for (let i = 0; i < n; i++) {
        const p = 12 + i * 16;
        const tag = String.fromCharCode(...bytes.subarray(p, p + 4));
        if (
          ['head', 'hhea', 'maxp', 'cmap', 'OS/2', 'post', 'name', 'GPOS', 'kern'].includes(tag)
        ) {
          const off = view.getUint32(p + 8);
          regions.push([off, Math.min(view.getUint32(p + 12), 4096)]);
        }
      }
      for (let round = 0; round < 150; round++) {
        const copy = bytes.slice();
        const flips = 1 + Math.floor(next() * 12);
        for (let k = 0; k < flips; k++) {
          const [start, len] = regions[Math.floor(next() * regions.length)]!;
          const at = start + Math.floor(next() * len);
          copy[at] = Math.floor(next() * 256);
        }
        loadOrFontError(copy);
      }
    }
  });

  it('survives corrupted synthetic fonts with every cmap format and GPOS', () => {
    const f0 = new ByteWriter().u16(0).u16(262).u16(0);
    for (let c = 0; c < 256; c++) f0.u8(c % 7);
    const font = synthFont({
      cmaps: [
        [
          0,
          4,
          cmapFormat12([
            [0x41, 0x5a, 1],
            [0x1f600, 0x1f600, 3],
          ]),
        ],
        [
          3,
          1,
          cmapFormat4([
            [0x41, 1],
            [0x56, 2],
            [0xe9, 3],
          ]),
        ],
        [1, 0, f0.done()],
      ],
      extra: { GPOS: gposPairFormat1(true), kern: new Uint8Array([0, 0, 0, 1, 0, 0, 0, 30, 0, 1]) },
    });
    loadOrFontError(font);
    const next = rng(7);
    for (let round = 0; round < 400; round++) {
      const copy = font.slice();
      const flips = 1 + Math.floor(next() * 6);
      for (let k = 0; k < flips; k++)
        copy[12 + Math.floor(next() * (copy.length - 12))] = Math.floor(next() * 256);
      loadOrFontError(copy);
    }
    for (let round = 0; round < 50; round++) {
      const garbage = new Uint8Array(64 + Math.floor(next() * 512));
      for (let i = 0; i < garbage.length; i++) garbage[i] = Math.floor(next() * 256);
      garbage.set([0, 1, 0, 0], 0);
      loadOrFontError(garbage);
      garbage.set([0x74, 0x74, 0x63, 0x66], 0); // ttcf
      loadOrFontError(garbage, Math.floor(next() * 3));
    }
  });
});
