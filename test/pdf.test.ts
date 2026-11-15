// PDF renderer: writer basics, text with an embedded subset TrueType font (ToUnicode checked by
// extracting the text again with poppler), standard-14 text, vectors, links, outline, tagging,
// determinism and validity (qpdf --check).

import { createHash } from 'node:crypto';
import { unzlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { baseFont, SubstituteFont, substituteMetricsOf } from '../src/layout/fonts/index.js';
import type { Font } from '../src/layout/types.js';
import { toUnicodeCMap } from '../src/render/pdf/fonts.js';
import { md5 } from '../src/render/pdf/md5.js';
import { renderPdf } from '../src/render/pdf/render.js';
import { faceInfo, subsetTrueType, tables } from '../src/render/pdf/sfnt.js';
import { name, num, str } from '../src/render/pdf/writer.js';
import { fixtures } from './helpers/fixtures.js';
import {
  dump,
  face,
  fontBytes,
  HAS_POPPLER,
  HAS_QPDF,
  layoutOf,
  pdfText,
  qpdfCheck,
  rasterize,
  runOn,
  standard,
  textItem,
} from './pdf-helpers.js';

const latin = (b: Uint8Array): string => Buffer.from(b).toString('latin1');

/** Every stream of an uncompressed-structure PDF, inflated when FlateDecode'd. */
function streams(pdf: Uint8Array): { dict: string; data: Uint8Array }[] {
  const s = latin(pdf);
  const out: { dict: string; data: Uint8Array }[] = [];
  const re = /(\d+) 0 obj\n<<(.*?)\/Length (\d+) >>\nstream\n/gs;
  for (let m = re.exec(s); m; m = re.exec(s)) {
    const start = m.index + m[0].length;
    const raw = pdf.subarray(start, start + Number(m[3]));
    out.push({ dict: m[2]!, data: m[2]!.includes('/FlateDecode') ? unzlibSync(raw) : raw });
  }
  return out;
}

describe('pdf writer basics', () => {
  it('md5 matches node:crypto', () => {
    for (const input of [
      '',
      'abc',
      'a'.repeat(55),
      'b'.repeat(56),
      'c'.repeat(64),
      'x'.repeat(1000),
    ]) {
      const bytes = new TextEncoder().encode(input);
      expect(Buffer.from(md5(bytes)).toString('hex')).toBe(
        createHash('md5').update(bytes).digest('hex'),
      );
    }
  });

  it('formats numbers, strings and names', () => {
    expect(num(1)).toBe('1');
    expect(num(-0.0001)).toBe('0');
    expect(num(1.23456)).toBe('1.235');
    expect(num(1e-7)).toBe('0');
    expect(num(12345678.5)).toBe('12345678.5');
    expect(num(Number.NaN)).toBe('0');
    expect(str('a(b)\\c')).toBe('(a\\(b\\)\\\\c)');
    expect(str('é')).toBe('<FEFF00E9>');
    expect(name('A B/C#')).toBe('/A#20B#2FC#23');
  });

  it('ToUnicode CMap maps ligatures and astral characters', () => {
    const cmap = toUnicodeCMap(
      new Map([
        [3, 'ffi'],
        [5, '😀'],
        [7, ''],
      ]),
    );
    expect(cmap).toContain('<0003> <006600660069>');
    expect(cmap).toContain('<0005> <d83dde00>');
    expect(cmap).not.toContain('<0007>');
    expect(cmap).toContain('2 beginbfchar');
  });
});

describe('TrueType subsetting', () => {
  const bytes = fontBytes('LiberationSans-Regular');
  const f = face('LiberationSans-Regular');

  it('keeps glyph ids, composite components and checksums', () => {
    const aacute = f.glyphId(0xc1);
    expect(aacute).toBeGreaterThan(0);
    const sub = subsetTrueType(bytes, [aacute]);
    expect(sub.length).toBeLessThan(bytes.length / 4);
    const t = tables(sub);
    expect([...t.keys()].sort()).toEqual(
      ['cvt ', 'fpgm', 'glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp', 'prep'].filter((k) =>
        tables(bytes).has(k),
      ),
    );
    const dv = new DataView(sub.buffer, sub.byteOffset, sub.byteLength);
    expect(dv.getUint16(t.get('maxp')!.offset + 4)).toBe(f.numGlyphs);
    // Whole-file checksum must be 0xB1B0AFBA.
    let sum = 0;
    for (let i = 0; i < sub.length; i += 4) sum = (sum + dv.getUint32(i)) >>> 0;
    expect(sum).toBe(0xb1b0afba);
    // Glyph lengths: Á is composite, its components (A and the accent) must be present too.
    const head = t.get('head')!.offset;
    const long = dv.getInt16(head + 50) === 1;
    const loca = t.get('loca')!.offset;
    const off = (g: number): number =>
      long ? dv.getUint32(loca + g * 4) : dv.getUint16(loca + g * 2) * 2;
    const len = (g: number): number => off(g + 1) - off(g);
    const glyf = t.get('glyf')!.offset;
    expect(len(aacute)).toBeGreaterThan(0);
    expect(dv.getInt16(glyf + off(aacute))).toBeLessThan(0); // composite
    const components: number[] = [];
    let p = glyf + off(aacute) + 10;
    for (;;) {
      const flags = dv.getUint16(p);
      components.push(dv.getUint16(p + 2));
      p += 4 + (flags & 1 ? 4 : 2) + (flags & 8 ? 2 : flags & 0x40 ? 4 : flags & 0x80 ? 8 : 0);
      if (!(flags & 0x20)) break;
    }
    expect(components).toContain(f.glyphId(0x41));
    for (const c of components) expect(len(c)).toBeGreaterThan(0);
    expect(len(f.glyphId(0x42))).toBe(0); // B not used
    expect(len(0)).toBeGreaterThan(0); // .notdef kept
  });

  it('reads FontDescriptor facts', () => {
    const info = faceInfo(bytes);
    expect(info.unitsPerEm).toBe(2048);
    expect(info.bbox[2]).toBeGreaterThan(info.bbox[0]);
    expect(info.fixedPitch).toBe(false);
    expect(faceInfo(fontBytes('LiberationMono-Regular')).fixedPitch).toBe(true);
  });
});

describe('text', () => {
  const sans = face('LiberationSans-Regular');
  const serif = face('LiberationSerif-Bold');
  const lines = ['Hello, PDF world!', 'Grüße — “quotes” & ½ €', 'Second font line'];

  function textLayout() {
    const items = [
      textItem(sans, 0, lines[0]!, 72, 100),
      textItem(sans, 0, lines[1]!, 72, 130, 14, { color: 'C00000' }),
      textItem(serif, 1, lines[2]!, 72, 160, 18, { fakeItalic: true }),
    ];
    return layoutOf(
      [items, [textItem(sans, 0, 'Page two', 72, 100, 12, { fakeBold: true, scale: 80 })]],
      [sans, serif],
      [],
      {
        metadata: {
          title: 'Täst',
          author: 'A. Author',
          lang: 'en-US',
          subject: 's',
          keywords: 'k',
          creator: 'c',
        },
      },
    );
  }

  it('embeds subset Type0 fonts with widths and ToUnicode', () => {
    const pdf = renderPdf(textLayout());
    const s = latin(pdf);
    expect(s.startsWith('%PDF-1.7\n')).toBe(true);
    expect(s).toMatch(/\/BaseFont \/[A-Z]{6}\+LiberationSans /);
    expect(s).toContain('/Subtype /CIDFontType2');
    expect(s).toContain('/CIDToGIDMap /Identity');
    expect(s).toContain('/Encoding /Identity-H');
    expect(s).toMatch(/\/W \[\d+ \[/);
    expect(s).toContain('/ToUnicode');
    expect(s).toContain('/Producer (turboword)');
    expect(s).toContain(
      '/Title <FEFF005400E400730074>', // "Täst" as UTF-16BE
    );
    expect(s).toContain('/Lang (en-US)');
    expect(s).not.toContain('/CreationDate');
    // The embedded subset is small (a few dozen glyphs out of ~2600).
    const fontFile = streams(pdf).find((x) => x.dict.includes('/Length1'))!;
    expect(fontFile.data.length).toBeLessThan(40_000);
  });

  it('positions glyphs with TJ adjustments from the layout advances', () => {
    const f = sans;
    const item = textItem(f, 0, 'AVA', 10, 20, 10);
    item.advances = [5, 6, 7]; // force adjustments
    const pdf = renderPdf(layoutOf([[item]], [f]), { compress: false });
    const content = streams(pdf)
      .map((x) => latin(x.data))
      .find((c) => c.includes(' TJ'))!;
    expect(content).toContain('/F0 10 Tf');
    expect(content).toContain('1 0 0 1 10 772 Tm');
    const tj = /\[(.*)\] TJ/.exec(content)![1]!;
    // Recompute: rendered pen after each glyph must land on the layout target within 1/1000 em.
    const parts = tj.match(/<[0-9a-f]+>|-?\d+/g)!;
    let pen = 0;
    const starts: number[] = [];
    for (const p of parts) {
      if (p.startsWith('<')) {
        for (let i = 1; i < p.length - 1; i += 4) {
          starts.push(pen);
          const gid = Number.parseInt(p.slice(i, i + 4), 16);
          pen += (Math.round((f.advance(gid) * 1000) / f.metrics.unitsPerEm) * 10) / 1000;
        }
      } else pen -= (Number(p) * 10) / 1000;
    }
    expect(starts.map((v) => Math.round(v * 100) / 100)).toEqual([0, 5, 11]);
  });

  it('emits text state for colour, scale, fake bold/italic and rise', () => {
    const item = textItem(sans, 0, 'x²', 0, 50, 12, {
      color: '336699',
      scale: 90,
      fakeBold: true,
      fakeItalic: true,
      dy: [0, -4],
    });
    const pdf = renderPdf(layoutOf([[item, textItem(sans, 0, 'plain', 0, 80)]], [sans]), {
      compress: false,
    });
    const c = streams(pdf)
      .map((x) => latin(x.data))
      .find((x) => x.includes(' TJ'))!;
    expect(c).toContain('0.2 0.4 0.6 rg');
    expect(c).toContain('90 Tz');
    expect(c).toContain('2 Tr');
    expect(c).toContain('0.4 w');
    expect(c).toContain('1 0 0.2126 1 0 742 Tm');
    expect(c).toContain('4 Ts');
    // The plain item resets the persistent text state.
    const plain = c.slice(c.lastIndexOf('BT'));
    expect(plain).toContain('0 0 0 rg');
    expect(plain).toContain('100 Tz');
    expect(plain).toContain('0 Tr');
    expect(plain).toContain('0 Ts');
  });

  it('never writes non-finite numbers', () => {
    const item = textItem(sans, 0, 'ab', Number.NaN, 10, 12, {
      advances: [Number.NaN, Number.POSITIVE_INFINITY],
    });
    const s = latin(
      renderPdf(
        layoutOf(
          [[item, { kind: 'rect', x: 0, y: Number.NaN, width: 5, height: 5, fill: '000000' }]],
          [sans],
        ),
        { compress: false },
      ),
    );
    expect(s).not.toMatch(/NaN|Infinity/);
  });

  it.skipIf(!HAS_QPDF)('passes qpdf --check', () => {
    const r = qpdfCheck(dump('text.pdf', renderPdf(textLayout())));
    expect(r.output).toContain('No syntax or stream encoding errors');
    expect(r.status).toBe(0);
  });

  it.skipIf(!HAS_POPPLER)('extracts the same text (ToUnicode) and lists embedded subsets', () => {
    const pdf = renderPdf(textLayout());
    const text = pdfText(pdf).trim().split('\n');
    expect(text).toEqual([...lines, 'Page two']);
    const fonts = runOn(pdf, 'pdffonts', []).stdout;
    expect(fonts).toMatch(/[A-Z]{6}\+LiberationSans\s+CID TrueType\s+Identity-H\s+yes yes yes/);
    expect(fonts).toMatch(/[A-Z]{6}\+LiberationSerif-Bold\s+CID TrueType/);
    const info = runOn(pdf, 'pdfinfo', []).stdout;
    expect(info).toMatch(/Title:\s+Täst/);
    expect(info).toMatch(/Author:\s+A\. Author/);
    expect(info).toMatch(/Producer:\s+turboword/);
    expect(info).toMatch(/Pages:\s+2/);
  });

  it.skipIf(!HAS_POPPLER)('maps ligature glyphs back to all their characters', () => {
    const fi = sans.glyphId(0xfb01);
    expect(fi).toBeGreaterThan(0);
    const item = textItem(sans, 0, 'office', 72, 100);
    // o f [fi] c e — the ligature glyph covers clusters 2..3.
    item.glyphs = [
      sans.glyphId(0x6f),
      sans.glyphId(0x66),
      fi,
      sans.glyphId(0x63),
      sans.glyphId(0x65),
    ];
    item.clusters = [0, 1, 2, 4, 5];
    item.advances = item.glyphs.map((g) => (sans.advance(g) * 12) / sans.metrics.unitsPerEm);
    expect(pdfText(renderPdf(layoutOf([[item]], [sans]))).trim()).toBe('office');
  });

  it.skipIf(!HAS_POPPLER)('renders subset glyphs (including composites) visibly', () => {
    const item = textItem(sans, 0, 'ÁÁÁ', 4, 60, 48);
    const r = rasterize(renderPdf(layoutOf([{ width: 120, height: 72, items: [item] }], [sans])));
    expect(r.stderr).toBe('');
    let dark = 0;
    for (let y = 0; y < r.height; y++)
      for (let x = 0; x < r.width; x++) if (r.px(x, y)[0] < 100) dark++;
    expect(dark).toBeGreaterThan(200);
  });

  it('writes standard-14 text as single-byte WinAnsi codes', () => {
    const helv = standard('Helvetica');
    const item = textItem(helv, 0, 'Hi there', 72, 72);
    const pdf = renderPdf(layoutOf([[item]], [helv]), { compress: false });
    const s = latin(pdf);
    expect(s).toContain('/Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding');
    expect(s).not.toContain('/FontFile');
    const c = streams(pdf)
      .map((x) => latin(x.data))
      .find((x) => x.includes(' TJ'))!;
    expect(c).toContain('<4869207468657265>');
    if (HAS_POPPLER) expect(pdfText(pdf).trim()).toBe('Hi there');
    if (HAS_QPDF) expect(qpdfCheck(pdf).status).toBe(0);
  });

  it('embeds CFF faces whole as FontFile3 /OpenType', () => {
    const cff: Font = {
      family: sans.family,
      subfamily: sans.subfamily,
      postscriptName: sans.postscriptName,
      bold: false,
      italic: false,
      metrics: sans.metrics,
      outlines: 'cff',
      glyphId: (c) => sans.glyphId(c),
      advance: (g) => sans.advance(g),
      kerning: () => 0,
      bytes: sans.bytes,
      numGlyphs: sans.numGlyphs,
    };
    const pdf = renderPdf(layoutOf([[textItem(cff, 0, 'cff', 10, 10)]], [cff]), {
      compress: false,
    });
    const s = latin(pdf);
    expect(s).toContain('/Subtype /CIDFontType0');
    expect(s).toContain('/FontFile3');
    expect(s).toContain('/Subtype /OpenType');
    expect(s).toContain('/BaseFont /LiberationSans-Identity-H');
  });

  it('embeds a face once for all its metric-substitute wrappers (union subset)', () => {
    const caladea = face('Caladea-Regular');
    const metrics = substituteMetricsOf('Cambria');
    expect(metrics).toBeDefined();
    const asCambria = new SubstituteFont(caladea, 'Cambria', metrics!);
    const again = new SubstituteFont(caladea, 'Cambria', metrics!);
    expect(baseFont(asCambria)).toBe(caladea);
    const layout = layoutOf(
      [
        [textItem(caladea, 0, 'plain', 72, 100), textItem(asCambria, 1, 'wrapped', 72, 130)],
        [textItem(again, 2, 'XYZ', 72, 100)],
      ],
      [caladea, asCambria, again],
    );
    const pdf = renderPdf(layout, { compress: false });
    const s = latin(pdf);
    expect(s.match(/\/Subtype \/Type0/g)?.length).toBe(1);
    expect(s.match(/\/FontFile2/g)?.length).toBe(1);
    expect(s).toMatch(/\/Font << \/F0 \d+ 0 R >>/);
    // Every glyph drawn through any of the three fonts is in the one subset and in /W.
    const program = streams(pdf).find((x) => x.dict.includes('/Length1'))!.data;
    const t = tables(program);
    const dv = new DataView(program.buffer, program.byteOffset, program.byteLength);
    const loca = t.get('loca')!.offset;
    const long = dv.getInt16(t.get('head')!.offset + 50) === 1;
    const off = (g: number): number =>
      long ? dv.getUint32(loca + g * 4) : dv.getUint16(loca + g * 2) * 2;
    for (const ch of 'plainwrappedXYZ') {
      const g = caladea.glyphId(ch.codePointAt(0)!);
      expect(off(g + 1) - off(g), ch).toBeGreaterThan(0);
    }
    expect(off(caladea.glyphId(0x51) + 1) - off(caladea.glyphId(0x51))).toBe(0); // Q unused
    const content = streams(pdf)
      .map((x) => latin(x.data))
      .filter((c) => c.includes(' TJ'));
    expect(content.every((c) => !c.includes('/F1 ') && !c.includes('/F2 '))).toBe(true);
    if (HAS_QPDF) expect(qpdfCheck(pdf).status).toBe(0);
    if (HAS_POPPLER)
      expect(pdfText(renderPdf(layout)).trim().split('\n')).toEqual(['plain', 'wrapped', 'XYZ']);
  });

  it('warns about unknown font indices and skips the item', () => {
    const warnings: string[] = [];
    const item = textItem(sans, 3, 'x', 0, 0);
    const pdf = renderPdf(layoutOf([[item]], [sans]), { onWarning: (w) => warnings.push(w.code) });
    expect(warnings).toEqual(['PDF_FONT_MISSING']);
    expect(latin(pdf)).not.toContain('/Type0');
  });
});

describe('vectors', () => {
  const layout = () =>
    layoutOf([
      [
        { kind: 'rect', x: 10, y: 20, width: 100, height: 50, fill: 'FF0000' },
        { kind: 'rect', x: 10, y: 80, width: 100, height: 50, stroke: '0000FF', strokeWidth: 2 },
        { kind: 'rect', x: 10, y: 140, width: 100, height: 50, fill: '#00FF00', stroke: '000000' },
        { kind: 'rect', x: 10, y: 200, width: 10, height: 10 }, // nothing to paint
        {
          kind: 'line',
          x1: 0,
          y1: 300,
          x2: 612,
          y2: 300,
          color: '808080',
          width: 0.5,
          dash: [3, 1.5],
        },
        {
          kind: 'path',
          d: 'M 200 200 h 100 v 100 h -100 Z M 225 225 h 50 v 50 h -50 z',
          fill: '000000',
          evenOdd: true,
        },
        { kind: 'path', d: 'M300 400 A 50 50 0 1 0 400 400', stroke: '000000' },
      ],
    ]);

  it('draws rects, lines and paths with the right operators', () => {
    const pdf = renderPdf(layout(), { compress: false });
    const c = streams(pdf)
      .map((x) => latin(x.data))
      .find((x) => x.includes(' re'))!;
    expect(c).toContain('1 0 0 rg\n10 722 100 50 re f');
    expect(c).toContain('0 0 1 RG 2 w\n10 662 100 50 re S');
    expect(c).toContain('0 1 0 rg\n0 0 0 RG 1 w\n10 602 100 50 re B');
    expect(c).not.toContain('10 582 10 10 re');
    expect(c).toContain('q 0.502 0.502 0.502 RG 0.5 w [3 1.5] 0 d 0 492 m 612 492 l S Q');
    expect(c).toContain('200 592 m\n300 592 l\n300 492 l\n200 492 l\nh');
    expect(c).toContain('f*');
    expect(c).toMatch(/300 392 m\n.* c\n/);
  });

  it.skipIf(!HAS_QPDF)('passes qpdf --check', () => {
    expect(qpdfCheck(dump('vectors.pdf', renderPdf(layout()))).status).toBe(0);
  });

  it.skipIf(!HAS_POPPLER)('paints where expected', () => {
    const r = rasterize(renderPdf(layout()), 72);
    expect(r.px(60, 45)).toEqual([255, 0, 0]); // filled rect
    expect(r.px(250, 210)).toEqual([0, 0, 0]); // even-odd ring
    expect(r.px(250, 250)).toEqual([255, 255, 255]); // even-odd hole
  });
});

describe('links, destinations, outline, tagging', () => {
  const sans = face('LiberationSans-Regular');
  const layout = () =>
    layoutOf(
      [
        [
          textItem(sans, 0, 'Go to chapter', 72, 100),
          { kind: 'link', x: 72, y: 88, width: 80, height: 14, dest: 'chap2' },
          { kind: 'link', x: 72, y: 120, width: 80, height: 14, url: 'https://example.com/ä?q=1' },
          { kind: 'link', x: 72, y: 140, width: 80, height: 14, dest: 'missing' },
          { kind: 'link', x: 72, y: 160, width: 80, height: 14, url: ' JavaScript:alert(1)' },
        ],
        [textItem(sans, 0, 'Chapter 2', 72, 100)],
      ],
      [sans],
      [],
      {
        destinations: new Map([
          ['chap2', { page: 1, y: 88 }],
          ['_Toc1', { page: 0, y: 10 }],
          ['é', { page: 0, y: 0 }],
        ]),
        outline: [
          { title: 'Intro', level: 0, page: 0, y: 72 },
          { title: 'Sub', level: 1, page: 0, y: 200 },
          { title: 'Deep (skipped level)', level: 3, page: 1, y: 50 },
          { title: 'Chapter 2', level: 0, page: 1, y: 88 },
        ],
        metadata: { title: 'Linked' },
      },
    );

  it('writes link annotations, a /Dests name tree and the outline', () => {
    const s = latin(renderPdf(layout(), { compress: false }));
    expect(s).toContain('/Subtype /Link /Rect [72 690 152 704] /Border [0 0 0] /Dest (chap2)');
    expect(s).toContain('/A << /S /URI /URI (https://example.com/%C3%A4?q=1) >>');
    expect(s).not.toContain('(missing)');
    expect(s).not.toMatch(/alert/);
    expect(s).toMatch(
      /\/Names \[\(_Toc1\) \[\d+ 0 R \/XYZ null 782 null\] \(chap2\) \[\d+ 0 R \/XYZ null 704 null\] <FEFF00E9>/,
    );
    expect(s).toContain('/PageMode /UseOutlines');
    expect(s).toMatch(/\/Type \/Outlines \/First \d+ 0 R \/Last \d+ 0 R \/Count 4/);
    expect(s).toContain('/Title (Deep \\(skipped level\\))');
    expect(s).toContain('/ViewerPreferences << /DisplayDocTitle true >>');
  });

  it.skipIf(!HAS_QPDF)('passes qpdf --check, and qpdf sees the outline', () => {
    const pdf = dump('links.pdf', renderPdf(layout()));
    expect(qpdfCheck(pdf).status).toBe(0);
    const json = JSON.parse(runOn(pdf, 'qpdf', ['--json', '--json-key=outlines']).stdout);
    const titles = (o: { title: string; kids: unknown[] }[]): unknown[] =>
      o.map((x) => [x.title, titles(x.kids as never)]);
    expect(titles(json.outlines)).toEqual([
      ['Intro', [['Sub', [['Deep (skipped level)', []]]]]],
      ['Chapter 2', []],
    ]);
    expect(json.outlines[1].destpageposfrom1).toBe(2);
  });

  it.skipIf(!HAS_POPPLER)('poppler resolves the named destinations', () => {
    const out = runOn(renderPdf(layout()), 'pdfinfo', ['-dests']).stdout;
    expect(out).toMatch(/2\s+\[ XYZ\s+null\s+704\s+null\s+\]\s+"chap2"/);
  });

  it('writes a basic structure tree when tagged', () => {
    const pdf = renderPdf(layout(), { tagged: true, compress: false });
    const s = latin(pdf);
    expect(s).toContain('/MarkInfo << /Marked true >>');
    expect(s).toMatch(
      /\/Type \/StructTreeRoot \/K \d+ 0 R \/ParentTree << \/Nums \[0 \[\d+ 0 R\] 1 \[\d+ 0 R\]\] >>/,
    );
    expect(s).toContain('/StructParents 1');
    const c = streams(pdf)
      .map((x) => latin(x.data))
      .find((x) => x.includes(' TJ'))!;
    expect(c.startsWith('/P << /MCID 0 >> BDC\nBT')).toBe(true);
    if (HAS_QPDF) expect(qpdfCheck(pdf).status).toBe(0);
  });
});

describe('document level', () => {
  const sans = face('LiberationSans-Regular');

  it('is deterministic and dates only on request', () => {
    const mk = () => layoutOf([[textItem(sans, 0, 'Same bytes', 72, 72)]], [sans]);
    const a = renderPdf(mk());
    const b = renderPdf(mk());
    expect(Buffer.from(a).equals(Buffer.from(b))).toBe(true);
    const dated = latin(
      renderPdf(mk(), {
        date: new Date(Date.UTC(2026, 10, 15, 8, 9, 10)),
        title: 'T',
        author: 'Me',
      }),
    );
    expect(dated).toContain('/CreationDate (D:20261115080910Z)');
    expect(dated).toContain('/Title (T)');
    expect(dated).toContain('/Author (Me)');
    // /ID is derived from the content.
    const id = (p: Uint8Array) => /\/ID \[<([0-9A-F]{32})>/.exec(latin(p))![1];
    expect(id(a)).not.toBe(id(renderPdf(layoutOf([[textItem(sans, 0, 'Other', 72, 72)]], [sans]))));
  });

  it('has a valid xref table', () => {
    const pdf = renderPdf(layoutOf([[textItem(sans, 0, 'xref', 72, 72)]], [sans]));
    const s = latin(pdf);
    const start = Number(/startxref\n(\d+)\n%%EOF\n$/.exec(s)![1]);
    expect(s.slice(start, start + 4)).toBe('xref');
    const [, count] = /xref\n0 (\d+)\n/.exec(s.slice(start))!;
    const entries = s
      .slice(start)
      .split('\n')
      .slice(3, 2 + Number(count));
    entries.forEach((e, i) => {
      expect(e).toMatch(/^\d{10} 00000 n $/);
      expect(s.slice(Number(e.slice(0, 10)), Number(e.slice(0, 10)) + 10)).toMatch(
        new RegExp(`^${i + 1} 0 obj`),
      );
    });
  });

  it('renders an empty layout as one blank page', () => {
    const pdf = renderPdf(layoutOf([]));
    expect(latin(pdf)).toContain('/Count 1');
    if (HAS_QPDF) expect(qpdfCheck(pdf).status).toBe(0);
  });

  it('toPdf paginates a real document (once layout lands)', async () => {
    const [fx] = fixtures();
    if (!fx) return;
    const { readDocx } = await import('../src/index.js');
    // Imported lazily: index.ts pulls in the layout engine, which other modules do not need.
    const { toPdf } = await import('../src/render/pdf/index.js');
    const doc = readDocx(fx.bytes);
    let pdf: Uint8Array;
    try {
      pdf = await toPdf(doc, { fonts: [() => fontBytes('LiberationSans-Regular')] });
    } catch (e) {
      if (String((e as Error).message).includes('not implemented')) return; // TODO: layout-B stub
      throw e;
    }
    expect(latin(pdf.subarray(0, 8))).toBe('%PDF-1.7');
    if (HAS_QPDF) expect(qpdfCheck(pdf).status).toBe(0);
  });
});
