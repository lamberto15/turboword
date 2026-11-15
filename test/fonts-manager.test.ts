// The font manager (createFontManager), Word's rFonts slot rules, substitutions and the fonts
// embedded in a docx.

import { describe, expect, it } from 'vitest';
import type { Warning } from '../src/core/warnings.js';
import {
  baseFont,
  createFontManager,
  EAST_ASIAN_FALLBACKS,
  embeddedFaces,
  embeddedFonts,
  FONT_SUBSTITUTIONS,
  fontTable,
  genericFamily,
  obfuscateFont,
  SUBSTITUTE_METRICS,
  SubstituteFont,
  slotFor,
  substituteMetricsOf,
  substitutesOf,
} from '../src/layout/fonts/index.js';
import type { FontRequest, FontResolver, LayoutOptions } from '../src/layout/types.js';
import type { ResolvedRunProps } from '../src/resolve/types.js';
import { createDocx } from '../src/wml/create.js';
import { readDocx, type WordDocument, writeDocx } from '../src/wml/document.js';
import { validatePackage } from '../src/wml/validate.js';
import { testFont, testFontBytes, testFontResolver } from './helpers/test-fonts.js';

type Fonts = ResolvedRunProps['fonts'] & { hint?: string };

function rp(over: Partial<ResolvedRunProps> = {}): ResolvedRunProps {
  return {
    fonts: {},
    size: 11,
    sizeCs: 11,
    bold: false,
    italic: false,
    boldCs: false,
    italicCs: false,
    strike: false,
    doubleStrike: false,
    color: 'auto',
    caps: false,
    smallCaps: false,
    hidden: false,
    vertAlign: 'baseline',
    spacing: 0,
    position: 0,
    scale: 100,
    kern: 0,
    rtl: false,
    cs: false,
    ...over,
  };
}

const all = (family: string, extra: Partial<Fonts> = {}): Fonts => ({
  ascii: family,
  hAnsi: family,
  eastAsia: family,
  cs: family,
  ...extra,
});

const cp = (s: string): number => s.codePointAt(0)!;

async function manager(doc: WordDocument, runs: ResolvedRunProps[], o: LayoutOptions = {}) {
  const m = createFontManager(doc, o);
  for (const r of runs) m.want(r);
  await m.load();
  return m;
}

/** A resolver that records every request it receives. */
function recording(inner: FontResolver = testFontResolver()): {
  resolver: FontResolver;
  requests: FontRequest[];
} {
  const requests: FontRequest[] = [];
  return {
    requests,
    resolver: (req) => {
      requests.push({ ...req });
      return inner(req);
    },
  };
}

// ---- slots --------------------------------------------------------------------------------------

describe('slotFor (Word rFonts rules)', () => {
  const plain = rp();
  const hinted = rp({ fonts: { hint: 'eastAsia' } as Fonts });
  const hintedZh = rp({ fonts: { hint: 'eastAsia' } as Fonts, langEastAsia: 'zh-CN' });

  it.each([
    ['A', 'ascii'],
    ['~', 'ascii'],
    ['é', 'hAnsi'],
    ['§', 'hAnsi'],
    ['α', 'hAnsi'],
    ['Ж', 'hAnsi'],
    ['—', 'hAnsi'],
    ['😀', 'hAnsi'],
    ['ā', 'hAnsi'],
    ['中', 'eastAsia'],
    ['あ', 'eastAsia'],
    ['カ', 'eastAsia'],
    ['한', 'eastAsia'],
    ['、', 'eastAsia'],
    ['Ａ', 'eastAsia'],
    ['𠀀', 'eastAsia'],
    ['א', 'cs'],
    ['ا', 'cs'],
    ['ก', 'cs'],
    ['क', 'cs'],
    ['ﻻ', 'cs'],
  ])('%s → %s', (ch, slot) => {
    expect(slotFor(plain, cp(ch))).toBe(slot);
  });

  it('w:cs and w:rtl force the cs slot', () => {
    expect(slotFor(rp({ rtl: true }), cp('A'))).toBe('cs');
    expect(slotFor(rp({ cs: true }), cp('中'))).toBe('cs');
  });

  it('hint="eastAsia" moves ambiguous characters to eastAsia', () => {
    expect(slotFor(hinted, cp('A'))).toBe('ascii');
    expect(slotFor(hinted, cp('§'))).toBe('eastAsia');
    expect(slotFor(hinted, cp('α'))).toBe('eastAsia');
    expect(slotFor(hinted, cp('—'))).toBe('eastAsia');
    expect(slotFor(hinted, 0xe000)).toBe('eastAsia');
    expect(slotFor(hinted, cp('ß'))).toBe('hAnsi'); // not in the hint list
    // Some Latin letters follow the hint only for Chinese.
    expect(slotFor(hinted, cp('é'))).toBe('hAnsi');
    expect(slotFor(hintedZh, cp('é'))).toBe('eastAsia');
    expect(slotFor(hinted, cp('ā'))).toBe('hAnsi');
    expect(slotFor(hintedZh, cp('ā'))).toBe('eastAsia');
    expect(slotFor(rp({ fonts: { hint: 'cs' } as Fonts }), cp('α'))).toBe('cs');
  });
});

// ---- substitution table ---------------------------------------------------------------------------

describe('substitutions', () => {
  it('lists the metric-compatible families', () => {
    expect(FONT_SUBSTITUTIONS.Calibri).toEqual(['Carlito']);
    expect(FONT_SUBSTITUTIONS.Cambria).toEqual(['Caladea']);
    expect(substitutesOf('arial')).toEqual(['Liberation Sans', 'Arimo']);
    expect(substitutesOf('TIMES NEW ROMAN')).toEqual(['Liberation Serif', 'Tinos']);
    expect(substitutesOf('Courier New')).toEqual(['Liberation Mono', 'Cousine']);
    expect(substitutesOf('Georgia')).toEqual(['Gelasio']);
    expect(substitutesOf('Arial Narrow')).toEqual(['Liberation Sans Narrow']);
    expect(substitutesOf('Segoe UI')).toEqual(['Selawik']);
    expect(substitutesOf('Symbol')).toEqual(['OpenSymbol']);
    expect(substitutesOf('Unknown')).toEqual([]);
  });

  it('classifies families', () => {
    expect(genericFamily('Garamond')).toBe('serif');
    expect(genericFamily('Consolas')).toBe('monospace');
    expect(genericFamily('Noto Sans Mono')).toBe('monospace');
    expect(genericFamily('Verdana')).toBe('sans-serif');
    expect(genericFamily('PT Sans')).toBe('sans-serif');
    expect(genericFamily('Wingdings')).toBe('symbol');
    expect(genericFamily('Mystery', { name: 'Mystery', family: 'roman' })).toBe('serif');
    expect(genericFamily('Mystery', { name: 'Mystery', pitch: 'fixed' })).toBe('monospace');
    expect(genericFamily('Mystery', { name: 'Mystery', charset: '02' })).toBe('symbol');
  });
});

// ---- the manager ----------------------------------------------------------------------------------

describe('createFontManager', () => {
  it('resolves Word families to their metric-compatible substitutes', async () => {
    const runs = [
      rp({ fonts: all('Calibri') }),
      rp({ fonts: all('Calibri'), bold: true }),
      rp({ fonts: all('Cambria') }),
      rp({ fonts: all('Arial') }),
      rp({ fonts: all('Times New Roman'), bold: true }),
      rp({ fonts: all('Courier New') }),
      rp(), // no fonts at all: Times New Roman
    ];
    const m = await manager(createDocx(), runs, { fonts: testFontResolver() });
    const ps = runs.map((r) => m.fontFor(r, cp('A')).postscriptName);
    expect(ps).toEqual([
      'Carlito-Regular',
      'Carlito-Bold',
      'Caladea-Regular',
      'LiberationSans',
      'LiberationSerif-Bold',
      'LiberationMono',
      'LiberationSerif',
    ]);
    expect(m.fontFor(runs[0]!, cp('A'))).toBe(testFont('Carlito'));
  });

  it('uses the regular face with synthetic styles when a styled face is missing', async () => {
    const italic = rp({ fonts: all('Calibri'), italic: true });
    const boldItalic = rp({ fonts: all('Calibri'), bold: true, italic: true });
    const bold = rp({ fonts: all('Calibri'), bold: true });
    const m = await manager(createDocx(), [italic, boldItalic, bold], {
      fonts: testFontResolver(),
    });
    const fi = m.fontFor(italic, cp('a'));
    expect(fi.postscriptName).toBe('Carlito-Regular');
    expect(m.synthetic(italic, fi)).toEqual({ bold: false, italic: true });
    const fbi = m.fontFor(boldItalic, cp('a'));
    expect(fbi.postscriptName).toBe('Carlito-Bold');
    expect(m.synthetic(boldItalic, fbi)).toEqual({ bold: false, italic: true });
    expect(m.synthetic(bold, m.fontFor(bold, cp('a')))).toEqual({ bold: false, italic: false });
  });

  it('uses boldCs/italicCs for the cs slot', async () => {
    const run = rp({ fonts: all('Arial'), cs: true, boldCs: true, bold: false });
    const m = await manager(createDocx(), [run], { fonts: testFontResolver() });
    const f = m.fontFor(run, cp('A'));
    expect(f.postscriptName).toBe('LiberationSans-Bold');
    expect(m.synthetic(run, f)).toEqual({ bold: false, italic: false });
    const plainRun = rp({ fonts: all('Arial'), boldCs: true });
    const regular = testFont('Liberation Sans');
    expect(m.synthetic(plainRun, regular, cp('A'))).toEqual({ bold: false, italic: false });
    expect(m.synthetic(plainRun, regular, cp('א'))).toEqual({ bold: true, italic: false });
  });

  it('falls back to generic families, then to the standard 14 fonts', async () => {
    const doc = createDocx();
    const ft = doc.mainPart.relatedByType('fontTable')[0]!;
    const xml = new TextDecoder()
      .decode(ft.bytes())
      .replace(
        '</w:fonts>',
        '<w:font w:name="Mystery Roman"><w:family w:val="roman"/></w:font>' +
          '<w:font w:name="Mystery Fixed"><w:family w:val="auto"/><w:pitch w:val="fixed"/></w:font></w:fonts>',
      );
    ft.setBytes(new TextEncoder().encode(xml));
    expect(fontTable(doc).get('mystery roman')).toEqual({ name: 'Mystery Roman', family: 'roman' });

    const runs = [
      rp({ fonts: all('Mystery Roman') }),
      rp({ fonts: all('Mystery Fixed') }),
      rp({ fonts: all('Frobnicate') }),
    ];
    const withFonts = await manager(doc, runs, { fonts: testFontResolver() });
    expect(runs.map((r) => withFonts.fontFor(r, cp('A')).postscriptName)).toEqual([
      'LiberationSerif',
      'LiberationMono',
      'LiberationSans',
    ]);

    const warnings: Warning[] = [];
    const none = await manager(
      doc,
      [
        ...runs,
        rp({ fonts: all('Calibri'), bold: true, italic: true }),
        rp({ fonts: all('Wingdings') }),
      ],
      { onWarning: (w) => warnings.push(w) },
    );
    expect(runs.map((r) => none.fontFor(r, cp('A')).postscriptName)).toEqual([
      'Times-Roman',
      'Courier',
      'Helvetica',
    ]);
    const bi = rp({ fonts: all('Calibri'), bold: true, italic: true });
    const f = none.fontFor(bi, cp('A'));
    expect(f.postscriptName).toBe('Helvetica-BoldOblique');
    expect(none.synthetic(bi, f)).toEqual({ bold: false, italic: false });
    expect(none.fontFor(rp({ fonts: all('Wingdings') }), 0xf0fc).postscriptName).toBe('Symbol');
    expect(warnings.some((w) => w.code === 'FONT_FALLBACK' && w.message.includes('Calibri'))).toBe(
      true,
    );
  });

  it('tries the caller substitutions and the font-table altName first', async () => {
    const run = rp({ fonts: all('Calibri') });
    const m = await manager(createDocx(), [run], {
      fonts: testFontResolver(),
      substitutions: { calibri: 'Liberation Sans' },
    });
    expect(m.fontFor(run, cp('A')).postscriptName).toBe('LiberationSans');
  });

  it('resolves the eastAsia slot separately even when all slots name one family', async () => {
    const cjk = testFont('Caladea'); // stands in for a CJK font: only asked as eastAsia
    const { resolver, requests } = recording((req) =>
      req.family === 'Noto Sans CJK SC' && req.script === 'eastAsia' ? cjk : undefined,
    );
    const run = rp({ fonts: all('MS Mincho') });
    const m = await manager(createDocx(), [run], { fonts: resolver });
    expect(m.fontFor(run, cp('中'))).toBe(cjk);
    expect(m.fontFor(run, cp('A')).postscriptName).toBe('Times-Roman'); // Mincho is serif
    expect(requests.filter((r) => r.family === 'MS Mincho').map((r) => r.script)).toEqual([
      'latin',
      'eastAsia',
      'cs',
    ]);
  });

  it('asks resolvers with the slot script and East Asian fallbacks', async () => {
    const { resolver, requests } = recording();
    const run = rp({
      fonts: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'MS Mincho', cs: 'Arial' },
    });
    const m = await manager(createDocx(), [run, rp({ fonts: all('Symbol') })], { fonts: resolver });
    expect(requests[0]).toEqual({ family: 'Arial', bold: false, italic: false, script: 'latin' });
    expect(requests).toContainEqual({
      family: 'MS Mincho',
      bold: false,
      italic: false,
      script: 'eastAsia',
    });
    expect(requests).toContainEqual({
      family: EAST_ASIAN_FALLBACKS[0],
      bold: false,
      italic: false,
      script: 'eastAsia',
    });
    expect(requests).toContainEqual({
      family: 'Symbol',
      bold: false,
      italic: false,
      script: 'symbol',
    });
    expect(requests).toContainEqual({
      family: 'OpenSymbol',
      bold: false,
      italic: false,
      script: 'symbol',
    });
    expect(m.fontFor(run, cp('A')).postscriptName).toBe('LiberationSans');
  });

  it('picks the eastAsia face for hinted ambiguous characters', async () => {
    const run = rp({
      fonts: {
        ascii: 'Arial',
        hAnsi: 'Arial',
        eastAsia: 'Cambria',
        cs: 'Arial',
        hint: 'eastAsia',
      } as Fonts,
    });
    const m = await manager(createDocx(), [run], { fonts: testFontResolver() });
    expect(m.fontFor(run, cp('A')).family).toBe('Liberation Sans');
    expect(m.fontFor(run, cp('§')).family).toBe('Caladea');
    expect(
      m.fontFor(rp({ fonts: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'Cambria' } }), cp('§'))
        .family,
    ).toBe('Liberation Sans');
  });

  it('falls back per glyph through the other faces, then standard 14', async () => {
    const calibri = rp({ fonts: all('Calibri') });
    const calibriBold = rp({ fonts: all('Calibri'), bold: true });
    const arial = rp({ fonts: all('Arial') });
    const m = await manager(createDocx(), [calibri, calibriBold, arial], {
      fonts: testFontResolver(),
    });
    expect(m.fontFor(calibri, cp('A')).family).toBe('Carlito');
    // ′ (U+2032) is missing from Carlito but present in Liberation Sans.
    const prime = m.fontFor(calibri, 0x2032);
    expect(prime.family).toBe('Liberation Sans');
    // The bold run gets the regular Liberation Sans (no bold face loaded) and a synthetic bold.
    const primeBold = m.fontFor(calibriBold, 0x2032);
    expect(primeBold.postscriptName).toBe('LiberationSans');
    expect(m.synthetic(calibriBold, primeBold)).toEqual({ bold: true, italic: false });
    // ∀ is in neither: the standard-14 Symbol font can encode it.
    const forall = m.fontFor(calibri, 0x2200);
    expect(forall.postscriptName).toBe('Symbol');
    expect(forall.glyphId(0x2200)).toBe(0x22);
    // Nothing has 中: the slot's own face (renders .notdef).
    expect(m.fontFor(calibri, cp('中')).family).toBe('Carlito');
    // Invisible characters never trigger a fallback.
    expect(m.fontFor(calibri, 0x200b).family).toBe('Carlito');
  });

  it('lists handed-out faces in first-use order and is deterministic', async () => {
    const runs = [
      rp({ fonts: all('Arial') }),
      rp({ fonts: all('Calibri') }),
      rp({ fonts: all('Cambria') }),
    ];
    const seq = async (): Promise<string[]> => {
      const m = await manager(createDocx(), runs, { fonts: testFontResolver() });
      m.fontFor(runs[1]!, cp('x'));
      m.fontFor(runs[0]!, cp('x'));
      m.fontFor(runs[1]!, cp('y'));
      m.fontFor(runs[1]!, 0x2032); // → Liberation Sans, already listed
      expect(m.indexOf(m.fontFor(runs[0]!, cp('x')))).toBe(1);
      const caladea = testFont('Caladea');
      expect(m.indexOf(caladea)).toBe(2); // not handed out yet: appended
      return m.fonts().map((f) => f.postscriptName);
    };
    const a = await seq();
    expect(a).toEqual(['Carlito-Regular', 'LiberationSans', 'Caladea-Regular']);
    expect(await seq()).toEqual(a);
  });

  it('caches resolver answers per request', async () => {
    const { resolver, requests } = recording();
    const runs = [
      rp({ fonts: all('Arial') }),
      rp({ fonts: all('Helvetica') }),
      rp({ fonts: all('Arial') }),
    ];
    const m = await manager(createDocx(), runs, { fonts: resolver });
    for (const r of runs) for (const c of 'Hello') m.fontFor(r, cp(c));
    const keys = requests.map((r) => `${r.family}|${r.bold}|${r.italic}|${r.script}`);
    expect(new Set(keys).size).toBe(keys.length);
    // Helvetica → Liberation Sans through the table, not asked twice.
    expect(keys.filter((k) => k.startsWith('Liberation Sans|false|false|latin')).length).toBe(1);
  });

  it('supports async resolvers, byte results and failing resolvers', async () => {
    const warnings: Warning[] = [];
    const bytes = testFontBytes('Caladea-Regular.ttf.gz');
    const resolvers: FontResolver[] = [
      (req) => {
        if (req.family === 'Boom') throw new Error('resolver exploded');
        return undefined;
      },
      async (req) => (req.family === 'Garbage' ? new Uint8Array([1, 2, 3]) : undefined),
      async (req) => (req.family === 'Brand' && !req.italic ? bytes : undefined),
      testFontResolver(),
    ];
    const brand = rp({ fonts: all('Brand') });
    const brand2 = rp({ fonts: all('Brand'), italic: true });
    const boom = rp({ fonts: all('Boom') });
    const garbage = rp({ fonts: all('Garbage') });
    const m = await manager(createDocx(), [brand, brand2, boom, garbage], {
      fonts: resolvers,
      onWarning: (w) => warnings.push(w),
    });
    const b = m.fontFor(brand, cp('A'));
    expect(b.family).toBe('Caladea');
    expect(m.fontFor(brand2, cp('A'))).toBe(b); // same bytes → same Font
    expect(m.fontFor(boom, cp('A')).family).toBe('Liberation Sans'); // generic sans
    expect(m.fontFor(garbage, cp('A')).family).toBe('Liberation Sans');
    const codes = warnings.map((w) => w.code);
    expect(codes).toContain('FONT_RESOLVER_ERROR');
    expect(codes).toContain('FONT_INVALID');
  });

  it('resolves synchronously what was not wanted before load()', () => {
    const warnings: Warning[] = [];
    const m = createFontManager(createDocx(), {
      fonts: [async () => testFont('Carlito'), testFontResolver()],
      onWarning: (w) => warnings.push(w),
    });
    // Sync resolvers still work; async ones are skipped with a warning.
    expect(m.fontFor(rp({ fonts: all('Liberation Mono') }), cp('A')).family).toBe(
      'Liberation Mono',
    );
    expect(warnings.map((w) => w.code)).toContain('FONT_NOT_LOADED');
  });

  it('works with no fonts option at all', async () => {
    const m = createFontManager(createDocx());
    m.want(rp({ fonts: all('Calibri') }));
    await m.load();
    expect(m.fontFor(rp({ fonts: all('Calibri') }), cp('A')).postscriptName).toBe('Helvetica');
  });
});

// ---- embedded fonts -------------------------------------------------------------------------------

const FONT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font';
const OBFUSCATED = 'application/vnd.openxmlformats-officedocument.obfuscatedFont';
const KEY = '{5D1E4D18-7E3B-4E0A-A3F1-2C9B8E7D6A54}';

/** A docx embedding Caladea under the name "Brand Font" (regular + bold), Word-style. */
function docWithEmbeddedFonts(extra = ''): WordDocument {
  const doc = createDocx();
  const pkg = doc.package;
  const ft = doc.mainPart.relatedByType('fontTable')[0]!;
  pkg.addPart(
    'word/fonts/font1.odttf',
    OBFUSCATED,
    obfuscateFont(testFontBytes('Caladea-Regular.ttf.gz'), KEY),
  );
  pkg.addPart(
    'word/fonts/font2.odttf',
    OBFUSCATED,
    obfuscateFont(testFontBytes('Caladea-Bold.ttf.gz'), KEY),
  );
  const r1 = ft.rels.add(FONT_REL, 'fonts/font1.odttf');
  const r2 = ft.rels.add(FONT_REL, 'fonts/font2.odttf');
  const xml = new TextDecoder()
    .decode(ft.bytes())
    .replace(
      '</w:fonts>',
      `<w:font w:name="Brand Font"><w:family w:val="roman"/>` +
        `<w:embedRegular r:id="${r1.id}" w:fontKey="${KEY}"/>` +
        `<w:embedBold r:id="${r2.id}" w:fontKey="${KEY}" w:subsetted="1"/></w:font>${extra}</w:fonts>`,
    );
  ft.setBytes(new TextEncoder().encode(xml));
  return doc;
}

describe('stand-in vertical metrics', () => {
  it('lists the Microsoft families', () => {
    expect(SUBSTITUTE_METRICS.Cambria).toMatchObject({
      unitsPerEm: 2048,
      ascender: 1946,
      descender: -455,
      lineGap: 0,
      winAscent: 1946,
      winDescent: 455,
    });
    expect(substituteMetricsOf('georgia')?.winAscent).toBe(1878);
    expect(substituteMetricsOf('Carlito')).toBeUndefined();
  });

  it("reports Cambria's vertical metrics on Caladea, keeping Caladea's advances", async () => {
    const cambria = rp({ fonts: all('Cambria') });
    const cambriaBold = rp({ fonts: all('Cambria'), bold: true });
    const caladeaRun = rp({ fonts: all('Caladea') });
    const m = await manager(createDocx(), [cambria, cambriaBold, caladeaRun], {
      fonts: testFontResolver(),
    });
    const f = m.fontFor(cambria, cp('A'));
    const caladea = testFont('Caladea');
    expect(f).toBeInstanceOf(SubstituteFont);
    expect(baseFont(f)).toBe(caladea);
    expect((f as SubstituteFont).substituteFor).toBe('Cambria');
    const k = 1000 / 2048;
    expect(f.metrics.unitsPerEm).toBe(1000);
    expect(f.metrics.winAscent).toBeCloseTo(1946 * k, 6);
    expect(f.metrics.winDescent).toBeCloseTo(455 * k, 6);
    expect(f.metrics.ascender).toBeCloseTo(1946 * k, 6);
    expect(f.metrics.descender).toBeCloseTo(-455 * k, 6);
    expect(f.metrics.lineGap).toBe(0);
    expect(f.metrics.underlinePosition).toBeCloseTo(-125 * k, 6);
    expect(f.metrics.capHeight).toBe(caladea.metrics.capHeight);
    // Line height 1.172 em instead of Caladea's own 1.30 em.
    expect((f.metrics.winAscent + f.metrics.winDescent) / 1000).toBeCloseTo(2401 / 2048, 6);
    expect((caladea.metrics.winAscent + caladea.metrics.winDescent) / 1000).toBe(1.3);
    // Glyphs, advances, kerning and bytes are Caladea's.
    const gA = f.glyphId(cp('A'));
    expect(gA).toBe(caladea.glyphId(cp('A')));
    expect(f.advance(gA)).toBe(caladea.advance(gA));
    expect(f.kerning(gA, f.glyphId(cp('V')))).toBe(-117);
    expect(f.bytes).toBe(caladea.bytes);
    expect(f.postscriptName).toBe('Caladea-Regular');
    // One wrapper per (face, family); the bold face gets its own.
    expect(m.fontFor(cambria, cp('b'))).toBe(f);
    const fb = m.fontFor(cambriaBold, cp('A'));
    expect(baseFont(fb)).toBe(testFont('Caladea', true));
    expect(fb.metrics.winAscent).toBeCloseTo(1946 * k, 6);
    // Asking for Caladea itself gives the plain face.
    expect(m.fontFor(caladeaRun, cp('A'))).toBe(caladea);
    expect(m.fonts()).toEqual([f, fb, caladea]);
  });

  it('leaves substitutes with matching line metrics unwrapped', async () => {
    const runs = [
      rp({ fonts: all('Calibri') }),
      rp({ fonts: all('Arial') }),
      rp({ fonts: all('Courier New') }),
      rp({ fonts: all('Times New Roman') }),
    ];
    const m = await manager(createDocx(), runs, { fonts: testFontResolver() });
    expect(m.fontFor(runs[0]!, cp('A'))).toBe(testFont('Carlito'));
    expect(m.fontFor(runs[1]!, cp('A'))).toBe(testFont('Liberation Sans'));
    expect(m.fontFor(runs[2]!, cp('A'))).toBe(testFont('Liberation Mono'));
    expect(m.fontFor(runs[3]!, cp('A'))).toBe(testFont('Liberation Serif'));
  });

  it('applies to generic and standard-14 fallbacks too', async () => {
    const run = rp({ fonts: all('Cambria') });
    const generic = await manager(createDocx(), [run], {
      fonts: (req) =>
        req.family === 'Liberation Serif' ? testFont('Liberation Serif') : undefined,
    });
    const g = generic.fontFor(run, cp('A'));
    expect(baseFont(g)).toBe(testFont('Liberation Serif'));
    expect(g.metrics.winAscent).toBe(1946);
    expect(g.metrics.winDescent).toBe(455);

    const none = await manager(createDocx(), [run]);
    const s = none.fontFor(run, cp('A'));
    expect(s.postscriptName).toBe('Times-Roman');
    expect(s.outlines).toBe('standard14');
    expect(s.metrics.winAscent).toBeCloseTo(1946 * (1000 / 2048), 6);
    expect(s.glyphId(cp('A'))).toBe(65);
  });
});

describe('embedded fonts', () => {
  it('finds and deobfuscates w:embed* fonts', () => {
    const doc = docWithEmbeddedFonts();
    const faces = embeddedFaces(doc);
    expect(faces.map((f) => [f.name, f.style, f.font.postscriptName, f.part, f.subsetted])).toEqual(
      [
        ['Brand Font', 'regular', 'Caladea-Regular', 'word/fonts/font1.odttf', false],
        ['Brand Font', 'bold', 'Caladea-Bold', 'word/fonts/font2.odttf', true],
      ],
    );
    expect(embeddedFonts(doc).map((f) => f.family)).toEqual(['Caladea', 'Caladea']);
    expect(embeddedFonts(createDocx())).toEqual([]);
    expect(validatePackage(doc).filter((i) => i.severity === 'error')).toEqual([]);
  });

  it('survives a save/load round trip', () => {
    const doc = readDocx(writeDocx(docWithEmbeddedFonts()));
    expect(embeddedFonts(doc).map((f) => f.postscriptName)).toEqual([
      'Caladea-Regular',
      'Caladea-Bold',
    ]);
  });

  it('are preferred by the manager, without any resolver', async () => {
    const doc = docWithEmbeddedFonts();
    const regular = rp({ fonts: all('Brand Font') });
    const bold = rp({ fonts: all('Brand Font'), bold: true });
    const italic = rp({ fonts: all('Brand Font'), italic: true });
    const m = await manager(doc, [regular, bold, italic]);
    expect(m.fontFor(regular, cp('A')).postscriptName).toBe('Caladea-Regular');
    expect(m.fontFor(bold, cp('A')).postscriptName).toBe('Caladea-Bold');
    const fi = m.fontFor(italic, cp('A'));
    expect(fi.postscriptName).toBe('Caladea-Regular');
    expect(m.synthetic(italic, fi).italic).toBe(true);
    // Also matched by the face's own family name.
    expect(m.fontFor(rp({ fonts: all('Caladea') }), cp('A')).postscriptName).toBe(
      'Caladea-Regular',
    );

    const off = await manager(doc, [regular], { useEmbeddedFonts: false });
    expect(off.fontFor(regular, cp('A')).postscriptName).toBe('Times-Roman'); // w:family roman
  });

  it('skips unusable embeddings with warnings', () => {
    const doc = docWithEmbeddedFonts(
      '<w:font w:name="Bad Key"><w:embedRegular r:id="rId1" w:fontKey="{nope}"/></w:font>' +
        '<w:font w:name="No Part"><w:embedRegular r:id="rId999" w:fontKey="{nope}"/></w:font>',
    );
    // rId1 of fontTable.xml.rels is font1.odttf; point a third font at junk bytes.
    const ft = doc.mainPart.relatedByType('fontTable')[0]!;
    doc.package.addPart('word/fonts/font3.odttf', OBFUSCATED, new Uint8Array(64));
    const r3 = ft.rels.add(FONT_REL, 'fonts/font3.odttf');
    const xml = new TextDecoder()
      .decode(ft.bytes())
      .replace(
        '</w:fonts>',
        `<w:font w:name="Junk"><w:embedRegular r:id="${r3.id}" w:fontKey="${KEY}"/></w:font></w:fonts>`,
      );
    ft.setBytes(new TextEncoder().encode(xml));
    const warnings: Warning[] = [];
    const faces = embeddedFaces(doc, (w) => warnings.push(w));
    expect(faces).toHaveLength(2);
    expect(warnings.map((w) => w.code)).toEqual([
      'FONT_EMBED_KEY',
      'FONT_EMBED_MISSING',
      'FONT_EMBED_INVALID',
    ]);
    expect(warnings[0]!.part).toBe('word/fontTable.xml');
  });
});
