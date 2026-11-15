// Generates src/layout/fonts/standard14-data.ts from pdf.js (Apache-2.0):
//   src/core/metrics.js   glyph widths of the 14 standard PDF fonts (Adobe Core14 AFMs)
//   src/core/encodings.js WinAnsiEncoding, SymbolSetEncoding, ZapfDingbatsEncoding
//   src/core/glyphlist.js Adobe Glyph List + ZapfDingbats glyph list (name -> Unicode)
//
// Usage: node test/fixtures/fonts/gen-standard14.mjs <dir with the three pdf.js files> [afm dir]
//        npx biome format --write src/layout/fonts/standard14-data.ts
// When an AFM directory is given (the Adobe Core14 AFMs, e.g. from Apache PDFBox), every width is
// cross-checked against the AFM `C code ; WX width` records and the script fails on a mismatch.
//
// Dev-only: parses the pdf.js sources textually (no eval) and writes a compact TS module.

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const [srcDir, afmDir] = process.argv.slice(2);
if (!srcDir) throw new Error('usage: gen-standard14.mjs <pdfjs src/core dir> [afm dir]');
const read = (f) => readFileSync(join(srcDir, f), 'utf8');

// ---- metrics.js: t.Name = getLookupTableFactory(function (t) { t.glyph = 123; ... }) ----------
const metricsSrc = read('metrics.js');
const widths = {};
{
  const blockRe = /^  t(?:\.([A-Za-z]+)|\["([A-Za-z-]+)"\]) = (?:(\d+);|getLookupTableFactory)/gm;
  const marks = [];
  for (let m; (m = blockRe.exec(metricsSrc)); ) {
    marks.push({ name: m[1] ?? m[2], fixed: m[3] ? Number(m[3]) : undefined, at: m.index });
  }
  const basic = metricsSrc.indexOf('getFontBasicMetrics');
  for (let i = 0; i < marks.length; i++) {
    const mk = marks[i];
    if (mk.at > basic) break;
    if (mk.fixed !== undefined) {
      widths[mk.name] = mk.fixed;
      continue;
    }
    const end = i + 1 < marks.length ? marks[i + 1].at : basic;
    const body = metricsSrc.slice(mk.at, end);
    const table = {};
    for (const g of body.matchAll(/t(?:\.([A-Za-z0-9_]+)|\["([^"]+)"\]) = (\d+);/g)) {
      table[g[1] ?? g[2]] = Number(g[3]);
    }
    widths[mk.name] = table;
  }
}

// ---- encodings.js ------------------------------------------------------------------------------
const encSrc = read('encodings.js');
function encoding(name) {
  const start = encSrc.indexOf(`const ${name} = [`);
  const end = encSrc.indexOf('];', start);
  return [...encSrc.slice(start, end).matchAll(/"([^"]*)"/g)].map((m) => m[1]);
}
const winAnsi = encoding('WinAnsiEncoding');
const symbolEnc = encoding('SymbolSetEncoding');
const dingbatsEnc = encoding('ZapfDingbatsEncoding');
for (const [n, e] of [
  ['WinAnsi', winAnsi],
  ['Symbol', symbolEnc],
  ['ZapfDingbats', dingbatsEnc],
]) {
  if (e.length !== 256) throw new Error(`${n} encoding has ${e.length} entries`);
}

// ---- glyphlist.js ------------------------------------------------------------------------------
const glSrc = read('glyphlist.js');
function glyphList(fn) {
  const start = glSrc.indexOf(`const ${fn} = getLookupTableFactory`);
  const end = glSrc.indexOf('});', start);
  const out = new Map();
  for (const g of glSrc.slice(start, end).matchAll(/t(?:\.([A-Za-z0-9_]+)|\["([^"]+)"\]) = 0x([0-9a-fA-F]+);/g)) {
    out.set(g[1] ?? g[2], Number.parseInt(g[3], 16));
  }
  return out;
}
const agl = glyphList('getGlyphsUnicode');
const dingbatsGl = glyphList('getDingbatsGlyphsUnicode');

// ---- AFM cross-check ---------------------------------------------------------------------------
function afmCodes(font) {
  const out = new Map();
  for (const line of readFileSync(join(afmDir, `${font}.afm`), 'latin1').split(/\r?\n/)) {
    const m = /^C (-?\d+) ; WX (\d+) ; N ([^ ;]+)/.exec(line);
    if (m) out.set(m[3], Number(m[2]));
  }
  return out;
}

// ---- encode ------------------------------------------------------------------------------------
const FONTS = [
  'Courier', 'Courier-Bold', 'Courier-Oblique', 'Courier-BoldOblique',
  'Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique', 'Helvetica-BoldOblique',
  'Times-Roman', 'Times-Bold', 'Times-Italic', 'Times-BoldItalic',
  'Symbol', 'ZapfDingbats',
];
const b36 = (n) => {
  if (n < 0 || n >= 36 * 36) throw new Error(`width ${n} out of range`);
  return n.toString(36).padStart(2, '0');
};
const out = {};
for (const font of FONTS) {
  const enc = font === 'Symbol' ? symbolEnc : font === 'ZapfDingbats' ? dingbatsEnc : winAnsi;
  const table = widths[font];
  const afm = afmDir ? afmCodes(font) : undefined;
  let s = '';
  for (let code = 32; code < 256; code++) {
    const name = enc[code];
    let w = 0;
    if (name) {
      w = typeof table === 'number' ? table : (table[name] ?? 0);
      if (w === 0) throw new Error(`${font}: no width for ${name}`);
      if (afm && afm.get(name) !== w) {
        throw new Error(`${font}: ${name} pdf.js ${w} vs AFM ${afm.get(name)}`);
      }
    }
    s += b36(w);
  }
  out[font] = s;
}

// Unicode -> code maps for the two symbolic fonts (their built-in encodings).
function unicodeMap(enc, list) {
  const pairs = [];
  const seen = new Set();
  for (let code = 32; code < 256; code++) {
    const name = enc[code];
    if (!name) continue;
    const u = list.get(name);
    if (u === undefined) throw new Error(`no Unicode for ${name}`);
    if (seen.has(u)) continue;
    seen.add(u);
    pairs.push(u, code);
  }
  return pairs;
}
const symbolMap = unicodeMap(symbolEnc, agl);
const dingbatsMap = unicodeMap(dingbatsEnc, dingbatsGl);

const here = fileURLToPath(new URL('.', import.meta.url));
const target = join(here, '../../../src/layout/fonts/standard14-data.ts');
const chunk = (s, n) => s.match(new RegExp(`.{1,${n}}`, 'g'));
const lines = [];
lines.push(
  '// GENERATED by test/fixtures/fonts/gen-standard14.mjs - do not edit by hand.',
  '//',
  '// Glyph widths of the 14 standard PDF fonts, from the Adobe Core14 AFM files as tabulated by',
  '// pdf.js (src/core/metrics.js, encodings.js, glyphlist.js):',
  '//   Copyright 2012 Mozilla Foundation. Licensed under the Apache License, Version 2.0',
  '//   (http://www.apache.org/licenses/LICENSE-2.0). Modified: re-encoded per character code.',
  '// The AFM data itself: Copyright (c) 1985-1997 Adobe Systems Incorporated; Adobe permits use,',
  '// copying and distribution of the Core14 AFM files.',
  '//',
  '// WIDTHS: per font, codes 32..255 of its PDF encoding (WinAnsiEncoding for Courier, Helvetica and',
  '// Times; the built-in encoding for Symbol and ZapfDingbats), 2 base-36 digits each, 00 = unused.',
  '// *_UNICODE: flat [unicode, code, ...] pairs for the built-in encodings of the symbolic fonts.',
  '',
  '/** Per-font widths, see the header. */',
  'export const WIDTHS: Readonly<Record<string, string>> = {',
);
for (const font of FONTS) {
  const parts = chunk(out[font], 88);
  lines.push(/^[A-Za-z]+$/.test(font) ? `  ${font}:` : `  '${font}':`);
  parts.forEach((p, i) => lines.push(`    '${p}'${i === parts.length - 1 ? ',' : ' +'}`));
}
lines.push('};', '');
for (const [name, pairs] of [
  ['SYMBOL_UNICODE', symbolMap],
  ['DINGBATS_UNICODE', dingbatsMap],
]) {
  lines.push(`/** Unicode -> code pairs of the built-in encoding. */`);
  lines.push(`export const ${name}: readonly number[] = [`);
  for (let i = 0; i < pairs.length; i += 12) {
    lines.push(`  ${pairs.slice(i, i + 12).map((n) => `0x${n.toString(16)}`).join(', ')},`);
  }
  lines.push('];', '');
}
writeFileSync(target, lines.join('\n'));
console.log(`wrote ${target}`);
