#!/usr/bin/env node
// Generates the Unicode property tables of src/layout/text/data/*.ts from the Unicode Character
// Database (UCD) 16.0.0. The tables are checked in; run this only to move to a new Unicode version.
//
//   node scripts/gen-unicode.mjs [cacheDir]
//
// cacheDir is where the UCD files are downloaded (default: <os tmp>/turboword-ucd-16.0.0); files
// already there are reused. The output is deterministic: the same UCD files give byte-identical
// tables.
//
// Why range tables: layout needs a handful of per-code-point properties (Line_Break with the bits
// UAX #14 rules consult, Bidi_Class, Script, bracket pairs, mirroring, default ignorables). Each
// property is stored as the run-length list of its values over U+0000..U+10FFFF, packed into a
// base64url string of 5-bit varints (value-palette index, run length - 1). The runtime decodes a
// table once, lazily, into two typed arrays and binary-searches them (src/layout/text/table.ts).
// That keeps the whole set to a few KB brotli'd, against ~1 MB for flat arrays.

import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { brotliCompressSync, constants as zc } from 'node:zlib';

const VERSION = '16.0.0';
const BASE = `https://www.unicode.org/Public/${VERSION}/ucd/`;
const ROOT = fileURLToPath(new URL('..', import.meta.url));
const OUT = join(ROOT, 'src', 'layout', 'text', 'data');
const CACHE = process.argv[2] ?? join(tmpdir(), `turboword-ucd-${VERSION}`);
const MAX = 0x110000;

const FILES = {
  LineBreak: 'LineBreak.txt',
  EastAsianWidth: 'EastAsianWidth.txt',
  Scripts: 'Scripts.txt',
  DerivedBidiClass: 'extracted/DerivedBidiClass.txt',
  BidiBrackets: 'BidiBrackets.txt',
  BidiMirroring: 'BidiMirroring.txt',
  EmojiData: 'emoji/emoji-data.txt',
  // Also needed: General_Category (LB1 resolves SA by Mn/Mc; LB15a/b/19 need Pi/Pf; LB30b needs
  // Cn), the short script codes, and Default_Ignorable_Code_Point for the shaper.
  DerivedGeneralCategory: 'extracted/DerivedGeneralCategory.txt',
  PropertyValueAliases: 'PropertyValueAliases.txt',
  DerivedCoreProperties: 'DerivedCoreProperties.txt',
};

// ---- download -----------------------------------------------------------------------------------

async function load(name) {
  const rel = FILES[name];
  const file = join(CACHE, rel.split('/').pop());
  if (!existsSync(file)) {
    mkdirSync(CACHE, { recursive: true });
    const res = await fetch(BASE + rel);
    if (!res.ok) throw new Error(`GET ${BASE + rel}: ${res.status}`);
    writeFileSync(file, Buffer.from(await res.arrayBuffer()));
  }
  const text = readFileSync(file, 'utf8');
  // Every UCD file names its version in the header (emoji-data.txt as "Emoji Version 16.0").
  const head = text.slice(0, 600);
  if (!head.includes(VERSION) && !head.includes(`Emoji Version ${VERSION.slice(0, -2)}`))
    throw new Error(`${file}: not UCD ${VERSION}`);
  return text;
}

// ---- parsing --------------------------------------------------------------------------------------

/**
 * Calls `fn(start, end, fields)` for every data line of a UCD file, after first applying its
 * `# @missing:` defaults (in file order, so specific ranges override the catch-all one).
 */
function eachRange(text, fn) {
  const lines = text.split('\n');
  for (const line of lines) {
    const m = /^# @missing: ([0-9A-F]+)\.\.([0-9A-F]+); (.*)$/.exec(line);
    if (m) fn(Number.parseInt(m[1], 16), Number.parseInt(m[2], 16), splitFields(m[3]));
  }
  for (const raw of lines) {
    const line = raw.replace(/#.*/, '').trim();
    if (!line) continue;
    const [range, ...fields] = line.split(';').map((s) => s.trim());
    const [a, b] = range.split('..');
    const start = Number.parseInt(a, 16);
    fn(start, b === undefined ? start : Number.parseInt(b, 16), fields);
  }
}

const splitFields = (s) => s.split(';').map((x) => x.trim());

function fill(arr, start, end, v) {
  for (let i = start; i <= end; i++) arr[i] = v;
}

// ---- encoding ---------------------------------------------------------------------------------------

const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';

/** Big-endian 5-bit groups; digits 32..63 mean "more groups follow". */
function varint(n) {
  if (!Number.isInteger(n) || n < 0) throw new Error(`varint(${n})`);
  const groups = [n & 31];
  for (n = Math.floor(n / 32); n > 0; n = Math.floor(n / 32)) groups.unshift(32 + (n & 31));
  return groups.map((g) => ALPHABET[g]).join('');
}

const zigzag = (n) => (n >= 0 ? n * 2 : -n * 2 - 1);

/**
 * Run-length encodes `values` (one per code point): palette (`first`, then the other values by
 * run frequency) + varint string.
 */
function encodeRuns(values, first = []) {
  const runs = [];
  let start = 0;
  for (let cp = 1; cp <= MAX; cp++) {
    if (cp === MAX || values[cp] !== values[start]) {
      runs.push([values[start], cp - start]);
      start = cp;
    }
  }
  const freq = new Map();
  for (const [v] of runs) freq.set(v, (freq.get(v) ?? 0) + 1);
  const rest = [...freq.keys()].filter((v) => !first.includes(v));
  const palette = [...first, ...rest.sort((a, b) => freq.get(b) - freq.get(a) || a - b)];
  const index = new Map(palette.map((v, i) => [v, i]));
  let data = '';
  for (const [v, len] of runs) data += varint(index.get(v)) + varint(len - 1);
  return { palette, data, runs: runs.length };
}

// ---- output -------------------------------------------------------------------------------------------

const LICENSE = `// Derived from the Unicode Character Database ${VERSION} (https://www.unicode.org/Public/${VERSION}/ucd/)
// by scripts/gen-unicode.mjs. GENERATED FILE: do not edit; re-run the script instead.
//
// UNICODE LICENSE V3
//
// COPYRIGHT AND PERMISSION NOTICE
//
// Copyright © 1991-2024 Unicode, Inc.
//
// NOTICE TO USER: Carefully read the following legal agreement. BY DOWNLOADING, INSTALLING,
// COPYING OR OTHERWISE USING DATA FILES, AND/OR SOFTWARE, YOU UNEQUIVOCALLY ACCEPT, AND AGREE TO
// BE BOUND BY, ALL OF THE TERMS AND CONDITIONS OF THIS AGREEMENT. IF YOU DO NOT AGREE, DO NOT
// DOWNLOAD, INSTALL, COPY, DISTRIBUTE OR USE THE DATA FILES OR SOFTWARE.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of data files and
// any associated documentation (the "Data Files") or software and any associated documentation
// (the "Software") to deal in the Data Files or Software without restriction, including without
// limitation the rights to use, copy, modify, merge, publish, distribute, and/or sell copies of
// the Data Files or Software, and to permit persons to whom the Data Files or Software are
// furnished to do so, provided that either (a) this copyright and permission notice appear with
// all copies of the Data Files or Software, or (b) this copyright and permission notice appear in
// associated Documentation.
//
// THE DATA FILES AND SOFTWARE ARE PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
// PARTICULAR PURPOSE AND NONINFRINGEMENT OF THIRD PARTY RIGHTS.
//
// IN NO EVENT SHALL THE COPYRIGHT HOLDER OR HOLDERS INCLUDED IN THIS NOTICE BE LIABLE FOR ANY
// CLAIM, OR ANY SPECIAL INDIRECT OR CONSEQUENTIAL DAMAGES, OR ANY DAMAGES WHATSOEVER RESULTING
// FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER
// TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THE DATA FILES
// OR SOFTWARE.
//
// Except as contained in this notice, the name of a copyright holder shall not be used in
// advertising or otherwise to promote the sale, use or other dealings in these Data Files or
// Software without prior written authorization of the copyright holder.
`;

const sizes = [];

/** Writes one generated module; `consts` = [[name, value (string), doc]]. */
function emit(file, what, consts) {
  let src = `${LICENSE}//\n// ${what}\n`;
  let payload = '';
  for (const [name, value, doc] of consts) {
    if (!/^[\w -]*$/.test(value)) throw new Error(`${name}: unexpected characters`);
    payload += `'${value}';`;
    src += `\n/** ${doc} */\nexport const ${name}: string =\n  '${value}';\n`;
  }
  writeFileSync(join(OUT, file), src);
  const br = brotliCompressSync(Buffer.from(payload), {
    params: { [zc.BROTLI_PARAM_QUALITY]: 11 },
  }).length;
  sizes.push([file, payload.length, br]);
}

/** A list of small non-negative integers as one varint string. */
const varints = (list) => list.map(varint).join('');

// ---- properties -------------------------------------------------------------------------------------

/** Resolved Line_Break classes (after LB1), in the order src/layout/text/linebreak.ts numbers them. */
const LB_CLASSES = [
  'BK',
  'CR',
  'LF',
  'NL',
  'SP',
  'ZW',
  'ZWJ',
  'CM',
  'WJ',
  'GL',
  'BA',
  'HY',
  'CL',
  'CP',
  'EX',
  'SY',
  'OP',
  'QU',
  'IS',
  'NS',
  'B2',
  'CB',
  'BB',
  'HL',
  'AL',
  'NU',
  'PR',
  'PO',
  'ID',
  'EB',
  'EM',
  'IN',
  'JL',
  'JV',
  'JT',
  'H2',
  'H3',
  'RI',
  'AK',
  'AP',
  'AS',
  'VF',
  'VI',
];
const LB_EA = 0x40; // East_Asian_Width F/W/H ($EastAsian in UAX #14)
const LB_PI = 0x80; // QU with General_Category Pi
const LB_PF = 0x100; // QU with General_Category Pf
const LB_EPCN = 0x200; // Extended_Pictographic and unassigned (LB30b)

/** Bidi_Class values in the order src/layout/text/bidi.ts numbers them. */
const BIDI_CLASSES = [
  'L',
  'R',
  'AL',
  'EN',
  'ES',
  'ET',
  'AN',
  'CS',
  'NSM',
  'BN',
  'B',
  'S',
  'WS',
  'ON',
  'LRE',
  'LRO',
  'RLE',
  'RLO',
  'PDF',
  'LRI',
  'RLI',
  'FSI',
  'PDI',
];

async function main() {
  mkdirSync(OUT, { recursive: true });
  const aliases = await load('PropertyValueAliases');
  const alias = { bc: new Map(), sc: new Map() };
  for (const raw of aliases.split('\n')) {
    const line = raw.replace(/#.*/, '').trim();
    if (!line) continue;
    const [prop, short, long, ...more] = line.split(';').map((s) => s.trim());
    if (prop in alias) for (const n of [short, long, ...more]) alias[prop].set(n, short);
  }

  // General_Category (default Cn)
  const gcNames = ['Cn'];
  const gc = new Uint8Array(MAX);
  eachRange(await load('DerivedGeneralCategory'), (s, e, [v]) => {
    if (v === 'Unassigned') v = 'Cn';
    let i = gcNames.indexOf(v);
    if (i < 0) i = gcNames.push(v) - 1;
    fill(gc, s, e, i);
  });
  const gcIs = (cp, ...names) => names.includes(gcNames[gc[cp]]);

  // East_Asian_Width: just the $EastAsian bit
  const ea = new Uint8Array(MAX);
  eachRange(await load('EastAsianWidth'), (s, e, [v]) =>
    fill(ea, s, e, v === 'F' || v === 'W' || v === 'H' ? 1 : 0),
  );

  const extPict = new Uint8Array(MAX);
  eachRange(await load('EmojiData'), (s, e, [v]) => {
    if (v === 'Extended_Pictographic') fill(extPict, s, e, 1);
  });

  // Line_Break, resolved per LB1, plus the bits the pair rules consult
  const lbRaw = new Array(MAX).fill('XX');
  eachRange(await load('LineBreak'), (s, e, [v]) => fill(lbRaw, s, e, v));
  const lb = new Uint16Array(MAX);
  for (let cp = 0; cp < MAX; cp++) {
    let c = lbRaw[cp];
    if (c === 'AI' || c === 'SG' || c === 'XX') c = 'AL';
    else if (c === 'SA') c = gcIs(cp, 'Mn', 'Mc') ? 'CM' : 'AL';
    else if (c === 'CJ') c = 'NS';
    const cls = LB_CLASSES.indexOf(c);
    if (cls < 0) throw new Error(`unknown Line_Break ${c} at ${cp.toString(16)}`);
    let v = cls | (ea[cp] ? LB_EA : 0);
    if (c === 'QU' && gcIs(cp, 'Pi')) v |= LB_PI;
    if (c === 'QU' && gcIs(cp, 'Pf')) v |= LB_PF;
    if (extPict[cp] && gcIs(cp, 'Cn')) v |= LB_EPCN;
    lb[cp] = v;
  }
  const lbEnc = encodeRuns(lb);
  emit(
    'linebreak.ts',
    'Line_Break (resolved per UAX #14 LB1) with East_Asian_Width and the Pi/Pf/Cn bits.',
    [
      [
        'LB_CLASSES',
        LB_CLASSES.join(' '),
        'Resolved Line_Break class names (space-separated); a value & 63 indexes them.',
      ],
      [
        'LB_PALETTE',
        varints(lbEnc.palette),
        'Value palette (varints): class | 0x40 East Asian | 0x80 Pi | 0x100 Pf | 0x200 Extended_Pictographic & Cn.',
      ],
      [
        'LB_RUNS',
        lbEnc.data,
        'Runs over U+0000..U+10FFFF: varint palette index, varint length - 1.',
      ],
    ],
  );

  // Bidi_Class
  const bidi = new Uint8Array(MAX);
  eachRange(await load('DerivedBidiClass'), (s, e, [v]) => {
    const i = BIDI_CLASSES.indexOf(alias.bc.get(v) ?? v);
    if (i < 0) throw new Error(`unknown Bidi_Class ${v}`);
    fill(bidi, s, e, i);
  });
  const bidiEnc = encodeRuns(bidi);

  // Bidi_Paired_Bracket (+ type): opening brackets with their closing pair
  const brackets = [];
  eachRange(await load('BidiBrackets'), (s, _e, [pair, type]) => {
    if (type === 'o') brackets.push([s, Number.parseInt(pair, 16)]);
  });
  let prev = 0;
  let brData = '';
  for (const [open, close] of brackets) {
    brData += varint(open - prev) + varint(zigzag(close - open));
    prev = open;
  }

  // Bidi_Mirroring_Glyph
  const mirrors = [];
  eachRange(await load('BidiMirroring'), (s, _e, [m]) => {
    if (/^[0-9A-F]+$/.test(m)) mirrors.push([s, Number.parseInt(m, 16)]); // skip @missing <none>
  });
  mirrors.sort((a, b) => a[0] - b[0]);
  prev = 0;
  let mirData = '';
  for (const [cp, m] of mirrors) {
    mirData += varint(cp - prev) + varint(zigzag(m - cp));
    prev = cp;
  }
  emit('bidi.ts', 'Bidi_Class, Bidi_Paired_Bracket (opening brackets) and Bidi_Mirroring_Glyph.', [
    [
      'BIDI_CLASSES',
      BIDI_CLASSES.join(' '),
      'Bidi_Class names (space-separated); a value indexes them.',
    ],
    ['BIDI_PALETTE', varints(bidiEnc.palette), 'Value palette (varints of Bidi_Class indexes).'],
    [
      'BIDI_RUNS',
      bidiEnc.data,
      'Runs over U+0000..U+10FFFF: varint palette index, varint length - 1.',
    ],
    [
      'BIDI_BRACKETS',
      brData,
      'Opening paired brackets: varint delta from the previous one, varint zigzag(closing - opening).',
    ],
    [
      'BIDI_MIRRORS',
      mirData,
      'Bidi_Mirroring_Glyph: varint delta from the previous code point, varint zigzag(mirror - code point).',
    ],
  ]);

  // Script (ISO 15924 codes)
  const scNames = ['Zzzz'];
  const sc = new Uint8Array(MAX);
  eachRange(await load('Scripts'), (s, e, [v]) => {
    const code = alias.sc.get(v);
    if (!code) throw new Error(`unknown script ${v}`);
    let i = scNames.indexOf(code);
    if (i < 0) i = scNames.push(code) - 1;
    fill(sc, s, e, i);
  });
  // Unknown, Common and Inherited first (the runtime hard-codes 0, 1, 2), then by frequency;
  // names are stored in palette order so the palette itself is the identity.
  const specials = ['Zzzz', 'Zyyy', 'Zinh'].map((c) => scNames.indexOf(c));
  const scEnc = encodeRuns(sc, specials);
  const scOrdered = scEnc.palette.map((i) => scNames[i]);
  const remap = new Uint8Array(scNames.length);
  scEnc.palette.forEach((old, i) => {
    remap[old] = i;
  });
  const scEnc2 = encodeRuns(
    sc.map((v) => remap[v]),
    [0, 1, 2],
  );
  if (scEnc2.palette.some((v, i) => v !== i)) throw new Error('script palette not identity');

  // Default_Ignorable_Code_Point ranges
  const ignorable = new Uint8Array(MAX);
  eachRange(await load('DerivedCoreProperties'), (s, e, [v]) => {
    if (v === 'Default_Ignorable_Code_Point') fill(ignorable, s, e, 1);
  });
  const ignEnc = encodeRuns(ignorable);
  if (ignEnc.palette[0] === undefined) throw new Error('no ignorables');

  emit('scripts.ts', 'Script (ISO 15924 codes) and Default_Ignorable_Code_Point.', [
    [
      'SCRIPT_CODES',
      scOrdered.join(''),
      'ISO 15924 codes, 4 letters each; run value i is codes.slice(4i, 4i + 4) (Zzzz Unknown, Zyyy Common, Zinh Inherited).',
    ],
    [
      'SCRIPT_RUNS',
      scEnc2.data,
      'Runs over U+0000..U+10FFFF: varint script index, varint length - 1.',
    ],
    [
      'IGNORABLE_PALETTE',
      varints(ignEnc.palette),
      'Value palette (varints; 1 = Default_Ignorable_Code_Point).',
    ],
    [
      'IGNORABLE_RUNS',
      ignEnc.data,
      'Runs over U+0000..U+10FFFF: varint palette index, varint length - 1.',
    ],
  ]);

  // Same layout as `biome format` (one string literal per export), but run it to be sure.
  const biome = join(
    ROOT,
    'node_modules',
    '.bin',
    process.platform === 'win32' ? 'biome.cmd' : 'biome',
  );
  if (existsSync(biome)) execFileSync(biome, ['format', '--write', OUT], { stdio: 'inherit' });

  let total = 0;
  for (const [file, raw, br] of sizes) {
    total += br;
    console.log(
      `${file.padEnd(14)} ${String(raw).padStart(7)} B raw  ${String(br).padStart(6)} B brotli`,
    );
  }
  console.log(
    `runs: lb ${lbEnc.runs}, bidi ${bidiEnc.runs}, script ${scEnc2.runs}, ignorable ${ignEnc.runs}; ` +
      `brackets ${brackets.length}, mirrors ${mirrors.length}; total ${total} B brotli`,
  );
}

await main();
