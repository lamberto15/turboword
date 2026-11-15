#!/usr/bin/env node
// Records the layout oracle data for test/layout-oracle.test.ts.
//
//   node test/fixtures/layout/record-libreoffice.mjs
//
// 1. corpus.json — every Word-saved .docx under test/fixtures (deduplicated by content) whose
//    docProps/app.xml records Word's page count (`<Pages>`), minus files whose app.xml is known
//    to be stale (edited by another tool after Word saved it).
// 2. libreoffice.json — the same files converted to PDF by LibreOffice (headless, throwaway
//    profile), read back with `pdftotext -bbox-layout`: the page count and, per page, one hash
//    per text row (words grouped by baseline, joined in x order, letters and digits only;
//    FNV-1a 32-bit). The secondary oracle for line breaks: LibreOffice is not Word, but it uses
//    the same metric-compatible fonts as the tests (Carlito, Caladea, Liberation), so where both
//    engines are right their rows agree.
//
// Needs LibreOffice (SOFFICE=/path/to/soffice, else found on PATH / in the usual places) and
// poppler's pdftotext. Plain Node (>= 22) plus fflate (a dependency of the repo).

import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, relative, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';

const HERE = fileURLToPath(new URL('.', import.meta.url));
const FIXTURES = join(HERE, '..');
/** app.xml `<Pages>` not written by Word for the current content. */
const STALE = new Set([
  // docx4j rewrote the body (one line and a picture) after Word saved "2 pages"
  'forms/docx4j/picture.docx',
]);

function walk(dir, out = []) {
  for (const n of readdirSync(dir).sort()) {
    const p = join(dir, n);
    if (statSync(p).isDirectory()) walk(p, out);
    else if (/\.docx$/i.test(n)) out.push(p);
  }
  return out;
}

function corpus() {
  const seen = new Set();
  const out = [];
  for (const p of walk(FIXTURES)) {
    const name = relative(FIXTURES, p).split(sep).join('/');
    const bytes = readFileSync(p);
    const sha1 = createHash('sha1').update(bytes).digest('hex');
    if (seen.has(sha1)) continue;
    let app = '';
    try {
      const z = unzipSync(new Uint8Array(bytes), { filter: (f) => f.name === 'docProps/app.xml' });
      app = z['docProps/app.xml'] ? strFromU8(z['docProps/app.xml']) : '';
    } catch {
      continue; // encrypted or not a zip
    }
    const pages = /<Pages>(\d+)<\/Pages>/.exec(app)?.[1];
    const application = /<Application>([^<]*)<\/Application>/.exec(app)?.[1] ?? '';
    if (pages === undefined || !/Microsoft.*Word/.test(application) || STALE.has(name)) continue;
    seen.add(sha1);
    out.push({ name, sha1, pages: Number(pages) });
  }
  return out;
}

function soffice() {
  if (process.env.SOFFICE) return process.env.SOFFICE;
  for (const p of [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    join(process.env.HOME ?? '', 'Applications/LibreOffice.app/Contents/MacOS/soffice'),
    '/usr/bin/soffice',
    '/usr/lib/libreoffice/program/soffice',
  ]) {
    if (existsSync(p)) return p;
  }
  return 'soffice';
}

/** FNV-1a 32-bit of a string's UTF-16 code units, 8 hex digits. */
export function fnv(s) {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, '0');
}

const decode = (s) =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&');

/** Row text as compared: letters and digits only (bullets, quotes, dashes, spaces dropped). */
export function normalize(s) {
  return s.normalize('NFC').replace(/[^\p{L}\p{N}]/gu, '');
}

/** Rows of each page: words grouped by baseline (yMax within 2pt), x order, normalized. */
function rows(bboxHtml) {
  const pages = [];
  let words = null;
  for (const line of bboxHtml.split('\n')) {
    if (line.includes('<page ')) {
      words = [];
      pages.push(words);
    }
    const m = /<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="([\d.]+)">(.*)<\/word>/.exec(line);
    if (m && words) words.push({ x: Number(m[1]), y: Number(m[3]), t: decode(m[4]) });
  }
  return pages.map((ws) => {
    ws.sort((a, b) => a.y - b.y || a.x - b.x);
    const out = [];
    let cur = null;
    for (const w of ws) {
      if (cur === null || Math.abs(w.y - cur.y) > 2) {
        cur = { y: w.y, ws: [] };
        out.push(cur);
      }
      cur.ws.push(w);
    }
    return out
      .map((r) => normalize(r.ws.sort((a, b) => a.x - b.x).map((w) => w.t).join('')))
      .filter((t) => t !== '')
      .map(fnv);
  });
}

function main() {
  const list = corpus();
  writeFileSync(join(HERE, 'corpus.json'), `${JSON.stringify(list, null, 1)}\n`);
  const work = mkdtempSync(join(tmpdir(), 'turboword-oracle-'));
  const inDir = join(work, 'in');
  const outDir = join(work, 'out');
  mkdirSync(inDir);
  mkdirSync(outDir);
  list.forEach((f, i) => {
    copyFileSync(join(FIXTURES, f.name), join(inDir, `${i}.docx`));
  });
  const files = list.map((_, i) => join(inDir, `${i}.docx`));
  const r = spawnSync(
    soffice(),
    [
      `-env:UserInstallation=${pathToFileURL(join(work, 'profile')).href}`,
      '--headless',
      '--norestore',
      '--convert-to',
      'pdf',
      '--outdir',
      outDir,
      ...files,
    ],
    { encoding: 'utf8', timeout: 1_800_000 },
  );
  if (r.error) throw r.error;
  const result = {};
  list.forEach((f, i) => {
    const pdf = join(outDir, `${i}.pdf`);
    if (!existsSync(pdf)) {
      console.warn(`no PDF for ${f.name}`);
      return;
    }
    const t = spawnSync('pdftotext', ['-bbox-layout', pdf, '-'], { encoding: 'utf8' });
    if (t.status !== 0) throw new Error(`pdftotext failed on ${f.name}: ${t.stderr}`);
    const pages = rows(t.stdout);
    result[f.name] = { pages: pages.length, rows: pages };
  });
  writeFileSync(join(HERE, 'libreoffice.json'), `${JSON.stringify(result)}\n`);
  rmSync(work, { recursive: true, force: true });
  const agree = list.filter((f) => result[f.name]?.pages === f.pages).length;
  console.log(`${list.length} files; LibreOffice page count = Word's for ${agree}`);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
