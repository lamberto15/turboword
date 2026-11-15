#!/usr/bin/env node
// Sanity check with LibreOffice: converts each .docx to PDF headless and counts the pages.
//
//   node scripts/libreoffice-check.mjs [--keep] [files or dirs...]   (default: test/fixtures)
//
// - A conversion failure (LibreOffice could not open the file, or produced no PDF) exits 1: it
//   usually means we wrote a broken package.
// - When test/fixtures/expected.json lists `pages` for a file with the same basename, the page
//   count is compared, but a mismatch is only a warning: fonts (and so pagination) differ by
//   machine unless the metric-compatible fonts (Carlito, Caladea, Liberation) are installed.
// - LibreOffice runs with a throwaway profile, so it works while another instance is open and
//   never touches the user's settings. `--keep` keeps the PDFs (the path is printed).
// - Set SOFFICE=/path/to/soffice when LibreOffice is not found automatically.
// Plain Node (>= 22), no dependencies.

import { spawnSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, delimiter, extname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const EXPECTED = join(ROOT, 'test', 'fixtures', 'expected.json');
const TIMEOUT_MS = 180_000;

/** First existing executable named `names` on PATH (with PATHEXT extensions on Windows). */
function onPath(names) {
  const exts =
    process.platform === 'win32' ? (process.env.PATHEXT ?? '.EXE;.CMD').split(';') : [''];
  for (const dir of (process.env.PATH ?? '').split(delimiter)) {
    if (!dir) continue;
    for (const name of names) {
      for (const ext of exts) {
        const p = join(dir, name + ext.toLowerCase());
        if (existsSync(p) && statSync(p).isFile()) return p;
      }
    }
  }
  return undefined;
}

/** Program-folder installs (globbing `C:\Program Files\LibreOffice*` and `/opt/libreoffice*`). */
function installed() {
  const out = [];
  if (process.platform === 'darwin') {
    out.push('/Applications/LibreOffice.app/Contents/MacOS/soffice');
    out.push(join(process.env.HOME ?? '', 'Applications/LibreOffice.app/Contents/MacOS/soffice'));
  } else if (process.platform === 'win32') {
    for (const base of [process.env.ProgramFiles, process.env['ProgramFiles(x86)']]) {
      if (!base || !existsSync(base)) continue;
      for (const d of readdirSync(base)) {
        if (/^libreoffice/i.test(d)) out.push(join(base, d, 'program', 'soffice.exe'));
      }
    }
  } else {
    out.push('/usr/bin/soffice', '/usr/lib/libreoffice/program/soffice', '/snap/bin/libreoffice');
    if (existsSync('/opt')) {
      for (const d of readdirSync('/opt')) {
        if (/^libreoffice/i.test(d)) out.push(join('/opt', d, 'program', 'soffice'));
      }
    }
  }
  return out.filter((p) => existsSync(p));
}

function findSoffice() {
  if (process.env.SOFFICE) return process.env.SOFFICE;
  // macOS: prefer the app bundle (the Homebrew `soffice` shim points at it anyway).
  return installed()[0] ?? onPath(['soffice', 'libreoffice']);
}

/** `.docx` files named by the arguments (directories are searched recursively). */
function collect(args) {
  const out = [];
  const walk = (p) => {
    const st = statSync(p);
    if (st.isDirectory()) {
      for (const name of readdirSync(p).sort()) walk(join(p, name));
    } else if (extname(p).toLowerCase() === '.docx' && !basename(p).startsWith('~$')) {
      out.push(p);
    }
  };
  for (const a of args) {
    if (!existsSync(a)) throw new Error(`not found: ${a}`);
    walk(resolve(a));
  }
  return out;
}

/** basename -> expected page count, from test/fixtures/expected.json (object or array form). */
function expectedPages() {
  const map = new Map();
  if (!existsSync(EXPECTED)) return map;
  const data = JSON.parse(readFileSync(EXPECTED, 'utf8'));
  const add = (name, entry) => {
    const pages = typeof entry === 'number' ? entry : entry?.pages;
    if (typeof name === 'string' && typeof pages === 'number') map.set(basename(name), pages);
  };
  if (Array.isArray(data)) {
    for (const e of data) add(e?.file ?? e?.name ?? e?.path, e);
  } else {
    const entries = data?.files && typeof data.files === 'object' ? data.files : data;
    for (const [name, e] of Object.entries(entries)) add(name, e);
  }
  return map;
}

/** Page count of a PDF: `qpdf --show-npages`, else counting `/Type /Page` objects. */
function countPages(pdf) {
  const q = spawnSync('qpdf', ['--show-npages', pdf], { encoding: 'utf8' });
  if (q.status === 0 && /^\d+\s*$/.test(q.stdout)) return Number(q.stdout.trim());
  const text = readFileSync(pdf).toString('latin1');
  return (text.match(/\/Type\s*\/Page(?![a-zA-Z])/g) ?? []).length;
}

/** Converts `file` into `outDir`; returns the PDF path or throws with LibreOffice's output. */
function convert(soffice, profile, file, outDir) {
  mkdirSync(outDir, { recursive: true });
  const r = spawnSync(
    soffice,
    [
      `-env:UserInstallation=${pathToFileURL(profile).href}`,
      '--headless',
      '--norestore',
      '--nologo',
      '--convert-to',
      'pdf',
      '--outdir',
      outDir,
      file,
    ],
    { encoding: 'utf8', timeout: TIMEOUT_MS },
  );
  const pdf = join(outDir, `${basename(file, extname(file))}.pdf`);
  if (r.error) throw new Error(`${r.error.message}`);
  if (!existsSync(pdf) || statSync(pdf).size === 0) {
    const log = `${r.stdout ?? ''}${r.stderr ?? ''}`.trim();
    throw new Error(`no PDF produced (exit ${r.status})${log ? `: ${log}` : ''}`);
  }
  return pdf;
}

function table(rows) {
  const widths = rows[0].map((_, c) => Math.max(...rows.map((r) => String(r[c]).length)));
  return rows.map((r) => r.map((v, c) => String(v).padEnd(widths[c])).join('  ')).join('\n');
}

/** Encrypted Office documents are OLE compound files (signature D0 CF 11 E0 A1 B1 1A E1). */
function isEncrypted(file) {
  const head = readFileSync(file).subarray(0, 8);
  const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  return head.length === 8 && sig.every((b, i) => head[i] === b);
}

function main() {
  const argv = process.argv.slice(2);
  const keep = argv.includes('--keep');
  const args = argv.filter((a) => a !== '--keep');
  const files = collect(args.length > 0 ? args : [join(ROOT, 'test', 'fixtures')]);
  if (files.length === 0) {
    console.error('libreoffice-check: no .docx files found');
    return 1;
  }
  const soffice = findSoffice();
  if (soffice === undefined) {
    console.error('libreoffice-check: LibreOffice (soffice) not found; install it or set SOFFICE');
    return 1;
  }
  const expected = expectedPages();
  const work = mkdtempSync(join(tmpdir(), 'turboword-lo-'));
  const profile = join(work, 'profile');
  console.log(`LibreOffice: ${soffice}\nconverting ${files.length} file(s)...\n`);

  const rows = [['file', 'pages', 'expected', 'status']];
  let failures = 0;
  let mismatches = 0;
  let skipped = 0;
  files.forEach((file, i) => {
    const name = file.startsWith(ROOT) ? file.slice(ROOT.length) : file;
    if (isEncrypted(file)) {
      // password-protected: LibreOffice cannot open it headless; decryptDocx's tests cover these
      skipped++;
      rows.push([name, '-', '-', 'skipped (encrypted)']);
      return;
    }
    try {
      // one output directory per file: same basenames in different folders must not collide
      const pdf = convert(soffice, profile, file, join(work, 'pdf', String(i)));
      const pages = countPages(pdf);
      const want = expected.get(basename(file));
      const ok = want === undefined || want === pages;
      if (!ok) mismatches++;
      rows.push([name, pages, want ?? '-', ok ? 'ok' : 'WARN page count differs']);
    } catch (e) {
      failures++;
      rows.push([name, '-', expected.get(basename(file)) ?? '-', 'FAIL']);
      console.error(`FAIL ${name}: ${e instanceof Error ? e.message : e}`);
    }
  });

  console.log(`\n${table(rows)}\n`);
  if (mismatches > 0) {
    console.warn(
      `warning: ${mismatches} page-count mismatch(es); fonts differ by machine, so this is not ` +
        'an error (install Carlito, Caladea and Liberation for comparable results)',
    );
  }
  const tried = files.length - skipped;
  console.log(
    `${tried - failures}/${tried} converted${skipped > 0 ? ` (${skipped} encrypted skipped)` : ''}`,
  );
  if (keep) console.log(`PDFs kept in ${join(work, 'pdf')}`);
  else rmSync(work, { recursive: true, force: true });
  return failures > 0 ? 1 : 0;
}

process.exitCode = main();
