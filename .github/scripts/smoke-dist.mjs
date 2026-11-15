// Runtime smoke test of the BUILT package: loads `dist/` (the ESM and the CommonJS build) and
// round-trips fixtures. CI runs it under Node, Bun and Deno, so the files users actually install
// are exercised on every runtime the package claims to support, including the runtime-specific
// compression paths (`deflate: 'native'` and `'fflate'`).
//
//   pnpm run build
//   node .github/scripts/smoke-dist.mjs [file.docx | dir]          (default dir: test/fixtures)
//   bun  .github/scripts/smoke-dist.mjs test/fixtures
//   deno run --allow-read .github/scripts/smoke-dist.mjs test/fixtures
//
// Given a directory, it round-trips every `.docx` below it (sorted), so it needs no fixture names
// and picks up new fixtures automatically. Given a file, only that file.
//
// Checks, for each fixture and each build:
// 1. a no-op save is byte-identical to the input (the lossless guarantee);
// 2. after marking the main part dirty, the file is written with each deflate backend, read back
//    (from bytes and from a Blob), and its text is unchanged.

import { readdir, readFile, stat } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { join, relative } from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../../', import.meta.url));
const esmEntry = new URL('../../dist/index.js', import.meta.url);
const cjsEntry = fileURLToPath(new URL('../../dist/index.cjs', import.meta.url));

/** Every .docx under `dir` (explicit stack, no recursion), sorted for a stable order. */
async function findDocx(dir) {
  const found = [];
  const stack = [dir];
  while (stack.length > 0) {
    const current = stack.pop();
    for (const e of await readdir(current, { withFileTypes: true })) {
      const path = join(current, e.name);
      if (e.isDirectory()) stack.push(path);
      else if (e.isFile() && e.name.toLowerCase().endsWith('.docx')) found.push(path);
    }
  }
  if (found.length === 0) throw new Error(`no .docx file found under ${dir}`);
  return found.sort();
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function sameBytes(a, b) {
  if (a.byteLength !== b.byteLength) return false;
  for (let i = 0; i < a.byteLength; i++) if (a[i] !== b[i]) return false;
  return true;
}

async function roundTrip(label, api, input) {
  for (const name of ['readDocx', 'loadDocx', 'writeDocx']) {
    assert(typeof api[name] === 'function', `${label}: export ${name} is missing`);
  }

  const doc = api.readDocx(input);
  const text = doc.text();
  assert(sameBytes(api.writeDocx(doc), input), `${label}: a no-op save is not byte-identical`);

  for (const deflate of ['native', 'fflate']) {
    doc.mainPart.xml().markDirty(); // forces a real serialize + deflate instead of a raw copy
    const out = api.writeDocx(doc, { deflate });
    assert(api.readDocx(out).text() === text, `${label}: text changed after save (${deflate})`);
    const loaded = await api.loadDocx(new Blob([out]));
    assert(loaded.text() === text, `${label}: text changed after loadDocx(Blob) (${deflate})`);
  }
}

const target = process.argv[2] ?? join(root, 'test', 'fixtures');
const fixtures = (await stat(target)).isDirectory() ? await findDocx(target) : [target];
const builds = [
  ['ESM', await import(esmEntry.href)],
  ['CJS', createRequire(import.meta.url)(cjsEntry)],
];
const runtime = globalThis.Deno
  ? `Deno ${globalThis.Deno.version.deno}`
  : globalThis.Bun
    ? `Bun ${globalThis.Bun.version}`
    : `Node ${process.versions.node}`;

for (const fixture of fixtures) {
  const name = relative(root, fixture) || fixture;
  const input = new Uint8Array(await readFile(fixture));
  for (const [build, api] of builds) await roundTrip(`${build} ${name}`, api, input);
  console.log(`ok ${name}`);
}
console.log(`${runtime}: dist (ESM and CJS) round-trips ${fixtures.length} fixture(s)`);
