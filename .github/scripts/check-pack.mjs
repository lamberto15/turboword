// Pack whitelist: checks that the tarball `pnpm publish` would upload contains exactly what should
// ship (the build, the docs, both licence files and NOTICE) and nothing else (sources, tests,
// fixtures, configs). Reads the output of `pnpm pack --dry-run --json` from stdin:
//
//   pnpm pack --dry-run --json | node .github/scripts/check-pack.mjs
//
// It is part of `pnpm run check:package`, which CI and the release workflow run after a build.

import { readFileSync } from 'node:fs';

const REQUIRED = [
  'package.json',
  'README.md',
  'LICENSE-MIT',
  'LICENSE-APACHE',
  'NOTICE',
  'dist/index.js',
  'dist/index.cjs',
  'dist/index.d.ts',
  'dist/index.d.cts',
];
const ALLOWED =
  /^(dist\/.+\.(js|cjs|d\.ts|d\.cts)|README\.md|LICENSE-MIT|LICENSE-APACHE|NOTICE|CHANGELOG\.md|package\.json)$/;

const input = readFileSync(0, 'utf8');
let pack;
try {
  pack = JSON.parse(input);
} catch (e) {
  throw new Error(`expected the JSON output of \`pnpm pack --dry-run --json\`, got:\n${input}`, {
    cause: e,
  });
}

const files = pack.files.map((f) => f.path.replaceAll('\\', '/'));
const unexpected = files.filter((f) => !ALLOWED.test(f));
const missing = REQUIRED.filter((f) => !files.includes(f));

if (unexpected.length > 0) console.error('Files that should not be published:', unexpected);
if (missing.length > 0) console.error('Files missing from the package:', missing);
if (unexpected.length > 0 || missing.length > 0) process.exit(1);
console.log(`${pack.filename}: ${files.length} files, whitelist ok`);
