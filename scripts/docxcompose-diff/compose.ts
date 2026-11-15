// turboword side of the docxcompose differential test (see README.md): composes every scenario of
// scenarios.json with composeDocuments() and writes `<out>/<name>.docx`. Run with tsx (it
// resolves the `.js` imports of src/ to their `.ts` files):
//
//   npx --yes tsx scripts/docxcompose-diff/compose.ts scripts/docxcompose-diff/scenarios.json out/dir
//
// Prints one JSON line per scenario: { name, ok, warnings, error? }.

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { composeDocuments } from '../../src/compose/index.js';
import { readDocx, writeDocx } from '../../src/wml/document.js';

interface Scenario {
  name: string;
  inputs: string[];
}

const [scenariosPath, outDir] = process.argv.slice(2);
if (scenariosPath === undefined || outDir === undefined) {
  console.error('usage: compose.ts <scenarios.json> <out-dir>');
  process.exit(2);
}
const root = resolve(scenariosPath, '../../..');
const { scenarios } = JSON.parse(readFileSync(scenariosPath, 'utf8')) as { scenarios: Scenario[] };
mkdirSync(outDir, { recursive: true });
for (const s of scenarios) {
  const warnings: string[] = [];
  try {
    const docs = s.inputs.map((p) => readDocx(readFileSync(join(root, p))));
    const out = composeDocuments(docs, { onWarning: (w) => warnings.push(w.code) });
    writeFileSync(join(outDir, `${s.name}.docx`), writeDocx(out));
    console.log(JSON.stringify({ name: s.name, ok: true, warnings }));
  } catch (e) {
    console.log(JSON.stringify({ name: s.name, ok: false, warnings, error: String(e) }));
  }
}
