// Generates a pack of .docx files for a manual check in real Microsoft Word (stricter than
// LibreOffice: it shows "Word found unreadable content..." where LibreOffice silently copes).
// Run after `pnpm build`:
//   node test/smoke/word-pack.mjs [outDir]      (default: out/pack)
// It writes, for every fixture in test/fixtures, a no-op round-trip (noop/<fixture path>) and a
// force-dirty round-trip (dirty/<fixture path>: every XML part re-serialized from the DOM), plus a
// new document from createDocx() and a properties-edited one, and a CHECKLIST.md saying what to
// look at in each file. Record the results in docs/word-verification.md.
// `node scripts/libreoffice-check.mjs out/pack` is the automated (weaker) version of the same
// check; fixture basenames are kept so it compares page counts with test/fixtures/expected.json.
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { basename, dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDocx, readDocx, writeDocx } from '../../dist/index.js';

const ROOT = fileURLToPath(new URL('../../', import.meta.url));
const FIXTURES = join(ROOT, 'test', 'fixtures');
const out = process.argv[2] ?? join(ROOT, 'out', 'pack');
mkdirSync(out, { recursive: true });
// stale files from an earlier run would otherwise be checked too
for (const d of ['noop', 'dirty']) rmSync(join(out, d), { recursive: true, force: true });
for (const f of readdirSync(out)) if (f.endsWith('.docx')) rmSync(join(out, f));

const expected = existsSync(join(FIXTURES, 'expected.json'))
  ? JSON.parse(readFileSync(join(FIXTURES, 'expected.json'), 'utf8'))
  : {};
const fixtures = readdirSync(FIXTURES, { recursive: true, encoding: 'utf8' })
  .filter((f) => f.endsWith('.docx'))
  .map((f) => f.split(sep).join('/'))
  .sort();

const OPENS = 'opens in Word without a repair prompt';
const sections = { noop: [], dirty: [], created: [] };
let failures = 0;

const save = (group, name, bytes, checks) => {
  mkdirSync(dirname(join(out, name)), { recursive: true });
  writeFileSync(join(out, name), bytes);
  sections[group].push(`- [ ] **${name}** — ${checks.join('; ')}`);
};
const fail = (group, name, e) => {
  failures++;
  sections[group].push(`- [x] **${name}** — NOT GENERATED: ${e instanceof Error ? e.message : e}`);
  console.error(`FAIL ${name}:`, e);
};

/** What the checker should see: header/footer texts and the page count Word reported. */
function expectations(doc, fixture) {
  const hf = doc.stories(['header', 'footer']);
  const texts = hf
    .map((s) => s.text().replace(/\s+/g, ' ').trim())
    .filter((t) => t !== '')
    .map((t) => `"${t.length > 40 ? `${t.slice(0, 40)}…` : t}"`);
  const out = [];
  if (hf.length > 0) {
    const what = texts.length > 0 ? `, e.g. ${[...new Set(texts)].slice(0, 3).join(', ')}` : '';
    out.push(`headers/footers present (${hf.length} part${hf.length === 1 ? '' : 's'}${what})`);
  } else out.push('no headers/footers expected');
  const pages = fixture === undefined ? undefined : expected[fixture]?.pages;
  if (typeof pages === 'number') out.push(`${pages} page(s)`);
  return out;
}

// 1. no-op round-trips: must be the input bytes, so these mostly prove the reader accepts the file
for (const f of fixtures) {
  const name = `noop/${f}`;
  try {
    const input = new Uint8Array(readFileSync(join(FIXTURES, f)));
    const doc = readDocx(input);
    const bytes = writeDocx(doc);
    const same = bytes.length === input.length && bytes.every((b, i) => b === input[i]);
    if (!same) throw new Error('a no-op save did not return the input bytes');
    save('noop', name, bytes, [
      OPENS,
      `identical to \`test/fixtures/${f}\` (byte-identical)`,
      ...expectations(doc, f),
    ]);
  } catch (e) {
    fail('noop', name, e);
  }
}

// 2. force-dirty round-trips: every XML part re-serialized from the DOM (the real test of the writer)
for (const f of fixtures) {
  const name = `dirty/${f}`;
  try {
    const input = new Uint8Array(readFileSync(join(FIXTURES, f)));
    const doc = readDocx(input);
    let parts = 0;
    for (const part of doc.package.parts()) {
      if (!part.isXml) continue;
      part.xml().markDirty();
      parts++;
    }
    const bytes = writeDocx(doc);
    const before = readDocx(input).text();
    if (readDocx(bytes).text() !== before) throw new Error('body text changed across the save');
    save('dirty', name, bytes, [
      OPENS,
      `looks identical to \`test/fixtures/${f}\` (${parts} XML parts re-serialized)`,
      ...expectations(doc, f),
    ]);
  } catch (e) {
    fail('dirty', name, e);
  }
}

// 3. a new document
try {
  const doc = createDocx();
  doc.body.appendParagraph('turboword smoke pack', { style: 'Title' });
  doc.body.appendParagraph('Created with createDocx()', { style: 'Heading1' });
  const p = doc.body.appendParagraph('Plain text, then ');
  p.appendRun('bold', { bold: true });
  p.appendRun(', ');
  p.appendRun('italic', { italic: true });
  p.appendRun(' and ');
  p.appendRun('both', { bold: true, italic: true });
  p.appendRun('.');
  doc.body.appendParagraph('Escaping: a < b & c > d, "double" and \'single\' quotes.');
  doc.body.appendParagraph('Unicode: émoji 🚀, 中文, עברית, Ελληνικά.');
  doc.body.appendParagraph('Leading and trailing spaces are kept:   [   three   ]   ');
  save('created', 'created-paragraphs.docx', writeDocx(doc), [
    OPENS,
    'after the empty first paragraph: Title, Heading 1 and 4 body paragraphs (text as in the script)',
    'the bold/italic words are formatted; quotes, < & > and the non-Latin scripts show correctly',
    'File > Info shows no errors; US Letter, 1" margins',
  ]);
} catch (e) {
  fail('created', 'created-paragraphs.docx', e);
}

// 4. edited document properties (on a fixture with headers/footers when there is one)
try {
  const f =
    fixtures.find((x) => basename(x) === 'docproperties_header_footer_3_sections.docx') ??
    fixtures[0];
  const doc = f ? readDocx(new Uint8Array(readFileSync(join(FIXTURES, f)))) : createDocx();
  const props = doc.properties;
  props.title = 'Smoke pack title';
  props.subject = 'Subject with <markup> & "quotes"';
  props.creator = 'turboword';
  props.keywords = 'smoke, pack, word';
  props.description = 'Edited by test/smoke/word-pack.mjs';
  props.lastModifiedBy = 'turboword smoke pack';
  props.created = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));
  props.modified = new Date(Date.UTC(2026, 10, 15, 12, 0, 0));
  props.company = 'turboword contributors';
  props.setCustom('Smoke text', 'hello');
  props.setCustom('Smoke number', 42);
  props.setCustom('Smoke flag', true);
  props.setCustom('Smoke date', new Date(Date.UTC(2025, 5, 30)));
  save('created', 'properties-edited.docx', writeDocx(doc), [
    OPENS,
    `content looks identical to ${f ? `\`test/fixtures/${f}\`` : 'a blank document'}`,
    'File > Info / Properties: title "Smoke pack title", subject `Subject with <markup> & "quotes"`, ' +
      'author "turboword", company "turboword contributors", created 2024-01-02, modified 2026-11-15',
    'Advanced/Custom properties: "Smoke text" = hello, "Smoke number" = 42, "Smoke flag" = yes, ' +
      '"Smoke date" = 2025-06-30',
    ...(f ? expectations(doc, f).slice(0, 1) : []),
  ]);
} catch (e) {
  fail('created', 'properties-edited.docx', e);
}

const checklist = `# Word smoke pack — manual checklist

Generated by \`node test/smoke/word-pack.mjs\` from \`dist/\`. Open every file in Microsoft Word
(Windows and/or Mac; note the version) and tick it when **all** of its checks hold. Any
"Word found unreadable content" / repair prompt is a release blocker: note the file and the
Word version in \`docs/word-verification.md\`. Compare against the original fixture side by side
(same fonts, page count, headers/footers on every page, images, tables, comments, footnotes).

Word version: ____________   Date: ____________   Checked by: ____________

## No-op round-trips (must be byte-identical to the fixture)

${sections.noop.join('\n') || '_no fixtures found_'}

## Force-dirty round-trips (every XML part re-serialized)

${sections.dirty.join('\n') || '_no fixtures found_'}

## Created and edited documents

${sections.created.join('\n')}
`;
writeFileSync(join(out, 'CHECKLIST.md'), checklist);

const count = Object.values(sections).reduce((n, s) => n + s.length, 0) - failures;
console.log(`wrote ${count} file(s) and CHECKLIST.md to ${relative(process.cwd(), out) || '.'}`);
if (failures > 0) {
  console.error(`${failures} file(s) could not be generated (see CHECKLIST.md)`);
  process.exitCode = 1;
}
