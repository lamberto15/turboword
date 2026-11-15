# turboword

[![npm](https://img.shields.io/npm/v/turboword)](https://www.npmjs.com/package/turboword) [![license](https://img.shields.io/badge/license-MIT%20OR%20Apache--2.0-blue)](#license)

Lossless editing and composition of existing .docx files, in pure TypeScript, with one runtime
dependency ([fflate](https://github.com/101arrowz/fflate)), on every JavaScript runtime.

> **Status: v0.1, pre-release.** The API may change in any minor release before 1.0.
>
> **Works today:** open and save losslessly; navigate stories, paragraphs, runs, tables (read-only
> logical grid), sections, and headers and footers (including inherited ones); visible text;
> document properties (read and write); basic edits (run text, bold and italic, paragraph style,
> appending paragraphs and runs); `validatePackage`; `createDocx`.
>
> **Planned:** combining documents with their headers, footers, sections, styles and numbering
> (v0.2), find and replace across runs, inserting documents and editing tables (v0.3), fields,
> content controls and encryption (v0.4), layout and PDF (v0.5 onwards, experimental). See
> [ROADMAP.md](./ROADMAP.md).

## Install

```sh
pnpm add turboword     # or npm install turboword
```

ESM and CommonJS builds, TypeScript types included. Requires Node 22 or newer, or another modern
runtime (see [Runtime support](#runtime-support)).

## Quick start

### Read, inspect, edit, write

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { readDocx, writeDocx } from 'turboword';

const doc = readDocx(readFileSync('report.docx')); // Uint8Array | ArrayBuffer, synchronous, lazy

// Inspect
doc.text(); // visible text of the body, one line per paragraph
for (const p of doc.body.paragraphs()) {
  console.log(p.style, p.text); // style id, visible text
  for (const run of p.runs()) console.log('  run:', run.text, run.bold);
}
const table = doc.body.tables()[0];
table?.cell(1, 2)?.text(); // grid slot row 1, column 2 (0-based); merged cells resolve to their origin

// Edit
const first = doc.body.paragraphs()[0]?.runs()[0];
if (first) first.text = first.text.toUpperCase();
doc.body.appendParagraph('Approved.', { style: 'Heading2' }).appendRun(' (draft)', { italic: true });

// Write. Untouched parts are copied raw; a save with no edits returns the input bytes.
writeFileSync('out.docx', writeDocx(doc));
```

`doc.stories()` lists every flow of text: the body, headers, footers, footnotes, endnotes and
comments. Anything turboword does not model yet is still reachable as XML: `paragraph.element`,
`story.element` and `part.xml()` are the live DOM, and changes made there are saved too.

### Sections and effective headers

```ts
for (const section of doc.sections()) {
  // Effective header: the section's own, or the one inherited from an earlier section
  // ("link to previous"). ownHeader() ignores inheritance.
  console.log(section.index, section.type, section.pageSize?.orientation);
  console.log('  header:', section.header('default')?.text());
  console.log('  footer:', section.footer('default')?.text());
  if (section.titlePage) console.log('  first-page header:', section.header('first')?.text());
}
```

### Document properties

```ts
const { title, creator, modified } = doc.properties; // docProps/core.xml
doc.properties.app.pages; // page count Word stored in docProps/app.xml

doc.properties.title = 'Quarterly report'; // creates docProps/core.xml if missing
doc.properties.setCustom('Reviewed', true); // docProps/custom.xml
```

### Validate

```ts
import { validatePackage } from 'turboword';

for (const issue of validatePackage(doc)) {
  console.log(issue.severity, issue.code, issue.message, issue.part ?? '');
}
```

`validatePackage` is not a schema validator. It checks the structural rules that make Word refuse
a file or offer to repair it: content types, relationship targets and the relationship IDs markup
uses, the main document part, section-properties placement, table cells ending with a paragraph,
duplicate drawing IDs and `mc:Ignorable` declarations. `writeDocx(doc, { validate: true })` runs it
first and throws a `TurbowordError` with code `'VALIDATION'` on errors.

### Create

```ts
import { createDocx, writeDocx } from 'turboword';

const doc = createDocx(); // like Word's blank document: one empty paragraph, US Letter
doc.body.paragraphs()[0]!.appendRun('Hello, world.');
doc.body.appendParagraph('Second paragraph.', { style: 'Heading1' });
const bytes = writeDocx(doc); // deterministic: no timestamps unless you set them
```

turboword is not a document generator: for building rich documents from code, see
[How it compares](#how-it-compares).

### In a browser

```ts
import { docxBlob, loadDocx } from 'turboword';

// loadDocx accepts bytes, a Blob or File, a ReadableStream or an AsyncIterable<Uint8Array>.
const file = (document.querySelector('input[type=file]') as HTMLInputElement).files![0]!;
const doc = await loadDocx(file);
// or: const doc = await loadDocx((await fetch('/report.docx')).body!);

const url = URL.createObjectURL(docxBlob(doc));
```

`docxStream(doc)` returns a web `ReadableStream<Uint8Array>`, handy for `new Response(stream)`.

## Why turboword

A library that loads a document into its own model and writes a new file from it can only keep
what the model covers. turboword edits the file's own XML instead:

- **A no-op save returns the input bytes.** Parts you do not touch are copied with their original
  compressed bytes.
- **Edited parts keep everything.** Unknown elements, both branches of `mc:AlternateContent`,
  extension namespaces and comments survive re-serialization.
- **Nothing is dropped behind your back.** Unreferenced parts stay until you call `doc.gc()`.
- **Deterministic output.** Same input and edits, same bytes: no random IDs, a fixed zip date.
- **Safe on untrusted input.** Zip-bomb and path-traversal checks, an XML depth limit, no DTD or
  entity expansion. See [SECURITY.md](./SECURITY.md).

The few documented losses (DOCTYPE, CDATA delimiters, ...) are listed in
[docs/round-trip.md](./docs/round-trip.md).

## Runtime support

| Runtime | Status |
| --- | --- |
| Node 22, 24, 26 | Tested in CI on Linux and Windows |
| Bun, Deno | Built package smoke-tested in CI |
| Browsers, Cloudflare Workers, other edge runtimes | Expected to work, not tested in CI yet |

`src/` uses web-standard APIs only (`Uint8Array`, Web Streams, `TextEncoder`, `Blob`). Where
`process.getBuiltinModule('node:zlib')` exists (Node 22.3+, Bun, Deno), deflate, inflate and CRC-32
use it; everywhere else, fflate and a table-driven CRC-32.

## How it compares

- **[docx](https://github.com/dolanmiu/docx)** (dolanmiu, TypeScript) creates documents from scratch
  with a declarative API, and can patch placeholders in an existing file. Merging documents is out
  of its scope. turboword focuses on existing files, so the two combine: generate with docx, then
  open and edit the result with turboword (and, from v0.2, compose it with other documents).
- **[docxtemplater](https://docxtemplater.com)** (JavaScript) fills `{tags}` in templates. Inserting
  or merging whole documents is done by its paid Subtemplate and Subsection modules. turboword is
  not a templating engine; free composition is planned for v0.2.
- **[docxcompose](https://github.com/4teamwork/docxcompose)** (Python) appends documents; per its
  documentation, headers and footers from the appended documents are ignored. turboword's planned
  `appendDocument` (v0.2) keeps each document's headers, footers, sections, styles and numbering.

## Documentation

- [docs/round-trip.md](./docs/round-trip.md): what is preserved, and what is lost by design
- [ARCHITECTURE.md](./ARCHITECTURE.md): layers, design decisions, rules
- [ROADMAP.md](./ROADMAP.md): versions, deliverables, exit criteria
- [CONTRIBUTING.md](./CONTRIBUTING.md): development setup, fixtures, conventions
- [SECURITY.md](./SECURITY.md): threat model and how to report a vulnerability

## License

Licensed under either of

- Apache License, Version 2.0 ([LICENSE-APACHE](./LICENSE-APACHE) or
  <https://www.apache.org/licenses/LICENSE-2.0>)
- MIT license ([LICENSE-MIT](./LICENSE-MIT) or <https://opensource.org/licenses/MIT>)

at your option. Third-party attributions are in [NOTICE](./NOTICE).

Unless you explicitly state otherwise, any contribution intentionally submitted for inclusion in
the work by you, as defined in the Apache-2.0 license, shall be dual licensed as above, without any
additional terms or conditions.
