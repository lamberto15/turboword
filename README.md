# turboword

[![npm](https://img.shields.io/npm/v/turboword)](https://www.npmjs.com/package/turboword) [![license](https://img.shields.io/badge/license-MIT%20OR%20Apache--2.0-blue)](#license)

Lossless editing, composition and rendering of Word .docx files, in pure TypeScript, with one
runtime dependency ([fflate](https://github.com/101arrowz/fflate)), on every JavaScript runtime.

> **Status: v0.1.0, first release.** Everything below ships in v0.1.0. The API may change in any
> minor release before 1.0.
>
> **Experimental:** layout, `pageCount`, PDF and SVG output. They work, but they do not match Word
> page for page yet; accuracy is measured and published in [docs/fidelity.md](./docs/fidelity.md).
>
> **Not yet verified in Microsoft Word.** Every output is checked with `validatePackage` and opened
> in LibreOffice in CI; the hand check in Word (the smoke pack in `test/smoke/word-pack.mjs`) is
> pending. Known gaps are listed in [ROADMAP.md](./ROADMAP.md).

## What it does

| Entry point | Features |
| --- | --- |
| `turboword` | Open and save losslessly (`readDocx`, `loadDocx`, `writeDocx`, `createDocx`); stories, paragraphs, runs, tables, sections with effective headers and footers; editing (text and formatting, paragraph split/insert/clone/remove, images, hyperlinks, bookmarks, table rows/columns/merges, section and header/footer editing, page setup); document properties; `validatePackage`; find and replace across runs, tracked or not; accept / reject revisions; resolved formatting (style cascade, theme, table styles) and list labels; document protection |
| `turboword/compose` | `appendDocument`, `insertDocument`, `composeDocuments`, `importBlocks`: merge documents keeping each one's sections, headers and footers, styles, numbering, notes, comments and bookmarks |
| `turboword/fields` | List, evaluate and update fields (MERGEFIELD, IF, `=`, DATE, DOCPROPERTY, REF, PAGEREF, SEQ, ...) with Word's `\*` `\#` `\@` formatting; rebuild tables of contents |
| `turboword/mailmerge` | Mail merge with Word's own merge fields: one document per record or all records in one, table rows repeated from arrays, NEXT / NEXTIF / SKIPIF, merged fields turned into plain text as Word does |
| `turboword/forms` | Content controls (typed values, repeating sections), custom XML data binding, legacy form fields |
| `turboword/crypto` | Decrypt (Agile and Standard) and encrypt (Agile AES-256) password-protected documents |
| `turboword/layout` | `paginate`, `pageCount`: pages, lines and glyph positions as plain data (experimental) |
| `turboword/pdf` | `toPdf`, `renderPdf`: PDF with embedded, subset fonts (experimental) |
| `turboword/svg` | `renderSvg`: one SVG per page, for previews and debugging (experimental) |

Each entry point is a separate import, so a bundler only ships what you use.

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
doc.body.tables()[0]?.cell(1, 2)?.text(); // grid row 1, column 2; merged cells resolve to their origin

// Edit
const first = doc.body.paragraphs()[0]?.runs()[0];
if (first) first.text = first.text.toUpperCase();
const p = doc.body.appendParagraph('Approved.', { style: 'Heading2' });
p.appendRun(' (draft)', { italic: true });
p.appendHyperlink(' details', 'https://example.com');
p.appendImage(readFileSync('logo.png'), { width: 72 }); // points; height keeps the aspect ratio

const table = doc.body.appendTable(3, 4);
table.merge(0, 0, 0, 3); // one header cell across the first row
table.cloneRow(1, { after: 1 });

// Write. Untouched parts are copied raw; a save with no edits returns the input bytes.
writeFileSync('out.docx', writeDocx(doc));
```

`doc.stories()` lists every flow of text: the body, headers, footers, footnotes, endnotes and
comments. Anything turboword does not model is still reachable as XML: `paragraph.element`,
`story.element` and `part.xml()` are the live DOM, and changes made there are saved too.

### Sections, headers and footers

```ts
import { readFileSync } from 'node:fs';
import { readDocx } from 'turboword';

const doc = readDocx(readFileSync('report.docx'));
for (const section of doc.sections()) {
  // The effective header: the section's own, or the one inherited from an earlier section.
  console.log(section.index, section.type, section.pageSize?.orientation);
  console.log('  header:', section.header('default')?.text());
  if (section.titlePage) console.log('  first-page header:', section.header('first')?.text());
}

// Give the last section its own header instead of inheriting one, then edit it.
const last = doc.sections().at(-1)!;
const header = last.unlinkFromPrevious('header', 'default');
header.story.appendParagraph('Appendix');
```

### Merge documents

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { readDocx, writeDocx } from 'turboword';
import { appendDocument, composeDocuments, insertDocument } from 'turboword/compose';

const load = (path: string) => readDocx(readFileSync(path));

// Append: the appended document keeps its own sections, headers, footers, styles and lists.
const target = load('cover.docx');
appendDocument(target, load('chapter-1.docx'));
appendDocument(target, load('chapter-2.docx'), { styles: 'keepDifferentStyles' });
writeFileSync('book.docx', writeDocx(target));

// Several at once, into a new document (the inputs are not modified).
const book = composeDocuments([load('cover.docx'), load('chapter-1.docx'), load('chapter-2.docx')]);

// Insert at a placeholder: the paragraph is split and the document goes in between.
const letter = load('letter.docx');
insertDocument(letter, { placeholder: '{{terms}}' }, load('terms.docx'), {
  keepSourceSections: false, // adopt the letter's page setup and headers
});
```

Options follow Aspose's `appendDocument` and Syncfusion's list modes: `styles`
(`useDestinationStyles`, `keepSourceFormatting`, `keepDifferentStyles`), `lists`
(`keepSource`, `mergeWithDestination`), `restartListsAtImport`, `headersFooters`,
`pageNumbering`, `sectionStart`, `revisions` (keep, accept or reject first), and more.

### Find and replace, revisions

```ts
import { readFileSync } from 'node:fs';
import { acceptAllRevisions, findText, readDocx, replaceText } from 'turboword';

const doc = readDocx(readFileSync('contract.docx'));

// Matches span split runs; the replacement takes the formatting of the match's first run.
findText(doc, /ACME (Corp|Inc)\.?/g, { in: 'all' }); // body, headers, footers, notes, comments, text boxes
replaceText(doc, 'ACME Corp', 'Globex Inc');

// Replace as a tracked change, and use regex groups.
replaceText(doc, /(\d+) days/g, '$1 business days', { track: { author: 'Legal' } });

acceptAllRevisions(doc); // or rejectAllRevisions(doc)
```

Bookmarks and comment ranges stay balanced. A match that crosses a field boundary is skipped with a
warning (or throws, with `onFieldCrossing: 'throw'`).

### Fields and tables of contents

```ts
import { readFileSync } from 'node:fs';
import { readDocx } from 'turboword';
import { listFields, updateFields, updateToc } from 'turboword/fields';

const doc = readDocx(readFileSync('template.docx'));
listFields(doc, { type: 'MERGEFIELD' }).map((f) => f.args[0]);

updateFields(doc, {
  values: { FirstName: 'Ada', Total: 1234.5 }, // MERGEFIELD data
  now: new Date('2026-11-15'), // DATE/TIME are only updated when you pass a clock
  locale: 'en-GB',
});
updateToc(doc); // headings and bookmarks; page numbers too when you pass `pages` (see Layout)
```

Field codes are never removed: only the results change, so Word can update the fields again.

### Mail merge

Templates use Word's own merge fields (Mailings → Insert Merge Field), not a custom tag syntax.

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { readDocx, writeDocx } from 'turboword';
import { mergeAll, mergeEach, mergeRecord } from 'turboword/mailmerge';

const template = readDocx(readFileSync('letter.docx'));

// One record → one document (undefined when a SKIPIF field drops the record)
const letter = mergeRecord(template, {
  name: 'Ada Lovelace',
  total: 1234.5, // MERGEFIELD total \# "#,##0.00" → 1,234.50
  customer: { city: 'London' }, // MERGEFIELD customer.city
  items: [
    // an array repeats the table row that holds MERGEFIELD items
    { items: 'Engine', qty: 1 },
    { items: 'Cards', qty: 200 },
  ],
});

// One document per record, streamed: records are read only as they are merged
for (const { doc, records } of mergeEach(template, customers)) {
  writeFileSync(`out/letter-${records[0]}.docx`, writeDocx(doc));
}

// Every record in one document, each starting a new page with its own headers and footers
writeFileSync('all.docx', writeDocx(mergeAll(template, customers)));
```

As in Word's merged documents, MERGEFIELD, IF, NEXT, NEXTIF, SKIPIF, MERGEREC, MERGESEQ and FILLIN
become plain text (`fields: 'keep'` keeps them live), other fields such as PAGE stay live, and the
mail-merge data-source setup is removed so Word opens the result without a query prompt. NEXT and
NEXTIF move to the next record inside one document (labels, lists), SKIPIF drops a record. Missing
values merge as empty; `missing: 'keep'` leaves those fields, `missing: 'error'` throws.
`mergeAll` separators: `'section'` (default), a section type such as `'oddPage'`, `'pageBreak'`,
or `'none'` (the default for Word directory/catalog templates).

### Content controls and forms

```ts
import { readFileSync } from 'node:fs';
import { readDocx } from 'turboword';
import { bindContentControls, fillContentControls } from 'turboword/forms';

const doc = readDocx(readFileSync('form.docx'));

fillContentControls(doc, {
  customer: 'Ada Lovelace', // text
  approved: true, // checkbox
  due: new Date('2026-12-01'), // date picker, formatted with the control's date format
  items: [{ name: 'Engine', qty: '1' }, { name: 'Cards', qty: '200' }], // repeating section
});

// Or write the custom XML part the controls are bound to, and refresh every bound control.
bindContentControls(doc, '<order><customer>Ada Lovelace</customer></order>');
```

### Protection and encryption

```ts
import { readFileSync } from 'node:fs';
import { protectDocument, readDocx, writeDocx } from 'turboword';
import { decryptDocx, encryptDocx, isEncrypted } from 'turboword/crypto';

const pass = process.env.DOC_PASSWORD ?? ''; // never hard-code real passwords

const doc = readDocx(readFileSync('policy.docx'));
protectDocument(doc, { mode: 'readOnly', password: pass }); // restricts editing in Word

// Encryption is different: the whole file needs the password to open.
const locked = await encryptDocx(writeDocx(doc), pass); // Agile, AES-256, SHA-512
isEncrypted(locked); // true
const plain = readDocx(await decryptDocx(locked, pass));
```

Protection is an editing restriction that Word enforces, not a security boundary: anyone can remove
it. Use `encryptDocx` to keep content confidential.

### Layout, PDF and SVG (experimental)

turboword ships no font files ([why](./ARCHITECTURE.md#decisions)). Fonts are looked up in this
order: fonts embedded in the .docx, then your `FontResolver`s, then metric-compatible substitutes
(Calibri → Carlito, Cambria → Caladea, Arial → Liberation Sans, ...) asked from your resolvers
again, then the PDF base-14 fonts (Helvetica, Times, Courier) as a last resort. With the
substitutes available, line breaks come close to Word's.

```ts
import { readFileSync, writeFileSync } from 'node:fs';
import { readDocx, writeDocx } from 'turboword';
import { updateFields } from 'turboword/fields';
import { type FontResolver, pageCount, pageInfo, paginate } from 'turboword/layout';
import { toPdf } from 'turboword/pdf';
import { renderSvg } from 'turboword/svg';

// Map family names to font files you are licensed to use.
const files: Record<string, string> = {
  Carlito: 'fonts/Carlito-Regular.ttf',
  'Carlito/bold': 'fonts/Carlito-Bold.ttf',
  'Liberation Serif': 'fonts/LiberationSerif-Regular.ttf',
};
const fonts: FontResolver = ({ family, bold }) => {
  const path = files[bold ? `${family}/bold` : family] ?? files[family];
  return path ? new Uint8Array(readFileSync(path)) : undefined;
};

const doc = readDocx(readFileSync('report.docx'));
await pageCount(doc, { fonts });

const layout = await paginate(doc, { fonts }); // pages of positioned text, images and shapes
updateFields(doc, { pages: pageInfo(layout) }); // PAGE, NUMPAGES, PAGEREF and TOC page numbers
writeFileSync('report-updated.docx', writeDocx(doc));

writeFileSync('report.pdf', await toPdf(doc, { fonts }));
writeFileSync('page-1.svg', renderSvg(layout.pages[0]!, layout));
```

The layout handles paragraphs, tabs, justification, lists, tables (merged cells, repeated header
rows, rows split across pages), sections and columns, headers and footers, footnotes, floating
objects with text wrap, shapes, charts drawn from their cached values, EMF/WMF pictures, equations
and bidirectional text. The PDF is deterministic, with subset TrueType fonts, links, an outline and
searchable text.

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

### Validate

```ts
import { readFileSync } from 'node:fs';
import { readDocx, validatePackage } from 'turboword';

const doc = readDocx(readFileSync('report.docx'));
for (const issue of validatePackage(doc)) {
  console.log(issue.severity, issue.code, issue.message, issue.part ?? '');
}
```

`validatePackage` is not a schema validator. It checks the structural rules that make Word refuse
a file or offer to repair it: content types, relationship targets and the relationship IDs markup
uses, the main document part, section-properties placement, table cells ending with a paragraph,
duplicate drawing IDs and `mc:Ignorable` declarations. `writeDocx(doc, { validate: true })` runs it
first and throws a `TurbowordError` with code `'VALIDATION'` on errors.

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
  entity expansion, bounded font and image parsers. See [SECURITY.md](./SECURITY.md).

The few documented losses (DOCTYPE, CDATA delimiters, ...) are listed in
[docs/round-trip.md](./docs/round-trip.md).

## Runtime support

| Runtime | Status |
| --- | --- |
| Node 22, 24, 26 | Tested in CI on Linux and Windows |
| Bun, Deno | Built package smoke-tested in CI |
| Browsers, Cloudflare Workers, other edge runtimes | Expected to work, not tested in CI yet |

`src/` uses web-standard APIs only (`Uint8Array`, Web Streams, `TextEncoder`, `Blob`, Web Crypto).
Where `process.getBuiltinModule('node:zlib')` exists (Node 22.3+, Bun, Deno), deflate, inflate and
CRC-32 use it; everywhere else, fflate and a table-driven CRC-32.

## How it compares

- **[docxcompose](https://github.com/4teamwork/docxcompose)** (Python) appends documents; per its
  documentation, headers and footers from the appended documents are ignored. We ran both on the
  same 59 merge scenarios, built from docxcompose's own test documents and Open-XML-PowerTools'
  DocumentBuilder test files (`scripts/docxcompose-diff`). turboword kept 338 of 338 header and
  footer slots, docxcompose 172. Page counts measured in LibreOffice reached 98.8% of the inputs'
  total with turboword, against 65.5% with docxcompose.
- **[docx](https://github.com/dolanmiu/docx)** (dolanmiu, TypeScript) creates documents from scratch
  with a declarative API, and can patch placeholders in an existing file. Merging documents is out
  of its scope. turboword focuses on existing files, so the two combine: generate with docx, then
  edit, merge or render the result with turboword.
- **[docxtemplater](https://docxtemplater.com)** (JavaScript) fills `{tags}` in templates. Inserting
  or merging whole documents takes its paid Subtemplate and Subsection modules. turboword is not a
  templating engine, but its compose, find/replace, fields and content-control APIs cover much of
  the same ground.
- **DOCX → PDF.** Commercial SDKs and LibreOffice convert with high fidelity today. turboword's
  layout is pure TypeScript with no native dependency, and is experimental: see
  [docs/fidelity.md](./docs/fidelity.md) for how close it gets.

## Documentation

- [docs/round-trip.md](./docs/round-trip.md): what is preserved, and what is lost by design
- [docs/fidelity.md](./docs/fidelity.md): layout accuracy, measured
- [ARCHITECTURE.md](./ARCHITECTURE.md): layers, design decisions, rules
- [ROADMAP.md](./ROADMAP.md): what v0.1.0 contains, known gaps, what comes next
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
