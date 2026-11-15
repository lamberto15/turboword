# Architecture

turboword is one published package with subpath exports. A .docx is a zip of XML parts; turboword
keeps every part exactly as it found it until you change it, and gives you typed views over the XML
rather than a separate document model. The build is unbundled (tsdown `unbundle`: one `dist` file
per source module) and the package has `sideEffects: false`, so bundlers drop what you do not
import.

## Layers

Imports go downward only: a file may import its own directory or a directory of a strictly lower
layer, so modules on the same layer do not import each other (`zip` and `xml` know nothing of one
another). Entry points directly in `src/` may import anything; nothing imports them.
`test/conventions.test.ts` enforces this, and a new directory must be given a layer there before it
can be imported.

```
L0 core     errors, limits (incl. maxXmlDepth), units, warnings, SHA-1/SHA-2 hashes
L1 zip      lazy reader, raw-copy writer, Budget / assertSafePath, native zlib
L1 xml      parser → DOM → serializer, namespace aliases, mc: helpers
L2 opc      OpcPackage, Part (lazy, dirty), Relationships, ContentTypes, naming, fork(), gc()
L3 wml      WordDocument + views: Story, Paragraph, Run, Table (logical grid + editing),
            Section, HeaderFooter, properties, bookmarks, hyperlinks, images (media),
            DocumentIds (document-wide id registry), validatePackage, createDocx
L4 resolve  style cascade, theme, table-style conditional formatting, borders, list labels
L4 text     flatten ↔ DOM positions, find, replace (tracked or not), accept / reject revisions
L5 compose  ImportContext, PartCopier, namespaces, styles, numbering, notes, comments,
            bookmarks, relationships, id renumbering, sections, settings
L5 fields   field index, switch formatting, evaluators, TOC
L5 forms    content controls, custom XML data binding, legacy form fields
L5 protect  documentProtection hashes
L5 crypto   ECMA-376 Agile / Standard encryption over CFB
L6 layout   fonts (sfnt, embedded, substitutes, base-14), text (UAX #14, UAX #9, scripts,
            shaper), IR, line (paragraph layout), page (pagination, tables, floats, notes),
            graphics (shapes, charts, EMF/WMF, OMML) → PageModel
L7 render   pdf (writer, TrueType subsetting, ToUnicode, images, links, outline), svg
```

`test/conventions.test.ts` holds the exact ranks: `core` 0; `zip`, `xml` 1; `opc` 2; `wml` 3;
`resolve`, `text` 4; `compose`, `fields`, `forms`, `protect`, `crypto` 5; `layout` 6; `render` 7.
Same-rank siblings never import each other: `fields` cannot import `forms`, and `crypto` cannot
import `protect` (which is why the SHA hashes both need live in `core`).

## Decisions

| | Decision | Why |
| --- | --- | --- |
| D1 | **Own lightweight DOM**, with a parser derived from the turbosheet tokenizer, not xmldom. It keeps comments, PIs, whitespace, attribute order and namespace declarations. Prefixes are kept as written; matching uses the namespace URI, with Strict OOXML URIs as aliases of their transitional twins. It understands markup compatibility (`mc:`) and tracks dirty state. | Lossless round-trips, speed, size. xmldom is a dev dependency only, for differential tests and benchmarks. |
| D2 | **Typed views over the DOM** (python-docx-style proxies), not a parallel model. A `Paragraph` wraps its `w:p` element. | Anything not modelled survives untouched, because there is nothing to convert back. |
| D3 | **Raw-copy zip.** Untouched entries are written back with their original compressed bytes; a no-op save returns the input bytes. | Unchanged parts cannot be damaged; saving is fast. |
| D4 | **One package, subpaths.** Light classes cover navigation and small edits; heavy operations (compose, fields, layout) are free functions in their own subpaths. | Tree-shaking, with a size budget per subpath. |
| D5 | **Compose through a generic part-graph copier** plus a document-wide ID registry. | Charts, SmartArt, OLE objects, customXml and picture bullets need no special cases. |
| D6 | **Layout hooks from day one:** a `resolve/` style cascade, one visible-text flattener shared by `text()`, search, list labels and layout, and a `PageInfoProvider` interface in fields. | Layout can be added later without reworking the core. |
| D7 | **No fonts in the core.** Fonts embedded in the docx first, then user `FontResolver`s, then a metric-compatible substitution table (Carlito, Caladea, Liberation, Gelasio; a substitute reports the original family's line metrics), then the PDF base-14 fonts (metrics only) as a last resort. A font package (`@turboword/fonts`) is planned. | Size and font licensing. |
| D8 | **Free layout oracles.** Word-saved fixtures record Word's page count (`docProps/app.xml`) and page breaks (`w:lastRenderedPageBreak`). | Layout is measured against Word without needing Word in CI. |
| D9 | **Validation in CI.** `validatePackage` with the rule "no new errors compared with the inputs", plus LibreOffice headless conversion of outputs. The Open XML SDK validator is planned (it needs .NET in CI). | Catches files Word would call unreadable. |
| D10 | **Multi-runtime.** `src/` uses web-standard APIs only (see below). | One build for Node, Bun, Deno, browsers and workers. |

## Lossless round-trip

1. A no-op save returns the input bytes.
2. Clean entries keep their original compressed payload and CRC.
3. Dirty parts are re-serialized from the full DOM: unknown elements, both `mc:AlternateContent`
   branches, extension namespaces and comments survive.
4. Parts are never dropped implicitly; cleanup is the explicit `gc()`.
5. Output is deterministic (see below).
6. Before a dirty part is serialized, every prefix it uses is declared and `mc:Ignorable` is kept
   consistent.
7. Losses by design: DOCTYPE, the spelling of character references, CDATA delimiters, CRLF inside the
   root element, literal whitespace in attribute values.

Details and the exact list of losses: [docs/round-trip.md](./docs/round-trip.md).

## Multi-runtime

`src/` uses web-standard APIs only: `Uint8Array`, `TextEncoder` / `TextDecoder`, Web Streams, `Blob`,
async iterators. No `Buffer`, no static `node:*` imports.

The one exception is `src/zip/node.ts`. It looks up `node:zlib` at runtime with
`process.getBuiltinModule` (Node 22.3+, Bun, Deno) for faster deflate, inflate and CRC-32. When the
lookup fails, callers fall back to fflate (deflate, inflate) and a table-driven CRC-32. fflate is
the only runtime dependency.

## Determinism

The same input and the same calls produce the same bytes:

- no `Math.random`, `Date.now()` or `new Date()` without arguments anywhere in `src/` (checked by the
  conventions test);
- new IDs come from counters or a seeded PRNG (`DocumentIds` in `wml/ids.ts`); list `nsid`s from a
  content hash plus a counter; custom XML store GUIDs from a content hash;
- new and re-serialized zip entries carry a fixed date, 1980-01-01 00:00, unless `zipDate` is given;
- entries keep the source order, and new parts are appended in creation order;
- DATE / TIME fields are only updated when the caller passes `now`; PDFs carry no date unless asked;
- layout and PDF output are deterministic too (font subset tags and the PDF `/ID` are content
  hashes).

The deliberate exceptions are cryptographic: a protection salt (unless you pass `salt`) and the
keys and salts of `encryptDocx` come from `crypto.getRandomValues`, as they must.

## No recursion over XML trees

Documents can be nested arbitrarily deep (tables in tables, content controls in content controls), and
crafted files can be deeper still. Parsing, serializing, cloning, equality and traversal use explicit
stacks, so depth never overflows the call stack. `maxXmlDepth` (default 1024) bounds parsing; see
[SECURITY.md](./SECURITY.md).

## Errors and warnings

Every deliberate error is a `TurbowordError` with a stable `code`: `DocxFormatError`
(`DOCX_FORMAT`), `XmlParseError` (`XML_PARSE`), `UsageError` (`USAGE`), `ZipLimitError`
(`ZIP_LIMIT`, with subclasses `ZipPathError` `ZIP_PATH` and `ZipFormatError` `ZIP_MALFORMED`), and
codes such as `VALIDATION`, `FONT_FORMAT` and `BAD_PASSWORD`. Recoverable problems go to the optional `onWarning` callback as
`{ code, message, part? }` instead of being thrown or ignored.

## Subpaths and size budgets

| Import | Contents |
| --- | --- |
| `turboword` | document model, editing, text search/replace, revisions, resolve, protection |
| `turboword/compose` | `appendDocument`, `insertDocument`, `composeDocuments`, `importBlocks` |
| `turboword/fields` | field index, evaluators, TOC |
| `turboword/forms` | content controls, data binding, legacy form fields |
| `turboword/crypto` | encrypted documents |
| `turboword/layout` | `paginate`, `pageCount`, `pageInfo`, fonts and Unicode helpers (experimental) |
| `turboword/pdf` | `toPdf`, `renderPdf` (experimental; `toPdf` pulls in the layout engine) |
| `turboword/svg` | `renderSvg`, `renderSvgPages` (experimental) |

Each subpath is its own tsdown entry (`dist/<name>.js`). Size budgets (min + brotli) are set from
measured sizes in the `size-limit` section of `package.json` and enforced by `pnpm size`; they
guard against creep rather than set targets. Subpath measurements include the shared core code
they import.

Later, separate packages: `@turboword/fonts` and `@turboword/harfbuzz`. The repository becomes a pnpm
workspace when they arrive.

## Layout pipeline

```
WordDocument ─buildIR→ IR (resolved props, list labels, images, notes, anchors)
  → FontManager.want() for every run → await load() (embedded → resolvers → substitutes → base-14)
  → layoutParagraph (segment by font/script/bidi, shape, greedy line break, tabs, justify)
  → paginate (sections, columns, keep rules, tables, floats, headers/footers, footnotes,
    PAGE/NUMPAGES/PAGEREF fixed point, ≤ 3 passes)
  → DocumentLayout (PageModel: plain data in points, origin top-left)
  → renderPdf / renderSvg
```

Layout never reads XML after `buildIR`, and renderers never read the IR or the document. The
PageModel is a list of pages holding text items (glyph ids and advances), images, rects, lines,
paths and links, plus fonts, images, the outline, named destinations and a paragraph → page map
(`pageInfo` turns it into the `PageInfoProvider` that `updateFields` and `updateToc` use).

## Source map

```
src/
  index.ts             root entry: names every public export explicitly (no `export *`)
  core/            L0  errors, limits, warnings, units (twips, half-points, EMUs), sha512 (SHA-1/2)
  zip/             L1  lazy.ts (central-directory reader, raw()), writer.ts (rawEntry, makeEntry,
                       writeZip), limits.ts (Budget, assertSafePath), stream.ts, crc32.ts,
                       node.ts (the only node:zlib lookup in src/)
  xml/             L1  parse.ts, dom.ts (XDocument, XElement, ..., domEqual), serialize.ts
                       (namespace and mc:Ignorable fix-up), ns.ts, mc.ts, escape.ts
  opc/             L2  package.ts (OpcPackage: lazy parts, dirty tracking, raw-copy save, fork, gc),
                       part.ts, rels.ts, content-types.ts, rel-types.ts, uri.ts, xml-util.ts
  wml/             L3  document.ts (WordDocument, read/load/write/blob/stream), create.ts,
                       story.ts, paragraph.ts (Paragraph, Run), table.ts (grid + editing),
                       section.ts, header-footer.ts, properties.ts, bookmarks.ts, hyperlinks.ts,
                       media.ts (images), ids.ts (DocumentIds), text.ts (visible-text walker),
                       order.ts (schema order), validate.ts
  resolve/         L4  styles, ppr / rpr / tblpr cascades, theme, table (conditional formatting,
                       borders), numbering + labels, format (numFmt), context (caches)
  text/            L4  flatten, stories, find, replace, runs, revisions
  compose/         L5  index.ts (orchestrator), context.ts, part-copier.ts, namespaces.ts,
                       renumber.ts, content.ts, core-insert.ts, styles.ts, styles-props.ts,
                       numbering.ts, notes.ts, comments.ts, bookmarks.ts, sections.ts, settings.ts
  fields/          L5  tokenize, scan, model, format (switches), formula, evaluate, docinfo,
                       write, update, toc, settings
  forms/           L5  sdt (classification, values), content, fill, bind, xpath, customxml,
                       legacy, date-format
  protect/         L5  documentProtection hashes and settings
  crypto/          L5  cfb (compound file), agile, standard, hash
  layout/          L6  types.ts (IR, PageModel), engine.ts (internal seams), ir.ts, paragraph.ts,
                       paginate.ts, table.ts, index.ts
    fonts/             sfnt parser, cmap, kerning, names, embedded (.odttf), substitutions,
                       substitute (Word's vertical metrics), standard14, slots, manager
    text/              linebreak (UAX #14), bidi (UAX #9), script, shaper, data/ (generated
                       from Unicode 16.0 by scripts/gen-unicode.mjs)
    line/              IR builders (ir-*.ts), pieces, shape, breaker, tabs, emit, metrics,
                       decorations
    page/              flow, strip, blocks, floats, fields, emit
    graphics/          presets, geometry, shape, paint, color, vml, chart*, emf, wmf, dib, png,
                       math (OMML), images
  render/          L7  pdf/ (writer, render, content, fonts + sfnt subsetting, path, md5,
                       images/: jpeg, png, gif, bmp, tiff), svg/
```

Outside `src/`: `test/` (Vitest; fixtures in `test/fixtures/`), `bench/` (benchmarks), `scripts/`
(LibreOffice check, docxcompose differential, Unicode data generator), `api/report.txt` (the public
API of every entry, checked by `test/api-surface.test.ts`), `docs/`.
