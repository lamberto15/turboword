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
L0 core     errors, limits (incl. maxXmlDepth), units, warnings, hash, seeded PRNG
L1 zip      lazy reader, raw-copy writer, Budget / assertSafePath, native zlib
L1 xml      parser → DOM → serializer, namespace aliases, mc: helpers, schema order, mutate
L2 opc      OpcPackage, Part (lazy, dirty), Relationships, ContentTypes, naming, media, props, fork()
L3 wml      WordDocument + views: Story, Paragraph, Run, Table (+ logical grid), Section,
            HeaderFooter, notes, comments, styles, numbering, settings, fonts, theme, bookmarks,
            images, content controls
L3 ids      DocumentIds: docPr, paraId/textId, bookmark, comment, note, revision, sdt, VML
L4 resolve  style cascade, theme, table-style conditional formatting, list labels
L4 text     flatten ↔ DOM positions, find, replace, extract
L4 validate package invariants
L5 compose  ImportContext, PartCopier, styles / numbering / notes / comments / bookmarks /
            sections / settings
L5 fields · forms · protect · crypto
L6 layout   sfnt fonts, UAX #14, bidi, Shaper interface, layout IR, paragraph / table / float /
            page → PageModel
L7 render   pdf (TrueType subset, ToUnicode, JPEG/PNG, links, outline), svg (debug)
```

v0.1 builds L0 to L3 (`core`, `zip`, `xml`, `opc`, `wml`) plus the first `validatePackage`. The
rest is planned; see [ROADMAP.md](./ROADMAP.md).

## Decisions

| | Decision | Why |
| --- | --- | --- |
| D1 | **Own lightweight DOM**, with a parser derived from the turbosheet tokenizer, not xmldom. It keeps comments, PIs, whitespace, attribute order and namespace declarations. Prefixes are kept as written; matching uses the namespace URI, with Strict OOXML URIs as aliases of their transitional twins. It understands markup compatibility (`mc:`) and tracks dirty state. | Lossless round-trips, speed, size. xmldom is a dev dependency only, for differential tests and benchmarks. |
| D2 | **Typed views over the DOM** (python-docx-style proxies), not a parallel model. A `Paragraph` wraps its `w:p` element. | Anything not modelled survives untouched, because there is nothing to convert back. |
| D3 | **Raw-copy zip.** Untouched entries are written back with their original compressed bytes; a no-op save returns the input bytes. | Unchanged parts cannot be damaged; saving is fast. |
| D4 | **One package, subpaths.** Light classes cover navigation and small edits; heavy operations (compose, fields, layout) are free functions in their own subpaths. | Tree-shaking, with a size budget per subpath. |
| D5 | **Compose through a generic part-graph copier** plus a document-wide ID registry. | Charts, SmartArt, OLE objects, customXml and picture bullets need no special cases. |
| D6 | **Layout hooks from day one:** a `resolve/` style cascade, one visible-text flattener shared by `text()`, search, list labels and layout, and a `PageInfoProvider` interface in fields. | Layout can be added later without reworking the core. |
| D7 | **No fonts in the core.** Fonts embedded in the docx first, then user `FontResolver`s, then a metric-compatible substitution table (Carlito, Caladea, Liberation, Gelasio). Fonts ship separately in `@turboword/fonts`. | Size and font licensing. |
| D8 | **Free layout oracles.** Word-saved fixtures record Word's page count (`docProps/app.xml`) and page breaks (`w:lastRenderedPageBreak`). | Layout is measured against Word without needing Word in CI. |
| D9 | **Validation in CI.** The Open XML SDK validator runs with the rule "no new errors compared with the inputs"; LibreOffice headless is a sanity check. | Catches files Word would call unreadable. |
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
- new IDs come from counters or a seeded PRNG (`core`);
- new and re-serialized zip entries carry a fixed date, 1980-01-01 00:00, unless `zipDate` is given;
- entries keep the source order, and new parts are appended in creation order.

## No recursion over XML trees

Documents can be nested arbitrarily deep (tables in tables, content controls in content controls), and
crafted files can be deeper still. Parsing, serializing, cloning, equality and traversal use explicit
stacks, so depth never overflows the call stack. `maxXmlDepth` (default 1024) bounds parsing; see
[SECURITY.md](./SECURITY.md).

## Errors and warnings

Every deliberate error is a `TurbowordError` with a stable `code` (`DocxFormatError`, `XmlParseError`,
`UsageError`, `ZipLimitError`, ...). Recoverable problems go to the optional `onWarning` callback as
`{ code, message, part? }` instead of being thrown or ignored.

## Subpaths and size budgets

Budgets are min + brotli, enforced by size-limit (`pnpm size`).

| Import | Contents | Budget | Status |
| --- | --- | --- | --- |
| `turboword` | open and save (`readDocx` + `writeDocx`) | 15 KB | v0.1 |
| `turboword` | everything in the root entry | 75 KB | v0.1 |
| `turboword/compose` | `appendDocument`, `insertDocument`, `composeDocuments` | 30 KB | planned (v0.2) |
| `turboword/fields` | field index, evaluators, TOC | 25 KB | planned (v0.4) |
| `turboword/forms` | content controls, legacy form fields | 10 KB | planned (v0.4) |
| `turboword/crypto` | encrypted documents | 8 KB | planned (v0.4) |
| `turboword/layout` | pagination (experimental) | 110 KB | planned (v0.5) |
| `turboword/pdf` | PDF renderer (experimental) | 45 KB | planned (v0.6) |
| `turboword/svg` | SVG debug renderer (experimental) | 8 KB | planned (v0.5) |

Later, separate packages: `@turboword/fonts` and `@turboword/harfbuzz`. The repository becomes a pnpm
workspace when they arrive.

## Source map

```
src/
  index.ts             root entry: names every public export explicitly (no `export *`)
  core/            L0  errors.ts, limits.ts, warnings.ts, units.ts (twips, half-points, EMUs)
  zip/             L1  lazy.ts        central-directory reader; raw() hands out stored payloads
                       writer.ts      raw-copy writer (rawEntry, makeEntry, writeZip)
                       limits.ts      Budget and assertSafePath (zip-bomb and path checks)
                       stream.ts      push inflater, fallback when node:zlib is missing
                       crc32.ts, types.ts
                       node.ts        the only node:zlib lookup in src/
  xml/             L1  parse.ts       whole-string parser (derived from the turbosheet tokenizer)
                       dom.ts         XDocument, XElement, XText, XComment, ..., domEqual
                       serialize.ts   serializer with namespace and mc:Ignorable fix-up
                       ns.ts          namespace URIs, Strict aliases, conventional prefixes
                       escape.ts
  opc/             L2  package.ts     OpcPackage: lazy parts, dirty tracking, raw-copy save, fork, gc
                       part.ts, rels.ts, content-types.ts, rel-types.ts, xml-util.ts
                       uri.ts         part names and relationship targets
  wml/             L3  document.ts    WordDocument, readDocx / loadDocx / writeDocx / docxBlob /
                                      docxStream
                       create.ts      createDocx (a fixed, deterministic template)
                       story.ts       Story (body, headers, footers, notes, comments)
                       paragraph.ts   Paragraph and Run
                       table.ts       Table, rows, cells, logical grid
                       section.ts     Section, effective headers and footers
                       header-footer.ts, properties.ts
                       text.ts        the visible-text flattener
                       order.ts       schema order of property children
                       validate.ts    validatePackage
```

Planned directories, one per layer module: `ids/`, `resolve/`, `text/`, `validate/`, `compose/`,
`fields/`, `forms/`, `protect/`, `crypto/`, `layout/`, `render/`.

Outside `src/`: `test/` (Vitest, fixtures in `test/fixtures/`), `bench/` (benchmarks), `api/report.txt`
(the public API, checked by `test/api-surface.test.ts`), `docs/`.
