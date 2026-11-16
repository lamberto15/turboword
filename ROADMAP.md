# Roadmap

Where turboword stands and what comes next. Update this file with every release.

**Layout, `pageCount`, PDF and SVG are experimental** until the scoreboard in
[docs/fidelity.md](./docs/fidelity.md) meets its targets (exact page count ≥ 85% and line match
≥ 95% on prose documents, measured against Word's own output), not until a particular version
number. Until then their APIs and output may change in any minor release.

## What v0.1.0 contains

The first release ships every feature planned for 1.0, instead of the phased v0.1–v0.8 plan this
project started with.

| Area | Delivered | Evidence |
| --- | --- | --- |
| Lossless core | zip with raw entry copy, XML DOM, OPC package, `readDocx` / `loadDocx` / `writeDocx` / `createDocx`, stories, sections, effective headers and footers, properties, `validatePackage` | No-op saves byte-identical on the whole corpus; force-dirtied parts reparse DOM-equal; DOM 2.3–5.5× faster than xmldom (parse + serialize); fuzzing at ~300 iterations per target |
| Editing | text and formatting, paragraph split / insert / clone / remove, images (9 formats incl. SVG), hyperlinks, bookmarks, table rows / columns / merges / clones, section and header / footer editing, page setup | Table invariants checked after randomized operation sequences; outputs validated and converted in LibreOffice |
| Text | find / replace across split runs, tracked replace, accept / reject every revision type | Split-run corpus 100% (5,450 paragraphs × 4 replace modes) |
| Formatting | style cascade, theme, table conditional formatting, border conflicts, list labels | List labels match LibreOffice on 97.9% of 521 paragraphs; every difference is a case where LibreOffice departs from Word |
| Compose | `appendDocument`, `insertDocument`, `composeDocuments`, `importBlocks`; three style modes, three list modes, notes, comments (five parts), bookmarks, sections with headers and footers, settings, fonts, theme | 59-scenario differential against docxcompose: 338 / 338 header and footer slots kept (docxcompose 172); 98.8% of input pages (docxcompose 65.5%) |
| Fields | field index, `\*` `\#` `\@` formatting, evaluators, TOC rebuild | TOC rebuild is stable on a second run; field codes preserved |
| Forms | content controls (all types, repeating sections), data binding, legacy form fields | Real fixtures from docx4j, Open-XML-SDK and Apache POI |
| Protection, encryption | Word-compatible protection hashes; Agile / Standard decryption, Agile encryption | Hash checked against a Word-written file; msoffcrypto-tool decrypts our output |
| Layout (experimental) | font stack, UAX #14, UAX #9, paragraph layout, pagination, tables, floats, headers / footers, footnotes, page fields | Unicode conformance files pass 100%; line breaks match LibreOffice on 97.9% of lines (compatibility mode 15); page counts in [docs/fidelity.md](./docs/fidelity.md) |
| Rendering (experimental) | PDF (subset fonts, images, links, outline), SVG, shapes, charts, EMF / WMF, OMML | `qpdf --check` clean; extracted text equals the laid-out text |

## Since v0.1.0 (unreleased)

| Area | Delivered | Evidence |
| --- | --- | --- |
| Mail merge | `turboword/mailmerge`: `mergeRecord`, `mergeEach` (streamed), `mergeAll`; table rows repeated from arrays (nested tables too), NEXT / NEXTIF / SKIPIF / MERGEREC / MERGESEQ, merged fields unlinked as Word does, data-source setup removed | The docx-mailmerge test suite's NEXT, NEXTIF, SKIPIF and IF expectations; every output validated; ~1 ms per document with `mergeEach`; `mergeAll` of 1,000 records with six headers/footers each in ~5 s |
| Compose at scale | Relationship, content-type, part-name and paragraph-id lookups indexed | Appending 1,000 documents with headers: from 432 s to 5 s |

## Not done yet

- **Verified in Microsoft Word.** An automated open-only pass of the Word smoke pack in Word for
  Mac 16.114 found no repair prompt caused by turboword (`docs/word-verification.md`). Still to
  do: Word for Windows, and a side-by-side visual comparison with the fixtures.
- **The Open XML SDK validator in CI** (needs .NET), with the rule "no new errors compared with the
  inputs".

## Known gaps

Grouped by area; each is a candidate for a minor release.

- **Editing.** `hyperlinks()` lists `w:hyperlink` elements only, not HYPERLINK fields. Bookmark
  rename rewrites a field instruction only when the name sits in one `w:instrText`. Images are not
  shrunk to the page width and JPEG EXIF orientation is ignored. Content inside
  `mc:AlternateContent` is edited in one branch only (tables, paragraph text). VML shape ids are not
  refreshed in cloned rows. Changing a cell width needs `recomputeGrid()`.
- **Text and revisions.** Matches touching math or ruby text are skipped. On reject,
  `numberingChange` is dropped (its original value cannot be rebuilt). Removing content can leave a
  comment or note without its reference.
- **Formatting.** `beforeLines` / `afterLines` are not converted; legacy `w:hMerge` is not used in
  border resolution.
- **Compose.** `webSettings` stays the target's; `wp14:anchorId` / `editId` are not renumbered;
  `pageNumbering: 'continue' | 'restart'` affects the first imported section only; with
  `keepSourceFormatting`, table-style conditional formatting is not expanded; the `validate` option
  throws after the target was modified.
- **Fields.** Number words (CardText, OrdText, DollarText) are English only; SET / ASK do not rewrite
  their own result; STYLEREF, SECTION and TOC work in the body only; `=` formula defaults follow
  Word's documentation, not a measured Word result.
- **Forms.** Controls inside `mc:Fallback` are not updated; abbreviated month and day names come from
  `Intl` and can differ slightly from Word's.
- **Protection and encryption.** No RC4, extensible or IRM decryption; MD2 / MD4 / MD5 protection
  hashes never verify; `w:writeProtection` is not handled.
- **Fonts and shaping.** WOFF / WOFF2 are rejected; GPOS covers pair kerning only; no ligatures or
  mark positioning (the planned HarfBuzz plugin); Thai, Lao and similar scripts get no dictionary
  line breaking.
- **Layout.** Soft-hyphen width is ignored when fitting a line; `docGrid` line pitch is not applied;
  per-page footnote numbering restarts are not handled. Measured page-count misses are listed in
  [docs/fidelity.md](./docs/fidelity.md).
- **Graphics and PDF.** Text in shapes and charts never rotates; gradients are flattened to one
  colour; shadows and 3D effects are ignored; EMF clipping and RLE bitmaps are unsupported; bubble
  charts and log axes are missing; OMML uses fixed proportions instead of the font's MATH table; CFF
  fonts are embedded whole; tagged PDF output is basic.

## After v0.1.0

1. **Word verification** of the smoke pack, and fixes for anything Word repairs or rejects.
2. **Fidelity work**, driven by the oracle scoreboard: the known layout gaps above, then floats,
   frames and CJK. Each release must not lower the scores in `docs/fidelity.md`.
3. **`@turboword/harfbuzz`:** a `Shaper` backed by HarfBuzz (WASM) for ligatures, marks and complex
   scripts.
4. **`@turboword/fonts`:** the metric-compatible fonts (Carlito, Caladea, Liberation, Gelasio)
   packaged with a ready `FontResolver`. The repository becomes a pnpm workspace.
5. **The Open XML SDK validator in CI.**
6. **Experimental label off** for layout and PDF as soon as the scoreboard meets its targets.
7. **1.0:** API freeze, documentation site, benchmarks, security review and verification in Word
   for Windows.

## Risks

| Risk | Mitigation |
| --- | --- |
| Layout does not match Word closely enough | Labelled experimental; published scoreboard with free oracles (Word's stored page counts and `w:lastRenderedPageBreak` markers) plus LibreOffice; metric-compatible fonts with Word's vertical metrics; the table grid Word stored is reused |
| "Unreadable content" errors in Word from the long tail of compose and edit cases | `validatePackage` and LibreOffice on every output in the tests, the PowerTools and docxcompose fixtures, the docxcompose differential, and the Word smoke pack before every release; the Open XML SDK validator next |
| Not yet opened in Word | The smoke pack is a release gate for any Word-compatibility claim |
| Font licensing | No fonts in the core: a `FontResolver` abstraction; fonts shipped separately |
| Bugs in the custom XML DOM | Differential tests against xmldom, plus fuzzing |
| Bundle size creep | size-limit budgets per subpath (set from measured sizes) and a layer test that enforces import direction |
