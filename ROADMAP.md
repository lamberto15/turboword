# Roadmap

Where turboword stands and what comes next. Update this file with every release.

The order is deliberate: a lossless core first, then composition (the most requested missing feature
in permissive libraries), then editing, fields and forms, then layout and PDF.

**Layout and PDF (v0.5 onwards) stay experimental until 1.0.** Their APIs and output may change in any
minor release.

## Versions

| Version | Deliverables | Exit criteria |
| --- | --- | --- |
| **v0.1 Lossless core** (in progress) | Scaffolding, CI and release; zip reader and raw-copy writer; XML DOM; OPC package; `readDocx` / `loadDocx` / `writeDocx` / `createDocx`; stories, paragraphs and runs; read-only table grid; sections with effective headers and footers; visible text (`text()`); document properties; `validatePackage` v1. Published to npm to claim the name. | Every fixture saves byte-identical on a no-op; force-dirty parts reparse DOM-equal and open in LibreOffice; fuzzing, 300 iterations per target; DOM at least 2x faster than xmldom; the Word smoke pack opens with no repair prompt |
| **v0.2 Compose I** | `ImportContext`, part-graph copier, document-wide ID registry, `mc:` prefix reconciliation; three style modes; numbering keep / merge / restart; footnotes and endnotes; comments (all five parts); bookmarks; sections with headers and footers; settings, fonts and theme. `appendDocument`, `composeDocuments`, `importBlocks` in `turboword/compose`. | Every docxcompose and Open-XML-PowerTools DocumentBuilder fixture combination passes; Open XML SDK validator reports no new errors; LibreOffice page count equals the sum of the inputs in at least 95% of cases; header and footer text on each page matches its source; beats docxcompose in the differential test (headers kept) |
| **v0.3 Edit and insert** | Text engine (find and replace across split runs); `insertDocument` at a placeholder, bookmark, content control, paragraph or table cell; table operations; images in any story; hyperlink and bookmark API; header and footer editing; style cascade and list labels; accept / reject revisions. | Split-run corpus passes 100%; list labels match LibreOffice in at least 98% of cases |
| **v0.4 Fields, forms, protection, encryption** | Field index; switch formatting; evaluators (MERGEFIELD, IF, DATE, DOCPROPERTY, REF, SEQ, `=`, FILLIN, ...); TOC structure; content controls and data binding; legacy form fields; document protection; `turboword/crypto`. | Field results equal Word's after F9; files encrypted by Word round-trip |
| *Layout spike* (during v0.2 and v0.3) | Font parsing and greedy line breaking on 20 prose documents, compared with PDFs exported by Word. | At least 95% identical lines decides go / no-go for v0.5 |
| **v0.5 Layout foundations** (experimental) | Font stack, UAX #14 line breaking, layout IR, paragraph layout, pagination (keep rules, section types), headers and footers with PAGE / NUMPAGES, tables from the stored grid, `paginate` / `pageCount`, SVG debug renderer, fidelity scoreboard. | Tier A documents: exact page count in at least 85%, line match at least 95% |
| **v0.6 PDF I** (experimental) | PDF writer, font subsetting, images, links, outline, TOC fixed point, `toPdf`. | `qpdf --check` clean; text extracted with pdf.js equals the visible text; deterministic output |
| **v0.7 Layout II** (experimental) | Footnotes, floating objects, text boxes, frames, columns, table autofit, vertical merges and border conflicts, table conditional formatting, document grid, compatibility flags, `@turboword/fonts`. | Tier B at least 75%, tier A at least 92% |
| **v0.8 Rendering breadth** (experimental) | DrawingML shapes, charts from cached values, an EMF subset, OMML equations, bidirectional text, CJK, column balancing, a HarfBuzz plugin. | Tier C renders with no placeholders; RTL and CJK documents pass the oracle |
| **v1.0** | API freeze, documentation, benchmarks, security review, verification in Word for Windows. | Fidelity scoreboard published |

Fidelity tiers group test documents by difficulty, from prose and simple tables (tier A) to shapes,
charts and equations (tier C); the exact sets are fixed with the scoreboard in v0.5. The oracle is
Word itself: the page count in `docProps/app.xml` and the `w:lastRenderedPageBreak` markers of
Word-saved files, plus PDFs exported by Word.

## Risks

| Risk | Mitigation |
| --- | --- |
| Layout does not match Word closely enough | Go / no-go spike before v0.5; published fidelity tiers; metric-compatible fonts; reuse of the table grid Word stored; free oracles from Word-saved files |
| "Unreadable content" errors in Word from the long tail of compose cases | Open XML SDK validator in CI, the PowerTools fixtures, the docxcompose differential test, and the Word smoke pack before every release |
| Font licensing | No fonts in the core: a `FontResolver` abstraction, with fonts shipped separately in `@turboword/fonts` |
| Bugs in the custom XML DOM | Differential tests against xmldom, plus fuzzing |
| Bundle size creep | size-limit budgets per subpath and a layer test that enforces import direction |
