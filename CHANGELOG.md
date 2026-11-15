# turboword

## 0.1.0

### Minor Changes

- 9052672: First release: lossless editing and composition of Word `.docx` files in pure TypeScript. One runtime dependency (`fflate`), ESM and CommonJS, Node 22+, Bun, Deno, browsers and workers.
  
  **Highlights**
  
  - **Merge documents without losing their headers and footers.** `appendDocument` / `composeDocuments` keep each source's sections, headers, footers, styles, numbering, notes, comments and bookmarks. In a 59-scenario differential test against docxcompose (inputs from docxcompose's and Open-XML-PowerTools' own test suites), turboword kept 338 of 338 header/footer slots (docxcompose 172) and 98.8% of the input pages in LibreOffice (docxcompose 65.5%).
  - **Lossless round trips.** A save without changes returns the exact input bytes. After an edit, every untouched zip entry keeps its original compressed bytes and CRC, and unknown elements, both `mc:AlternateContent` branches, comments and namespace declarations survive in edited parts.
  - **Find and replace across split runs**, keeping formatting, bookmarks and comment ranges, optionally as tracked changes: 100% on a 5,450-paragraph split-run corpus in four replace modes.
  - **Everything in one package:** editing, fields and TOCs, content controls, protection and Office encryption, plus experimental layout, PDF and SVG.
  
  **`turboword`: documents, editing, text, formatting**
  
  - Open and save: `readDocx`, `loadDocx` (Uint8Array, ArrayBuffer, Blob, ReadableStream, async iterable), `createDocx`, `writeDocx`, `docxBlob`, `docxStream`. Parts load lazily; output is deterministic.
  - Document model: stories (body, headers, footers, footnotes, endnotes, comments), paragraphs, runs, tables with a logical grid that resolves merged cells, sections with effective headers and footers (following "link to previous"), and core, app and custom document properties.
  - Editing:
    - text and run formatting;
    - paragraph split, insert, clone, replace and remove;
    - images in any story (PNG, JPEG, GIF, BMP, TIFF, WebP, SVG with raster fallback, EMF, WMF);
    - hyperlinks and bookmarks;
    - table rows and columns (insert, delete, clone), merge and unmerge, widths, heights, header rows;
    - section breaks and page setup;
    - header/footer creation, copying, unlinking, relinking, and even/odd pages.
  - Text: `findText` and `replaceText` across run boundaries, in any story including text boxes; matches that cross a field boundary are skipped (or throw). `acceptAllRevisions` and `rejectAllRevisions` handle every revision type, including moves, table rows and cells, and property changes.
  - Formatting resolution:
    - the full style cascade, with Word's toggle-property semantics;
    - theme fonts and colours;
    - table-style conditional formatting and border conflicts;
    - list labels counted the way Word counts them. They agree with LibreOffice on 97.9% of paragraphs, and every difference is a case where LibreOffice departs from Word.
  - Protection: `protectDocument` writes ISO/IEC 29500 `w:algorithmName` hashes (Word's legacy attributes are available as an option); `verifyProtectionPassword` reads both forms.
  - `validatePackage`: 13 checks for the usual causes of Word's "unreadable content" error, including unresolved relationships, duplicate drawing IDs, misplaced section properties and table cells without a paragraph.
  - Low-level access for anything not modelled yet: the OPC package, parts, relationships, content types, a lossless XML DOM (2.3–5.5× faster than xmldom) and a document-wide ID registry.
  
  **`turboword/compose`**
  
  - `appendDocument`, `composeDocuments`, `importBlocks`, and `insertDocument` at a text placeholder, bookmark, content control, paragraph, story position, table cell or header.
  - Styles: use the destination's, keep the source formatting, or keep different styles as renamed copies. Matching works by ID or name, so localized style IDs such as a German "Überschrift1" match Heading 1.
  - Lists: keep the source's (with a fresh `w:nsid`), merge with matching destination lists, or restart. Picture bullets are copied.
  - Footnotes and endnotes; comments, with all five comment parts kept in sync; bookmarks, with conflicting names renamed and REF, PAGEREF and hyperlink references rewritten.
  - Sections, with header/footer inheritance materialized so imported sections never inherit the target's headers. Document-wide even/odd headers are reconciled. Settings, fonts (including embedded ones) and the theme are merged.
  - Strict and transitional OOXML can be mixed.
  
  **`turboword/fields`**
  
  - `listFields`, `updateFields` and `updateToc`. Field codes are never altered; only the results are rewritten.
  - Evaluators: MERGEFIELD, IF, COMPARE, `=` formulas (functions, bookmark and table references), DATE/TIME, CREATEDATE/SAVEDATE/PRINTDATE, DOCPROPERTY, DOCVARIABLE, REF, NOTEREF, PAGEREF/PAGE/NUMPAGES (with page information from layout), SEQ, STYLEREF, FILLIN, ASK, SET, QUOTE, SYMBOL and more, plus custom evaluators.
  - Word's `\*`, `\#` and `\@` switch formatting, locale-aware.
  
  **`turboword/forms`**
  
  - Content controls of every type: text, rich text, checkbox, dropdown, combo box, date, picture, group and repeating sections.
  - Typed values, `fillContentControls` by tag or alias, data binding to custom XML through an XPath subset, and legacy FORMTEXT / FORMCHECKBOX / FORMDROPDOWN fields.
  
  **`turboword/crypto`**
  
  - `isEncrypted`, `decryptDocx` (Agile and Standard encryption) and `encryptDocx` (Agile, AES-256 / SHA-512), on Web Crypto only.
  - Interoperable with msoffcrypto-tool in both directions.
  
  **Experimental: `turboword/layout`, `turboword/pdf`, `turboword/svg`**
  
  - Fonts: TTF, OTF and TTC parsing; fonts embedded in the document; metric-compatible substitutes that use Word's vertical metrics; PDF base-14 fallback. No fonts are bundled: pass a `FontResolver`.
  - Text: Unicode 16 line breaking and bidi, passing 100% of Unicode's conformance tests.
  - Paragraph layout with Word's tab, justification and line-height rules. On the paragraph-level test set, line breaks match LibreOffice on 97.9% of lines in Word 2013+ documents.
  - Pagination: keep rules, widow and orphan control, columns, section types, headers and footers per page, footnotes and endnotes, tables (borders, merges, repeated header rows, rows split across pages), floating objects with text wrap, and page-number fields.
  - Graphics: 62 DrawingML preset shapes, charts drawn from their cached values, EMF/WMF pictures as vectors, and equations (OMML).
  - Rendering:
    - `toPdf` / `renderPdf`: deterministic PDF with subset fonts, searchable text, images, links and an outline;
    - `renderSvg`: SVG pages.
  - Measured accuracy on 165 Word-saved files: exact page count on 100% of prose and table/float documents and 95.2% of documents with graphics; line match against LibreOffice across whole paginated documents 93.4% on prose. The experimental label comes off when the scoreboard in `docs/fidelity.md` reaches its targets.
  
  **Quality**
  
  - 1,862 tests, including fuzzing of the zip, XML and document readers.
  - Every output in the test suite passes `validatePackage` and converts in LibreOffice.
  - CI runs on Node 22, 24 and 26 on Linux and Windows, plus Bun and Deno.
  - A Word smoke pack (`test/smoke/word-pack.mjs`) is provided for manual checks in Microsoft Word.
  
  Known gaps are listed per area in `ROADMAP.md`.
