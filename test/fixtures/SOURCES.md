# Test fixtures: sources and licences

Every `.docx` here is a real file copied byte for byte from a permissively licensed open-source
project (MIT, BSD or Apache-2.0 only). Each source has its own folder, and its licence text sits
next to the files as `<source>/LICENSE` (plus `NOTICE` for Apache-2.0). The files are test data
only and are not part of the published npm package.

`expected.json` records, per file, Word's page count (`docProps/app.xml` `<Pages>`), the producing
application, and the features the file exercises. `test/roundtrip.test.ts` checks that this file,
`expected.json` and the folder contents stay in sync, and that the size budget holds
(each file ≤ 300 KB, ≤ 3 MB in total).

All source URLs are pinned to a commit. Every file was checked against the upstream blob
(`git hash-object` equals the GitHub blob SHA at that commit).

## Sources

| Folder | Project | Licence | Pinned commit |
| --- | --- | --- | --- |
| `docx-mailmerge/` | [iulica/docx-mailmerge](https://github.com/iulica/docx-mailmerge) | MIT | `7e2aeb29443f50025c3b67881b993281b3edb77d` |
| `docxcompose/` | [4teamwork/docxcompose](https://github.com/4teamwork/docxcompose) | MIT | `b8bac82507786aa350178ad39bfeb9555b3f7429` |
| `python-docx/` | [python-openxml/python-docx](https://github.com/python-openxml/python-docx) | MIT | `e45454602b53e8e572b179ccf1c91093ec9f4ed7` |
| `apache-poi/` | [apache/poi](https://github.com/apache/poi) | Apache-2.0 | `47dcff769bb395034dddfd59c4609659b6dbc6a4` |
| `open-xml-sdk/` | [dotnet/Open-XML-SDK](https://github.com/dotnet/Open-XML-SDK) | MIT | `431ab05cf160248cc3885a4a766026d4f8243792` |

## Files

| File | Source (pinned) | Licence | Producing app + version | Features covered |
| --- | --- | --- | --- | --- |
| `docx-mailmerge/test_footnote_header_footer.docx` | [tests/test_footnote_header_footer.docx](https://github.com/iulica/docx-mailmerge/blob/7e2aeb29443f50025c3b67881b993281b3edb77d/tests/test_footnote_header_footer.docx) | MIT | Microsoft Office Word 16.0000 | Default, first-page and even-page headers and footers (`titlePg`, `evenAndOddHeaders`); a footnote; MERGEFIELD as `fldSimple` in the body and as complex fields in every header and footer; customXml. 3 pages. |
| `docxcompose/docproperties_header_footer_3_sections.docx` | [tests/docs/docproperties_header_footer_3_sections.docx](https://github.com/4teamwork/docxcompose/blob/b8bac82507786aa350178ad39bfeb9555b3f7429/tests/docs/docproperties_header_footer_3_sections.docx) | MIT | Microsoft Office Word 16.0000 | 3 sections with different header and footer sets: section 1 has default, first and even; section 2 its own default pair; section 3 inherits. DOCPROPERTY complex fields in headers and footers; `docProps/custom.xml`. 2 pages. |
| `docxcompose/header_with_image.docx` | [tests/docs/header_with_image.docx](https://github.com/4teamwork/docxcompose/blob/b8bac82507786aa350178ad39bfeb9555b3f7429/tests/docs/header_with_image.docx) | MIT | Microsoft Office Word 16.0000 | DrawingML picture in a header (`word/_rels/header1.xml.rels` → `media/image1.jpg`); stored (method 0) zip entries. |
| `docxcompose/numberings.docx` | [tests/docs/numberings.docx](https://github.com/4teamwork/docxcompose/blob/b8bac82507786aa350178ad39bfeb9555b3f7429/tests/docs/numberings.docx) | MIT | Microsoft Macintosh Word 15.0000 | Several multi-level numbered lists (`numbering.xml`, `w:numPr` at levels 0-2); thumbnail; stored zip entries. |
| `docxcompose/content_controls.docx` | [tests/docs/content_controls.docx](https://github.com/4teamwork/docxcompose/blob/b8bac82507786aa350178ad39bfeb9555b3f7429/tests/docs/content_controls.docx) | MIT | Microsoft Office Word 16.0000 | 4 content controls (`w:sdt`): plain text (single-line, multiline, empty) and rich text, with aliases; glossary document part (`word/glossary/*`); German placeholder text. |
| `docxcompose/smart_art.docx` | [tests/docs/smart_art.docx](https://github.com/4teamwork/docxcompose/blob/b8bac82507786aa350178ad39bfeb9555b3f7429/tests/docs/smart_art.docx) | MIT | Microsoft Office Word 16.0000 | SmartArt: `dgm:relIds` (rIds in `r:dm`/`r:lo`/`r:qs`/`r:cs` attributes) and the five diagram parts, including the Microsoft `diagramDrawing` relationship. |
| `python-docx/tbl-cell-access.docx` | [features/steps/test_files/tbl-cell-access.docx](https://github.com/python-openxml/python-docx/blob/e45454602b53e8e572b179ccf1c91093ec9f4ed7/features/steps/test_files/tbl-cell-access.docx) | MIT | Microsoft Macintosh Word 14.0000 | Tables with merged cells: horizontal (`w:gridSpan`) and vertical (`w:vMerge`); thumbnail; stored zip entries. |
| `python-docx/comments-rich-para.docx` | [features/steps/test_files/comments-rich-para.docx](https://github.com/python-openxml/python-docx/blob/e45454602b53e8e572b179ccf1c91093ec9f4ed7/features/steps/test_files/comments-rich-para.docx) | MIT | Microsoft Office Word 16.0000 | 4 comments with rich content: a hyperlink, an inline image (`word/_rels/comments.xml.rels`), a character style and a heading style; `commentsExtended.xml`, `commentsIds.xml`. |
| `apache-poi/delins.docx` | [test-data/document/delins.docx](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/delins.docx) | Apache-2.0 | Microsoft Office Word 12.0000 | Tracked changes: `w:ins` (32), `w:del` (4) and `w:rPrChange` (46) format changes; numbering; hyperlinks; complex fields; customXml. |
| `apache-poi/shapes-with-text.docx` | [test-data/document/shapes-with-text.docx](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/shapes-with-text.docx) | Apache-2.0 | Microsoft Office Word 16.0000 | Text boxes and shapes with text: `wps:txbx` in `mc:Choice` with a VML `v:textbox` fallback in `mc:Fallback`; group shapes (`wpg:wgp` / `v:group`), an ellipse and a rectangle. |
| `apache-poi/bug69628.docx` | [test-data/document/bug69628.docx](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/bug69628.docx) | Apache-2.0 | LibreOffice 25.2.1.2 (macOS AArch64), AppVersion 15.0000 | LibreOffice-saved: numbered list and a table; zip entries use data descriptors (flag bit 3); `[Content_Types].xml` is the last entry. |
| `apache-poi/libreoffice-comment.docx` | [test-data/document/comment.docx](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/comment.docx) (renamed) | Apache-2.0 | LibreOffice 6.0.7.3 (Linux x86-64) | LibreOffice-saved: a two-paragraph comment on an empty paragraph; no theme part; data descriptors; `[Content_Types].xml` last. |
| `open-xml-sdk/TOC2.docx` | [test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/wordprocessing/Toc/TOC2.docx](https://github.com/dotnet/Open-XML-SDK/blob/431ab05cf160248cc3885a4a766026d4f8243792/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/wordprocessing/Toc/TOC2.docx) | MIT | Microsoft Office Word 12.0000 | Table of contents: `TOC \o "1-3" \h \z \u` complex field with `w:hyperlink w:anchor` entries and `PAGEREF _Toc… \h` fields; `_Toc` bookmarks on headings; CJK text; customXml. |
| `open-xml-sdk/TableFntEdnt.docx` | [test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/wordprocessing/table/TableFntEdnt/TableFntEdnt.docx](https://github.com/dotnet/Open-XML-SDK/blob/431ab05cf160248cc3885a4a766026d4f8243792/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/v2FxTestFiles/wordprocessing/table/TableFntEdnt/TableFntEdnt.docx) | MIT | Microsoft Office Word 12.0000 | 3 footnotes and 3 endnotes, all referenced from table cells; customXml. |
| `open-xml-sdk/strict-CommentRangeStart.docx` | [test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Word/CommentRangeStart.docx](https://github.com/dotnet/Open-XML-SDK/blob/431ab05cf160248cc3885a4a766026d4f8243792/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Word/CommentRangeStart.docx) (renamed) | MIT | Microsoft Office Word 14.0000, saved as Strict Open XML (see note) | **Strict OOXML** (`purl.oclc.org/ooxml` namespaces and relationship types, including `…/relationships/extendedProperties`); a comment with `commentRangeStart/End`; UTF-8 BOM on every part. |
| `open-xml-sdk/strict-footer-bookmarkStart.docx` | [test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Word/ftr - Footer - bookmarkStart - Bookmark Start.docx](https://github.com/dotnet/Open-XML-SDK/blob/431ab05cf160248cc3885a4a766026d4f8243792/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestDataStorage/O14ISOStrict/Word/ftr%20-%20Footer%20-%20bookmarkStart%20-%20Bookmark%20Start.docx) (renamed) | MIT | Microsoft Office Word 14.0000, saved as Strict Open XML (see note) | **Strict OOXML** with default, first and even headers and footers (Strict `header`/`footer` relationship types); a bookmark inside a footer; `stylesWithEffects.xml`; UTF-8 BOM on every part. |

**Note on the Strict files.** Word 2010 saved them as Strict Open XML (Word 2010 does not write
`w:conformance="strict"`). The package was then rewritten by .NET `System.IO.Packaging`, which
added a UTF-8 BOM to every part and produced the `.rels` and `[Content_Types].xml` layout. The
part content is Word's. Word 2010's Strict mode only changes namespaces and relationship types, so
for that purpose these files are representative.

## Synthetic Strict variants

`test/helpers/fixtures.ts` also exports `toStrict()` and `strictFixtures()`. They rewrite a
transitional fixture into Strict OOXML: the namespace URIs are swapped for their `purl.oclc.org`
twins (the same table as `src/xml/ns.ts`), the officeDocument relationship types become Strict
ones (`extended-properties` → `extendedProperties`, `custom-properties` → `customProperties`), and
`w:conformance="strict"` is added. The results are **synthetic** and named
`synthetic-strict/<original>`. They are not schema-valid Strict, because attribute value types
such as `ST_OnOff` and percentages are left as they are. They exist to run richer content (header
images, numbering, merged cells, sections) through the Strict code paths. They are never written
to disk.

## Known gaps

No permissively licensed file was found for these, so they are not covered yet:

- a Word 2013+ Strict file (`w:conformance="strict"`); the synthetic variants approximate one;
- a modern TOC wrapped in a `docPartGallery="Table of Contents"` content control (the TOC fixture
  is Word 2007, a bare field);
- tracked moves (`w:moveFrom`/`w:moveTo`) and table row/cell revisions;
- a VML header image or watermark (the header image is DrawingML only);
- UTF-16 encoded parts, `altChunk`, and encrypted or password-protected packages.

## Adding a fixture

1. Use only MIT, BSD or Apache-2.0 sources. Copy the licence (and any NOTICE) into
   `<source>/LICENSE`.
2. Keep each file ≤ 300 KB and the corpus ≤ 3 MB. Do not edit the file: fixtures must stay
   byte-identical to upstream.
3. Add a row to the table above (pinned URL, licence, app + version from `docProps/app.xml`,
   features) and an entry to `expected.json`.
