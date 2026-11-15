# Layout fidelity

`turboword/layout` (and the PDF/SVG renderers built on it) is **experimental until this scoreboard
meets its targets** (exact page count ≥ 85% and line match ≥ 95% on Tier A), not until a version
number. This page
states how close its pagination is to Microsoft Word, how that is measured, and where it is known
to differ. The numbers below are asserted by `test/layout-oracle.test.ts`, so a change that makes
them worse fails CI.

## How it is measured

There is no Word in CI, so the scoreboard uses free oracles (plan decision D8):

- **Word's page count.** Every Word-saved .docx records the page count Word computed when it saved
  the file (`docProps/app.xml` `<Pages>`). The score is the share of files where turboword makes
  exactly that many pages.
- **Word's page breaks.** Word writes `w:lastRenderedPageBreak` into the run where each page
  started. The score is the share of these (soft) breaks that turboword makes at the same paragraph:
  before it, or inside it when Word broke inside it. Word writes the marker only into runs, so a
  page that starts with an empty paragraph has no marker; markers right after a hard break are not
  counted.
- **LibreOffice line match** (the secondary oracle). LibreOffice converts each file to PDF with
  the same metric-compatible fonts the tests use (Carlito for Calibri, Caladea for Cambria,
  Liberation Sans, Serif and Mono for Arial, Times New Roman and Courier New). The text rows of
  each page are compared: words grouped by baseline, letters and digits only, matched on the same
  page. Score = matched rows / max(our rows, LibreOffice's rows). LibreOffice is not Word, but on
  this corpus it reproduces Word's page count for 163 of 165 files. Where both engines are right,
  their rows agree.

Layout runs with the open-source test fonts only (`testFontResolver()`: Carlito, Caladea,
Liberation; regular and bold, no italic faces). Real Calibri or Cambria are never used.

**Corpus.** `test/fixtures/layout/corpus.json` lists every Word-saved fixture in the repository,
deduplicated by content, whose app.xml records a page count: 165 files from docxcompose,
Open-XML-PowerTools, Apache POI, Open XML SDK, python-docx, docx-mailmerge, docx4j, Tika, mammoth
and msoffcrypto-tool. One file is excluded because its `<Pages>` is stale (docx4j rewrote the body
after Word saved it). Files are tiered by content (body, headers and footers):

| Tier | Content | Files |
|---|---|---|
| A | prose: paragraphs, lists, headers/footers, notes, fields, sections | 104 |
| B | tables, pictures, floating objects, frames, multiple columns | 40 |
| C | DrawingML/VML shapes, charts, SmartArt, text boxes, math, OLE objects | 21 |

## Scoreboard

Measured on 2026-11-15 with all layout modules at their v0.1.0 state:

| Tier | Page count exact | Word's page breaks reproduced | Line match vs LibreOffice |
|---|---|---|---|
| A | **100%** (104/104) | 100% (3/3) | **93.4%** (846 rows) |
| B | 100% (40/40) | 45.5% (5/11) | 90.1% (646 rows) |
| C | 95.2% (20/21) | 50% (2/4) | 78.2% (357 rows) |

**Plan targets (v0.5 exit):** Tier A page count exact ≥ 85% and line match ≥ 95%.

- **Page count: met, but the test is weak.** Most fixtures are one or two pages long, so a page
  count is easy to get right. Only 18 Word-recorded soft page breaks exist in the whole corpus.
  The page-count score will drop when longer documents are added, and adding them (permissively
  licensed, Word-saved, multi-page) is the most useful next step for this scoreboard.
- **Line match: not met (93.4%).** The misses are mostly in measurement and content, not in
  pagination. Each item below is shared with another layout module:
  - Horizontal measurement differences make a line wrap one word earlier or later. Small
    condensed fonts and kerning are typical causes.
  - Field result text differs. turboword shows the stored result of IF, REF and DATE fields,
    while LibreOffice sometimes shows a different one.
  - Tracked changes are shown differently (LibreOffice displays deleted text).
  - A row that differs only in a list label glyph still matches, because only letters and digits
    are compared.

## Word behaviours implemented

Pagination (`src/layout/paginate.ts`, `src/layout/page/`, `src/layout/table.ts`) follows these
Word rules. The ones marked (oracle) were confirmed against Word's own page-break markers or page
counts in the corpus:

- **Paragraph spacing collapses.** The gap between two paragraphs is the larger of the first
  one's space after and the second one's space before, not their sum (oracle). Contextual spacing
  drops both between paragraphs of the same style. HTML autospacing (`beforeAutospacing`,
  `afterAutospacing`) is dropped between two list paragraphs and at the top and bottom of a table
  cell (oracle: apache-poi/delins.docx).
- **Space before at the top of a page** is dropped after an automatic page or column break and kept
  after a hard break: a page break, `pageBreakBefore`, a section break or the first page.
  LibreOffice also drops it after `pageBreakBefore`; Word does so only with the
  `suppressSpBfAfterPgBrk` compatibility option, which is not read yet.
- **Keeping lines together.**
  - `keepNext` chains, `keepLines` and widow/orphan control (2 lines) all become "avoid a break
    here" between lines.
  - When a strip does not fit, the paginator moves back to the last line a break may precede.
  - A chain longer than a page breaks where it overflows, as Word does.
- **Hard breaks and sections.**
  - `pageBreakBefore` at the top of a page adds no page.
  - Section types: nextPage, continuous (a new band of columns on the same page), evenPage and
    oddPage (a blank page is inserted when the parity is wrong), nextColumn.
- **Columns.**
  - Equal or explicit widths, with an optional separator line.
  - Columns are balanced before a continuous section break.
  - Like Word, the last page of the document is *not* balanced: Word balances it only when the
    section ends with a continuous section break.
- **Headers and footers.**
  - Header selection: first page when `titlePg`, even pages when `evenAndOddHeaders`, blank when
    the selected header does not exist.
  - The body starts at max(top margin, header distance + header height). The footer pushes the
    bottom up the same way. A negative margin is "exact": headers and footers do not push it.
  - PAGE, NUMPAGES and SECTIONPAGES in headers and footers are laid out per page, with that page's
    values.
- **Page numbers.** `pgNumType/@w:start` restarts numbering, otherwise it continues. The format
  comes from `pgNumType/@w:fmt` (through resolve's `formatNumber`) or from the field's `\*` switch.
- **Page fields in the body.** PAGE, NUMPAGES, SECTIONPAGES and PAGEREF are evaluated where they
  land. When a value's text differs from what was measured, the document is laid out again, up to
  `maxPasses` (default 3). After the last pass, any remaining difference is re-shaped in place.
- **Footnotes and endnotes.**
  - A line's footnotes are reserved on its page: the line moves on when its notes do not fit.
    Notes that cannot fit even on an empty page continue on the next page below a full-width
    continuation separator. The normal separator is 144pt wide.
  - Endnotes come after the last section's content, in reference order, below a separator.
- **Tables.**
  - The stored grid is used as is. Autofit (min/max content widths) applies only when there is no
    usable grid.
  - Row height = content + cell margins + the horizontal border (oracle: rows are border-width
    taller than their text in Word and LibreOffice). `trHeight` atLeast/exact are honoured.
  - vMerge spans grow their last row. Borders collapse: the heavier of two adjacent borders wins,
    and table borders fill in.
  - Rows split across pages unless `cantSplit` or an exact height is set. Header rows repeat on
    every page and column.
  - Tables nest up to 32 levels; deeper levels are skipped with a `LAYOUT_TABLE_DEPTH` warning.
- **Floating objects.**
  - Positioned relative to page, margin, column, paragraph, line, character or the margins (inside
    and outside by page parity), by offset or by alignment.
  - `topAndBottom` pushes text below the object. `square`, `tight` and `through` are rectangular:
    lines beside the object use its wider free side (per line, through `layoutParagraph`'s
    `insets`), and gaps narrower than 36pt are skipped.
  - Objects behind text are drawn first, objects in front of text last. Text boxes are laid out
    inside their box with DrawingML's default insets. Floating EMF/WMF pictures are drawn as
    vectors.
- **Page `vAlign`:** top, center, bottom, and both (the free space is spread between paragraphs).

## Known gaps

- Lines beside a square-wrapped float are positioned from estimated line heights, and a page break
  inside such a paragraph does not re-wrap it.
- `tight` and `through` wrap use the bounding box, not the wrap polygon.
- Floating tables (`tblpPr`), frames (`framePr`), line numbering, page borders and document grid
  snapping (`docGrid`) are not laid out specially: frames and floating tables flow inline.
- A float anchored in a paragraph that moves to the next page moves with it, but the text laid out
  around it is not re-wrapped.
- Footnotes in multi-column pages sit at the bottom of their column. A page with several column
  bands puts all its notes in one area.
- RTL tables (`bidiVisual`) are not mirrored.
- Fields inside text boxes keep their stored result.

## Reproducing the numbers

```sh
npx vitest run test/layout-oracle.test.ts                     # scoreboard + assertions
FIDELITY_REPORT=1 npx vitest run test/layout-oracle.test.ts   # ... and every miss
FX=compose/powertools/DB007-Notes.docx npx vitest run test/layout-oracle.test.ts -t dumps
node test/fixtures/layout/record-libreoffice.mjs              # re-record corpus + LibreOffice data
```

Re-record after adding fixtures (it needs LibreOffice and poppler's `pdftotext`), then update the
baseline in the test and the scoreboard above.
