# Microsoft Word verification

Results of opening the Word smoke pack (`pnpm build && node test/smoke/word-pack.mjs`, written to
`out/pack/`) in Microsoft Word. Record every "Word found unreadable content" / repair prompt here
with the file and the Word version; any such prompt is a release blocker.

## 2026-11-16 — Word for Mac 16.114.1, turboword 0.1.0

Automated open-only pass: every file was opened read-only through AppleScript, its paragraphs and
pages were counted, and it was closed. A modal dialog (a repair prompt included) blocks AppleScript,
so any file that did not open within 60 s was recorded as blocked. Side-by-side visual comparison
with the fixtures was **not** done in this pass.

| Result | Files |
| --- | ---: |
| Opened, no dialog | 402 (+2 slow, confirmed by hand) |
| Blocked by a dialog (defective fixtures, see below) | 16 |
| Not generated (encrypted fixtures, by design) | 8 |

**No repair prompts were attributable to turboword.** Each of the 16 blocked files belongs to a
pair: its `noop/` copy is byte-identical to the original fixture, and that copy is blocked the same
way as the `dirty/` copy (re-serialized by turboword). The cause is in the fixtures themselves:

- Mail-merge main documents whose `w:mailMerge` data source is an external file on the fixture
  author's machine. Word asks for confirmation before running the SQL query:
  `fields/test_field_with_paragraph`, `fields/test_field_with_paragraph_and_nested_field`,
  `fields/test_skipif`, `fields/test_winword2010`.
- Open XML PowerTools DocumentBuilder templates that carry PowerTools' private
  `<pt:Insert Id="Templafy"/>` marker (not OOXML) directly in a header or footer; Word says "Word
  experienced an error trying to open the file":
  `compose/powertools/DB__HeadersFooters__Dest__Fax` (`footer2.xml`), `…__Letter` and
  `…__Letter_(run-level_insert_element)` (`header1.xml`). The compose tests replace the marker
  before validating (`test/compose.test.ts`, PowerTools DB009).
- `forms/open-xml-sdk/sdtRow`: a plain-text content control (`w:text`) wraps a whole table row,
  which plain-text controls can't do; Word offers "Word found unreadable content … recover?".

Checked by hand afterwards: `resolve/poi-bug59058` is only slow to open (a long table, beyond the
60 s limit). Both the original and turboword's `dirty/` copy open without a prompt and look the
same.

The two created/edited outputs opened cleanly: `created-paragraphs.docx` (7 paragraphs, 1 page) and
`properties-edited.docx` (12 paragraphs, 2 pages).

Still to do by hand: compare a sample of the `dirty/` files side by side with their fixtures.
