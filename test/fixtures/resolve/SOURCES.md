# test/fixtures/resolve — sources

Numbering-heavy documents for the style/numbering resolver (`test/list-labels.test.ts`,
`test/resolve-sweep.test.ts`), copied unmodified from permissively licensed projects (only the
file names were prefixed with the source). They are loaded with `fixture('resolve/<file>')` and
are not part of the curated corpus in `../expected.json`.

Each file is pinned to the upstream commit it was taken from. "Verified" means `git hash-object
<file>` equals the blob SHA GitHub reports for that path at that commit (checked with the GitHub
trees API on 2026-11-15).

## Commits

| Upstream | Ref fetched | Commit |
|---|---|---|
| [apache/poi](https://github.com/apache/poi) | `trunk` | `47dcff769bb395034dddfd59c4609659b6dbc6a4` |
| [plutext/docx4j](https://github.com/plutext/docx4j) | `VERSION_17_3_7` | `836eab29ea8b47536ba9320e9bdf3da434de1d2f` |
| [apache/tika](https://github.com/apache/tika) | `main` | `eedcbd14a72f536f8bc1ef3d669037540a2cc196` |
| [ShayHill/docx2python](https://github.com/ShayHill/docx2python) | `master` | `66b1ae3e24ffa7040ed5244490ce2872ac2e86bb` |
| [mwilliamson/mammoth.js](https://github.com/mwilliamson/mammoth.js) | `master` | `5171a9072f8b7343ab1c9ed3a15d6d372cb59b67` |

## Documents

| File | Upstream (pinned) | Licence | Blob SHA | Verified |
|---|---|---|---|---|
| `docx2python-created-in-pages-bulleted-lists.docx` | [ShayHill/docx2python: `created-in-pages-bulleted-lists.docx`](https://github.com/ShayHill/docx2python/blob/66b1ae3e24ffa7040ed5244490ce2872ac2e86bb/tests/resources/created-in-pages-bulleted-lists.docx) | MIT | `48a746c53308` | yes |
| `docx2python-example_numbering.docx` | [ShayHill/docx2python: `example_numbering.docx`](https://github.com/ShayHill/docx2python/blob/66b1ae3e24ffa7040ed5244490ce2872ac2e86bb/tests/resources/example_numbering.docx) | MIT | `3702d6dbb571` | yes |
| `docx4j-25a-bullets.docx` | [plutext/docx4j: `25a-bullets.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/resources/revisions/core-ts/check25/25a-bullets.docx) | Apache-2.0 | `395d1b3bc4b9` | yes |
| `docx4j-Eastern_Arabic_numbering.docx` | [plutext/docx4j: `Eastern Arabic numbering.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docs/developer/Eastern%20Arabic%20numbering.docx) | Apache-2.0 | `52526728d69e` | yes |
| `docx4j-NumberingImplicitNumId.docx` | [plutext/docx4j: `NumberingImplicitNumId.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/resources/NumberingImplicitNumId.docx) | Apache-2.0 | `1a4bf87c4500` | yes |
| `docx4j-NumberingIndents.docx` | [plutext/docx4j: `NumberingIndents.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/NumberingIndents.docx) | Apache-2.0 | `46e8d577990b` | yes |
| `docx4j-article-section-NotIsLgl.docx` | [plutext/docx4j: `article-section-NotIsLgl.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/article-section-NotIsLgl.docx) | Apache-2.0 | `ac665583f3a1` | yes |
| `docx4j-article-section-isLgl.docx` | [plutext/docx4j: `article-section-isLgl.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/article-section-isLgl.docx) | Apache-2.0 | `adbafb696be5` | yes |
| `docx4j-cn-1-labels-word365.docx` | [plutext/docx4j: `cn-1-labels-word365.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/resources/clause-numbers/cn-1-labels-word365.docx) | Apache-2.0 | `1beaad664c6f` | yes |
| `docx4j-numbering-linked-level-other-style.docx` | [plutext/docx4j: `numbering-linked-level-other-style.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/numbering-linked-level-other-style.docx) | Apache-2.0 | `8ee1ec5d486d` | yes |
| `docx4j-numbering-shared-abstract-styles.docx` | [plutext/docx4j: `numbering-shared-abstract-styles.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/numbering-shared-abstract-styles.docx) | Apache-2.0 | `fdf61a77005c` | yes |
| `docx4j-numbering-stories.docx` | [plutext/docx4j: `numbering-stories.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/numbering-stories.docx) | Apache-2.0 | `0d7562f43d91` | yes |
| `docx4j-numbering-unlinked-level-claim.docx` | [plutext/docx4j: `numbering-unlinked-level-claim.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/numbering-unlinked-level-claim.docx) | Apache-2.0 | `33dec9eff0be` | yes |
| `docx4j-numbering_indentation.docx` | [plutext/docx4j: `numbering_indentation.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/resources/numbering_indentation.docx) | Apache-2.0 | `381ccaab3368` | yes |
| `docx4j-numbering_indentation_firstline.docx` | [plutext/docx4j: `numbering_indentation_firstline.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/resources/numbering_indentation_firstline.docx) | Apache-2.0 | `dab90be315a1` | yes |
| `docx4j-startOverride.docx` | [plutext/docx4j: `startOverride.docx`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/docx4j-core-tests/src/test/java/org/docx4j/model/listnumbering/startOverride.docx) | Apache-2.0 | `c5aa2a39ea42` | yes |
| `mammoth-simple-list.docx` | [mwilliamson/mammoth.js: `simple-list.docx`](https://github.com/mwilliamson/mammoth.js/blob/5171a9072f8b7343ab1c9ed3a15d6d372cb59b67/test/test-data/simple-list.docx) | BSD-2-Clause | `c2cd5ae96eba` | yes |
| `poi-60316.docx` | [apache/poi: `60316.docx`](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/60316.docx) | Apache-2.0 | `78b2b1d846cb` | yes |
| `poi-61787.docx` | [apache/poi: `61787.docx`](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/61787.docx) | Apache-2.0 | `0c4c9353cffe` | yes |
| `poi-ComplexNumberedLists.docx` | [apache/poi: `ComplexNumberedLists.docx`](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/ComplexNumberedLists.docx) | Apache-2.0 | `b4409c9d1595` | yes |
| `poi-Numbering.docx` | [apache/poi: `Numbering.docx`](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/Numbering.docx) | Apache-2.0 | `d5605c9f7861` | yes |
| `poi-NumberingWOverrides.docx` | [apache/poi: `NumberingWOverrides.docx`](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/NumberingWOverrides.docx) | Apache-2.0 | `48a7e2c25191` | yes |
| `poi-bug59058.docx` | [apache/poi: `bug59058.docx`](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/bug59058.docx) | Apache-2.0 | `08808daec82a` | yes |
| `poi-bug65738.docx` | [apache/poi: `bug65738.docx`](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/bug65738.docx) | Apache-2.0 | `5ef040e98817` | yes |
| `poi-form_footnotes.docx` | [apache/poi: `form_footnotes.docx`](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/test-data/document/form_footnotes.docx) | Apache-2.0 | `70abb60c9a8f` | yes |
| `tika-testWORD_numbered_list.docx` | [apache/tika: `testWORD_numbered_list.docx`](https://github.com/apache/tika/blob/eedcbd14a72f536f8bc1ef3d669037540a2cc196/tika-parsers/tika-parsers-standard/tika-parsers-standard-modules/tika-parser-microsoft-module/src/test/resources/test-documents/testWORD_numbered_list.docx) | Apache-2.0 | `51d31de55f0b` | yes |
| `tika-testWORD_override_list_numbering.docx` | [apache/tika: `testWORD_override_list_numbering.docx`](https://github.com/apache/tika/blob/eedcbd14a72f536f8bc1ef3d669037540a2cc196/tika-parsers/tika-parsers-standard/tika-parsers-standard-modules/tika-parser-microsoft-module/src/test/resources/test-documents/testWORD_override_list_numbering.docx) | Apache-2.0 | `23368f4bfb31` | yes |

All 28 documents verified. The docx4j probes `docx4j-numbering-*.docx` carry Word 365 goldens in
docx4j's own tests (`LevelOwnerTest`, `NumberingStoriesTest`); `docx4j-startOverride.docx` and
`docx4j-article-section-*.docx` carry `[expect]…[/expect]` labels in their text.

## Licence files

| File | Upstream (pinned) | Verified |
|---|---|---|
| `LICENSE.apache-2.0` | [plutext/docx4j `legals/LICENSE`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/legals/LICENSE) — the Apache-2.0 text, covering the POI, docx4j and Tika files | yes |
| `NOTICE.apache-poi` | [apache/poi `legal/NOTICE`](https://github.com/apache/poi/blob/47dcff769bb395034dddfd59c4609659b6dbc6a4/legal/NOTICE) | yes |
| `NOTICE.docx4j` | [plutext/docx4j `legals/NOTICE`](https://github.com/plutext/docx4j/blob/836eab29ea8b47536ba9320e9bdf3da434de1d2f/legals/NOTICE) | yes |
| `NOTICE.tika` | [apache/tika `NOTICE.txt`](https://github.com/apache/tika/blob/eedcbd14a72f536f8bc1ef3d669037540a2cc196/NOTICE.txt) | yes |
| `LICENSE.docx2python.mit` | [ShayHill/docx2python `LICENSE.txt`](https://github.com/ShayHill/docx2python/blob/66b1ae3e24ffa7040ed5244490ce2872ac2e86bb/LICENSE.txt) | yes |
| `LICENSE.mammoth.bsd-2-clause` | [mwilliamson/mammoth.js `LICENSE`](https://github.com/mwilliamson/mammoth.js/blob/5171a9072f8b7343ab1c9ed3a15d6d372cb59b67/LICENSE) | yes |

Apache POI's own `legal/LICENSE` appends third-party notices to the Apache-2.0 text; only the
Apache-2.0 text is kept here, since the test documents are covered by it alone.

## `lo/` — LibreOffice references

`lo/*.txt` are LibreOffice plain-text exports used as the reference for list labels. Generated
on macOS with **LibreOffice 26.8.3.1** (build `861effc5289d539f15c4c18d1ba8579ffd78728b`,
Homebrew `/opt/homebrew/bin/soffice`), with a throwaway user profile:

```sh
soffice -env:UserInstallation=file:///<scratch>/lo-profile --headless \
  --convert-to "txt:Text (encoded):UTF8" --outdir <out> <file>.docx
```

LibreOffice names the output `<file>.txt`; it was renamed `<a>+<b>+<file>.txt` for the input
`test/fixtures/<a>/<b>/<file>.docx` (for example `resolve+poi-Numbering.txt`,
`compose+powertools+DB007-Spec.txt`). The text is unedited (UTF-8 with BOM, LF line ends).
Besides this folder, references exist for list documents in `apache-poi/`, `compose-styles/`,
`docxcompose/`, `compose/docxcompose/` and `compose/powertools/`; the test skips a reference
whose document has moved. Exports of `compose-styles/*.docx` were kept only when they contain a
list line.
