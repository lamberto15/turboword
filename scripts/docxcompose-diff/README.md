# docxcompose differential test

Composes the same documents with turboword and with
[docxcompose](https://github.com/4teamwork/docxcompose) (the reference permissive Python
library) and compares the results:

- **Header/footer retention.** For every section of every input, the header and footer text
  each page slot *shows* (default; first page when `titlePg`; even pages when
  `evenAndOddHeaders`), after inheritance, must be what the composed document shows for the
  corresponding section.
- **Page count** (LibreOffice): composed pages against the sum of the input pages (goal ≥ 95%).

`scenarios.json` lists the append scenarios (first input = target, the others appended in order):
docxcompose's own test combinations, Open-XML-PowerTools DocumentBuilder combinations, the
PowerTools header/footer sources appended to letter/fax templates, and a few corpus documents
with first-page/even headers and Strict OOXML. Inputs live in `test/fixtures/compose/` (see its
`SOURCES.md`) and `test/fixtures/`.

## Running

docxcompose is not a dependency of this repository; use a throwaway virtual environment:

```sh
python3 -m venv /tmp/dcx-venv
/tmp/dcx-venv/bin/pip install docxcompose
/tmp/dcx-venv/bin/python -I scripts/docxcompose-diff/run.py
```

Options: `--out DIR` (default `out/docxcompose-diff/`, which ends up holding both tools'
outputs and `summary.json`), `--no-libreoffice` (skip page counts), `--only SUBSTR` (run the
scenarios whose name contains SUBSTR), `--scenarios FILE`.

The turboword side runs `npx --yes tsx scripts/docxcompose-diff/compose.ts` (tsx is fetched on
demand; it is not a dependency either). LibreOffice is found like `scripts/libreoffice-check.mjs`
does (`SOFFICE` overrides); `qpdf` is used for page counts when installed.

The exit code is 1 when turboword fails a scenario or keeps fewer header/footer slots than
docxcompose in any scenario.

## Last result (2026-11-15, docxcompose 2.2.0, LibreOffice 25.x, macOS)

| | docxcompose | turboword |
| --- | --- | --- |
| header/footer slots kept (59 scenarios, 338 slots) | 172 (50.9%) | **338 (100%)** |
| scenarios with page count ≥ 95% of the inputs | 3 / 59 | **57 / 59** |
| total pages (inputs: 165) | 108 (65.5%) | **163 (98.8%)** |
| failures | 1 (Strict OOXML input) | 0 |

docxcompose keeps only the target's sections: every appended document loses its own section
properties, so its pages show the target's last header/footer (and its page size/orientation).
turboword keeps more headers/footers in all 59 scenarios and fewer in none. The two scenarios
under 95% (PowerTools `DB005` + `DB002-Landscape-Section`, mail-merge footnote/header document +
docproperties 3 sections) come out one page short in both tools: both targets end with an empty
or continuous final section, which LibreOffice lays out differently once more content follows.
