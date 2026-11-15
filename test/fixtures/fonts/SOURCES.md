# Test fonts

The oracle font set for font and layout tests: metric-compatible open replacements for the
Microsoft fonts Word documents name most often (Calibri → Carlito, Cambria → Caladea,
Arial → Liberation Sans, Times New Roman → Liberation Serif, Courier New → Liberation Mono).

All files are licensed under the **SIL Open Font License 1.1** (licence texts beside them). They
are stored **gzip-compressed and otherwise unmodified**: `gunzip -c X.ttf.gz` gives back the
original file byte for byte (SHA-256 below). Compressing keeps the repository at ~1.8 MB instead
of ~3.7 MB; it is not a modification of the Font Software. Load them with
`test/helpers/test-fonts.ts` (`testFontBytes`, `testFont`, `testFontResolver`).

| File | Version | Source | SHA-256 of the .ttf |
|---|---|---|---|
| Carlito-Regular.ttf.gz | 1.104 | https://github.com/google/fonts/tree/main/ofl/carlito | f6418f708baede9789daef5d458c0f53d2a888af9820e8062934e504fedc6595 |
| Carlito-Bold.ttf.gz | 1.104 | https://github.com/google/fonts/tree/main/ofl/carlito | bb5d20f79b82599ec72983597437373a80f2d2085fa91fc144fd74e876a594db |
| Caladea-Regular.ttf.gz | 1.001 | https://github.com/google/fonts/tree/main/ofl/caladea | f1e899278b7b4491aba5b6a8253c4b04c050cc59b21865be5c37559a775153cd |
| Caladea-Bold.ttf.gz | 1.001 | https://github.com/google/fonts/tree/main/ofl/caladea | ae3cb2dcbc925809dd29d2a44e9802211cab66be541bacbfc9c08c74b27c3742 |
| LiberationSans-Regular.ttf.gz | 2.1.5 | https://github.com/liberationfonts/liberation-fonts/releases/tag/2.1.5 (liberation-fonts-ttf-2.1.5.tar.gz) | 76d04c18ea243f426b7de1f3ad208e927008f961dc5945e5aad352d0dfde8ee8 |
| LiberationSans-Bold.ttf.gz | 2.1.5 | same | 788abee4c806d660e8aee46689dd8540cd4bb98da03dcc9d171ce3efd99a9173 |
| LiberationSerif-Regular.ttf.gz | 2.1.5 | same | 058ea80864aef09a23f45cbec2bb5400bc3dfbdea01c3f10538a21fcb497fb74 |
| LiberationSerif-Bold.ttf.gz | 2.1.5 | same | d754ba427cfe0bca54ae052384baa8f842da5bd6550ad4da024ac441e7a7d5ce |
| LiberationMono-Regular.ttf.gz | 2.1.5 | same | f2b83c763e8afd21709333370bed4774337fae82267937e2b5aea7e2fbd922c1 |
| LiberationMono-Bold.ttf.gz | 2.1.5 | same | bd62a0672d0b9b6710b01df434c80ad54fa5f0835207eb7b17b7a761463067bb |

Licences: `OFL-Carlito.txt` (Copyright 2013 The Carlito Project Authors, Reserved Font Name
"Carlito"), `OFL-Caladea.txt` (Copyright 2012 The Caladea Project Authors),
`OFL-Liberation.txt` (Digitized data copyright (c) 2010 Google Corporation, Reserved Font Arimo,
Tinos and Cousine; Copyright (c) 2012 Red Hat, Inc., Reserved Font Name Liberation).

LibreOffice ships its own copies of these families; when it is used as a layout oracle, its
bundled versions may differ slightly from the ones here.

## gen-standard14.mjs

Dev-only generator of `src/layout/fonts/standard14-data.ts` (standard-14 PDF font widths) from
pdf.js `src/core/metrics.js`, `encodings.js` and `glyphlist.js` (Copyright 2012 Mozilla
Foundation, Apache-2.0, https://github.com/mozilla/pdf.js), cross-checked against the Adobe Core14
AFM files (as shipped by Apache PDFBox, `pdfbox/src/main/resources/org/apache/pdfbox/resources/afm`).
Neither source is stored in the repository.
