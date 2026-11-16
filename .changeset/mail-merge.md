---
"turboword": minor
---

Add `turboword/mailmerge`: Word mail merge without Word. `mergeRecord` merges one record, `mergeEach` streams one document per record, `mergeAll` puts every record in one document (each in its own section with its own headers and footers, or separated by page breaks, or run on). Array values repeat table rows, nested tables included. NEXT, NEXTIF, SKIPIF, MERGEREC and MERGESEQ work as in Word. Merged fields become plain text as in Word's output (`fields: 'keep'` keeps them live), and the mail-merge data-source setup is removed so Word opens the result without a query prompt. Missing values merge as empty, or are kept or reported (`missing`).

Composing many documents is much faster: relationship, content-type, part-name and paragraph-id lookups are now indexed, so appending 1,000 documents that have headers and footers takes seconds instead of minutes.
