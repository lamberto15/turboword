---
"turboword": minor
---

First release. Lossless editing and composition of Word .docx files in pure TypeScript:

- open, edit and save without losing anything (a no-op save returns the input bytes; untouched parts keep their compressed bytes)
- document model: stories, paragraphs, runs, tables, sections, headers/footers, images, hyperlinks, bookmarks, properties, `validatePackage`
- find/replace across split runs (optionally tracked), accept/reject revisions, style cascade and list labels
- `turboword/compose`: `appendDocument`, `insertDocument`, `composeDocuments` keeping each document's sections, headers/footers, styles, numbering, notes, comments and bookmarks
- `turboword/fields`, `turboword/forms`: field updates and TOCs; content controls, data binding and legacy forms
- document protection and `turboword/crypto` (Office encryption)
- experimental: `turboword/layout` (pagination, page count), `turboword/pdf` and `turboword/svg`
