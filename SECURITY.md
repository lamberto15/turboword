# Security

turboword reads .docx files that may come from untrusted sources. This page lists what it defends
against and how to report a vulnerability.

## What turboword defends against

All limits live in `Limits` (`src/core/limits.ts`) and can be changed per call with
`readDocx(bytes, { limits })` / `loadDocx(src, { limits })`. Breaking one throws a `TurbowordError`:
`ZipLimitError` (code `ZIP_LIMIT`), `ZipPathError` for an unsafe entry name (`ZIP_PATH`), or
`XmlParseError` for XML depth (`XML_PARSE`).

| Threat | Defence | Default |
| --- | --- | --- |
| Zip bombs | Decompressed size is measured on the actual inflated output, never trusted from the zip headers. Checked per entry, across all entries, and as a compression ratio. Parts are inflated lazily, only when read. | 1 GiB per entry (`maxEntryBytes`), 2 GiB in total (`maxTotalBytes`), ratio 1000 checked past 1 MiB of output (`maxCompressionRatio`) |
| Archives with huge entry counts | The central directory is rejected past the entry limit before anything is inflated. | 10 000 entries (`maxEntries`) |
| Path traversal | Entry names with a `..` segment, a leading `/`, a drive letter, a backslash or a NUL byte, and empty names, are rejected (`ZipPathError`). Part names are never used as file-system paths. | always on |
| Deeply nested XML | Element nesting deeper than the limit fails the parse. The parser, serializer, clone and equality code use explicit stacks, so depth cannot overflow the call stack. | 1024 levels (`maxXmlDepth`) |
| XXE, billion laughs | DTDs are never read and entities are never expanded; `<!DOCTYPE ...>` is skipped. Only the five predefined entities and numeric character references are decoded. | always on |
| Malformed fonts | The sfnt parser (`loadFont`, embedded `.odttf` fonts, fonts from resolvers) bounds-checks every table offset and length and throws `TurbowordError` with code `FONT_FORMAT`; a broken optional table is skipped instead. Fuzzed with corrupted fonts. | always on |
| Malformed images and pictures | The PDF image decoders (PNG, JPEG, GIF, BMP, TIFF) and the EMF/WMF and DIB readers reject absurd dimensions before allocating; an image that fails to decode is drawn as a grey box with a `PDF_IMAGE_UNSUPPORTED` warning instead of failing the render. | at most 2^28 pixels per image in PDF decoding, 2^24 for EMF/WMF bitmaps |
| Deeply nested layout input | Nested tables and text boxes are capped in the layout IR and the paginator, with a warning past the cap. | 32 levels |
| Encrypted files | `decryptDocx` rejects files asking for more password-hash iterations than `maxSpinCount`, verifies the Agile `dataIntegrity` HMAC (rejecting tampered files), caps every container stream and the decrypted package at `limits.maxEntryBytes`, and throws `TurbowordError` with code `BAD_PASSWORD` on a wrong password, `DocxFormatError` otherwise. | `maxSpinCount` 10,000,000; `verifyIntegrity: true` |
| Script links in PDFs | Links to `javascript:`, `vbscript:` and `data:` URIs are dropped from PDF output with a `PDF_LINK_UNSAFE` warning. | always on |
| Unexpected I/O | `src/` makes no network requests and never touches the file system. Input arrives as bytes, a `Blob` or a stream; output leaves as bytes, a `Blob` or a stream. Fonts come only from the document or from the `FontResolver`s you pass. | always on |

**Protection is not encryption.** `protectDocument` writes Word's editing restriction
(`w:documentProtection`): Word enforces it, but anyone can remove it, and the content is stored in
the clear. To keep content confidential, use `encryptDocx` (`turboword/crypto`, Agile AES-256 with
SHA-512 key derivation and a random salt from Web Crypto). The password is never stored.

Out of scope: turboword does not sanitize document content. Macros (`.docm`), embedded OLE objects,
external links and field codes are preserved as data, never executed or followed. Treat a document
from an untrusted source as untrusted after turboword has processed it, too.

## Reporting a vulnerability

Please report privately through
[GitHub private security advisories](https://github.com/lamberto15/turboword/security/advisories/new)
on `lamberto15/turboword`. Do not open a public issue.

Include the turboword version, the runtime, and a minimal file or script that reproduces the problem.
A crash, hang or excessive memory use on a crafted .docx counts.

Every confirmed vulnerability gets fixed, whichever version it was reported against, and the fix is
released as soon as it is ready. Fixes ship in a new release of the current line; if you depend on an
older release line and cannot upgrade, say so in the report and we will publish a fix for that line
too. You will be credited in the advisory unless you prefer not to be.
