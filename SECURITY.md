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
| Unexpected I/O | `src/` makes no network requests and never touches the file system. Input arrives as bytes, a `Blob` or a stream; output leaves as bytes, a `Blob` or a stream. | always on |

Out of scope: turboword does not sanitize document content. Macros (`.docm`), embedded OLE objects,
external links and field codes are preserved as data, never executed or followed. Treat a document
from an untrusted source as untrusted after turboword has processed it, too.

## Reporting a vulnerability

Please report privately through
[GitHub private security advisories](https://github.com/lamberto15/turboword/security/advisories/new)
on `lamberto15/turboword`. Do not open a public issue.

Include the turboword version, the runtime, and a minimal file or script that reproduces the problem.
A crash, hang or excessive memory use on a crafted .docx counts.

Until 1.0, security fixes go into the latest release only.
