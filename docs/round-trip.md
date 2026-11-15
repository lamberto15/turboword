# Round-trip guarantees

turboword is built to open a .docx, change part of it and save it, without touching anything it was
not asked to change. This page states what is guaranteed and what is lost by design.

Terms used below:

- **Part:** one file inside the .docx zip (`word/document.xml`, `word/header1.xml`, an image, ...).
- **Clean part:** a part you did not modify.
- **Dirty part:** a part you modified (its DOM was mutated, or its bytes were replaced), or a new part.

## Guarantees

1. **A no-op save is byte-identical.** `writeDocx(readDocx(bytes))` returns the input bytes when
   nothing was modified.
2. **Clean parts are copied raw.** Their compressed payload, CRC-32, compression method and timestamp
   are written back as stored, never recompressed, so editing one part leaves every other part bit
   for bit the same. Reading a part (parsing it to navigate) does not make it dirty.
3. **Dirty parts are re-serialized from the full DOM.** The DOM keeps everything the file contains,
   modelled or not:
   - unknown elements and attributes, including vendor and future extension namespaces;
   - both branches of `mc:AlternateContent` (`mc:Choice` and `mc:Fallback`);
   - comments and processing instructions;
   - whitespace text, attribute order, namespace prefixes and declarations as written.

   A re-serialized part reparses to a DOM equal to the original (`domEqual`), apart from your edits.
4. **No implicit part dropping.** Parts nothing refers to anymore stay in the package. Removing them
   is an explicit call: `doc.gc()` deletes unreachable parts and returns their names.
5. **Deterministic output.** The same input and the same edits give the same bytes:
   - entries keep the source order; new parts are appended after them;
   - new IDs come from counters or a seeded PRNG, never from `Math.random`;
   - new and re-serialized entries carry a fixed zip date (1980-01-01 00:00, override with
     `writeDocx(doc, { zipDate })`);
   - `src/` never reads the clock (`Date.now()`, `new Date()` without arguments).

   Identical bytes hold for one deflate backend. Dirty and new entries are compressed with native
   zlib where the runtime has it (Node, Bun, Deno) and with fflate elsewhere, and the two produce
   different (equally valid) compressed bytes. For output that is identical across runtimes, pass
   `writeDocx(doc, { deflate: 'fflate' })`.
6. **Namespaces and `mc:Ignorable` are fixed up.** Before a dirty part is written, every prefix used by
   an element or attribute is declared on the root element if it is not in scope. Extension
   namespaces that older Word versions do not understand (`w14`, `w15`, `w16*`, `wp14`, ...) are added
   to `mc:Ignorable`, so a file that gains a `w14:` attribute still opens in Word 2007. Prefixes
   that are not declared on the root are removed from `mc:Ignorable` and `mc:ProcessContent`, which
   consumers reject otherwise. A prefix the source used without declaring it (some generators do)
   is read as its conventional OOXML namespace (`w` → WordprocessingML) and declared on save. The
   fix-up changes only the output, never the DOM.

## The zip container

- Clean entries keep their compressed bytes, CRC-32, compression method and timestamp. Dirty and new
  entries get a fresh local header with the fixed date, no data descriptor and the UTF-8 name flag.
- New media that are already compressed (PNG, JPEG, GIF, nested Office files, audio, video, ...)
  are stored rather than deflated, and any entry is stored when deflating would not make it
  smaller.
- **Directory entries** (zip entries ending in `/`, which some producers add) are not parts. A no-op
  save keeps them, since it returns the input bytes; any save after an edit leaves them out.
- ZIP64 archives are read; writing is limited to 4 GiB and 65,534 entries.

## Losses by design

These apply only to dirty parts; clean parts are copied raw and lose nothing. Apart from removing
characters that XML forbids, none of them changes what the XML means.

| Lost | What happens | Why |
| --- | --- | --- |
| DOCTYPE | The `<!DOCTYPE ...>` declaration is dropped. | DTDs are never read or expanded (see [SECURITY.md](../SECURITY.md)). Office never writes one. |
| Spelling of character references | `&#65;`, `&#x41;` and `A` are all written as `A`; only characters that must be escaped are escaped. | References are decoded at parse time; the DOM stores text, not markup. |
| CDATA delimiters | `<![CDATA[a<b]]>` is written as `a&lt;b`. | CDATA content becomes ordinary text. |
| CRLF inside the root element | `\r\n` (and a lone `\r`) inside the root element becomes `\n`. Whitespace outside the root (after the XML declaration) is kept as written. A CR you put into text through the API is written as `&#13;`, so it survives the next parse. | XML end-of-line handling (XML 1.0, section 2.11). |
| Literal whitespace in attribute values | A literal tab, CR or LF in an attribute value becomes a space. Escaped ones (`&#10;`) are kept. | Attribute-value normalization (XML 1.0, section 3.3.3). |
| Characters not allowed in XML 1.0 | C0 control characters other than tab, LF and CR, U+FFFE, U+FFFF and unpaired surrogates are removed from text and attribute values, including text you set through the API. | They make a part unreadable. |
| Text encoding | A dirty part is always written as UTF-8 (UTF-16 and declared single-byte encodings are decoded on read); a declaration naming another encoding is rewritten. A UTF-8 byte-order mark is kept. | One encoding on output keeps the writer simple. |

Markup spellings the DOM does not record are written in Word's own style: `<x></x>` becomes
`<x/>`, attribute values are double-quoted, and attributes are separated by one space. Since Word
writes parts that way, a part Word saved usually re-serializes byte for byte outside your edits:
on the test corpus, about 84% of XML parts come back byte-identical, and the rest differ only in
the spellings above (mostly `<x></x>` in `docProps` and custom XML parts written by other tools).

Namespace edge cases: a prefix bound to a different URI deeper in the tree is re-declared on that
element rather than on the root, and an empty prefixed declaration (`xmlns:p=""`, which XML 1.0
namespaces forbid) is rejected as an `XmlParseError`.

## Where this is tested

- No-op saves are byte-identical across the whole fixture corpus.
- After editing one part, every other entry's compressed payload and CRC are unchanged.
- Every XML part is forced dirty, saved, reparsed and compared with `domEqual`.
- The same checks run on synthetic Strict OOXML copies of the corpus, and the fuzz test re-saves
  mutated documents.
- Edited and composed outputs pass `validatePackage` with no new errors and convert in LibreOffice
  headless (in the tests and in CI). The Word smoke pack (`test/smoke/word-pack.mjs`) is meant to be
  opened in Word by hand before each release; for v0.1.0 that check is still pending.
