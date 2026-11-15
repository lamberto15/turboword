# Contributing

Run everything below from the repository root. Read [ARCHITECTURE.md](./ARCHITECTURE.md) before
changing code in `src/`.

## Setup

Requires Node 22 or newer and [pnpm](https://pnpm.io) 12 (the exact version is pinned in
`package.json`).

```sh
pnpm install --frozen-lockfile
```

## Scripts

| Script | What it does |
| --- | --- |
| `pnpm test` | Vitest suite (`pnpm test:watch` to iterate) |
| `pnpm typecheck` | `tsc --noEmit` over `src/` and `test/` |
| `pnpm lint` | Biome check (`pnpm format` to fix) |
| `pnpm build` | Build `dist/` with tsdown (ESM and CJS, one file per module) |
| `pnpm check:dts` | Type-checks the built `dist/index.d.ts` without Node types, so no Node type leaks into the public API |
| `pnpm verify` | All of the above: typecheck, lint, test, build, check:dts |
| `pnpm check:package` | publint (strict), arethetypeswrong, and a pack whitelist: the tarball holds the build, the docs, both licence files and NOTICE, nothing else (build first) |
| `pnpm smoke` | Round-trips every fixture through the built `dist/`, ESM and CJS (build first; CI also runs it under Bun and Deno) |
| `pnpm size` | Bundle-size budgets per entry point (build first) |
| `pnpm bench` | Benchmarks in `bench/` (the DOM against xmldom) |

Before opening a PR, `pnpm verify` and `pnpm check:package` must pass. CI also runs `verify` on
Node 22, 24 and 26 on Linux and Windows, the smoke test under Bun and Deno, and opens the fixtures
in LibreOffice headless.

## Tests

- `test/<area>.test.ts`: one or more files per area (`zip`, `xml`, `opc`, `wml`, `compose*`,
  `text*`, `fields*`, `forms*`, `layout*`, `pdf*`, ...).
- `test/conventions.test.ts`: the rules below that a machine can check (layer imports, no `node:`
  imports outside `src/zip/node.ts`, no clock or `Math.random` in `src/`, `.js` import extensions,
  licence files in the package).
- `test/api-surface.test.ts`: compares the exports of every entry point (`turboword` and each
  subpath) with the committed `api/report.txt`. When you change the API on purpose, update it with
  `npx vitest run test/api-surface.test.ts -u` and commit the diff.
- `test/roundtrip.test.ts`: byte-identical no-op saves, force-dirty round-trips and raw-copy checks
  over the curated corpus.
- `test/fuzz.test.ts`: mutated XML, zips and documents must only ever throw a `TurbowordError`.
- `bench/`: benchmarks (`pnpm bench`).

Tests must pass on Windows: build paths with `fileURLToPath(new URL('./x', import.meta.url))`, never
`URL.pathname`.

### Opt-in and environment switches

| Variable | Effect |
| --- | --- |
| `FUZZ_SEED`, `FUZZ_ITER` | Reproduce or lengthen a fuzz run (`FUZZ_STATS`, `FUZZ_TRACE`, `FUZZ_DUMP` print details) |
| `TURBOWORD_PDF_CORPUS=1` | Lay out and render every corpus fixture to PDF and check it with qpdf and pdf.js (`TURBOWORD_PDF_OUT` keeps the files) |
| `TURBOWORD_SKIP_LIBREOFFICE=1` | Skip the LibreOffice conversion tests in `test/compose.test.ts`; `SOFFICE` points the LibreOffice scripts and tests at a custom binary |
| `TURBOWORD_MSOFFCRYPTO_PYTHON` | A Python with msoffcrypto-tool, to cross-check encrypted output |

### Oracles and tools

| Tool | What it does |
| --- | --- |
| `node scripts/libreoffice-check.mjs <files or dirs>` | Converts .docx files to PDF with LibreOffice headless (isolated profile), counts pages with qpdf and compares them with `test/fixtures/expected.json`; fails on any conversion error |
| `scripts/docxcompose-diff/` | Composes the same scenarios with turboword and [docxcompose](https://github.com/4teamwork/docxcompose) and compares header/footer retention and page counts. Needs a throwaway Python venv with docxcompose; see its README |
| `node scripts/gen-unicode.mjs` | Regenerates `src/layout/text/data/*.ts` from the Unicode Character Database (byte-identical on re-run) |
| `node test/smoke/word-pack.mjs` | Writes `out/pack/*.docx` (no-op and force-dirty round-trips of every fixture, created and edited samples) plus `CHECKLIST.md` for opening in real Microsoft Word. Build first |

### Word smoke pack

LibreOffice is more forgiving than Word. Before each release, open the smoke pack in Microsoft Word
and work through `out/pack/CHECKLIST.md`, looking for "unreadable content" prompts and rendering
differences. Record the results in `docs/word-verification.md`. Run it whenever you change how parts
are written.

## Fixtures

Two kinds of fixtures live in `test/fixtures/`:

- **The curated corpus**: the documents listed in `test/fixtures/expected.json` (page count from
  `docProps/app.xml`, producing application, feature tags). `fixtures()` in
  `test/helpers/fixtures.ts` returns exactly these, and the general loops (round-trip, fuzz, smoke
  pack, layout oracle) run over them. Keep each file ≤ 300 KB and the corpus ≤ 3 MB.
- **Feature folders** (`compose/`, `forms/`, `crypto/`, `fonts/`, `unicode/`, ...): inputs for one
  feature's tests, loaded by name with `fixture('<folder>/<file>')`. They stay out of the general
  loops until vetted and added to `expected.json`.

Rules for both:

- **Permissive licences only:** files you made yourself, or files under MIT, BSD, Apache-2.0, CC0 or a
  compatible licence (OFL for test fonts, the Unicode licence for Unicode data). Nothing from GPL or
  LGPL projects, nothing with unclear origin.
- **A licence file in every folder** (`LICENSE`, `OFL*` or `COPYING*`); `test/roundtrip.test.ts`
  fails otherwise.
- **Record provenance in `test/fixtures/SOURCES.md`** (or the folder's own `SOURCES.md`): origin URL
  pinned to a commit, licence, and what the file covers. Verify downloads against the upstream blob
  SHA.
- **Keep files small:** one feature per file where possible.
- **No personal or confidential content.**
- Save fixtures with Word when you can: Word records its page count (`docProps/app.xml`) and its page
  breaks (`w:lastRenderedPageBreak`), which serve as layout oracles.

## Conventions

- **Formatting:** Biome, 2 spaces, single quotes, line width 100. Run `pnpm format`.
- **TypeScript:** strict, with `exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`,
  `isolatedDeclarations` (explicit types on everything exported) and `verbatimModuleSyntax` (use
  `import type`). ESM, with `.js` extensions in relative imports.
- **Layers:** a file may import its own directory or a directory of strictly lower rank:
  `core` 0; `zip`, `xml` 1; `opc` 2; `wml` 3; `resolve`, `text` 4; `compose`, `fields`, `forms`,
  `protect`, `crypto` 5; `layout` 6; `render` 7. Same-rank siblings never import each other. A new
  top-level directory needs a rank in `test/conventions.test.ts` first. See
  [ARCHITECTURE.md](./ARCHITECTURE.md#layers).
- **Web APIs only in `src/`:** no `Buffer`, no `node:*` imports. The only exception is
  `src/zip/node.ts`, which looks up `node:zlib` with `process.getBuiltinModule` and has a fallback.
- **No recursion over XML trees:** documents can be arbitrarily deep; use an explicit stack.
- **Determinism:** no `Math.random`, no `Date.now()` or `new Date()` without arguments in `src/`. New
  IDs come from `DocumentIds` (counters or a seeded PRNG); random bytes only where cryptography
  needs them (`crypto.getRandomValues`).
- **No new runtime dependencies:** fflate is the only one. Dev-only tools run ad hoc
  (`npx --yes`, a Python venv) instead of being added to `package.json`.
- **Errors:** throw a `TurbowordError` subclass with a stable `code`. Report recoverable problems
  through the `onWarning` callback instead of ignoring them.
- **Comments:** each file starts with a short comment explaining why it exists; exported functions
  and classes have JSDoc.
- **Lossless first:** never drop content you do not model. See [docs/round-trip.md](./docs/round-trip.md).

## Changesets

User-visible changes need a changeset:

```sh
pnpm changeset
```

Pick the bump, write a one-line summary, and commit the generated `.changeset/*.md`. Docs-only,
test-only and CI changes do not need one.

## Pull requests

- Use conventional commit messages (`feat:`, `fix:`, `docs:`, ...).
- Open PRs against `main`. Keep them focused, with tests for behavior changes.

## Licensing of contributions

turboword is dual-licensed under [MIT](./LICENSE-MIT) OR [Apache-2.0](./LICENSE-APACHE).
Contributions are inbound = outbound: unless you explicitly state otherwise, anything you
intentionally submit for inclusion is licensed under the same terms, without any additional terms or
conditions. There is no CLA.

Only bring in third-party code under a compatible permissive licence, and add its attribution to
[NOTICE](./NOTICE) in the same PR.
