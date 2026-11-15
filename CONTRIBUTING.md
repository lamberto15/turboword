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

- `test/<area>.test.ts`: one file per area (`zip`, `xml-dom`, `opc`, `round-trip`, ...).
- `test/conventions.test.ts`: the rules below that a machine can check (layer imports, no `node:`
  imports outside `src/zip/node.ts`, no clock or `Math.random` in `src/`).
- `test/api-surface.test.ts`: compares the public exports with the committed `api/report.txt`. When
  you change the API on purpose, update it with `npx vitest run test/api-surface.test.ts -u` and
  commit the diff.
- `test/fixtures/`: .docx files used by the tests, listed in `test/fixtures/SOURCES.md`.
- `bench/`: benchmarks (`pnpm bench`).

Tests must pass on Windows: build paths with `fileURLToPath(new URL('./x', import.meta.url))`, never
`URL.pathname`.

### Word smoke pack

LibreOffice is more forgiving than Word. The Word smoke pack (`test/smoke/word-pack.mjs`, being
added for the v0.1 release) writes a set of files plus a checklist to open in real Microsoft Word
before each release, checking for "unreadable content" prompts. Results go in
`docs/word-verification.md`. Run it whenever you change how parts are written.

## Fixtures

- **Permissive licences only:** files you made yourself, or files under MIT, BSD, Apache-2.0, CC0 or a
  compatible licence. Nothing from GPL or LGPL projects, nothing with unclear origin.
- **Record every file in `test/fixtures/SOURCES.md`:** origin (URL or "made in Word 16.x for Mac"),
  licence, and what the file covers. Third-party sets keep their licence file next to them.
- **Keep files small:** one feature per file where possible, no large images or embedded fonts unless
  they are what the test is about.
- **No personal or confidential content.**
- Save fixtures with Word when you can: Word records its page count (`docProps/app.xml`) and its page
  breaks (`w:lastRenderedPageBreak`), which later serve as layout oracles.

## Conventions

- **Formatting:** Biome, 2 spaces, single quotes, line width 100. Run `pnpm format`.
- **TypeScript:** strict, with `exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`,
  `isolatedDeclarations` (explicit types on everything exported) and `verbatimModuleSyntax` (use
  `import type`). ESM, with `.js` extensions in relative imports.
- **Layers:** imports only go downward (`core` ← `zip`, `xml` ← `opc` ← `wml` ← ...). See
  [ARCHITECTURE.md](./ARCHITECTURE.md#layers). The conventions test fails otherwise.
- **Web APIs only in `src/`:** no `Buffer`, no `node:*` imports. The only exception is
  `src/zip/node.ts`, which looks up `node:zlib` with `process.getBuiltinModule` and has a fallback.
- **No recursion over XML trees:** documents can be arbitrarily deep; use an explicit stack.
- **Determinism:** no `Math.random`, no `Date.now()` or `new Date()` without arguments in `src/`. New
  IDs come from counters or a seeded PRNG.
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
