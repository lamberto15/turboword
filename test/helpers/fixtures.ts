// The fixture corpus: real .docx files under test/fixtures/<source>/, copied unmodified from
// permissively licensed projects (provenance and licences in test/fixtures/SOURCES.md), plus
// what each one is expected to contain (test/fixtures/expected.json).
//
// It also builds *synthetic* Strict OOXML variants of the transitional fixtures. The corpus has
// two real Strict files, but both are small; rewriting richer files (header images, numbering,
// merged cells, sections) into Strict sends that content through the Strict code paths too.
// The variants are built in memory, never written to disk, and always named
// `synthetic-strict/...` so a failure on one is never mistaken for a failure on a real file.
//
// Test-only code: it uses node:fs and fflate's sync zip API, not the library under test, so the
// corpus can be listed and checked even while the library is broken.

import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, type Zippable, zipSync } from 'fflate';

/** Absolute path of test/fixtures/ (platform separators, trailing separator). */
export const FIXTURES_DIR: string = fileURLToPath(new URL('../fixtures/', import.meta.url));

/** One fixture file. */
export interface Fixture {
  /** `<source>/<file>` with forward slashes, e.g. `docxcompose/numberings.docx`. */
  readonly name: string;
  /** Absolute path on disk. For a synthetic variant, the path of the file it was built from. */
  readonly path: string;
  /** File contents (a fresh copy each call, so tests may mutate it). */
  readonly bytes: Uint8Array;
  /** True for in-memory variants such as `synthetic-strict/...` (no such file exists). */
  readonly synthetic?: boolean;
}

/** One entry of test/fixtures/expected.json. */
export interface ExpectedFixture {
  /** Word's page count from `docProps/app.xml` `<Pages>`, when the file records one. */
  pages?: number;
  /** `docProps/app.xml` `<Application>`, verbatim. */
  application?: string;
  /** Feature tags (kebab-case), e.g. `header-first`, `merged-cells`, `strict`, `libreoffice`. */
  features: string[];
}

const DOCX_FILE = /\.(docx|docm|dotx|dotm)$/i;

let cache: { name: string; path: string; bytes: Uint8Array }[] | undefined;

function load(): { name: string; path: string; bytes: Uint8Array }[] {
  if (cache) return cache;
  const out: { name: string; path: string; bytes: Uint8Array }[] = [];
  for (const dir of readdirSync(FIXTURES_DIR, { withFileTypes: true })) {
    if (!dir.isDirectory()) continue;
    for (const file of readdirSync(join(FIXTURES_DIR, dir.name))) {
      if (!DOCX_FILE.test(file)) continue;
      const path = join(FIXTURES_DIR, dir.name, file);
      out.push({ name: `${dir.name}/${file}`, path, bytes: new Uint8Array(readFileSync(path)) });
    }
  }
  out.sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  cache = out;
  return out;
}

/**
 * The curated corpus: every fixture listed in test/fixtures/expected.json, sorted by name. Other
 * .docx files under test/fixtures (feature-specific inputs such as encrypted files or compose
 * sources) are reachable through {@link fixture} but stay out of the general loops until they are
 * vetted and added to expected.json.
 */
export function fixtures(): Fixture[] {
  const curated = new Set(Object.keys(expectedFixtures()));
  return load()
    .filter((f) => curated.has(f.name))
    .map((f) => ({ name: f.name, path: f.path, bytes: f.bytes.slice() }));
}

/** One fixture by name (`<source>/<file>`). Throws when there is no such fixture. */
export function fixture(name: string): Fixture {
  const f = load().find((x) => x.name === name);
  if (!f) throw new Error(`No fixture named ${name}`);
  return { name: f.name, path: f.path, bytes: f.bytes.slice() };
}

/** test/fixtures/expected.json, keyed by fixture name. */
export function expectedFixtures(): Record<string, ExpectedFixture> {
  const text = readFileSync(join(FIXTURES_DIR, 'expected.json'), 'utf8');
  return JSON.parse(text) as Record<string, ExpectedFixture>;
}

/** Real fixtures whose expected.json entry lists `feature`. */
export function fixturesWith(feature: string): Fixture[] {
  const expected = expectedFixtures();
  return fixtures().filter((f) => expected[f.name]?.features.includes(feature) === true);
}

// ---- synthetic Strict OOXML ---------------------------------------------------------------

/**
 * Transitional namespace URI → Strict twin. The same table as `src/xml/ns.ts` (STRICT), written
 * the other way round; roundtrip.test.ts checks that `canonicalNs` maps every Strict URI back.
 */
export const STRICT_NAMESPACES: readonly (readonly [transitional: string, strict: string])[] = [
  [
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'http://purl.oclc.org/ooxml/wordprocessingml/main',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'http://purl.oclc.org/ooxml/officeDocument/relationships',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'http://purl.oclc.org/ooxml/officeDocument/math',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/main',
    'http://purl.oclc.org/ooxml/drawingml/main',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'http://purl.oclc.org/ooxml/drawingml/picture',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'http://purl.oclc.org/ooxml/drawingml/chart',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/diagram',
    'http://purl.oclc.org/ooxml/drawingml/diagram',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing',
    'http://purl.oclc.org/ooxml/drawingml/chartDrawing',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas',
    'http://purl.oclc.org/ooxml/drawingml/lockedCanvas',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/compatibility',
    'http://purl.oclc.org/ooxml/drawingml/compatibility',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'http://purl.oclc.org/ooxml/officeDocument/extendedProperties',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
    'http://purl.oclc.org/ooxml/officeDocument/customProperties',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    'http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes',
    'http://purl.oclc.org/ooxml/officeDocument/sharedTypes',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/bibliography',
    'http://purl.oclc.org/ooxml/officeDocument/bibliography',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/customXml',
    'http://purl.oclc.org/ooxml/officeDocument/customXml',
  ],
];

const TRANSITIONAL_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const STRICT_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/';
/** Relationship types whose last segment is spelled differently in Strict. */
const STRICT_REL_RENAMES: Readonly<Record<string, string>> = {
  'extended-properties': 'extendedProperties',
  'custom-properties': 'customProperties',
};

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Matches a whole quoted transitional URI (so `…/chart` never rewrites `…/chartDrawing`). */
const NS_PATTERN = new RegExp(
  `(["'])(${STRICT_NAMESPACES.map(([t]) => escapeRegExp(t)).join('|')})\\1`,
  'g',
);
const STRICT_OF: ReadonlyMap<string, string> = new Map(STRICT_NAMESPACES);
const REL_TYPE_PATTERN = new RegExp(
  `(Type=)(["'])${escapeRegExp(TRANSITIONAL_REL)}([^"']+)\\2`,
  'g',
);

/** Rewrites one transitional XML part (as text) to its Strict form. */
export function toStrictXml(xml: string): string {
  return xml
    .replace(NS_PATTERN, (_m, q: string, uri: string) => `${q}${STRICT_OF.get(uri) ?? uri}${q}`)
    .replace(
      REL_TYPE_PATTERN,
      (_m, attr: string, q: string, kind: string) =>
        `${attr}${q}${STRICT_REL}${STRICT_REL_RENAMES[kind] ?? kind}${q}`,
    );
}

/** Target of the package-level officeDocument relationship in `_rels/.rels` text. */
function mainPartName(rels: string): string | undefined {
  for (const tag of rels.match(/<Relationship\b[^>]*>/g) ?? []) {
    const type = /\bType=["']([^"']*)["']/.exec(tag)?.[1];
    if (type?.endsWith('/officeDocument') !== true) continue;
    return /\bTarget=["']([^"']*)["']/.exec(tag)?.[1]?.replace(/^\//, '');
  }
  return undefined;
}

/** Fixed local-time zip date, so the synthetic bytes are the same on every run and machine. */
const ZIP_DATE = new Date(1980, 0, 1, 0, 0, 0);

/**
 * **Synthetic** Strict OOXML from a transitional .docx. In every `.xml`/`.rels` part:
 * - each transitional namespace URI in {@link STRICT_NAMESPACES} becomes its Strict twin;
 * - each officeDocument relationship type becomes the Strict type (package-level types such as
 *   core-properties and thumbnail are the same in both flavours, so they stay as they are);
 * - the main document root gets `w:conformance="strict"`, as Word 2013+ writes it.
 *
 * Attribute value types (`ST_OnOff`, percentages, ...) are left alone, so the output is not
 * schema-valid Strict. It is good for exercising namespace and relationship-type aliasing.
 * Entry order is kept. XML parts are deflated, other entries are stored, and dates are fixed.
 */
export function toStrict(docx: Uint8Array): Uint8Array {
  const entries = unzipSync(docx);
  const rels = entries['_rels/.rels'];
  const main = rels ? mainPartName(strFromU8(rels)) : undefined;
  const out: Zippable = {};
  for (const [name, data] of Object.entries(entries)) {
    const utf16 =
      data.length >= 2 &&
      ((data[0] === 0xff && data[1] === 0xfe) || (data[0] === 0xfe && data[1] === 0xff));
    if (!/\.(xml|rels)$/i.test(name) || utf16) {
      out[name] = [data, { level: 0 }];
      continue;
    }
    let xml = toStrictXml(strFromU8(data));
    if (name === main && !/:conformance=/.test(xml)) {
      xml = xml.replace(/<([A-Za-z_][\w.-]*):document\b/, '<$1:document $1:conformance="strict"');
    }
    out[name] = [strToU8(xml), { level: 6 }];
  }
  return zipSync(out, { mtime: ZIP_DATE });
}

/**
 * Synthetic Strict variants (see {@link toStrict}) of the named real fixtures, or of every
 * transitional one when `names` is omitted. Each is named `synthetic-strict/<name>`.
 */
export function strictFixtures(names?: readonly string[]): Fixture[] {
  const expected = expectedFixtures();
  const picked = names
    ? names.map((n) => fixture(n))
    : fixtures().filter((f) => expected[f.name]?.features.includes('strict') !== true);
  return picked.map((f) => ({
    name: `synthetic-strict/${f.name}`,
    path: f.path,
    bytes: toStrict(f.bytes),
    synthetic: true,
  }));
}
