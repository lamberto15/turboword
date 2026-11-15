// Round-trip guarantees over the real fixture corpus (test/fixtures, see SOURCES.md):
//   (a) a no-op save returns the input bytes;
//   (b) force-dirtying every XML part keeps the entry names and order, and every part re-parses
//       to an equal DOM;
//   (c) editing one body paragraph leaves every other entry's compressed bytes and CRC alone;
//   (d) serializeXml(parseXml(x)) is idempotent; the share of byte-identical parts is logged;
//   (e) validatePackage finds no error code in the output that the input did not already have.
// Synthetic Strict variants of a few rich fixtures run through the same checks. The corpus
// itself is checked first: licences, size budget, and expected.json in sync with the files.

import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { afterAll, describe, expect, it } from 'vitest';
import {
  readDocx,
  type ValidationIssue,
  validatePackage,
  type WordDocument,
  writeDocx,
} from '../src/index.js';
import {
  domEqual,
  NS,
  parseXml,
  parseXmlBytes,
  serializeXml,
  type XDocument,
} from '../src/xml/index.js';
import { canonicalNs, isStrictNs } from '../src/xml/ns.js';
import { type LazyEntry, unzipLazy } from '../src/zip/index.js';
import {
  expectedFixtures,
  FIXTURES_DIR,
  type Fixture,
  fixtures,
  STRICT_NAMESPACES,
  strictFixtures,
  toStrict,
} from './helpers/fixtures.js';

const MAX_FIXTURE_BYTES = 300 * 1024;
const MAX_CORPUS_BYTES = 3 * 1024 * 1024;

/** Features the corpus must keep covering (see the fixtures brief in SOURCES.md). */
const REQUIRED_FEATURES = [
  'sections',
  'header-first',
  'header-even',
  'footer-first',
  'footer-even',
  'numbering',
  'merged-cells',
  'grid-span',
  'v-merge',
  'header-image',
  'comments',
  'footnotes',
  'endnotes',
  'textbox',
  'tracked-changes',
  'content-controls',
  'fields',
  'toc',
  'libreoffice',
  'strict',
];

/** Transitional fixtures with rich content, also run as synthetic Strict OOXML. */
const SYNTHETIC_STRICT_SOURCES = [
  'docxcompose/docproperties_header_footer_3_sections.docx',
  'docxcompose/header_with_image.docx',
  'python-docx/comments-rich-para.docx',
  'python-docx/tbl-cell-access.docx',
];

const MARKER = 'turboword round-trip marker';
const utf8 = new TextDecoder('utf-8');

// ---- helpers ------------------------------------------------------------------------------

/** Index of the first differing byte, or -1 when equal (cheaper and clearer than toEqual). */
function firstDifference(a: Uint8Array, b: Uint8Array): number {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) if (a[i] !== b[i]) return i;
  return a.length === b.length ? -1 : n;
}

function isXmlEntry(name: string): boolean {
  return /\.(xml|rels)$/i.test(name);
}

/** Part bytes as a string: UTF-8 (BOM stripped by TextDecoder) or UTF-16 by BOM. */
function decodeXml(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  return utf8.decode(bytes);
}

/** Marks the DOM of every XML part dirty; returns the names of the parts touched. */
function forceDirty(doc: WordDocument): string[] {
  const touched: string[] = [];
  for (const part of doc.package.parts()) {
    if (!part.isXml) continue;
    part.xml().markDirty();
    touched.push(part.name);
  }
  return touched;
}

/**
 * Puts MARKER into one body paragraph of the main part: replaces the first `w:t` of the first
 * paragraph that has one, or appends a run to the first paragraph when the body has no text.
 */
function editOneBodyParagraph(xml: XDocument): void {
  const body = xml.root.element(NS.w, 'body');
  if (!body) throw new Error('main part has no w:body');
  const paragraphs = body.descendants(NS.w, 'p');
  const withText = paragraphs.find((p) => p.descendants(NS.w, 't').length > 0);
  if (withText) {
    withText.descendants(NS.w, 't')[0]!.text = MARKER;
    return;
  }
  const p = paragraphs[0];
  if (!p) throw new Error('main part has no paragraph');
  const t = xml.createElement(NS.w, 't');
  t.text = MARKER;
  p.append(xml.createElement(NS.w, 'r').append(t));
}

function entryNames(entries: Map<string, LazyEntry>): string[] {
  return [...entries.keys()];
}

function errorCodes(issues: readonly ValidationIssue[]): Set<string> {
  return new Set(issues.filter((i) => i.severity === 'error').map((i) => i.code));
}

// ---- the corpus itself --------------------------------------------------------------------

describe('fixture corpus', () => {
  const all = fixtures();
  const expected = expectedFixtures();
  const sources = readFileSync(join(FIXTURES_DIR, 'SOURCES.md'), 'utf8');

  it('has 10-16 real fixtures within the size budget', () => {
    expect(all.length).toBeGreaterThanOrEqual(10);
    expect(all.length).toBeLessThanOrEqual(16);
    const tooBig = all.filter((f) => f.bytes.length > MAX_FIXTURE_BYTES).map((f) => f.name);
    expect(tooBig).toEqual([]);
    expect(all.reduce((n, f) => n + f.bytes.length, 0)).toBeLessThanOrEqual(MAX_CORPUS_BYTES);
  });

  it('keeps expected.json and SOURCES.md in sync with the files', () => {
    expect(Object.keys(expected).sort()).toEqual(all.map((f) => f.name));
    const unlisted = all.filter((f) => !sources.includes(`\`${f.name}\``)).map((f) => f.name);
    expect(unlisted).toEqual([]);
  });

  it('ships a licence next to every source', () => {
    const dirs = readdirSync(FIXTURES_DIR, { withFileTypes: true }).filter((d) => d.isDirectory());
    expect(dirs.length).toBeGreaterThan(0);
    for (const d of dirs) {
      const files = readdirSync(join(FIXTURES_DIR, d.name));
      const licence = files.some((f) => /^(LICEN[CS]E|COPYING|OFL|UNICODE-LICENSE)/i.test(f));
      expect(licence, `${d.name} needs a LICENSE (or OFL/COPYING) file`).toBe(true);
    }
  });

  it('records the page count and application of docProps/app.xml', () => {
    for (const f of all) {
      const app = unzipSync(f.bytes, { filter: (e) => e.name === 'docProps/app.xml' })[
        'docProps/app.xml'
      ];
      const text = app ? strFromU8(app) : '';
      const pages = /<(?:\w+:)?Pages>(\d+)</.exec(text)?.[1];
      const application = /<(?:\w+:)?Application>([^<]*)</.exec(text)?.[1];
      const want = expected[f.name];
      expect(want?.pages, `${f.name} pages`).toBe(pages === undefined ? undefined : Number(pages));
      expect(want?.application, `${f.name} application`).toBe(application);
    }
  });

  it('covers every required feature', () => {
    const covered = new Set(Object.values(expected).flatMap((e) => e.features));
    expect(REQUIRED_FEATURES.filter((f) => !covered.has(f))).toEqual([]);
  });

  it('builds synthetic Strict variants that ns.ts maps back to transitional', () => {
    for (const [transitional, strict] of STRICT_NAMESPACES) {
      expect(isStrictNs(strict), strict).toBe(true);
      expect(canonicalNs(strict)).toBe(transitional);
    }
    const [synthetic] = strictFixtures(['docxcompose/header_with_image.docx']);
    const files = unzipSync(synthetic!.bytes);
    const doc = strFromU8(files['word/document.xml']!);
    expect(doc).toContain('xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"');
    expect(doc).toContain('w:conformance="strict"');
    expect(doc).not.toContain('schemas.openxmlformats.org/wordprocessingml');
    const rels = strFromU8(files['_rels/.rels']!);
    expect(rels).toContain(
      'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument',
    );
    expect(rels).toContain(
      'http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties',
    );
    expect(rels).toContain('http://schemas.openxmlformats.org/package/2006/relationships/metadata');
    // Entry order is kept, and the rewrite is deterministic.
    expect(Object.keys(files)).toEqual(
      Object.keys(
        unzipSync(fixtures().find((f) => f.name.endsWith('header_with_image.docx'))!.bytes),
      ),
    );
    expect(firstDifference(toStrict(fixtures()[0]!.bytes), toStrict(fixtures()[0]!.bytes))).toBe(
      -1,
    );
  });
});

// ---- round-trip guarantees ----------------------------------------------------------------

/** Per-fixture serializer identity stats for (d), printed once at the end. */
const identity: { name: string; identical: number; total: number; differing: string[] }[] = [];

afterAll(() => {
  if (identity.length === 0) return;
  const identical = identity.reduce((n, s) => n + s.identical, 0);
  const total = identity.reduce((n, s) => n + s.total, 0);
  const lines = identity.map(
    (s) =>
      `  ${s.name}: ${s.identical}/${s.total}` +
      (s.differing.length > 0 ? ` (differ: ${s.differing.join(', ')})` : ''),
  );
  console.info(
    `serializeXml(parseXml(x)) === x for ${identical}/${total} XML parts ` +
      `(${((100 * identical) / Math.max(total, 1)).toFixed(1)}%)\n${lines.join('\n')}`,
  );
});

const corpus: Fixture[] = [...fixtures(), ...strictFixtures(SYNTHETIC_STRICT_SOURCES)];

describe.each(corpus.map((f) => [f.name, f] as const))('round-trip %s', (_name, f) => {
  it('(a) no-op save is byte-identical', () => {
    const out = writeDocx(readDocx(f.bytes));
    expect(firstDifference(out, f.bytes)).toBe(-1);
  });

  it('(b) force-dirty save keeps entry order and DOM-equal XML parts', () => {
    const doc = readDocx(f.bytes);
    const touched = forceDirty(doc);
    expect(touched.length).toBeGreaterThan(0);
    expect(doc.package.dirty).toBe(true);
    const out = writeDocx(doc);
    // The parts really were re-serialized (otherwise this test would prove nothing).
    expect(firstDifference(out, f.bytes)).not.toBe(-1);

    const before = unzipLazy(f.bytes);
    const after = unzipLazy(out);
    expect(entryNames(after)).toEqual(entryNames(before));

    const differing: string[] = [];
    for (const [name, e] of before) {
      const o = after.get(name)!;
      if (isXmlEntry(name)) {
        const a = parseXmlBytes(e.bytes(), { part: name });
        const b = parseXmlBytes(o.bytes(), { part: name });
        if (!domEqual(a, b)) differing.push(name);
      } else if (firstDifference(e.bytes(), o.bytes()) !== -1) {
        differing.push(name);
      }
    }
    expect(differing).toEqual([]);
  });

  it('(c) editing one body paragraph leaves every other entry untouched', () => {
    const doc = readDocx(f.bytes);
    const main = doc.mainPart;
    editOneBodyParagraph(main.xml());
    const out = writeDocx(doc);

    const before = unzipLazy(f.bytes);
    const after = unzipLazy(out);
    expect(entryNames(after)).toEqual(entryNames(before));

    const changed: string[] = [];
    for (const [name, e] of before) {
      if (name === main.name) continue;
      const o = after.get(name)!;
      if (o.crc32 !== e.crc32 || firstDifference(o.raw(), e.raw()) !== -1) changed.push(name);
    }
    expect(changed).toEqual([]);

    const edited = after.get(main.name)!;
    expect(edited.crc32).not.toBe(before.get(main.name)!.crc32);
    expect(decodeXml(edited.bytes())).toContain(MARKER);
  });

  it('(d) serializeXml(parseXml(x)) is idempotent', () => {
    let identical = 0;
    let total = 0;
    const differing: string[] = [];
    const notIdempotent: string[] = [];
    for (const [name, e] of unzipLazy(f.bytes)) {
      if (!isXmlEntry(name)) continue;
      const x = decodeXml(e.bytes());
      const once = serializeXml(parseXml(x, { part: name }));
      const twice = serializeXml(parseXml(once, { part: name }));
      if (twice !== once) notIdempotent.push(name);
      total++;
      if (once === x) identical++;
      else differing.push(name);
    }
    identity.push({ name: f.name, identical, total, differing });
    expect(total).toBeGreaterThan(0);
    expect(notIdempotent).toEqual([]);
  });

  it('(e) validatePackage reports no new error codes after saving', () => {
    const inputCodes = errorCodes(validatePackage(readDocx(f.bytes)));

    const dirty = readDocx(f.bytes);
    forceDirty(dirty);
    const edited = readDocx(f.bytes);
    editOneBodyParagraph(edited.mainPart.xml());

    for (const [label, doc] of [
      ['force-dirty', dirty],
      ['edited', edited],
    ] as const) {
      const outCodes = errorCodes(validatePackage(readDocx(writeDocx(doc))));
      const added = [...outCodes].filter((c) => !inputCodes.has(c));
      expect(added, label).toEqual([]);
    }
  });
});
