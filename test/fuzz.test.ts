// Seeded mutation fuzzing of the XML parser, the zip reader and the docx read/write path.
// Property: every input either succeeds or throws a TurbowordError (any subclass): never a raw
// TypeError/RangeError/undefined access, never a stack overflow on deep input, never a hang (each
// case has a time budget). Where a step succeeds, a few round-trip properties are checked too.
// `FUZZ_ITER=N` sets the iterations per target (default 300), `FUZZ_SEED=N` shifts every seed;
// a failure message names the target, iteration and seed so it can be reproduced, and
// `FUZZ_TRACE=1` / `FUZZ_DUMP=dir` (see timed()) find and save the input of a crash or hang.

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterAll, describe, expect, it } from 'vitest';
import { TurbowordError, XmlParseError } from '../src/core/errors.js';
import type { Limits } from '../src/core/limits.js';
import { readDocx, type WordDocument, writeDocx } from '../src/wml/document.js';
import { validatePackage } from '../src/wml/validate.js';
import { parseXml, parseXmlBytes } from '../src/xml/parse.js';
import { serializeXml } from '../src/xml/serialize.js';
import { unzipLazy } from '../src/zip/index.js';
import { fixtures, strictFixtures } from './helpers/fixtures.js';

const ITER = Number(process.env.FUZZ_ITER) || 300;
const SEED = (Number(process.env.FUZZ_SEED) || 0) * 1000;
/** Time budget of one case (generous: CI runners are several times slower than laptops). */
const CASE_MS = 5000;
const TEST_MS = 300_000;

// ------------------------------------------------------------------ PRNG + helpers

/** mulberry32: small seeded PRNG so every run sees the same inputs. */
function prng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
type Rng = () => number;
const int = (r: Rng, n: number): number => Math.floor(r() * n);
const pick = <T>(r: Rng, a: readonly T[]): T => a[int(r, a.length)] as T;

function describeErr(e: unknown): string {
  return e instanceof Error ? `${e.name}: ${e.message}\n${e.stack ?? ''}` : String(e);
}

/** Outcomes per step (`FUZZ_STATS=1` prints them): shows how deep the mutations get. */
const STATS = new Map<string, { ok: number; rejected: number }>();
afterAll(() => {
  if (!process.env.FUZZ_STATS) return;
  const rows = [...STATS].map(([k, v]) => `${k.padEnd(18)} ok ${v.ok}  rejected ${v.rejected}`);
  process.stderr.write(`${rows.join('\n')}\n`);
});

/**
 * Runs step `what` of case `name`. A TurbowordError is an accepted outcome (returns undefined);
 * anything else fails the case. Returns `fn`'s result on success.
 */
function step<T>(name: string, what: string, fn: () => T): T | undefined {
  const key = what.replace(/\(.+\)$/, '()');
  const s = STATS.get(key) ?? { ok: 0, rejected: 0 };
  STATS.set(key, s);
  try {
    const v = fn();
    s.ok++;
    return v;
  } catch (e) {
    if (e instanceof TurbowordError) {
      s.rejected++;
      return undefined;
    }
    throw new Error(`${name}: ${what} threw a non-TurbowordError: ${describeErr(e)}`);
  }
}

/**
 * Runs one fuzz case and fails it when it exceeds the time budget. `FUZZ_TRACE=1` prints its name
 * first; `FUZZ_DUMP=dir` saves its input as `dir/last.<ext>` before running (the culprit of a
 * crash) and as `dir/<case>.<ext>` when it fails.
 */
function timed(name: string, input: string | Uint8Array, ext: string, fn: () => void): void {
  const dump = process.env.FUZZ_DUMP;
  const save = (base: string): void => {
    if (dump === undefined) return;
    mkdirSync(dump, { recursive: true });
    writeFileSync(join(dump, `${base}.${ext}`), input);
  };
  save('last');
  if (process.env.FUZZ_TRACE) process.stderr.write(`${name}\n`);
  const t0 = performance.now();
  try {
    fn();
  } catch (e) {
    save(name.replace(/\W+/g, '_'));
    throw e;
  }
  const dt = performance.now() - t0;
  if (dt > CASE_MS) throw new Error(`${name} took ${Math.round(dt)} ms (budget ${CASE_MS} ms)`);
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

// ------------------------------------------------------------------ corpus

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_W = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const story = (tag: string, body: string): string =>
  `${XML_DECL}<w:${tag} xmlns:w="${W}" xmlns:r="${R}">${body}</w:${tag}>`;
const para = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** A hand-built docx touching every v0.1 feature (independent of the library's own writer). */
function synthetic(): Uint8Array {
  const document = story(
    'document',
    `<w:body>${para('Hello &amp; welcome')}` +
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Head</w:t></w:r><w:r><w:tab/>' +
      '<w:t>ing</w:t><w:br/></w:r><w:r><w:footnoteReference w:id="1"/></w:r>' +
      '<w:commentRangeStart w:id="0"/><w:r><w:t>noted</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>' +
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>' +
      `${para('wide')}</w:tc><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${para('tall')}</w:tc></w:tr>` +
      `<w:tr><w:tc>${para('a')}</w:tc><w:tc>${para('b')}</w:tc>` +
      '<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId2"/>' +
      '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" ' +
      'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p>' +
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      `<mc:Choice Requires="w14">${para('choice')}</mc:Choice><mc:Fallback>${para('fallback')}` +
      '</mc:Fallback></mc:AlternateContent>' +
      `${para('section two')}<w:sectPr><w:footerReference w:type="default" r:id="rId3"/>` +
      '<w:type w:val="continuous"/><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>' +
      '<w:cols w:num="2"/><w:pgNumType w:start="5"/><w:titlePg/></w:sectPr></w:body>',
  );
  const rel = (id: string, type: string, target: string): string =>
    `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
  const files: Record<string, string> = {
    '[Content_Types].xml':
      `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${CT_W}.document.main+xml"/>` +
      `<Override PartName="/word/header1.xml" ContentType="${CT_W}.header+xml"/>` +
      `<Override PartName="/word/footer1.xml" ContentType="${CT_W}.footer+xml"/>` +
      `<Override PartName="/word/footnotes.xml" ContentType="${CT_W}.footnotes+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="${CT_W}.comments+xml"/>` +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      '</Types>',
    '_rels/.rels':
      `${XML_DECL}<Relationships xmlns="${PR}">${rel('rId1', 'officeDocument', 'word/document.xml')}` +
      `<Relationship Id="rId2" Type="${PR}/metadata/core-properties" Target="docProps/core.xml"/>` +
      `${rel('rId3', 'extended-properties', 'docProps/app.xml')}</Relationships>`,
    'word/document.xml': document,
    'word/_rels/document.xml.rels':
      `${XML_DECL}<Relationships xmlns="${PR}">${rel('rId2', 'header', 'header1.xml')}` +
      `${rel('rId3', 'footer', 'footer1.xml')}${rel('rId4', 'footnotes', 'footnotes.xml')}` +
      `${rel('rId5', 'comments', 'comments.xml')}` +
      '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
      'Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/></Relationships>',
    'word/header1.xml': story('hdr', para('Header text')),
    'word/footer1.xml': story('ftr', para('Footer text')),
    'word/footnotes.xml': story(
      'footnotes',
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
        `<w:footnote w:id="1">${para('A footnote')}</w:footnote>`,
    ),
    'word/comments.xml': story(
      'comments',
      `<w:comment w:id="0" w:author="fuzz" w:date="2020-01-01T00:00:00Z">${para('A comment')}</w:comment>`,
    ),
    'docProps/core.xml':
      `${XML_DECL}<cp:coreProperties xmlns:cp="${PR.replace('relationships', 'metadata/core-properties')}" ` +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Fuzz</dc:title>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>' +
      '</cp:coreProperties>',
    'docProps/app.xml':
      `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
      '<Application>fuzz</Application><Pages>2</Pages></Properties>',
  };
  const out: Record<string, Uint8Array> = {};
  for (const [name, text] of Object.entries(files)) out[name] = strToU8(text);
  return zipSync(out);
}

/** The fixture corpus plus two synthetic Strict variants (Strict namespace/rel-type paths). */
const DOCX: readonly Uint8Array[] = [
  synthetic(),
  ...[
    ...fixtures(),
    ...strictFixtures(['docxcompose/numberings.docx', 'docxcompose/header_with_image.docx']),
  ]
    .map((f) => f.bytes)
    .filter((b) => b.length < 300_000), // mutation cost grows with size; big files add no shapes
];

const XML_SNIPPETS: readonly string[] = [
  '<?xml version="1.0"?><!DOCTYPE r [<!ENTITY e "x"><!ELEMENT r ANY>]><r>&e;&amp;&#x41;&#66;</r>',
  '\uFEFF<?xml version="1.0" encoding="UTF-8"?>\r\n<a xmlns="urn:d" xmlns:p="urn:p"><p:b p:c="1"/></a>',
  '<a><![CDATA[ <not> & markup ]]]]><![CDATA[> ]]><!-- c --><?pi data?></a>',
  '<w:p xmlns:w="urn:w" xml:space="preserve">  <w:t>\t x \r\n y </w:t>  </w:p>',
  '<a b="&lt;&gt;&quot;&apos;&#9;&#10;&#13;" c=\'"\'>&#x1F600;</a>',
  '<x:a xmlns:x="urn:x"><x:b xmlns:x="urn:y"><x:c/></x:b><x:d/></x:a>',
  '<root><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    '<mc:Choice Requires="a"><a/></mc:Choice><mc:Fallback><b/></mc:Fallback></mc:AlternateContent></root>',
];

/** Every XML part of the corpus (and the snippets): the parser's corpus. */
function xmlCorpus(): string[] {
  const out = [...XML_SNIPPETS];
  for (const d of DOCX) {
    for (const [name, data] of Object.entries(unzipSync(d))) {
      if (/\.(xml|rels)$/i.test(name) && data.length < 200_000) out.push(strFromU8(data));
    }
  }
  return out;
}

// ------------------------------------------------------------------ mutators

function mutateBytes(r: Rng, src: Uint8Array): Uint8Array {
  let b = Array.from(src);
  const n = 1 + int(r, 8);
  for (let i = 0; i < n; i++) {
    const p = int(r, b.length);
    switch (int(r, 4)) {
      case 0:
        b[p] = int(r, 256);
        break;
      case 1:
        b[p] = (b[p] as number) ^ (1 << int(r, 8));
        break;
      case 2:
        b.splice(p, 0, int(r, 256));
        break;
      default:
        b.splice(p, 1 + int(r, 4));
    }
  }
  if (r() < 0.1) b = b.slice(0, int(r, b.length));
  return Uint8Array.from(b);
}

const SIGS = [0x04034b50, 0x02014b50, 0x06054b50, 0x08074b50]; // local, central, end, descriptor
const ODD_U32 = [0, 1, 0xffff, 0xffffffff, 0x7fffffff, 0x80000000];

/** Overwrites one field of a zip header (local/central/EOCD) with a random or boundary value. */
function mutateZipHeader(r: Rng, src: Uint8Array): Uint8Array {
  const b = src.slice();
  const dv = new DataView(b.buffer);
  const at: number[] = [];
  for (let i = 0; i + 4 <= b.length; i++) {
    if (SIGS.includes(dv.getUint32(i, true))) at.push(i);
  }
  if (at.length === 0) return mutateBytes(r, src);
  const n = 1 + int(r, 3);
  for (let k = 0; k < n; k++) {
    const off = pick(r, at) + 4 + 2 * int(r, 20); // any 16-bit-aligned field of the header
    if (off + 4 > b.length) continue;
    if (r() < 0.5)
      dv.setUint16(off, r() < 0.5 ? int(r, 0x10000) : pick(r, [0, 0xffff, 8, 99]), true);
    else dv.setUint32(off, r() < 0.5 ? int(r, 0x100000000) : pick(r, ODD_U32), true);
  }
  return r() < 0.2 ? b.slice(0, int(r, b.length)) : b;
}

const ENTITIES = [
  '&amp;',
  '&lt;',
  '&#0;',
  '&#x110000;',
  '&#xD800;',
  '&bogus;',
  '&#99999999999;',
  '&',
  '&#;',
  '&#x;',
  '&#x41',
];
const INSERTS = [
  '<',
  '>',
  '&',
  '"',
  "'",
  '</',
  '/>',
  '=',
  ']]>',
  '<!--',
  '-->',
  '<?',
  '?>',
  '<![CDATA[',
  '<!DOCTYPE x>',
  '\0',
  '\uD800',
  '\uFFFE',
  '\r',
  ' xmlns:w=""',
  ' xmlns=""',
  ' xmlns:xml="urn:bad"',
  ' a="1" a="2"',
  '<zz:q/>',
  '<:a/>',
  '<a:/>',
  '<a b=c>',
];
const FRAGS = [
  // relationships and content types
  `<Relationship Id="rIdF" Type="${R}/officeDocument" Target="../../../outside.xml"/>`,
  `<Relationship Id="rId1" Type="${R}/header" Target="/word/document.xml" TargetMode="External"/>`,
  `<Relationship Id="rId2" Type="${R}/footer" Target=""/>`,
  `<Relationship Id="rId3" Type="${R}/footnotes" Target="%2e%2e/%2e%2e/x.xml"/>`,
  `<Relationship Id="rId4" Type="${R}/comments" Target="word/missing.xml"/>`,
  '<Relationship Id="rId5"/>',
  '<Override PartName="/word/document.xml" ContentType="text/plain"/>',
  '<Override PartName="" ContentType=""/><Default Extension="" ContentType="x"/>',
  // body structure
  '<w:headerReference w:type="default" r:id="rId999"/>',
  '<w:footerReference w:type="first" r:id="rId1"/>',
  '<w:footnoteReference w:id="99999999"/><w:endnoteReference w:id="-5"/>',
  '<w:sectPr><w:pgSz w:w="-1" w:h="NaN"/><w:cols w:num="99999999"/><w:pgNumType w:start="x"/></w:sectPr>',
  '<w:sectPr><w:sectPr/></w:sectPr>',
  '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="99999999"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>',
  '<w:tbl><w:tr><w:trPr><w:gridBefore w:val="-3"/><w:gridAfter w:val="1e9"/></w:trPr><w:tc><w:tcPr><w:vMerge/></w:tcPr></w:tc></w:tr></w:tbl>',
  '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="0"/><w:hMerge w:val="continue"/></w:tcPr></w:tc></w:tr><w:tr/></w:tbl>',
  '<w:tbl/><w:tr/><w:tc/>',
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/>',
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice/><mc:Fallback><mc:AlternateContent/></mc:Fallback></mc:AlternateContent>',
  '<w:body/><w:body><w:p/></w:body>',
  '<w:p><w:r><w:t/><w:tab/><w:br w:type="page"/><w:sym w:char="ZZZZ"/><w:noBreakHyphen/></w:r></w:p>',
  '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple><w:hyperlink r:id="rId999"/>',
  '<w:footnote w:id="1" w:type="separator"/><w:comment w:id="x"/>',
  '<dc:title>x</dc:title><dcterms:created>not a date</dcterms:created><Pages>-1e400</Pages>',
];

/** Text-level mutations of an XML string. */
function mutateXml(r: Rng, src: string): string {
  let xml = src;
  const n = 1 + int(r, 4);
  for (let i = 0; i < n; i++) {
    const p = int(r, xml.length + 1);
    switch (int(r, 12)) {
      case 0: // truncation
        xml = xml.slice(0, p);
        break;
      case 1: // flip one character (special, random code unit, or an ASCII neighbour)
        if (p < xml.length) {
          const c =
            r() < 0.5
              ? pick(r, ['<', '>', '&', '"', "'", '=', '/', '!', '?', '[', ']', '-', ':', ' '])
              : r() < 0.5
                ? String.fromCharCode(int(r, 0x10000))
                : String.fromCharCode(xml.charCodeAt(p) ^ (1 << int(r, 7)));
          xml = xml.slice(0, p) + c + xml.slice(p + 1);
        }
        break;
      case 2:
        xml = xml.slice(0, p) + pick(r, INSERTS) + xml.slice(p);
        break;
      case 3:
        xml = xml.slice(0, p) + pick(r, ENTITIES) + xml.slice(p);
        break;
      case 4: {
        // duplicate a tag
        const s = xml.indexOf('<', p);
        const e = s < 0 ? -1 : xml.indexOf('>', s);
        if (e > 0) xml = xml.slice(0, e + 1) + xml.slice(s, e + 1) + xml.slice(e + 1);
        break;
      }
      case 5: {
        // drop a closing tag
        const s = xml.indexOf('</', p);
        const e = s < 0 ? -1 : xml.indexOf('>', s);
        if (e > 0) xml = xml.slice(0, s) + xml.slice(e + 1);
        break;
      }
      case 6: // delete a span
        xml = xml.slice(0, p) + xml.slice(p + 1 + int(r, 40));
        break;
      case 7: // a structural fragment where a tag starts
      case 8: {
        const s = xml.indexOf('<', p);
        const at = s < 0 ? p : s;
        xml = xml.slice(0, at) + pick(r, FRAGS) + xml.slice(at);
        break;
      }
      case 9: // deep nesting (beyond small maxDepth limits; sometimes closed again)
        {
          const k = 50 + int(r, 3000);
          const close = r() < 0.5 ? '</a>'.repeat(k) : '';
          xml = `${xml.slice(0, p)}${'<a>'.repeat(k)}${close}${xml.slice(p)}`;
        }
        break;
      case 10: // a long text run
        xml = `${xml.slice(0, p)}${'x'.repeat(50_000)}${xml.slice(p)}`;
        break;
      default: // absurd number in an attribute
        xml = xml.replace(
          /="(-?\d+)"/,
          (_m, d: string) => `="${pick(r, [Number(d) + 1, -1, 1e9, -1e9, 'NaN', '1e400', ''])}"`,
        );
    }
  }
  return xml;
}

/** Mutates 1-2 XML parts of a docx (sometimes drops or renames an entry) and re-zips with fflate. */
function mutateParts(r: Rng, src: Uint8Array): Uint8Array {
  const files = unzipSync(src);
  const names = Object.keys(files).filter((n) => /\.(xml|rels)$/i.test(n));
  const k = 1 + int(r, 2);
  for (let i = 0; i < k && names.length > 0; i++) {
    const name = pick(r, names);
    files[name] = strToU8(mutateXml(r, strFromU8(files[name]!)));
  }
  const roll = r();
  const all = Object.keys(files);
  if (roll < 0.05) delete files[pick(r, all)];
  else if (roll < 0.08) {
    const from = pick(r, all);
    files[from.toUpperCase()] = files[from]!;
    delete files[from];
  }
  return zipSync(files, { level: pick(r, [0, 6]) });
}

// ------------------------------------------------------------------ runners

const noWarnings = (): void => {};

/** Touches the read API of a document (beyond text()), every call guarded. */
function explore(name: string, doc: WordDocument): void {
  step(name, 'stories', () => {
    for (const s of doc.stories()) {
      s.text();
      for (const t of s.tables()) t.grid();
    }
  });
  step(name, 'sections', () => {
    for (const s of doc.sections()) {
      void [s.type, s.pageSize, s.margins, s.titlePage, s.columns, s.pageNumberStart];
      for (const t of ['default', 'first', 'even'] as const) {
        s.header(t)?.text();
        s.footer(t)?.text();
      }
    }
  });
  step(name, 'properties', () => {
    const p = doc.properties;
    void [p.title, p.creator, p.created, p.modified, p.app];
    p.custom();
  });
  step(name, 'validatePackage', () => validatePackage(doc));
}

/**
 * readDocx -> text() -> writeDocx (must return the input bytes: nothing was modified), then the
 * rest of the read API, then optionally a force-dirty save whose output must read back with the
 * same text.
 */
function runDocx(
  name: string,
  data: Uint8Array,
  forceDirty: boolean,
  limits?: Partial<Limits>,
): void {
  const doc = step(name, 'readDocx', () =>
    readDocx(
      data,
      limits === undefined ? { onWarning: noWarnings } : { limits, onWarning: noWarnings },
    ),
  );
  if (doc === undefined) return;
  const text = step(name, 'text()', () => doc.text());
  const out = step(name, 'writeDocx', () => writeDocx(doc));
  if (out !== undefined && !sameBytes(out, data)) {
    throw new Error(`${name}: a no-op writeDocx did not return the input bytes`);
  }
  explore(name, doc);
  if (!forceDirty) return;
  for (const part of doc.package.parts()) {
    if (part.isXml) step(name, `xml(${part.name})`, () => part.xml().markDirty());
  }
  const dirty = step(name, 'writeDocx dirty', () => writeDocx(doc));
  if (dirty === undefined) return;
  let again: string;
  try {
    again = readDocx(dirty, { onWarning: noWarnings }).text();
  } catch (e) {
    if (text === undefined && e instanceof TurbowordError) return; // was unreadable before, too
    throw new Error(`${name}: re-reading a force-dirty save failed: ${describeErr(e)}`);
  }
  if (text !== undefined && again !== text) {
    throw new Error(`${name}: text changed across a force-dirty save`);
  }
}

/** unzipLazy, then raw() and bytes() of every entry. */
function runZip(name: string, data: Uint8Array, limits?: Partial<Limits>): void {
  const entries = step(name, 'unzipLazy', () => unzipLazy(data, limits));
  if (entries === undefined) return;
  for (const e of entries.values()) {
    step(name, `raw(${e.path})`, () => e.raw());
    step(name, `bytes(${e.path})`, () => e.bytes());
  }
}

/** parseXml; when it parses, serialize -> parse -> serialize must be a fixpoint. */
function runXml(name: string, text: string, maxDepth: number | undefined): void {
  const options = maxDepth === undefined ? {} : { maxDepth };
  const doc = step(name, 'parseXml', () => parseXml(text, options));
  if (doc === undefined) return;
  const s1 = step(name, 'serializeXml', () => serializeXml(doc));
  if (s1 === undefined) return;
  let s2: string;
  try {
    s2 = serializeXml(parseXml(s1, options));
  } catch (e) {
    throw new Error(`${name}: serialized output does not parse back: ${describeErr(e)}`);
  }
  if (s2 !== s1)
    throw new Error(`${name}: serialize(parse(serialize(doc))) differs from serialize(doc)`);
}

const BOMS: readonly number[][] = [[], [], [0xef, 0xbb, 0xbf], [0xff, 0xfe], [0xfe, 0xff]];
const BAD_UTF8: readonly number[][] = [
  [0xff],
  [0xc0, 0x80],
  [0x80],
  [0xed, 0xa0, 0x80],
  [0xf4, 0x90, 0x80, 0x80],
  [0xe2, 0x82],
];

/** Mutated text as bytes: invalid UTF-8 sequences, sometimes behind a UTF-8 or UTF-16 BOM. */
function xmlBytes(r: Rng, text: string): Uint8Array {
  const body = Array.from(strToU8(text));
  for (let i = int(r, 3); i > 0; i--) body.splice(int(r, body.length + 1), 0, ...pick(r, BAD_UTF8));
  return Uint8Array.from([...pick(r, BOMS), ...body]);
}

// ------------------------------------------------------------------ targets

describe('fuzz corpus', () => {
  it('reads cleanly before mutation', () => {
    expect(DOCX.length).toBeGreaterThan(0);
    for (const [i, d] of DOCX.entries()) {
      expect(() => readDocx(d).text(), `corpus#${i}`).not.toThrow();
      timed(`corpus#${i}`, d, 'docx', () => runDocx(`corpus#${i}`, d, true));
    }
  });
});

describe('fuzz: parseXml', () => {
  const corpus = xmlCorpus();

  it(
    'mutated XML strings only throw XmlParseError; parsed ones serialize to a fixpoint',
    () => {
      const r = prng(11 + SEED);
      for (let i = 0; i < ITER; i++) {
        const name = `parseXml#${i} (seed ${11 + SEED})`;
        const text = mutateXml(r, pick(r, corpus));
        const maxDepth = r() < 0.5 ? 64 : undefined;
        timed(name, text, 'xml', () => runXml(name, text, maxDepth));
      }
    },
    TEST_MS,
  );

  it(
    'mutated XML bytes (BOMs, invalid UTF-8) only throw XmlParseError',
    () => {
      const r = prng(12 + SEED);
      for (let i = 0; i < ITER; i++) {
        const name = `parseXmlBytes#${i} (seed ${12 + SEED})`;
        const bytes = xmlBytes(r, mutateXml(r, pick(r, corpus)));
        timed(name, bytes, 'xml', () => {
          step(name, 'parseXmlBytes', () => parseXmlBytes(bytes, { maxDepth: 64 }));
        });
      }
    },
    TEST_MS,
  );

  it('nesting beyond maxDepth throws XmlParseError (never a stack overflow)', () => {
    expect(() => parseXml('<a>'.repeat(100_000))).toThrow(XmlParseError);
    expect(() => parseXml(`${'<a>'.repeat(65)}${'</a>'.repeat(65)}`, { maxDepth: 64 })).toThrow(
      XmlParseError,
    );
    expect(() =>
      parseXml(`${'<a>'.repeat(64)}${'</a>'.repeat(64)}`, { maxDepth: 64 }),
    ).not.toThrow();
  });

  it('deep trees within a raised limit parse and serialize without recursion', () => {
    const depth = 50_000;
    const xml = `<a>${'<b>'.repeat(depth)}x${'</b>'.repeat(depth)}</a>`;
    const doc = parseXml(xml, { maxDepth: depth + 1 });
    expect(serializeXml(doc)).toContain('<b><b>x</b></b>');
  });

  // Cases the fuzzer found once, kept so they stay fixed whatever the seed.
  it.each([
    // an (illegal) prefix undeclaration on a descendant of a declaring element
    '<w:a xmlns:w="urn:w"><w:b xmlns:w="" w:c="1"/></w:a>',
    '<w:a xmlns:w="urn:w"><w:b w:c="1" xmlns:w=""/></w:a>',
  ])('found by fuzzing: %s', (xml) => {
    runXml(xml, xml, undefined);
  });
});

describe('fuzz: unzipLazy', () => {
  const zips = [...DOCX, zipSync({ 'a.txt': strToU8('x'.repeat(5000)), 'b/': new Uint8Array(0) })];

  it(
    'mutated and truncated zips only throw TurbowordError',
    () => {
      const r = prng(21 + SEED);
      for (let i = 0; i < ITER; i++) {
        const name = `unzipLazy#${i} (seed ${21 + SEED})`;
        const src = pick(r, zips);
        const data = r() < 0.5 ? mutateBytes(r, src) : mutateZipHeader(r, src);
        const limits = r() < 0.2 ? { maxEntryBytes: 4096, maxCompressionRatio: 5 } : undefined;
        timed(name, data, 'zip', () => runZip(name, data, limits));
      }
    },
    TEST_MS,
  );
});

describe('fuzz: readDocx -> text() -> writeDocx', () => {
  it(
    'mutated docx bytes only throw TurbowordError',
    () => {
      const r = prng(31 + SEED);
      for (let i = 0; i < ITER; i++) {
        const name = `docx-bytes#${i} (seed ${31 + SEED})`;
        const src = pick(r, DOCX);
        const data = r() < 0.5 ? mutateBytes(r, src) : mutateZipHeader(r, src);
        timed(name, data, 'docx', () => runDocx(name, data, false));
      }
    },
    TEST_MS,
  );

  it(
    'docx with mutated XML parts only throw TurbowordError; force-dirty saves read back',
    () => {
      const r = prng(32 + SEED);
      for (let i = 0; i < ITER; i++) {
        const name = `docx-parts#${i} (seed ${32 + SEED})`;
        const data = mutateParts(r, pick(r, DOCX));
        const limits = r() < 0.3 ? { maxXmlDepth: 64 } : undefined;
        timed(name, data, 'docx', () => runDocx(name, data, true, limits));
      }
    },
    TEST_MS,
  );

  it('a 100k-deep document.xml fails cleanly and a 20k-deep one works with a raised limit', () => {
    const files = unzipSync(synthetic());
    const deep = (n: number) =>
      strToU8(
        story(
          'document',
          `<w:body><w:p><w:r><w:t>${'<w:x>'.repeat(n)}deep${'</w:x>'.repeat(n)}</w:t></w:r></w:p></w:body>`,
        ),
      );
    files['word/document.xml'] = deep(100_000);
    expect(() => readDocx(zipSync(files)).text()).toThrow(TurbowordError);
    files['word/document.xml'] = deep(20_000);
    const doc = readDocx(zipSync(files), { limits: { maxXmlDepth: 30_000 } });
    expect(() => doc.text()).not.toThrow();
    doc.mainPart.xml().markDirty();
    expect(() =>
      readDocx(writeDocx(doc), { limits: { maxXmlDepth: 30_000 } }).text(),
    ).not.toThrow();
  });
});
