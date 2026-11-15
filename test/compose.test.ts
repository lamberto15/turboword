/// <reference types="node" />
// Compose engine, core: the orchestration (appendDocument, insertDocument, composeDocuments,
// importBlocks), PartCopier, reconcileNamespaces and renumberIds, checked on synthetic documents
// and on real fixture combinations (test/fixtures/compose: docxcompose and Open-XML-PowerTools
// DocumentBuilder test files, plus the main corpus).
//
// Every composed result is saved, re-opened and checked for the invariants Word cares about:
// unique ids, every relationship id resolving, complete content types, legal sectPr placement,
// cells ending with a paragraph, namespace/markup-compatibility consistency, comment and note
// parts in sync, validatePackage without new errors, and (when the inputs are plain appends)
// per-section header/footer text equal to the sources'. A final test converts a sample with
// LibreOffice (skipped when it is not installed) and compares page counts.

import { spawnSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { DOMParser } from '@xmldom/xmldom';
import { describe, expect, it } from 'vitest';
import {
  appendDocument,
  composeDocuments,
  type ImportOptions,
  importBlocks,
  insertDocument,
} from '../src/compose/index.js';
import { namePattern } from '../src/compose/part-copier.js';
import type { Warning } from '../src/core/warnings.js';
import type { Part } from '../src/opc/part.js';
import { relTypeIs } from '../src/opc/rel-types.js';
import { createDocx } from '../src/wml/create.js';
import { readDocx, type WordDocument, writeDocx } from '../src/wml/document.js';
import { validatePackage } from '../src/wml/validate.js';
import type { XElement } from '../src/xml/dom.js';
import { IGNORABLE_NS, NS } from '../src/xml/ns.js';
import { parseXmlBytes } from '../src/xml/parse.js';
import { fixture, strictFixtures } from './helpers/fixtures.js';

const DIR = fileURLToPath(new URL('./fixtures/compose/', import.meta.url));
const W = NS.w;
const enc = new TextEncoder();

// ---- loading ------------------------------------------------------------------------------

/** A compose fixture (`docxcompose/x.docx`, `powertools/y.docx`) or a main-corpus fixture. */
function bytesOf(name: string): Uint8Array {
  const p = join(DIR, name);
  return existsSync(p) ? new Uint8Array(readFileSync(p)) : fixture(name).bytes;
}
const open = (name: string): WordDocument => readDocx(bytesOf(name));

function composeFixtures(): string[] {
  const out: string[] = [];
  for (const d of ['docxcompose', 'powertools']) {
    for (const f of readdirSync(join(DIR, d)).sort())
      if (f.endsWith('.docx')) out.push(`${d}/${f}`);
  }
  return out;
}

// ---- invariants ---------------------------------------------------------------------------

/** Namespaces that must be ignorable (or required by an enclosing mc:Choice) where used. */
const EXTENSION_NS: ReadonlySet<string> = new Set([
  ...IGNORABLE_NS,
  NS.wps,
  NS.wpg,
  NS.wpc,
  NS.wpi,
]);
const STORY_RELS = ['header', 'footer', 'footnotes', 'endnotes', 'comments'];

function root(part: Part | undefined): XElement | undefined {
  if (part === undefined) return undefined;
  try {
    return part.xml().root;
  } catch {
    return undefined;
  }
}

/** The main part and its story parts (headers, footers, notes, comments), each once. */
function storyParts(doc: WordDocument): Part[] {
  const out: Part[] = [doc.mainPart];
  for (const rel of doc.mainPart.rels.all()) {
    if (rel.external || !STORY_RELS.some((k) => relTypeIs(rel.type, k))) continue;
    const p = doc.mainPart.related(rel.id);
    if (p !== undefined && !out.includes(p)) out.push(p);
  }
  return out;
}

const inFallback = (el: XElement): boolean => el.closest(NS.mc, 'Fallback') !== undefined;

function count<T>(values: readonly T[]): Map<T, number> {
  const m = new Map<T, number>();
  for (const v of values) m.set(v, (m.get(v) ?? 0) + 1);
  return m;
}
const dupes = (values: readonly string[]): number =>
  [...count(values).values()].reduce((n, c) => n + (c > 1 ? c - 1 : 0), 0);

/** Numbers of each kind of invariant violation (compared against the inputs' own). */
function violations(doc: WordDocument): Map<string, number> {
  const v = new Map<string, number>();
  const add = (k: string, n = 1): void => {
    if (n > 0) v.set(k, (v.get(k) ?? 0) + n);
  };
  // validatePackage errors, by code.
  for (const i of validatePackage(doc)) if (i.severity === 'error') add(`validate:${i.code}`);

  // Document-wide unique ids.
  const docPr: string[] = [];
  const paraIds: string[] = [];
  const sdtIds: string[] = [];
  const bookmarkIds: string[] = [];
  const bookmarkNames: string[] = [];
  for (const part of storyParts(doc)) {
    const r = root(part);
    if (r === undefined) continue;
    for (const el of [r, ...r.descendants()]) {
      if (inFallback(el)) continue;
      if (el.is(NS.wp, 'docPr')) docPr.push(el.attr('id') ?? '');
      const pid = el.attrNS(NS.w14, 'paraId');
      if (pid !== undefined) paraIds.push(pid.toUpperCase());
      if (el.is(W, 'id') && (el.parent as XElement).is(W, 'sdtPr'))
        sdtIds.push(el.attrNS(W, 'val') ?? '');
      if (el.is(W, 'bookmarkStart')) {
        bookmarkIds.push(el.attrNS(W, 'id') ?? '');
        bookmarkNames.push(el.attrNS(W, 'name') ?? '');
      }
    }
  }
  add('dup:docPr', dupes(docPr));
  add('dup:paraId', dupes(paraIds));
  add('dup:sdt', dupes(sdtIds));
  add('dup:bookmarkId', dupes(bookmarkIds));
  add('dup:bookmarkName', dupes(bookmarkNames));

  // o:relid resolves (validatePackage checks the r: namespace only).
  for (const part of doc.package.parts()) {
    if (!part.isXml || part.name.endsWith('.rels')) continue;
    const r = root(part);
    if (r === undefined) continue;
    for (const el of r.descendants(NS.v, 'imagedata')) {
      const id = el.attrNS(NS.o, 'relid');
      if (id !== undefined && id !== '' && part.rels.get(id) === undefined) add('rel:o:relid');
    }
  }

  // Notes and comments in sync with their references.
  const notes = (rel: string, local: string, ref: string): void => {
    const p = doc.mainPart.relatedByType(rel)[0];
    const ids = new Set((root(p)?.elements(W, local) ?? []).map((n) => n.attrNS(W, 'id')));
    add(
      `dup:${local}`,
      dupes((root(p)?.elements(W, local) ?? []).map((n) => n.attrNS(W, 'id') ?? '')),
    );
    for (const part of storyParts(doc)) {
      for (const r of root(part)?.descendants(W, ref) ?? []) {
        if (!ids.has(r.attrNS(W, 'id'))) add(`ref:${ref}`);
      }
    }
  };
  notes('footnotes', 'footnote', 'footnoteReference');
  notes('endnotes', 'endnote', 'endnoteReference');
  notes('comments', 'comment', 'commentReference');
  const comments = root(doc.mainPart.relatedByType('comments')[0]);
  const commentIds = new Set(
    (comments?.elements(W, 'comment') ?? []).map((c) => c.attrNS(W, 'id')),
  );
  for (const part of storyParts(doc)) {
    for (const local of ['commentRangeStart', 'commentRangeEnd']) {
      for (const r of root(part)?.descendants(W, local) ?? []) {
        if (!commentIds.has(r.attrNS(W, 'id'))) add(`ref:${local}`);
      }
    }
  }
  const commentParaIds = new Set(
    (comments?.descendants(W, 'p') ?? []).map((p) => p.attrNS(NS.w14, 'paraId')?.toUpperCase()),
  );
  const ext = root(doc.mainPart.relatedByType('commentsExtended')[0]);
  for (const c of ext?.elements(NS.w15, 'commentEx') ?? []) {
    if (!commentParaIds.has(c.attrNS(NS.w15, 'paraId')?.toUpperCase()))
      add('sync:commentsExtended');
  }
  const cids = root(doc.mainPart.relatedByType('commentsIds')[0]);
  const durable: string[] = [];
  for (const c of cids?.elements(NS.w16cid, 'commentId') ?? []) {
    if (!commentParaIds.has(c.attrNS(NS.w16cid, 'paraId')?.toUpperCase())) add('sync:commentsIds');
    durable.push(c.attrNS(NS.w16cid, 'durableId')?.toUpperCase() ?? '');
  }
  add('dup:durableId', dupes(durable));
  const cex = root(doc.mainPart.relatedByType('commentsExtensible')[0]);
  if (cids !== undefined) {
    for (const c of cex?.elements(NS.w16cex, 'commentExtensible') ?? []) {
      if (!durable.includes(c.attrNS(NS.w16cex, 'durableId')?.toUpperCase() ?? '')) {
        add('sync:commentsExtensible');
      }
    }
  }

  // Namespaces: every XML part parses with a namespace-aware parser; mc lists name declared
  // prefixes; extension namespaces are used only where ignorable or required.
  for (const part of doc.package.parts()) {
    if (!part.isXml) continue;
    const bytes = part.bytes();
    try {
      new DOMParser({ onError: () => undefined }).parseFromString(
        new TextDecoder().decode(bytes),
        'text/xml',
      );
    } catch {
      add('ns:unbound-prefix');
      continue;
    }
    if (part.name.endsWith('.rels') || part.name === '[Content_Types].xml') continue;
    const r = parseXmlBytes(bytes).root;
    add('mc', mcProblems(r));
  }
  return v;
}

/** Undeclared prefixes in mc lists + extension-namespace uses outside ignorable/required scope. */
function mcProblems(top: XElement): number {
  // A part whose vocabulary is an extension namespace (commentsIds, people, ...) is only read by
  // consumers that understand it.
  if (EXTENSION_NS.has(top.uri)) return 0;
  let n = 0;
  const stack: [XElement, ReadonlySet<string>][] = [[top, new Set()]];
  while (stack.length > 0) {
    const [el, outer] = stack.pop()!;
    let ok = outer;
    const resolve = (p: string): string | undefined => el.lookupNamespace(p);
    for (const a of el.attributes()) {
      const isList =
        (a.uri === NS.mc &&
          /(?:^|:)(?:Ignorable|MustUnderstand|ProcessContent|PreserveElements|PreserveAttributes)$/.test(
            a.name,
          )) ||
        (a.name === 'Requires' && el.is(NS.mc, 'Choice'));
      if (!isList) continue;
      for (const tok of a.value.split(/\s+/).filter(Boolean)) {
        const p = tok.includes(':') ? tok.slice(0, tok.indexOf(':')) : tok;
        const uri = resolve(p);
        if (uri === undefined) n++;
        else if (/(?:^|:)Ignorable$/.test(a.name) || a.name === 'Requires') {
          if (ok === outer) ok = new Set(outer);
          (ok as Set<string>).add(uri);
        }
      }
    }
    const used = [el.uri, ...el.attributes().map((a) => a.uri)];
    for (const u of used) if (EXTENSION_NS.has(u) && !ok.has(u)) n++;
    for (const k of el.elements()) stack.push([k, ok]);
  }
  return n;
}

/** Input violations by input bytes (inputs are re-opened per scenario; checking them is slow). */
const inputViolations = new Map<string, Map<string, number>>();

function violationsOfInput(d: WordDocument): Map<string, number> {
  if (d.package.dirty) return violations(d);
  const bytes = writeDocx(d);
  let h = 0;
  for (let i = 0; i < bytes.length; i += 7) h = (Math.imul(h, 31) + bytes[i]!) | 0;
  const key = `${bytes.length}:${h}`;
  let v = inputViolations.get(key);
  if (v === undefined) {
    v = violations(d);
    inputViolations.set(key, v);
  }
  return v;
}

/** Violations in `result` beyond the sum of the inputs' own (as a readable list). */
function newViolations(result: WordDocument, inputs: readonly WordDocument[]): string[] {
  const reread = readDocx(writeDocx(result));
  const base = new Map<string, number>();
  for (const d of inputs) {
    for (const [k, n] of violationsOfInput(d)) base.set(k, (base.get(k) ?? 0) + n);
  }
  const out: string[] = [];
  for (const [k, n] of violations(reread)) {
    if (n > (base.get(k) ?? 0)) out.push(`${k}: ${n} (inputs ${base.get(k) ?? 0})`);
  }
  if (out.some((s) => s.startsWith('validate:'))) {
    for (const i of validatePackage(reread))
      if (i.severity === 'error') out.push(`  ${i.code} ${i.part ?? ''} ${i.message}`);
  }
  return out;
}

// ---- headers and footers per section --------------------------------------------------------

function evenAndOdd(doc: WordDocument): boolean {
  const s = root(doc.mainPart.relatedByType('settings')[0])?.element(W, 'evenAndOddHeaders');
  if (s === undefined) return false;
  const v = s.attrNS(W, 'val');
  return v === undefined || !['0', 'false', 'off'].includes(v);
}

/** Effective header/footer texts of every section, for the slots the document shows. */
function headerTexts(doc: WordDocument, eo = evenAndOdd(doc)): string[] {
  const out: string[] = [];
  for (const s of doc.sections()) {
    const types: ('default' | 'first' | 'even')[] = ['default'];
    if (s.titlePage) types.push('first');
    if (eo) types.push('even');
    for (const t of types) {
      out.push(`${s.index}/${t}/header: ${(s.header(t)?.text() ?? '').trim()}`);
      out.push(`${s.index}/${t}/footer: ${(s.footer(t)?.text() ?? '').trim()}`);
    }
  }
  return out;
}

/** The header/footer texts each input section should still show in an append result. */
function expectedHeaderTexts(inputs: readonly WordDocument[]): string[] {
  const out: string[] = [];
  let offset = 0;
  for (const d of inputs) {
    for (const line of headerTexts(d)) {
      const [idx, ...rest] = line.split('/');
      out.push(`${Number(idx) + offset}/${rest.join('/')}`);
    }
    offset += d.sections().length;
  }
  return out;
}

/** Result texts restricted to the slots each input showed (the result may show more). */
function resultHeaderTexts(result: WordDocument, inputs: readonly WordDocument[]): string[] {
  const want = new Set(expectedHeaderTexts(inputs).map((l) => l.slice(0, l.indexOf(':'))));
  const all = headerTexts(result, true).concat(headerTexts(result, false));
  const seen = new Set<string>();
  return all.filter((l) => {
    const k = l.slice(0, l.indexOf(':'));
    if (!want.has(k) || seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

// ---- scenario helpers -----------------------------------------------------------------------

interface Composed {
  result: WordDocument;
  inputs: WordDocument[];
  warnings: Warning[];
}

/** composeDocuments over fresh copies of `names` (inputs re-opened for comparisons). */
function composeNames(names: readonly string[], o: ImportOptions = {}): Composed {
  const warnings: Warning[] = [];
  const docs = names.map(open);
  const result = composeDocuments(docs, { ...o, onWarning: (w) => warnings.push(w) });
  for (const d of docs) expect(d.package.dirty, 'inputs untouched').toBe(false);
  return { result, inputs: names.map(open), warnings };
}

function expectClean(c: Composed): void {
  expect(newViolations(c.result, c.inputs)).toEqual([]);
}

/** A document whose body is `body` (raw WordprocessingML, `w`/`r`/`w14` etc. declared). */
function docWithBody(body: string, rootAttrs = ''): WordDocument {
  const doc = createDocx();
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W}" xmlns:r="${NS.r}" xmlns:mc="${NS.mc}" xmlns:w14="${NS.w14}" ` +
    `xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" ${rootAttrs}><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>';
  doc.mainPart.setBytes(enc.encode(xml));
  return readDocx(writeDocx(doc));
}

const bodyXml = (doc: WordDocument): string => new TextDecoder().decode(doc.mainPart.bytes());

// ---- fixture combinations -------------------------------------------------------------------

/** The combinations docxcompose's own test-suite composes (tests/test_*.py). */
const DOCXCOMPOSE_SCENARIOS: readonly (readonly string[])[] = [
  ['master.docx', 'embedded_excel_chart.docx'],
  ['master.docx', 'footnote.docx'],
  ['master.docx', 'footnotes_with_hyperlinks.docx'],
  ['master_header_footer.docx', 'header_footer_sections.docx'],
  ['header_footer.docx', 'header_footer.docx'],
  ['header_footer_sections.docx', 'header_footer_sections.docx'],
  ['master.docx', 'hyperlinks.docx'],
  ['master.docx', 'images.docx'],
  ['master.docx', 'embedded_and_external_image.docx'],
  ['header_with_image.docx', 'image.docx'],
  ['common_stylename_different_id1.docx', 'common_stylename_different_id2.docx'],
  ['numberings.docx', 'numberings.docx'],
  ['numberings_restart.docx', 'numberings_restart.docx'],
  ['numberings_styles.docx', 'numberings_styles.docx'],
  ['numbering_reference_to_numbering_zero.docx', 'numbering_reference_to_numbering_zero.docx'],
  ['numbering_with_paragraphs_in_between.docx', 'numbering_with_paragraphs_in_between.docx'],
  ['broken_listing_master.docx', 'broken_listing.docx'],
  ['broken_listing_master.docx', 'broken_listing_many.docx'],
  ['broken_listing_master.docx', 'broken_listing_nested.docx'],
  ['continous_section_break.docx', 'continous_section_break.docx'],
  ['continous_section_break.docx', 'odd_section_break.docx'],
  ['next_page_section_break.docx', 'legacy_empty.docx'],
  ['master.docx', 'next_page_section_break_empty.docx'],
  ['master.docx', 'embedded_visio.docx'],
  ['master.docx', 'smart_art.docx'],
  ['styles_de.docx', 'styles_en.docx'],
  ['styles_en.docx', 'styles_de.docx'],
  ['master_switched_listing_style.docx', 'switched_listing_style.docx'],
  ['master.docx', 'table.docx'],
  [
    'docproperties_header_footer_3_sections.docx',
    'docproperties_different_odd_even_pages_1_section.docx',
  ],
  [
    'docproperties_different_first_page_1_section.docx',
    'docproperties_header_footer_3_sections.docx',
  ],
  [
    'docproperties_shared_header_footer_2_sections.docx',
    'master_header_footer.docx',
    'header_footer.docx',
  ],
].map((s) => s.map((n) => `docxcompose/${n}`));

/** Open-XML-PowerTools DocumentBuilder test combinations (DocumentBuilderTests.cs). */
const POWERTOOLS_SCENARIOS: readonly (readonly string[])[] = [
  ['DB001-Sections.docx', 'DB001-Sections.docx'],
  ['DB002-Sections-With-Headers.docx', 'DB002-Landscape-Section.docx'],
  ['DB003-Only-Default-Header.docx', 'DB002-Landscape-Section.docx'],
  ['DB004-No-Headers.docx', 'DB002-Landscape-Section.docx'],
  ['DB005-Headers-With-Images.docx', 'DB002-Landscape-Section.docx'],
  ['DB006-Source1.docx', 'DB006-Source3.docx'],
  ['DB007-Spec.docx', 'DB007-WhitePaper.docx', 'DB007-Abstract.docx', 'DB007-AuthorBiography.docx'],
  ['DB007-Notes.docx', 'DB007-Notes.docx'],
  ['DB011-Header-With-Shape.docx', 'DB011-Body-With-Shape.docx'],
  ['DB012-Lists-With-Different-Numberings.docx', 'DB012-Lists-With-Different-Numberings.docx'],
  ['DB013a-Red-Heading1-English.docx', 'DB013a-Green-Heading1-Danish.docx'],
  ['DB013b-Orange-List-Danish.docx', 'DB013b-Blue-List-English.docx'],
  ['DB014-WebExtensions.docx', 'DB014-WebExtensions.docx'],
  [
    'DB__GlossaryDocuments__BaseDocument.docx',
    'DB__GlossaryDocuments__CellLevelContentControl.docx',
    'DB__GlossaryDocuments__NestedContentControl.docx',
  ],
  ['DB__GlossaryDocuments__HeaderContent.docx', 'DB__GlossaryDocuments__FooterContent.docx'],
].map((s) => s.map((n) => `powertools/${n}`));

const HF_DEST = [
  'Fax',
  'Fax_(content_control)',
  'Letter',
  'Letter_(content_control)',
  'Letter_(run-level_insert_element)',
];
const HF_SRC = [
  'Content-Controls',
  'Disclaimer',
  'Footer',
  'Letterhead-with-Watermark',
  'Letterhead',
  'Logo',
  'Watermark-1',
  'Watermark-2',
];

describe('appendDocument: docxcompose and PowerTools combinations', () => {
  for (const names of [...DOCXCOMPOSE_SCENARIOS, ...POWERTOOLS_SCENARIOS]) {
    it(names.map((n) => basename(n, '.docx')).join(' + '), () => {
      const c = composeNames(names);
      expectClean(c);
      expect(resultHeaderTexts(c.result, c.inputs)).toEqual(expectedHeaderTexts(c.inputs));
      expect(c.result.sections().length).toBe(
        c.inputs.reduce((n, d) => n + d.sections().length, 0),
      );
      // Body text of every input is there, in order.
      const text = c.result.text();
      let at = 0;
      for (const d of c.inputs) {
        const t = d.text().trim();
        if (t === '') continue;
        const i = text.indexOf(t.slice(0, 200), at);
        expect(i, `text of ${names.join(',')}`).toBeGreaterThanOrEqual(at);
        at = i;
      }
    });
  }

  it('PowerTools header/footer sources appended to every destination (40 combinations)', () => {
    const problems: string[] = [];
    for (const d of HF_DEST) {
      for (const s of HF_SRC) {
        const names = [
          `powertools/DB__HeadersFooters__Dest__${d}.docx`,
          `powertools/DB__HeadersFooters__Src__${s}.docx`,
        ];
        const c = composeNames(names);
        for (const p of newViolations(c.result, c.inputs)) problems.push(`${d}+${s}: ${p}`);
        const got = resultHeaderTexts(c.result, c.inputs);
        const want = expectedHeaderTexts(c.inputs);
        if (JSON.stringify(got) !== JSON.stringify(want))
          problems.push(`${d}+${s}: headers ${JSON.stringify(got)} != ${JSON.stringify(want)}`);
      }
    }
    expect(problems).toEqual([]);
  });
});

describe('appendDocument: every fixture', () => {
  const all = [
    ...composeFixtures(),
    ...[
      'docx-mailmerge/test_footnote_header_footer.docx',
      'python-docx/tbl-cell-access.docx',
      'python-docx/comments-rich-para.docx',
      'apache-poi/delins.docx',
      'apache-poi/shapes-with-text.docx',
      'apache-poi/bug69628.docx',
      'apache-poi/libreoffice-comment.docx',
      'open-xml-sdk/TOC2.docx',
      'open-xml-sdk/TableFntEdnt.docx',
      'open-xml-sdk/strict-CommentRangeStart.docx',
      'open-xml-sdk/strict-footer-bookmarkStart.docx',
    ],
  ];

  it('appended to a master, to itself, and with a master appended: no new violations', () => {
    const problems: string[] = [];
    for (const name of all) {
      for (const pair of [
        ['docxcompose/master.docx', name],
        [name, name],
        [name, 'docxcompose/master_header_footer.docx'],
      ]) {
        let c: Composed;
        try {
          c = composeNames(pair);
        } catch (e) {
          problems.push(`${pair.join(' + ')}: threw ${e instanceof Error ? e.stack : e}`);
          continue;
        }
        for (const p of newViolations(c.result, c.inputs))
          problems.push(`${pair.join(' + ')}: ${p}`);
        const got = resultHeaderTexts(c.result, c.inputs);
        const want = expectedHeaderTexts(c.inputs);
        if (JSON.stringify(got) !== JSON.stringify(want)) {
          problems.push(
            `${pair.join(' + ')}: headers/footers differ\n    got  ${JSON.stringify(got)}\n    want ${JSON.stringify(want)}`,
          );
        }
      }
    }
    expect(problems).toEqual([]);
  });

  it('synthetic Strict documents compose with transitional ones both ways', () => {
    const strict = strictFixtures([
      'docxcompose/header_with_image.docx',
      'python-docx/comments-rich-para.docx',
    ]);
    for (const s of strict) {
      for (const [a, b] of [
        [readDocx(s.bytes), open('docxcompose/master.docx')],
        [open('docxcompose/master.docx'), readDocx(s.bytes)],
      ] as const) {
        const inputs = [readDocx(writeDocx(a)), readDocx(writeDocx(b))];
        const result = composeDocuments([a, b]);
        expect(newViolations(result, inputs), s.name).toEqual([]);
      }
    }
    // New declarations in a Strict target use Strict URIs.
    const target = readDocx(strict[0]!.bytes);
    appendDocument(target, open('docxcompose/smart_art.docx'));
    const decls = target.mainPart
      .xml()
      .root.attributes()
      .filter((a) => a.name.startsWith('xmlns:'));
    expect(
      decls.some((a) => a.value.includes('schemas.openxmlformats.org/drawingml/2006/main')),
    ).toBe(false);
  });
});

// ---- PartCopier ---------------------------------------------------------------------------

describe('PartCopier', () => {
  it('names copies after the source pattern', () => {
    expect(namePattern('word/media/image7.png')).toBe('word/media/image{n}.png');
    expect(namePattern('word/embeddings/Microsoft_Excel_Worksheet12.xlsx')).toBe(
      'word/embeddings/Microsoft_Excel_Worksheet{n}.xlsx',
    );
    expect(namePattern('customXml/item1.xml')).toBe('customXml/item{n}.xml');
    expect(namePattern('word/media/logo.png')).toBe('word/media/logo{n}.png');
    expect(namePattern('word/123.bin')).toBe('word/part{n}.bin');
  });

  it('copies images with their content types and dedupes identical media', () => {
    const target = open('docxcompose/master.docx');
    const media = (): string[] =>
      target.package
        .parts()
        .map((p) => p.name)
        .filter((n) => n.includes('/media/'));
    const r1 = appendDocument(target, open('docxcompose/images.docx'));
    const after1 = media();
    expect(after1.length).toBeGreaterThan(0);
    for (const n of after1) expect(target.package.contentTypes.get(n)).toMatch(/^image\//);
    expect(r1.addedParts.filter((n) => n.includes('/media/')).length).toBe(after1.length);
    const r2 = appendDocument(target, open('docxcompose/images.docx'));
    expect(media()).toEqual(after1);
    expect(r2.addedParts.filter((n) => n.includes('/media/'))).toEqual([]);
  });

  it('re-creates external relationships as external', () => {
    const target = open('docxcompose/master.docx');
    const src = open('docxcompose/hyperlinks.docx');
    const external = src.mainPart.rels.all().filter((r) => r.external);
    expect(external.length).toBeGreaterThan(0);
    appendDocument(target, src);
    const got = target.mainPart.rels
      .all()
      .filter((r) => r.external)
      .map((r) => r.target);
    for (const r of external) expect(got).toContain(r.target);
  });

  it('copies charts with their embedding, colors and style', () => {
    const target = open('docxcompose/master.docx');
    const r = appendDocument(target, open('docxcompose/embedded_excel_chart.docx'));
    const chart = r.addedParts.find((n) => /charts\/chart\d+\.xml$/.test(n));
    expect(chart).toBeDefined();
    const part = target.package.part(chart!)!;
    const kinds = part.rels.all().map((x) => x.type.slice(x.type.lastIndexOf('/') + 1));
    expect(kinds).toContain('package');
    for (const rel of part.rels.all()) {
      if (!rel.external) expect(target.package.part(rel.partName!), rel.target).toBeDefined();
    }
    expect(validatePackage(target).filter((i) => i.severity === 'error')).toEqual([]);
  });

  it('copies SmartArt parts and rewrites dataModelExt/@relId to the host relationship', () => {
    const target = open('docxcompose/master.docx');
    const r = appendDocument(target, open('docxcompose/smart_art.docx'));
    const data = r.addedParts.find((n) => n.includes('diagrams/data'));
    expect(data).toBeDefined();
    const ext = target.package
      .part(data!)!
      .xml()
      .root.descendants(
        'http://schemas.microsoft.com/office/drawing/2008/diagram',
        'dataModelExt',
      )[0];
    expect(ext).toBeDefined();
    const rel = target.mainPart.rels.get(ext!.attr('relId')!);
    expect(rel?.type).toMatch(/diagramDrawing$/);
    expect(r.addedParts).toContain(rel!.partName);
    for (const id of ['dm', 'lo', 'qs', 'cs']) {
      const relIds = target.mainPart.xml().root.descendants(NS.dgm, 'relIds');
      for (const el of relIds) expect(target.mainPart.rels.get(el.attrNS(NS.r, id)!)).toBeDefined();
    }
  });

  it('never copies a glossary document', () => {
    const target = open('docxcompose/master.docx');
    const r = appendDocument(target, open('docxcompose/content_controls.docx'));
    expect(r.addedParts.filter((n) => n.includes('glossary'))).toEqual([]);
    expect(target.package.parts().some((p) => p.name.includes('glossary'))).toBe(false);
  });

  it('brings customXml items used by data bindings once (same storeItemID)', () => {
    const storeId = '{6C3C8BC8-F283-45AE-878A-BAB7291924A1}';
    const src = docWithBody(
      `<w:sdt><w:sdtPr><w:id w:val="5"/><w:dataBinding w:xpath="/root/a" w:storeItemID="${storeId}"/></w:sdtPr>` +
        '<w:sdtContent><w:p><w:r><w:t>bound</w:t></w:r></w:p></w:sdtContent></w:sdt>',
    );
    const pkg = src.package;
    pkg.addPart('customXml/item1.xml', 'application/xml', enc.encode('<root><a>x</a></root>'));
    pkg.addPart(
      'customXml/itemProps1.xml',
      'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
      enc.encode(
        `<ds:datastoreItem ds:itemID="${storeId}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>`,
      ),
    );
    pkg
      .part('customXml/item1.xml')!
      .rels.add(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps',
        'itemProps1.xml',
      );
    src.mainPart.rels.add(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
      '../customXml/item1.xml',
    );
    const source = readDocx(writeDocx(src));

    const target = createDocx();
    const r1 = appendDocument(target, source);
    expect(r1.addedParts).toEqual(
      expect.arrayContaining(['customXml/item1.xml', 'customXml/itemProps1.xml']),
    );
    const r2 = appendDocument(target, source);
    expect(r2.addedParts.filter((n) => n.startsWith('customXml/'))).toEqual([]);
    expect(validatePackage(target).filter((i) => i.severity === 'error')).toEqual([]);
  });
});

// ---- namespaces ---------------------------------------------------------------------------

describe('reconcileNamespaces', () => {
  /** A source whose WordprocessingML prefix is `ns0` and whose extensions use odd prefixes. */
  function oddSource(): WordDocument {
    const doc = createDocx();
    const xml =
      `<ns0:document xmlns:ns0="${W}" xmlns:x14="${NS.w14}" xmlns:m2="${NS.mc}" xmlns:shp="${NS.wps}" ` +
      `xmlns:o="urn:example:not-office" xmlns:r="${NS.r}" m2:Ignorable="x14"><ns0:body>` +
      '<ns0:p x14:paraId="1234ABCD" x14:textId="77777777"><ns0:r><ns0:t>Hello</ns0:t></ns0:r></ns0:p>' +
      '<ns0:p><ns0:r><m2:AlternateContent><m2:Choice Requires="shp"><ns0:t>choice</ns0:t></m2:Choice>' +
      '<m2:Fallback><ns0:t>fallback</ns0:t></m2:Fallback></m2:AlternateContent></ns0:r></ns0:p>' +
      '<ns0:p><ns0:r xmlns:foo="urn:example:foo" m2:Ignorable="foo x14" foo:bar="1"><ns0:t>t</ns0:t></ns0:r>' +
      '<o:custom m2:ProcessContent="o:*"/></ns0:p>' +
      '<ns0:sectPr/></ns0:body></ns0:document>';
    doc.mainPart.setBytes(enc.encode(xml));
    return readDocx(writeDocx(doc));
  }

  it('renames prefixes to the target root bindings and rewrites mc prefix lists', () => {
    const target = createDocx();
    appendDocument(target, oddSource());
    const root = target.mainPart.xml().root;
    const ps = target.body.element.elements(W, 'p');
    const imported = ps.filter(
      (p) => p.text.includes('Hello') || p.text.includes('choice') || p.text.includes('t'),
    );
    for (const p of imported) {
      for (const el of [p, ...p.descendants()]) {
        if (el.uri === W) expect(el.name.startsWith('w:')).toBe(true);
        if (el.uri === NS.mc) expect(el.name.startsWith('mc:')).toBe(true);
      }
    }
    const hello = imported.find((p) => p.text === 'Hello')!;
    expect(hello.attr('w14:paraId')).toBe('1234ABCD');
    const choice = root.descendants(NS.mc, 'Choice')[0]!;
    expect(choice.attr('Requires')).toBe('wps');
    // urn:example:foo is declared on the root; the run's own mc:Ignorable names target prefixes.
    expect(root.attr('xmlns:foo')).toBe('urn:example:foo');
    const run = root.descendants(W, 'r').find((r) => r.attr('foo:bar') === '1')!;
    expect(run.attr('mc:Ignorable')).toBe('foo w14');
    expect(run.attributes().some((a) => a.name.startsWith('xmlns'))).toBe(false);
    // `o` is taken by VML in the target: the source's `o` namespace gets a new prefix.
    const custom = root.descendants('urn:example:not-office', 'custom')[0]!;
    const p = custom.name.slice(0, custom.name.indexOf(':'));
    expect(p).not.toBe('o');
    expect(root.attr(`xmlns:${p}`)).toBe('urn:example:not-office');
    expect(custom.attr('mc:ProcessContent')).toBe(`${p}:*`);
    expect(newViolations(target, [createDocx(), oddSource()])).toEqual([]);
  });

  it('declares ignorable extension namespaces on a target root that lacks them', () => {
    const target = createDocx();
    target.mainPart.setBytes(
      enc.encode(`<w:document xmlns:w="${W}"><w:body><w:p/><w:sectPr/></w:body></w:document>`),
    );
    const t = readDocx(writeDocx(target));
    appendDocument(t, oddSource());
    const root = t.mainPart.xml().root;
    expect(root.attr('xmlns:w14')).toBe(NS.w14);
    expect(root.attrNS(NS.mc, 'Ignorable')?.split(' ')).toContain('w14');
    const reread = readDocx(writeDocx(t));
    expect(mcProblems(reread.mainPart.xml().root)).toBe(0);
  });
});

// ---- ids ----------------------------------------------------------------------------------

describe('renumberIds', () => {
  it('makes drawing, paragraph, content-control, revision and VML ids unique', () => {
    for (const name of [
      'powertools/DB011-Header-With-Shape.docx',
      'docxcompose/content_controls.docx',
      'apache-poi/delins.docx',
      'apache-poi/shapes-with-text.docx',
      'python-docx/comments-rich-para.docx',
    ]) {
      const c = composeNames([name, name, name]);
      const v = violations(readDocx(writeDocx(c.result)));
      for (const k of ['dup:docPr', 'dup:paraId', 'dup:sdt'])
        expect(v.get(k) ?? 0, `${name} ${k}`).toBe(0);
      if (name.includes('delins')) {
        const ids = c.result.mainPart
          .xml()
          .root.descendants(W, 'ins')
          .map((e) => e.attrNS(W, 'id'));
        expect(new Set(ids).size).toBe(ids.length);
      }
      if (name.includes('shapes')) {
        // The source repeats some o:spid values itself; each copy must get its own set.
        const spids = (d: WordDocument): string[] =>
          d.mainPart
            .xml()
            .root.descendants()
            .map((e) => e.attrNS(NS.o, 'spid'))
            .filter((x): x is string => x !== undefined);
        const own = new Set(spids(open(name))).size;
        expect(new Set(spids(c.result)).size).toBe(3 * own);
      }
    }
  });

  it('keeps ids the target does not use; pic:cNvPr follows its docPr', () => {
    const body =
      '<w:p w14:paraId="0A0B0C0D"><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/>' +
      '<wp:docPr id="7" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="P"/><pic:cNvPicPr/></pic:nvPicPr></pic:pic>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const target = createDocx();
    appendDocument(target, docWithBody(body));
    const xml = bodyXml(target);
    expect(xml).toContain('w14:paraId="0A0B0C0D"');
    expect(xml).toContain('<wp:docPr id="7"');
    // Appending the same again must move both ids together.
    appendDocument(target, docWithBody(body));
    const root = target.mainPart.xml().root;
    const docPrs = root.descendants(NS.wp, 'docPr').map((d) => d.attr('id'));
    expect(new Set(docPrs).size).toBe(2);
    for (const inline of root.descendants(NS.wp, 'inline')) {
      const d = inline.element(NS.wp, 'docPr')!.attr('id');
      expect(inline.descendants(NS.pic, 'cNvPr')[0]!.attr('id')).toBe(d);
    }
    const pids = root
      .descendants(W, 'p')
      .map((p) => p.attrNS(NS.w14, 'paraId'))
      .filter(Boolean);
    expect(new Set(pids).size).toBe(pids.length);
  });
});

// ---- insertDocument -----------------------------------------------------------------------

describe('insertDocument', () => {
  const host = (): WordDocument =>
    docWithBody(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Before [[IN</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve">SERT]] after</w:t></w:r><w:bookmarkStart w:id="0" w:name="mark"/><w:bookmarkEnd w:id="0"/></w:p>' +
        '<w:p><w:r><w:t>Second</w:t></w:r></w:p>' +
        '<w:sdt><w:sdtPr><w:alias w:val="Slot"/><w:tag w:val="slot"/><w:id w:val="99"/><w:showingPlcHdr/></w:sdtPr>' +
        '<w:sdtContent><w:p><w:r><w:t>Click here</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>{{CELL}}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        '<w:p><w:r><w:t>End</w:t></w:r></w:p>',
    );
  const piece = (): WordDocument => open('docxcompose/table.docx');

  it('splits a placeholder paragraph across runs, both halves keeping their properties', () => {
    const target = host();
    const src = piece();
    const before = target.body.element.elements(W, 'p').length;
    const r = insertDocument(target, { placeholder: '[[INSERT]]' }, src);
    expect(r.blocks.length).toBeGreaterThan(0);
    const ps = target.body.element.elements(W, 'p');
    expect(ps.length).toBeGreaterThan(before);
    const left = ps.find((p) => p.text.startsWith('Before'))!;
    const right = ps.find((p) => p.text.endsWith('after'))!;
    expect(left.text).toBe('Before ');
    expect(right.text).toBe(' after');
    expect(target.text()).not.toContain('[[INSERT]]');
    for (const p of [left, right])
      expect(p.element(W, 'pPr')?.element(W, 'jc')?.attrNS(W, 'val')).toBe('center');
    expect(left.descendants(W, 'b').length).toBe(1);
    // The imported table sits between the halves.
    const kids = target.body.element.elements();
    expect(kids.indexOf(left)).toBeLessThan(kids.indexOf(r.blocks[0]!));
    expect(kids.indexOf(r.blocks.at(-1)!)).toBeLessThan(kids.indexOf(right));
    // The bookmark after the placeholder is still balanced.
    expect(right.descendants(W, 'bookmarkStart').length).toBe(1);
    expect(newViolations(target, [host(), piece()])).toEqual([]);
  });

  it('inserts at a bookmark, into a content control and around paragraphs', () => {
    const t1 = host();
    insertDocument(t1, { bookmark: 'mark' }, piece());
    expect(newViolations(t1, [host(), piece()])).toEqual([]);

    const t2 = host();
    const r2 = insertDocument(t2, { contentControl: { tag: 'slot' } }, piece());
    const sdt = t2.body.element.elements(W, 'sdt')[0]!;
    expect(sdt.element(W, 'sdtContent')!.elements().length).toBeGreaterThanOrEqual(
      r2.blocks.length,
    );
    expect(sdt.element(W, 'sdtPr')!.element(W, 'showingPlcHdr')).toBeUndefined();
    expect(t2.text()).not.toContain('Click here');
    expect(newViolations(t2, [host(), piece()])).toEqual([]);

    for (const mode of ['before', 'after', 'replace'] as const) {
      const t = host();
      const second = t.body.paragraphs().find((p) => p.text === 'Second')!;
      insertDocument(t, { paragraph: second, mode }, piece());
      expect(t.text().includes('Second')).toBe(mode !== 'replace');
      expect(newViolations(t, [host(), piece()]), mode).toEqual([]);
    }

    const t3 = host();
    const r3 = insertDocument(t3, { story: t3.body, index: 0 }, piece());
    expect(t3.body.element.firstElement()).toBe(r3.blocks[0]);
  });

  it('inserting into a table cell keeps "a cell ends with w:p" and drops sections', () => {
    const t = host();
    const warnings: Warning[] = [];
    insertDocument(
      t,
      { placeholder: '{{CELL}}' },
      open('powertools/DB002-Sections-With-Headers.docx'),
      {
        onWarning: (w) => warnings.push(w),
      },
    );
    const tc = t.body.element.descendants(W, 'tc')[0]!;
    expect(tc.lastElement()?.is(W, 'p')).toBe(true);
    expect(tc.descendants(W, 'sectPr')).toEqual([]);
    expect(warnings.map((w) => w.code)).toContain('SECTIONS_DROPPED');
    // A table as the last block of a cell gets a paragraph after it.
    const t2 = host();
    insertDocument(t2, { placeholder: '{{CELL}}' }, piece());
    const tc2 = t2.body.element.descendants(W, 'tc')[0]!;
    expect(tc2.lastElement()?.is(W, 'p')).toBe(true);
    expect(newViolations(t2, [host(), piece()])).toEqual([]);
  });

  it('keepSourceSections false adopts the host section and warns about dropped headers', () => {
    const t = host();
    const warnings: Warning[] = [];
    const r = insertDocument(
      t,
      { placeholder: '[[INSERT]]' },
      open('docxcompose/header_footer_sections.docx'),
      {
        keepSourceSections: false,
        onWarning: (w) => warnings.push(w),
      },
    );
    expect(r.sections).toBe(0);
    expect(t.sections().length).toBe(1);
    expect(warnings.map((w) => w.code)).toContain('SECTIONS_DROPPED');
  });

  it('keeps the source sections in the middle of the body by default', () => {
    const t = host();
    const src = open('docxcompose/header_footer_sections.docx');
    const n = src.sections().length;
    const r = insertDocument(t, { placeholder: '[[INSERT]]' }, src);
    expect(r.sections).toBe(n);
    // The host section is split around the inserted sections.
    expect(t.sections().length).toBe(2 + n);
    expect(newViolations(t, [host(), open('docxcompose/header_footer_sections.docx')])).toEqual([]);
  });

  it('inserts into headers and footers (PowerTools DB009: content controls and insert markers)', () => {
    const PT = 'http://powertools.codeplex.com/documentbuilder/2011/insert';
    const problems: string[] = [];
    for (const dest of HF_DEST) {
      for (const src of HF_SRC) {
        const names = [
          `powertools/DB__HeadersFooters__Dest__${dest}.docx`,
          `powertools/DB__HeadersFooters__Src__${src}.docx`,
        ];
        const t = open(names[0]!);
        const s = open(names[1]!);
        if (dest.includes('content_control')) {
          const has = t.headersFooters().some((h) =>
            h.part
              .xml()
              .root.descendants(W, 'tag')
              .some((g) => g.attrNS(W, 'val') === 'Templafy'),
          );
          if (!has) continue; // Fax (content control) has no insert point (PowerTools does not use it)
          insertDocument(t, { contentControl: { tag: 'Templafy' } }, s);
        } else {
          // Replace PowerTools' private <pt:Insert Id="Templafy"/> marker (not OOXML) with the source.
          const hf = t
            .headersFooters()
            .find((h) => h.part.xml().root.descendants(PT, 'Insert').length > 0);
          if (hf === undefined) {
            problems.push(`${dest}: no insert marker`);
            continue;
          }
          const marker = hf.part.xml().root.descendants(PT, 'Insert')[0]!;
          const p = marker.closest(W, 'p');
          if (p !== undefined) {
            marker.remove();
            const para = hf.story.paragraphs().find((x) => x.element === p)!;
            insertDocument(t, { paragraph: para, mode: 'after' }, s);
          } else {
            const index = hf.story.blocks().findIndex((b) => (b as XElement) === marker);
            insertDocument(t, { story: hf.story, index }, s);
            marker.remove();
          }
        }
        for (const v of newViolations(t, names.map(open))) problems.push(`${dest}+${src}: ${v}`);
      }
    }
    expect(problems).toEqual([]);
  });

  it('fills table cells and a text box of a template (PowerTools DB010)', () => {
    const t = open('powertools/DB010-Template.docx');
    const cells = t.body.element.descendants(W, 'tc');
    const cellPara = (i: number) =>
      t.body.paragraphs().find((p) => p.element === cells[i]!.elements(W, 'p')[0])!;
    insertDocument(
      t,
      { paragraph: cellPara(0), mode: 'replace' },
      open('powertools/DB010-Insert-01.docx'),
    );
    insertDocument(
      t,
      { paragraph: cellPara(2), mode: 'replace' },
      open('powertools/DB010-Insert-02.docx'),
    );
    const box = t.mainPart.xml().root.descendants(W, 'txbxContent')[0];
    if (box !== undefined) {
      const p = box.elements(W, 'p')[0]!;
      const para = t.body.paragraphs().find((x) => x.element === p) ?? {
        element: p,
        story: t.body,
      };
      insertDocument(
        t,
        { paragraph: para as never, mode: 'replace' },
        open('powertools/DB010-FrontMatter.docx'),
      );
    }
    for (const tc of t.body.element.descendants(W, 'tc')) {
      expect(tc.descendants(W, 'sectPr')).toEqual([]);
    }
    expect(
      newViolations(
        t,
        ['DB010-Template', 'DB010-Insert-01', 'DB010-Insert-02', 'DB010-FrontMatter'].map((n) =>
          open(`powertools/${n}.docx`),
        ),
      ),
    ).toEqual([]);
  });

  it('rejects unknown targets', () => {
    expect(() => insertDocument(host(), { placeholder: 'nope' }, piece())).toThrow(/not found/);
    expect(() => insertDocument(host(), { bookmark: 'nope' }, piece())).toThrow(/not found/);
    expect(() => insertDocument(host(), { contentControl: {} }, piece())).toThrow(
      /tag or an alias/,
    );
    expect(() => insertDocument(host(), { story: host().body, index: -1 }, piece())).toThrow(
      /index/,
    );
  });
});

// ---- composeDocuments / importBlocks / options ----------------------------------------------

describe('composeDocuments, importBlocks and options', () => {
  it('composeDocuments leaves its inputs untouched', () => {
    const docs = [
      'docxcompose/master.docx',
      'docxcompose/header_footer.docx',
      'docxcompose/footnote.docx',
    ].map(open);
    const before = docs.map((d) => writeDocx(d));
    const out = composeDocuments(docs);
    expect(out).not.toBe(docs[0]);
    for (const [i, d] of docs.entries()) expect(writeDocx(d)).toEqual(before[i]);
    expect(() => composeDocuments([])).toThrow(/no documents/);
  });

  it('importBlocks brings selected paragraphs with their dependencies (PowerTools DB006)', () => {
    const src = open('powertools/DB006-Source1.docx');
    const target = createDocx();
    const picked = src.body.blocks().slice(5, 15);
    const r = importBlocks(target, src, picked);
    expect(r.blocks.length).toBe(picked.length);
    expect(r.sections).toBe(0);
    expect(target.sections().length).toBe(1);
    expect(src.package.dirty).toBe(false);
    expect(newViolations(target, [createDocx(), open('powertools/DB006-Source1.docx')])).toEqual(
      [],
    );
    // Importing from the target itself works on a fork of it.
    const again = importBlocks(target, target, target.body.blocks().slice(0, 2));
    expect(again.blocks.length).toBe(2);
  });

  it('appends a document to itself', () => {
    const doc = open('docxcompose/header_footer_sections.docx');
    const n = doc.sections().length;
    appendDocument(doc, doc);
    expect(doc.sections().length).toBe(2 * n);
    const inputs = [
      open('docxcompose/header_footer_sections.docx'),
      open('docxcompose/header_footer_sections.docx'),
    ];
    expect(newViolations(doc, inputs)).toEqual([]);
  });

  it('validate: true throws on new errors only', () => {
    const t = open('docxcompose/master.docx');
    expect(() =>
      appendDocument(t, open('docxcompose/header_footer.docx'), { validate: true }),
    ).not.toThrow();
    // A source with a dangling relationship id produces a new error in the target.
    const bad = docWithBody(
      '<w:p><w:hyperlink r:id="rId999"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>',
    );
    const warnings: Warning[] = [];
    expect(() =>
      appendDocument(createDocx(), bad, { validate: true, onWarning: (w) => warnings.push(w) }),
    ).not.toThrow();
    expect(warnings.map((w) => w.code)).toContain('COMPOSE_DANGLING_REL');
  });

  it('accepts or rejects source revisions on a fork', () => {
    for (const revisions of ['accept', 'reject'] as const) {
      const src = open('apache-poi/delins.docx');
      const t = createDocx();
      appendDocument(t, src, { revisions });
      expect(src.package.dirty).toBe(false);
      const root = t.mainPart.xml().root;
      expect(root.descendants(W, 'ins').length + root.descendants(W, 'del').length).toBe(0);
    }
  });
});

// ---- LibreOffice ------------------------------------------------------------------------------

function findSoffice(): string | undefined {
  if (process.env.SOFFICE) return process.env.SOFFICE;
  for (const p of [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    join(process.env.HOME ?? '', 'Applications/LibreOffice.app/Contents/MacOS/soffice'),
    '/usr/bin/soffice',
    '/usr/lib/libreoffice/program/soffice',
  ]) {
    if (existsSync(p)) return p;
  }
  const w = spawnSync(process.platform === 'win32' ? 'where' : 'which', ['soffice'], {
    encoding: 'utf8',
  });
  const found = w.status === 0 ? w.stdout.split(/\r?\n/)[0]?.trim() : undefined;
  return found === '' ? undefined : found;
}

function pdfPages(pdf: string): number {
  const text = readFileSync(pdf).toString('latin1');
  return (text.match(/\/Type\s*\/Page(?![a-zA-Z])/g) ?? []).length;
}

const SOFFICE = process.env.TURBOWORD_SKIP_LIBREOFFICE ? undefined : findSoffice();

describe.skipIf(SOFFICE === undefined)('LibreOffice', () => {
  it('converts composed results; page count >= 95% of the inputs', { timeout: 300_000 }, () => {
    const work = mkdtempSync(join(tmpdir(), 'turboword-compose-'));
    try {
      const scenarios = [
        ['docxcompose/master_header_footer.docx', 'docxcompose/header_footer_sections.docx'],
        ['powertools/DB002-Sections-With-Headers.docx', 'powertools/DB002-Landscape-Section.docx'],
        [
          'powertools/DB007-Spec.docx',
          'powertools/DB007-WhitePaper.docx',
          'powertools/DB007-AuthorBiography.docx',
        ],
        ['docxcompose/master.docx', 'docxcompose/smart_art.docx'],
        ['docxcompose/master.docx', 'docxcompose/embedded_excel_chart.docx'],
        [
          'powertools/DB__HeadersFooters__Dest__Letter.docx',
          'powertools/DB__HeadersFooters__Src__Letterhead-with-Watermark.docx',
        ],
      ];
      const files: string[] = [];
      const inputs = new Set<string>();
      scenarios.forEach((names, i) => {
        const out = join(work, `composed-${i}.docx`);
        writeFileSync(out, writeDocx(composeNames(names).result));
        files.push(out);
        for (const n of names) inputs.add(n);
      });
      for (const n of inputs) {
        const p = join(work, `in-${n.replace(/[/\\]/g, '__')}`);
        writeFileSync(p, bytesOf(n));
        files.push(p);
      }
      const pdfDir = join(work, 'pdf');
      mkdirSync(pdfDir);
      const r = spawnSync(
        SOFFICE!,
        [
          `-env:UserInstallation=${pathToFileURL(join(work, 'profile')).href}`,
          '--headless',
          '--norestore',
          '--convert-to',
          'pdf',
          '--outdir',
          pdfDir,
          ...files,
        ],
        { encoding: 'utf8', timeout: 280_000 },
      );
      expect(r.error).toBeUndefined();
      const pages = (f: string): number => {
        const pdf = join(pdfDir, `${basename(f, '.docx')}.pdf`);
        expect(existsSync(pdf), `${basename(f)} converted`).toBe(true);
        return pdfPages(pdf);
      };
      scenarios.forEach((names, i) => {
        const got = pages(join(work, `composed-${i}.docx`));
        const sum = names.reduce(
          (n, name) => n + pages(join(work, `in-${name.replace(/[/\\]/g, '__')}`)),
          0,
        );
        expect(got, names.join(' + ')).toBeGreaterThanOrEqual(Math.floor(0.95 * sum));
      });
    } finally {
      rmSync(work, { recursive: true, force: true });
    }
  });
});
