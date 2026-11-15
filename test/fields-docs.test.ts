// src/fields on real documents: the docx-mailmerge (turbomerge) fixtures in
// test/fixtures/fields/, the corpus files with fields, and the table of contents
// (open-xml-sdk/TOC2.docx plus synthetic TOCs). Every updated document must keep its field
// codes, validate without new errors, and is written to out/fields/ for
// `node scripts/libreoffice-check.mjs out/fields`.

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterAll, describe, expect, it } from 'vitest';
import {
  type Field,
  listFields,
  type PageInfoProvider,
  updateFields,
  updateToc,
} from '../src/fields/index.js';
import {
  NS,
  Paragraph,
  readDocx,
  validatePackage,
  type WordDocument,
  writeDocx,
  type XElement,
} from '../src/index.js';
import { listLabel } from '../src/resolve/index.js';
import { serializeXml } from '../src/xml/serialize.js';
import { cf, docWith, heading, para, ppara, results } from './fields-helpers.js';
import { fixture } from './helpers/fixtures.js';

const OUT = fileURLToPath(new URL('../out/fields/', import.meta.url));
const outputs = new Map<string, Uint8Array>();

afterAll(() => {
  mkdirSync(OUT, { recursive: true });
  for (const [name, bytes] of outputs) writeFileSync(join(OUT, name), bytes);
});

const errors = (doc: WordDocument): string[] =>
  validatePackage(doc)
    .filter((i) => i.severity === 'error')
    .map((i) => `${i.code}: ${i.message}`);

/** Saves, re-reads, checks no new validation errors, and records the output. */
function roundTrip(name: string, before: string[], doc: WordDocument): WordDocument {
  const bytes = writeDocx(doc);
  const again = readDocx(bytes);
  expect(errors(again), name).toEqual(before);
  outputs.set(name, bytes);
  return again;
}

/** Field codes, in order. */
const codes = (fields: readonly Field[]): string[] => fields.map((f) => f.code ?? '');

/**
 * Codes of the fields an update keeps: all but those living in another field's result (an
 * updated parent replaces its result, nested fields included).
 */
const keptCodes = (fields: readonly Field[]): string[] =>
  codes(
    fields.filter((f) => {
      for (let c: Field = f; c.parent !== undefined; c = c.parent) {
        if (!c.parent.nested.includes(c)) return false;
      }
      return true;
    }),
  );

function load(name: string): { doc: WordDocument; before: string[] } {
  const doc = readDocx(fixture(name).bytes);
  return { doc, before: errors(doc) };
}

describe('docx-mailmerge fixtures', () => {
  it('merges a field whose instruction spans paragraphs (\\b and \\f with paragraph marks)', () => {
    const { doc, before } = load('fields/test_field_with_paragraph.docx');
    const old = codes(listFields(doc));
    expect(updateFields(doc, { values: { fieldname: 'X' } }).updated).toBe(1);
    const again = roundTrip('field_with_paragraph.docx', before, doc);
    const [f] = listFields(again);
    expect(codes([f!])).toEqual(old);
    expect(f!.result).toBe('before\nparX after \npar');
    // \b/\f paragraph marks become line breaks, inside the paragraph holding the result.
    expect(
      f!.separate!.parent?.nodeType === 1 &&
        (f!.separate!.parent as XElement).descendants(NS.w, 'br'),
    ).toHaveLength(2);
  });

  it('merges a nested MERGEFIELD inside \\b', () => {
    const { doc, before } = load('fields/test_field_with_paragraph_and_nested_field.docx');
    const old = codes(listFields(doc));
    updateFields(doc, { values: { fieldname: 'X' } });
    const again = roundTrip('field_with_paragraph_and_nested_field.docx', before, doc);
    expect(codes(listFields(again))).toEqual(old);
    expect(results(again)).toEqual(['before\npar X X after \npar', 'X']);
  });

  it('evaluates IF with nested fields and a multi-paragraph result', () => {
    for (const [value, want] of [
      ['one', 'one is \none'],
      ['two', 'two is not \none'],
    ] as const) {
      const { doc, before } = load('fields/test_if_with_paragraph.docx');
      const old = keptCodes(listFields(doc));
      updateFields(doc, { values: { fieldname: value } });
      const again = roundTrip(`if_with_paragraph_${value}.docx`, before, doc);
      const fields = listFields(again);
      expect(keptCodes(fields)).toEqual(old);
      expect(fields[1]!.result).toBe(want);
      const last = again.body.paragraphs().at(-1)!;
      expect(last.text).toBe('one');
    }
  });

  it('evaluates nested IFs inside and outside', () => {
    for (const [value, want] of [
      ['one', '- one -'],
      ['two', 'two'],
      ['three', 'more: three'],
    ] as const) {
      const { doc, before } = load('fields/test_nested_if_inside.docx');
      const old = keptCodes(listFields(doc));
      updateFields(doc, { values: { fieldname: value } });
      const again = roundTrip(`nested_if_inside_${value}.docx`, before, doc);
      const fields = listFields(again);
      // The outer IF and the fields of its instruction keep their codes (fields that lived in
      // the old results are replaced along with them).
      expect(keptCodes(fields)).toEqual(old);
      expect(fields[0]!.result).toBe(want);
      expect(again.text()).toBe(want);
    }
    const { doc, before } = load('fields/test_nested_if_outside.docx');
    updateFields(doc, { values: { fieldname: 'v' } });
    const again = roundTrip('nested_if_outside.docx', before, doc);
    expect(results(again)).toEqual(['true', 'v']);
  });

  it('merges simple and complex MERGEFIELDs, quoted names, tables, headers and footers', () => {
    const cases: [string, Record<string, string>, (doc: WordDocument) => void][] = [
      [
        'fields/test_spaces.docx',
        { Singleword: 'one', 'Hello world': 'two', 'More than one space': 'three' },
        (d) => expect(results(d)).toEqual(['one', 'two', 'three']),
      ],
      [
        'fields/test_keep_fields.docx',
        { first: '1', second: '2', three_simple: '3' },
        (d) => expect(results(d)).toEqual(['1', '2', '3', '«four_simple»']),
      ],
      [
        'fields/test_winword2010.docx',
        { Titel: 'Dhr.', Voornaam: 'Jan', Achternaam: 'Jansen', Plaats: 'Utrecht' },
        (d) => expect(d.text()).toContain('Jan Jansen'),
      ],
      [
        'fields/test_macword2011.docx',
        { first_name: 'Ada', last_name: 'Lovelace', city: 'London', date: '2024-03-09' },
        (d) => {
          expect(d.text()).toContain('Ada Lovelace');
          expect(results(d)).toContain('London');
        },
      ],
      [
        'fields/test_merge_table_rows.docx',
        { student_name: 'Bob', class_code: 'C1', thesis_grade: 'A' },
        (d) => {
          expect(results(d).slice(0, 1)).toEqual(['Bob']);
          expect(results(d).at(-1)).toBe('A');
        },
      ],
      [
        'fields/test_skipif.docx',
        { fieldname: 'one' },
        (d) => expect(results(d)).toEqual(['one', '«Next Record If»', 'one', 'one']),
      ],
      [
        'docx-mailmerge/test_footnote_header_footer.docx',
        {
          fieldname: 'F',
          headerfield: 'H',
          footerfield: 'FT',
          headerfirst: 'HF',
          footerfirst: 'FF',
          headereven: 'HE',
          footereven: 'FE',
        },
        (d) => {
          expect(results(d)).toEqual(['F', 'H', 'HE', 'FF', 'HF', 'FT', 'FE', 'F']);
          expect(listFields(d).map((f) => f.story.kind)).toEqual([
            'body',
            'header',
            'header',
            'footer',
            'header',
            'footer',
            'footer',
            'footnote',
          ]);
        },
      ],
    ];
    for (const [name, values, check] of cases) {
      const { doc, before } = load(name);
      const old = codes(listFields(doc));
      updateFields(doc, { values });
      const again = roundTrip(name.replace(/^.*\//, 'merged_'), before, doc);
      expect(codes(listFields(again)), name).toEqual(old);
      check(again);
    }
  });

  it('treats a DATE without a backslash before @ as having no picture', () => {
    const { doc, before } = load('fields/test_unbalanced_quote_date_field.docx');
    updateFields(doc, { now: new Date(Date.UTC(2024, 2, 9)), timeZone: 'UTC' });
    const again = roundTrip('unbalanced_quote_date.docx', before, doc);
    expect(results(again)).toEqual(['3/9/2024']);
  });
});

describe('corpus documents', () => {
  it('re-evaluates DOCPROPERTY fields in headers and footers to the same values', () => {
    const { doc, before } = load('docxcompose/docproperties_header_footer_3_sections.docx');
    const old = listFields(doc);
    updateFields(doc, { locale: 'de-DE', timeZone: 'Europe/Berlin' });
    const again = roundTrip('docproperties.docx', before, doc);
    const now = listFields(again);
    expect(codes(now)).toEqual(codes(old));
    expect(now.map((f) => f.result)).toEqual(['Y', '1,1', 'Foo Bar', '11.06.2019', '123']);
  });

  it('leaves HYPERLINK fields alone in a document with tracked changes', () => {
    const { doc, before } = load('apache-poi/delins.docx');
    const old = listFields(doc);
    const r = updateFields(doc);
    expect(r.updated).toBe(0);
    expect(r.skipped).toHaveLength(old.length);
    const again = roundTrip('delins.docx', before, doc);
    expect(results(again)).toEqual(old.map((f) => f.result));
  });

  it('numbers REF \\n from the list label of the bookmarked paragraph', () => {
    const { doc, before } = load('docxcompose/numberings.docx');
    const numbered = doc.body
      .paragraphs()
      .filter((p) => p.element.element(NS.w, 'pPr')?.element(NS.w, 'numPr') !== undefined);
    const target = numbered[3]!;
    const label = listLabel(target)!;
    expect(label).toBeTruthy();
    const pPr = target.element.element(NS.w, 'pPr')!;
    const x = doc.mainPart.xml();
    pPr.after(x.createElement(NS.w, 'bookmarkStart', { 'w:id': '900', 'w:name': '_Ref1' }));
    target.element.append(x.createElement(NS.w, 'bookmarkEnd', { 'w:id': '900' }));
    const p = doc.body.appendParagraph('See ');
    const src = docWith(
      para(cf('REF _Ref1 \\w \\h', '?')),
      para(cf('REF _Ref1 \\n', '?')),
      para(cf('REF _Ref1 \\p', '?')),
    );
    for (const q of src.body.paragraphs())
      for (const k of q.element.children.slice()) p.element.append(k);
    updateFields(doc);
    roundTrip('ref_numbers.docx', before, doc);
    expect(listFields(doc).map((f) => f.result)).toEqual([
      label,
      label.replace(/[.\s]+$/, ''),
      'above',
    ]);
  });

  it('numbers NOTEREF from the footnote reference inside the bookmark', () => {
    const { doc, before } = load('open-xml-sdk/TableFntEdnt.docx');
    const refs = doc.body.element.descendants(NS.w, 'footnoteReference');
    expect(refs.length).toBeGreaterThanOrEqual(2);
    const run = refs[1]!.parent as XElement;
    const x = doc.mainPart.xml();
    run.before(x.createElement(NS.w, 'bookmarkStart', { 'w:id': '901', 'w:name': '_Ref2' }));
    run.after(x.createElement(NS.w, 'bookmarkEnd', { 'w:id': '901' }));
    const p = doc.body.appendParagraph('Note ');
    for (const k of docWith(para(cf('NOTEREF _Ref2 \\h', '?')))
      .body.paragraphs()[0]!
      .element.children.slice()) {
      p.element.append(k);
    }
    updateFields(doc);
    roundTrip('noteref.docx', before, doc);
    expect(results(doc)).toEqual(['2']);
  });
});

/** Pages: paragraph i of the body is on page 1 + floor(i / 3). */
function fakePages(doc: WordDocument): PageInfoProvider {
  return {
    pageOf(node: XElement) {
      const p = node.is(NS.w, 'p') ? node : node.closest(NS.w, 'p');
      const i = doc.body.paragraphs().findIndex((q) => q.element === p);
      return i < 0 ? undefined : 1 + Math.floor(i / 3);
    },
    pageCount: () => 9,
  };
}

describe('table of contents', () => {
  it('rebuilds TOC2.docx keeping its code, styles, bookmarks and page numbers', () => {
    const { doc, before } = load('open-xml-sdk/TOC2.docx');
    const old = listFields(doc);
    const bookmarksBefore = doc.body.element.descendants(NS.w, 'bookmarkStart').length;
    expect(updateToc(doc)).toBe(1);
    const again = roundTrip('toc2.docx', before, doc);
    const now = listFields(again);
    expect(codes(now)).toEqual(codes(old));
    expect(now[0]!.result).toBe(old[0]!.result);
    expect(again.body.element.descendants(NS.w, 'bookmarkStart')).toHaveLength(bookmarksBefore);
    const entries = again.body.paragraphs().slice(0, 6);
    expect(entries.map((p) => p.style)).toEqual(['10', '20', '30', '10', '20', '30']);
    expect(
      entries.map((p) => p.element.descendants(NS.w, 'hyperlink')[0]?.attrNS(NS.w, 'anchor')),
    ).toEqual(old.slice(1).map((f) => f.args[0]));
    // The end marker keeps its own paragraph, as Word wrote it.
    expect(again.body.paragraphs()[6]!.text).toBe('');
    // Idempotent: a second rebuild gives the same XML.
    const xml1 = serializeXml(again.mainPart.xml());
    updateToc(again);
    expect(serializeXml(again.mainPart.xml())).toBe(xml1);
  });

  it('takes page numbers from a PageInfoProvider (also through updateFields)', () => {
    const { doc } = load('open-xml-sdk/TOC2.docx');
    const r = updateFields(doc, { pages: fakePages(doc) });
    expect(r.updated).toBe(1);
    const pagerefs = listFields(doc, { type: 'PAGEREF' });
    expect(pagerefs.map((f) => f.result)).toEqual(['3', '4', '4', '4', '5', '5']);
  });

  const tocDoc = (instr: string, ...extra: ReturnType<typeof para>[]): WordDocument =>
    docWith(
      para(cf(instr, '')),
      heading(1, 'Introduction'),
      para('Body text.'),
      heading(2, 'Background'),
      heading(3, 'Details'),
      ppara('<w:outlineLvl w:val="1"/>', 'Direct outline'),
      heading(1, 'Results'),
      ...extra,
    );

  it('builds a TOC from scratch: styles, bookmarks, hyperlinks, PAGEREFs, tab leaders', () => {
    const doc = tocDoc('TOC \\o "1-3" \\h \\z \\u');
    const before = errors(doc);
    expect(updateToc(doc, { pages: fakePages(doc) })).toBe(1);
    const again = roundTrip('toc_new.docx', before, doc);
    const paras = again.body.paragraphs();
    expect(paras.slice(0, 5).map((p) => p.text)).toEqual([
      'Introduction\t1',
      'Background\t2',
      'Details\t2',
      'Direct outline\t2',
      'Results\t3',
    ]);
    expect(paras.slice(0, 5).map((p) => p.style)).toEqual(['TOC1', 'TOC2', 'TOC3', 'TOC2', 'TOC1']);
    const styles = again.mainPart.relatedByType('styles')[0]!.xml().root;
    const toc2 = styles.elements(NS.w, 'style').find((s) => s.attrNS(NS.w, 'styleId') === 'TOC2')!;
    expect(toc2.element(NS.w, 'name')?.attrNS(NS.w, 'val')).toBe('toc 2');
    expect(toc2.descendants(NS.w, 'ind')[0]?.attrNS(NS.w, 'left')).toBe('220');
    const tab = paras[0]!.element.descendants(NS.w, 'tab')[0]!;
    expect([tab.attrNS(NS.w, 'val'), tab.attrNS(NS.w, 'leader'), tab.attrNS(NS.w, 'pos')]).toEqual([
      'right',
      'dot',
      '9350',
    ]);
    // _Toc bookmarks on the headings, unique ids, hyperlinks and PAGEREFs pointing at them.
    const starts = again.body.element.descendants(NS.w, 'bookmarkStart');
    const names = starts.map((b) => b.attrNS(NS.w, 'name')!);
    expect(names).toHaveLength(5);
    expect(names.every((n) => /^_Toc\d{9}$/.test(n))).toBe(true);
    expect(new Set(starts.map((b) => b.attrNS(NS.w, 'id'))).size).toBe(5);
    expect(starts[0]!.parent).toBe(paras[5]!.element);
    const anchors = again.body.element
      .descendants(NS.w, 'hyperlink')
      .map((h) => h.attrNS(NS.w, 'anchor'));
    expect(anchors).toEqual(names);
    expect(listFields(again, { type: 'PAGEREF' }).map((f) => f.args[0])).toEqual(names);
    expect(paras[0]!.element.descendants(NS.w, 'webHidden').length).toBeGreaterThan(0);
    // Without page information the next rebuild keeps the numbers it had.
    updateToc(again);
    expect(again.body.paragraphs()[4]!.text).toBe('Results\t3');
    // Deterministic: the same input gives the same bytes.
    const twin = tocDoc('TOC \\o "1-3" \\h \\z \\u');
    updateToc(twin, { pages: fakePages(twin) });
    expect(serializeXml(twin.mainPart.xml())).toBe(serializeXml(doc.mainPart.xml()));
  });

  it('honours \\o ranges, \\n, \\p, \\t, \\b, levels and empty TOCs', () => {
    const texts = (
      instr: string,
      o: Parameters<typeof updateToc>[1] = {},
      ...extra: ReturnType<typeof para>[]
    ): string[] => {
      const doc = tocDoc(instr, ...extra);
      updateToc(doc, { pages: fakePages(doc), ...o });
      return listFields(doc)[0]!.result.split('\n');
    };
    expect(texts('TOC \\o "1-1"')).toEqual(['Introduction\t1', 'Results\t3']);
    expect(texts('TOC \\o "2-3" \\n "3-3"')).toEqual(['Background\t2', 'Details']);
    expect(texts('TOC \\o "1-1" \\p "-"')).toEqual(['Introduction-1', 'Results-3']);
    expect(texts('TOC \\o "1-3"', { levels: [1, 2] })).toEqual([
      'Introduction\t1',
      'Background\t2',
      'Results\t3',
    ]);
    expect(texts('TOC \\o "1-9"')).toEqual([
      'Introduction\t1',
      'Background\t2',
      'Details\t2',
      'Results\t3',
    ]);
    expect(texts('TOC \\t "Title,1"', {}, ppara('<w:pStyle w:val="Title"/>', 'A Title'))).toEqual([
      'A Title\t3',
    ]);
    expect(texts('TOC \\o "1-3" \\n')).toEqual([
      'Introduction',
      'Background',
      'Details',
      'Results',
    ]);
    expect(texts('TOC \\o "4-5"')).toEqual(['No table of contents entries found.']);
  });

  it('collects TC fields with \\f and \\l, captions with \\c, and headings inside \\b', () => {
    const doc = docWith(
      para(cf('TOC \\f \\l "1-2"', '')),
      para(
        'Here ',
        cf('TC "Entry one" \\l 1', ''),
        cf('TC "Deep" \\l 3', ''),
        cf('TC "Other list" \\f X', ''),
      ),
      para(cf('TC "Entry two" \\l 2 \\n', '')),
      para(cf('TOC \\c "Figure"', '')),
      para('Figure ', cf('SEQ Figure', '1'), ': A cat'),
      para('Figure ', cf('SEQ Figure', '2'), ': A dog'),
      para(cf('TOC \\o "1-3" \\b part2', '')),
      heading(1, 'Outside'),
      { xml: '<w:bookmarkStart w:id="50" w:name="part2"/>' },
      heading(1, 'Inside'),
      { xml: '<w:bookmarkEnd w:id="50"/>' },
    );
    const before = errors(doc);
    expect(updateToc(doc)).toBe(3);
    const again = roundTrip('toc_tc_captions.docx', before, doc);
    const tocs = listFields(again, { type: 'TOC' });
    expect(tocs.map((f) => f.result.split('\n'))).toEqual([
      ['Entry one\t', 'Entry two'],
      ['Figure 1: A cat\t', 'Figure 2: A dog\t'],
      ['Inside\t'],
    ]);
  });

  it('creates TOC styles with unique ids', () => {
    const doc = tocDoc('TOC \\o "1-2"');
    const styles = doc.mainPart.relatedByType('styles')[0]!.xml();
    // A style already using the id TOC1 under another name.
    styles.root.append(
      styles.createElement(NS.w, 'style', { 'w:type': 'paragraph', 'w:styleId': 'TOC1' }),
    );
    styles.root
      .lastElement()!
      .append(styles.createElement(NS.w, 'name', { 'w:val': 'Something else' }));
    updateToc(doc);
    expect(
      doc.body
        .paragraphs()
        .slice(0, 3)
        .map((p) => p.style),
    ).toEqual(['TOC1_1', 'TOC2', 'TOC1_1']);
    expect(new Paragraph(doc.body.paragraphs()[3]!.element, doc.body).text).toBe('Introduction');
  });
});
