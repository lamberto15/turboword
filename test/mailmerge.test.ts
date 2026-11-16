// turboword/mailmerge: mergeRecord, mergeEach, mergeAll.
//
// The record-control expectations (NEXT, NEXTIF, SKIPIF, IF) are those of the docx-mailmerge test
// suite (test/fixtures/fields, MIT), the Python implementation the fixtures come from; every
// merged document is also written, re-read and validated.

import { describe, expect, it } from 'vitest';
import { listFields } from '../src/fields/index.js';
import {
  readDocx,
  TurbowordError,
  UsageError,
  validatePackage,
  type WordDocument,
  writeDocx,
} from '../src/index.js';
import { mergeAll, mergeEach, mergeRecord } from '../src/mailmerge/index.js';
import { NS } from '../src/xml/ns.js';
import { cf, docWith, para, table } from './fields-helpers.js';
import { fixture } from './helpers/fixtures.js';

const open = (name: string): WordDocument => readDocx(fixture(name).bytes);
const values = ['one', 'two', 'three', 'four', 'five'];

/** Writes and re-reads `doc`, failing on any validation error. */
function roundTrip(doc: WordDocument): WordDocument {
  const back = readDocx(writeDocx(doc));
  const errors = validatePackage(back).filter((i) => i.severity === 'error');
  expect(errors.map((e) => `${e.code}: ${e.message}`)).toEqual([]);
  return back;
}

/** Non-empty paragraph texts of the body. */
const lines = (doc: WordDocument): string[] =>
  doc
    .text()
    .split('\n')
    .filter((t) => t !== '');

const settingsXml = (doc: WordDocument): string => {
  const part = doc.mainPart.relatedByType('settings')[0];
  return part === undefined ? '' : new TextDecoder().decode(part.bytes());
};

describe('mergeRecord', () => {
  it('merges simple and complex fields into plain text and drops the data source', () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    const out = roundTrip(mergeRecord(tpl, { fieldname: 'hello' })!);
    expect(lines(out)).toEqual(['Merge : hello']);
    expect(listFields(out)).toEqual([]);
    expect(settingsXml(out)).not.toContain('mailMerge');
  });

  it('removes the mail-merge setup and its data-source relationship', () => {
    const tpl = open('fields/test_field_with_paragraph.docx');
    expect(settingsXml(tpl)).toContain('w:mailMerge');
    const out = roundTrip(mergeRecord(tpl, { fieldname: 'X' })!);
    expect(settingsXml(out)).not.toContain('mailMerge');
    const settings = out.mainPart.relatedByType('settings')[0]!;
    expect(settings.rels.all().filter((r) => /mailMerge/i.test(r.type))).toEqual([]);
    // \b and \f text, with its paragraph marks, around the value
    expect(out.text()).toContain('X');
  });

  it('never modifies the template', () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    const before = writeDocx(tpl);
    mergeRecord(tpl, { fieldname: 'a' });
    mergeRecord(tpl, { fieldname: 'b' });
    expect(writeDocx(tpl)).toEqual(before);
    expect(listFields(tpl, { type: 'MERGEFIELD' })).toHaveLength(1);
  });

  it('merges headers, footers and footnotes', () => {
    const tpl = open('docx-mailmerge/test_footnote_header_footer.docx');
    const names = [...new Set(listFields(tpl, { type: 'MERGEFIELD' }).map((f) => f.args[0]!))];
    expect(names.length).toBeGreaterThan(5);
    const out = roundTrip(mergeRecord(tpl, Object.fromEntries(names.map((n) => [n, `<${n}>`])))!);
    expect(listFields(out, { type: 'MERGEFIELD' })).toEqual([]);
    const text = out
      .stories()
      .map((s) => s.text())
      .join('|');
    for (const n of names) expect(text).toContain(`<${n}>`);
  });

  it('keeps other fields live (PAGE) while unlinking merge fields', () => {
    const tpl = open('docx-mailmerge/test_footnote_header_footer.docx');
    const out = mergeRecord(tpl, {})!;
    const types = new Set(listFields(out).map((f) => f.type));
    expect(types.has('MERGEFIELD')).toBe(false);
    expect([...types].every((t) => !['IF', 'NEXT', 'MERGEFIELD'].includes(t))).toBe(true);
  });

  it("fields: 'keep' leaves live fields showing the values", () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    const out = roundTrip(mergeRecord(tpl, { fieldname: 'kept' }, { fields: 'keep' })!);
    const f = listFields(out, { type: 'MERGEFIELD' });
    expect(f).toHaveLength(1);
    expect(f[0]!.result).toBe('kept');
    expect(settingsXml(out)).not.toContain('mailMerge');
  });

  it('formats values with the field switches', () => {
    const tpl = docWith(
      para(cf('MERGEFIELD Total \\# "#,##0.00"', '«Total»')),
      para(cf('MERGEFIELD Due \\@ "d MMMM yyyy"', '«Due»')),
      para(cf('MERGEFIELD Name \\* Upper \\b "Dear "', '«Name»')),
      para(cf('MERGEFIELD customer.city', '«city»')),
    );
    const out = roundTrip(
      mergeRecord(tpl, {
        Total: 1234.5,
        Due: new Date(Date.UTC(2026, 11, 1, 12)),
        Name: 'ada',
        customer: { city: 'London' },
      })!,
    );
    expect(lines(out)).toEqual(['1,234.50', '1 December 2026', 'Dear ADA', 'London']);
  });

  it('missing values: empty by default, kept, or an error', () => {
    const tpl = docWith(para('A', cf('MERGEFIELD Absent \\b "-"', '«Absent»'), 'B'));
    expect(lines(mergeRecord(tpl, {})!)).toEqual(['AB']);
    const kept = mergeRecord(tpl, {}, { missing: 'keep' })!;
    expect(lines(kept)).toEqual(['A«Absent»B']);
    expect(listFields(kept, { type: 'MERGEFIELD' })).toHaveLength(1);
    const err = (() => {
      try {
        mergeRecord(tpl, {}, { missing: 'error' });
      } catch (e) {
        return e;
      }
      return undefined;
    })();
    expect(err).toBeInstanceOf(TurbowordError);
    expect((err as TurbowordError).code).toBe('MERGE_MISSING_FIELD');
    expect((err as Error).message).toContain('"Absent"');
  });

  it('SKIPIF that holds gives no document', () => {
    const tpl = open('fields/test_skipif.docx');
    expect(mergeRecord(tpl, { fieldname: 'one' })).toBeUndefined();
    expect(lines(mergeRecord(tpl, { fieldname: 'two' })!)).toEqual(['two', 'two']);
  });

  it('NEXT with a single record leaves the following fields empty', () => {
    const out = mergeRecord(open('fields/test_next_record.docx'), { field: 'x' })!;
    expect(lines(out)).toEqual(['x/x', '/', '/', '/']);
  });

  it('evaluates IF on the merged values', () => {
    const tpl = docWith(
      para(cf(['IF ', cf('MERGEFIELD Qty', '', { nested: true }), ' > 1 "items" "item"'], '')),
    );
    expect(lines(mergeRecord(tpl, { Qty: 3 })!)).toEqual(['items']);
    expect(lines(mergeRecord(tpl, { Qty: 1 })!)).toEqual(['item']);
    expect(listFields(mergeRecord(tpl, { Qty: 1 })!)).toEqual([]);
  });
});

describe('table rows', () => {
  const rows = [
    { class_code: 'ECON101', class_name: 'Economics 101', class_grade: 'A' },
    { class_code: 'ECONADV', class_name: 'Economics Advanced', class_grade: 'B' },
    { class_code: 'OPRES', class_name: 'Operations Research', class_grade: 'A' },
  ];

  it('repeats the row holding the array key, once per item', () => {
    const tpl = open('fields/test_merge_table_rows.docx');
    const out = roundTrip(
      mergeRecord(tpl, { student_name: 'Bouke Haarsma', thesis_grade: 'A', class_code: rows })!,
    );
    const text = lines(out).join('|');
    expect(text).toContain(
      'ECON101|Economics 101|A|ECONADV|Economics Advanced|B|OPRES|Operations Research|A',
    );
    expect(text).toContain('Bouke Haarsma');
    expect(listFields(out)).toEqual([]);
  });

  it('rows read the item over the record', () => {
    const tpl = docWith(
      table([
        ['Item', 'Owner'],
        [[para(cf('MERGEFIELD sku', '«sku»'))], [para(cf('MERGEFIELD owner', '«owner»'))]],
      ]),
    );
    const out = roundTrip(
      mergeRecord(tpl, { owner: 'Ada', sku: [{ sku: 'A-1' }, { sku: 'B-2', owner: 'Bob' }] })!,
    );
    expect(lines(out)).toEqual(['Item', 'Owner', 'A-1', 'Ada', 'B-2', 'Bob']);
  });

  it('an empty array removes the row', () => {
    const tpl = open('fields/test_merge_table_rows.docx');
    const out = roundTrip(mergeRecord(tpl, { class_code: [] })!);
    expect(out.text()).not.toContain('class_');
    expect(out.body.element.descendants(NS.w, 'tbl').length).toBeGreaterThan(0);
  });

  it('several tables anchored in one document', () => {
    const tpl = open('fields/test_merge_multiple_table_rows.docx');
    const out = roundTrip(mergeRecord(tpl, { student_name: 'B', class_code: rows })!);
    const count = out.text().split('ECONADV').length - 1;
    expect(count).toBeGreaterThanOrEqual(1);
    expect(listFields(out)).toEqual([]);
  });

  it('rows of a table nested in a cell', () => {
    const tpl = open('fields/test_merge_nested_table_rows.docx');
    const data = {
      name: [
        { name: 'turboword', version: '1', note: 'n1', desc: 'd1' },
        { name: 'turbomerge', version: '2', note: 'n2', desc: 'd2' },
      ],
    };
    const out = roundTrip(mergeRecord(tpl, data)!);
    expect(lines(out).join('|')).toContain('turboword|1|n1|d1|turbomerge|2|n2|d2');
    expect(listFields(out)).toEqual([]);
  });

  it('arrays inside items expand the rows of nested tables', () => {
    const inner = table([[[para(cf('MERGEFIELD line', '«line»'))]]]);
    const tpl = docWith(
      table([[[para(cf('MERGEFIELD group', '«group»')), inner, para('')]], [[para('end')]]]),
    );
    const out = roundTrip(
      mergeRecord(tpl, {
        group: [
          { group: 'G1', line: ['a', 'b'] },
          { group: 'G2', line: [{ line: 'c' }] },
          { group: 'G3', line: [] },
        ],
      })!,
    );
    expect(lines(out)).toEqual(['G1', 'a', 'b', 'G2', 'c', 'G3', 'end']);
    expect(listFields(out)).toEqual([]);
  });

  it('drawing ids stay unique in repeated rows and records', () => {
    const tpl = open('fields/test_input_duplicate_id.docx');
    const out = roundTrip(mergeAll(tpl, [{ field: 'a' }, { field: 'b' }, { field: 'c' }]));
    const ids = out.body.element.descendants(NS.wp, 'docPr').map((d) => d.attr('id'));
    expect(ids.length).toBeGreaterThanOrEqual(3);
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe('record control (NEXT, NEXTIF, SKIPIF, MERGEREC, MERGESEQ)', () => {
  it('NEXT moves to the next record within one document (docx-mailmerge test_next_record)', () => {
    const tpl = open('fields/test_next_record.docx');
    const docs = [
      ...mergeEach(
        tpl,
        values.map((field) => ({ field })),
      ),
    ];
    expect(docs.map((d) => d.records)).toEqual([[0, 1, 2, 3], [4]]);
    const out = mergeAll(
      tpl,
      values.map((field) => ({ field })),
    );
    expect(lines(roundTrip(out))).toEqual([...values.map((v) => `${v}/${v}`), '/', '/', '/']);
  });

  it('NEXTIF moves on when its condition holds (test_next_if_record)', () => {
    const out = roundTrip(
      mergeAll(
        open('fields/test_nextif.docx'),
        values.map((fieldname) => ({ fieldname })),
      ),
    );
    expect(lines(out)).toEqual(['one', 'two', 'three', 'three', 'four', 'four', 'five', 'five']);
    expect(listFields(out)).toEqual([]);
  });

  it('SKIPIF drops the record (test_skipif)', () => {
    const tpl = open('fields/test_skipif.docx');
    const docs = [
      ...mergeEach(
        tpl,
        values.map((fieldname) => ({ fieldname })),
      ),
    ];
    expect(docs.map((d) => d.records)).toEqual([[1], [2], [3], [4]]);
    expect(docs.map((d) => d.sequence)).toEqual([1, 2, 3, 4]);
    const out = roundTrip(
      mergeAll(
        tpl,
        values.map((fieldname) => ({ fieldname })),
      ),
    );
    expect(lines(out)).toEqual(['two', 'two', 'three', 'three', 'four', 'four', 'five', 'five']);
  });

  it('IF with paragraphs in its branches (test_if_record)', () => {
    const out = roundTrip(
      mergeAll(
        open('fields/test_if_with_paragraph.docx'),
        values.map((fieldname) => ({ fieldname })),
      ),
    );
    expect(lines(out)).toEqual(
      values.flatMap((v) => [v, `${v}${v === 'one' ? ' is ' : ' is not '}`, 'one']),
    );
  });

  it('MERGEREC counts records, MERGESEQ merged documents', () => {
    const tpl = docWith(
      para(
        cf('SKIPIF ', ''),
        cf('MERGEREC', ''),
        '/',
        cf('MERGESEQ', ''),
        ' ',
        cf('MERGEFIELD n', ''),
      ),
    );
    // SKIPIF without a condition does nothing; skip "b" with a real one
    const skip = docWith(
      para(
        cf(['SKIPIF "', cf('MERGEFIELD n', '', { nested: true }), '" = "b"'], ''),
        cf('MERGEREC', ''),
        '/',
        cf('MERGESEQ', ''),
        ' ',
        cf('MERGEFIELD n', ''),
      ),
    );
    const recs = [{ n: 'a' }, { n: 'b' }, { n: 'c' }];
    expect([...mergeEach(tpl, recs)].map((d) => lines(d.doc)[0])).toEqual([
      '1/1 a',
      '2/2 b',
      '3/3 c',
    ]);
    expect([...mergeEach(skip, recs)].map((d) => lines(d.doc)[0])).toEqual(['1/1 a', '3/2 c']);
  });

  it('record-control fields are left alone outside a merge', () => {
    const tpl = open('fields/test_next_record.docx');
    expect(listFields(tpl, { type: 'NEXT' })).toHaveLength(3);
  });
});

describe('mergeEach', () => {
  it('reads records lazily', () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    let read = 0;
    function* source() {
      for (let i = 0; i < 1000; i++) {
        read++;
        yield { fieldname: `r${i}` };
      }
    }
    const it = mergeEach(tpl, source());
    const first = it.next();
    expect(first.done).toBe(false);
    expect(lines(first.value!.doc)).toEqual(['Merge : r0']);
    expect(read).toBeLessThanOrEqual(2);
  });

  it('reports the fields a record lacked', () => {
    const tpl = open('fields/test_merge_table_rows.docx');
    const [m] = [...mergeEach(tpl, [{ student_name: 'x' }])];
    expect(m!.missing).toContain('thesis_grade');
    expect(m!.missing).not.toContain('student_name');
  });

  it('rejects bad input', () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    expect(() => [...mergeEach(tpl, [null as never])]).toThrow(UsageError);
    expect(() => [...mergeEach(tpl, [['a'] as never])]).toThrow(UsageError);
    expect(() => [...mergeEach(tpl, [{}], { fields: 'nope' as never })]).toThrow(UsageError);
    expect(() => mergeAll(tpl, [])).toThrow(UsageError);
    expect(() => mergeAll(tpl, [{}], { separator: 'nope' as never })).toThrow(UsageError);
  });
});

describe('mergeAll', () => {
  it('one section per record, each with its own headers and footers', () => {
    const tpl = open('docx-mailmerge/test_footnote_header_footer.docx');
    const names = [...new Set(listFields(tpl, { type: 'MERGEFIELD' }).map((f) => f.args[0]!))];
    const records = [0, 1, 2].map((i) => Object.fromEntries(names.map((n) => [n, `${n}-${i}`])));
    const out = roundTrip(mergeAll(tpl, records));
    expect(out.sections()).toHaveLength(3 * tpl.sections().length);
    const headers = out.headersFooters().map((h) => h.story.text());
    for (let i = 0; i < 3; i++) {
      expect(headers.some((t) => t.includes(`headerfield-${i}`))).toBe(true);
      expect(headers.some((t) => t.includes(`footerfirst-${i}`))).toBe(true);
    }
    const notes = out.footnotes().map((n) => n.text());
    for (let i = 0; i < 3; i++) expect(notes.some((t) => t.includes(`fieldname-${i}`))).toBe(true);
    expect(listFields(out, { type: 'MERGEFIELD' })).toEqual([]);
  });

  it("separator 'pageBreak': one section, a page break between records", () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    const out = roundTrip(
      mergeAll(tpl, [{ fieldname: 'a' }, { fieldname: 'b' }, { fieldname: 'c' }], {
        separator: 'pageBreak',
      }),
    );
    expect(out.sections()).toHaveLength(1);
    const breaks = out.body.element
      .descendants(NS.w, 'br')
      .filter((b) => b.attrNS(NS.w, 'type') === 'page');
    expect(breaks).toHaveLength(2);
    expect(lines(out)).toEqual(['Merge : a', 'Merge : b', 'Merge : c']);
  });

  it("separator 'none' runs records on; a section type forces the break", () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    const recs = [{ fieldname: 'a' }, { fieldname: 'b' }];
    const none = roundTrip(mergeAll(tpl, recs, { separator: 'none' }));
    expect(none.sections()).toHaveLength(1);
    expect(lines(none)).toEqual(['Merge : a', 'Merge : b']);
    const odd = roundTrip(mergeAll(tpl, recs, { separator: 'oddPage' }));
    expect(odd.sections().map((s) => s.type)).toEqual([tpl.sections()[0]!.type, 'oddPage']);
  });

  it("separator 'section' starts every record on a new page, even after a continuous template", () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    expect(tpl.sections()[0]!.type).toBe('continuous');
    const recs = [{ fieldname: 'a' }, { fieldname: 'b' }, { fieldname: 'c' }];
    expect(
      mergeAll(tpl, recs)
        .sections()
        .map((s) => s.type),
    ).toEqual(['continuous', 'nextPage', 'nextPage']);
    tpl.sections()[0]!.element!.element(NS.w, 'type')!.setAttrNS(NS.w, 'val', 'oddPage');
    expect(
      mergeAll(tpl, recs)
        .sections()
        .map((s) => s.type),
    ).toEqual(['oddPage', 'oddPage', 'oddPage']);
  });

  it('a catalog (directory) template runs records on by default', () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    const settings = tpl.mainPart.relatedByType('settings')[0]!.xml().root;
    const mm = settings.element(NS.w, 'mailMerge');
    if (mm === undefined) return; // fixture without a mail-merge setup: nothing to check
    mm.element(NS.w, 'mainDocumentType')?.setAttrNS(NS.w, 'val', 'catalog');
    const out = mergeAll(tpl, [{ fieldname: 'a' }, { fieldname: 'b' }]);
    expect(out.sections()).toHaveLength(1);
  });

  it('ids stay unique across many records with headers, footers and notes', () => {
    const tpl = open('docx-mailmerge/test_footnote_header_footer.docx');
    const out = roundTrip(
      mergeAll(
        tpl,
        Array.from({ length: 40 }, (_, i) => ({ fieldname: i })),
      ),
    );
    const paraIds: string[] = [];
    const docPr: string[] = [];
    for (const story of out.stories()) {
      for (const el of [story.element, ...story.element.descendants()]) {
        const id = el.attrNS(NS.w14, 'paraId');
        if (id !== undefined) paraIds.push(id);
        if (el.is(NS.wp, 'docPr')) docPr.push(el.attr('id') ?? '');
      }
    }
    expect(paraIds.length).toBeGreaterThan(40);
    expect(new Set(paraIds).size).toBe(paraIds.length);
    expect(new Set(docPr).size).toBe(docPr.length);
    expect(out.headersFooters().length).toBe(40 * tpl.headersFooters().length);
  });

  it('merges many records', () => {
    const tpl = open('fields/test_merge_templates_simple.docx');
    const n = 200;
    const out = roundTrip(
      mergeAll(
        tpl,
        Array.from({ length: n }, (_, i) => ({ fieldname: `r${i}` })),
      ),
    );
    expect(lines(out)).toHaveLength(n);
    expect(lines(out)[n - 1]).toBe(`Merge : r${n - 1}`);
  });
});
