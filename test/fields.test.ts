// src/fields: the field index (listFields), the result writer and every evaluator of
// updateFields, on synthetic documents (test/fields-helpers.ts). Real files and the TOC are in
// fields-docs.test.ts; switch formatting in fields-format.test.ts.

import { describe, expect, it } from 'vitest';
import {
  type FieldEvaluator,
  listFields,
  type PageInfoProvider,
  setUpdateFieldsOnOpen,
  updateFields,
} from '../src/fields/index.js';
import { NS, readDocx, writeDocx, type XElement } from '../src/index.js';
import { serializeXml } from '../src/xml/serialize.js';
import {
  bookmark,
  cf,
  docWith,
  heading,
  para,
  ppara,
  results,
  sf,
  table,
} from './fields-helpers.js';

const xmlOf = (doc: ReturnType<typeof docWith>): string => serializeXml(doc.mainPart.xml());

describe('listFields', () => {
  it('finds complex fields spanning runs, with type, args and switches', () => {
    const doc = docWith(
      para(
        'Dear ',
        cf([' MERGE', 'FIELD "First Name" ', '\\b "Mr. \\"X\\" " \\* Upper '], 'JOHN'),
        '!',
      ),
    );
    const [f] = listFields(doc);
    expect(f).toMatchObject({
      kind: 'complex',
      type: 'MERGEFIELD',
      instruction: ' MERGEFIELD "First Name" \\b "Mr. \\"X\\" " \\* Upper ',
      args: ['First Name'],
      switches: [
        { name: '\\b', arg: 'Mr. "X" ' },
        { name: '\\*', arg: 'Upper' },
      ],
      result: 'JOHN',
      locked: false,
      dirty: false,
    });
    expect(f!.start.is(NS.w, 'r')).toBe(true);
    expect(f!.separate?.element(NS.w, 'fldChar')?.attrNS(NS.w, 'fldCharType')).toBe('separate');
    expect(f!.end.element(NS.w, 'fldChar')?.attrNS(NS.w, 'fldCharType')).toBe('end');
    expect(f!.story).toBe(doc.body);
  });

  it('finds simple fields, fldLock and dirty', () => {
    const doc = docWith(
      para(sf(' DATE \\@ "yyyy" ', '2020', { lock: true }), cf('PAGE', '3', { dirty: true })),
    );
    const [a, b] = listFields(doc);
    expect(a).toMatchObject({ kind: 'simple', type: 'DATE', result: '2020', locked: true });
    expect(a!.start).toBe(a!.end);
    expect(a!.start.is(NS.w, 'fldSimple')).toBe(true);
    expect(b).toMatchObject({ kind: 'complex', type: 'PAGE', dirty: true, locked: false });
  });

  it('follows fields across paragraphs and around tables', () => {
    const doc = docWith(para(cf(['IF 1 = 1 "first'], '')), { xml: '' }, para('x'));
    expect(listFields(doc)).toHaveLength(1);
    // begin in one paragraph, instruction continued in the next, result holding a table.
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
    const it1 = '<w:r><w:instrText>QUOTE "a</w:instrText></w:r>';
    const it2 =
      '<w:r><w:instrText>b"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>old</w:t></w:r>';
    const end = '<w:r><w:t>tail</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const doc2 = docWith(
      { xml: `<w:p>${begin}${it1}</w:p><w:p>${it2}</w:p>` },
      table([['1', '2']]),
      { xml: `<w:p>${end}</w:p>` },
    );
    const [f] = listFields(doc2);
    expect(f!.instruction).toBe('QUOTE "a\nb"');
    expect(f!.args).toEqual(['a\nb']);
    expect(f!.result).toBe('old\n1\n2\ntail');
  });

  it('nests fields in instructions and results', () => {
    const doc = docWith(
      para(
        cf(
          [
            'IF ',
            cf('MERGEFIELD a', 'x', { nested: true }),
            ' = "x" "',
            sf('MERGEFIELD b', 'B', { nested: true }),
            '" "no"',
          ],
          'B',
        ),
        cf('QUOTE "q"', ''),
      ),
    );
    const fields = listFields(doc);
    expect(fields.map((f) => f.type)).toEqual(['IF', 'MERGEFIELD', 'MERGEFIELD', 'QUOTE']);
    const [outer, a, b] = fields;
    expect(outer!.instruction).toBe('IF x = "x" "B" "no"');
    expect(outer!.code).toBe('IF {MERGEFIELD a} = "x" "{MERGEFIELD b}" "no"');
    expect(outer!.nested).toEqual([a, b]);
    expect(a!.parent).toBe(outer);
    expect(fields[3]!.parent).toBeUndefined();
    expect(listFields(doc, { type: 'mergefield' })).toHaveLength(2);
    expect(listFields(doc, (f) => f.type === 'QUOTE')).toHaveLength(1);
    expect(listFields(doc.body)).toHaveLength(4);
  });

  it('ignores deleted fields and unbalanced markers, and keeps text boxes separate', () => {
    const doc = docWith(
      {
        xml:
          '<w:p><w:del w:id="1" w:author="a"><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:delInstrText>PAGE</w:delInstrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>',
      },
      para(cf('QUOTE "ok"', 'ok')),
      {
        xml: '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r></w:p>',
      },
    );
    expect(listFields(doc).map((f) => f.type)).toEqual(['QUOTE']);
  });

  it('lists the fields of every story of a document', () => {
    const doc = readDocx(writeDocx(docWith(para(cf('PAGE', '1')))));
    expect(listFields(doc)).toHaveLength(1);
  });
});

describe('result writing', () => {
  it('keeps field codes and replaces only the result (MERGEFORMAT keeps the first result run)', () => {
    const doc = docWith(
      para(cf('MERGEFIELD name \\* MERGEFORMAT', '«name»', { rPr: '<w:b/>', codeRPr: '<w:i/>' })),
    );
    const before = listFields(doc)[0]!.instruction;
    const r = updateFields(doc, { values: { name: 'Ada' } });
    expect(r).toEqual({ updated: 1, skipped: [] });
    const [f] = listFields(doc);
    expect(f!.instruction).toBe(before);
    expect(f!.result).toBe('Ada');
    const resultRun = f!.separate!.nextElement()!;
    expect(resultRun.element(NS.w, 'rPr')?.element(NS.w, 'b')).toBeDefined();
    expect(resultRun.element(NS.w, 'rPr')?.element(NS.w, 'i')).toBeUndefined();
  });

  it('applies CHARFORMAT from the run holding the field type', () => {
    const doc = docWith(
      para(cf('QUOTE "x" \\* CHARFORMAT', 'old', { rPr: '<w:b/>', codeRPr: '<w:i/>' })),
    );
    updateFields(doc);
    const resultRun = listFields(doc)[0]!.separate!.nextElement()!;
    expect(resultRun.element(NS.w, 'rPr')?.element(NS.w, 'i')).toBeDefined();
    expect(resultRun.element(NS.w, 'rPr')?.element(NS.w, 'b')).toBeUndefined();
  });

  it('adds a result section to a field without one, and clears w:dirty', () => {
    const doc = docWith(para(cf('QUOTE "new"', '', { noSeparate: true, dirty: true })));
    updateFields(doc);
    const [f] = listFields(doc);
    expect(f!.result).toBe('new');
    expect(f!.dirty).toBe(false);
    expect(f!.separate).toBeDefined();
  });

  it('collapses a multi-paragraph result and writes multi-paragraph results', () => {
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
    const doc = docWith(
      {
        xml:
          `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>A </w:t></w:r>${begin}` +
          '<w:r><w:instrText>QUOTE "one"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
          '<w:r><w:t>x</w:t></w:r></w:p><w:p><w:r><w:t>y</w:t></w:r></w:p>',
      },
      table([['in result']]),
      {
        xml: '<w:p><w:r><w:t>z</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t> B</w:t></w:r></w:p>',
      },
    );
    updateFields(doc);
    expect(doc.body.paragraphs().map((p) => p.text)).toEqual(['A one B']);
    expect(doc.body.tables()).toHaveLength(0);
    // Now a result with paragraph marks (IF branches spanning paragraphs).
    const jc = '<w:pPr><w:jc w:val="center"/></w:pPr>';
    const doc2 = docWith({
      xml:
        `<w:p>${jc}<w:r><w:t xml:space="preserve">A </w:t></w:r>${begin}<w:r><w:instrText>IF 1 = 1 "p1</w:instrText></w:r></w:p>` +
        `<w:p>${jc}<w:r><w:instrText>p2</w:instrText></w:r></w:p>` +
        `<w:p>${jc}<w:r><w:instrText>p3" "no"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        '<w:r><w:t>old</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> B</w:t></w:r></w:p>',
    });
    updateFields(doc2);
    const paras = doc2.body.paragraphs();
    expect(paras.map((p) => p.text)).toEqual(['A ', '', 'p1', 'p2', 'p3 B']);
    for (const p of paras)
      expect(p.element.element(NS.w, 'pPr')?.element(NS.w, 'jc')).toBeDefined();
    expect(listFields(doc2)[0]!.result).toBe('p1\np2\np3');
  });

  it('keeps bookmark and comment markers whose partner is outside the result', () => {
    const doc = docWith({
      xml:
        '<w:p><w:bookmarkStart w:id="7" w:name="outer"/><w:r><w:t>a</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>QUOTE "new"</w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>o</w:t></w:r>' +
        '<w:bookmarkEnd w:id="7"/><w:bookmarkStart w:id="8" w:name="inner"/><w:r><w:t>l</w:t></w:r>' +
        '<w:bookmarkEnd w:id="8"/><w:commentRangeStart w:id="3"/><w:r><w:t>d</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>z</w:t></w:r><w:commentRangeEnd w:id="3"/></w:p>',
    });
    updateFields(doc);
    const x = xmlOf(doc);
    expect(x).toContain('<w:bookmarkStart w:id="7" w:name="outer"/>');
    expect(x).toContain('<w:bookmarkEnd w:id="7"/>');
    expect(x).not.toContain('inner');
    expect(x).toContain('<w:commentRangeStart w:id="3"/>');
    expect(listFields(doc)[0]!.result).toBe('new');
    expect(doc.text()).toBe('anewz');
  });

  it('writes nested results as instrText and handles tabs and line breaks', () => {
    const doc = docWith(
      para(cf(['QUOTE "', cf('QUOTE "a\tb\nc"', 'old', { nested: true }), '"'], 'old')),
    );
    updateFields(doc);
    const [outer, inner] = listFields(doc);
    expect(inner!.result).toBe('a\tb\nc');
    const innerRun = inner!.separate!.nextElement()!;
    expect(innerRun.elements().map((e) => e.local)).toEqual([
      'instrText',
      'tab',
      'instrText',
      'br',
      'instrText',
    ]);
    expect(outer!.result).toBe('a\tb\nc');
    expect(outer!.separate!.nextElement()!.element(NS.w, 't')).toBeDefined();
  });

  it('splits runs that hold several field markers', () => {
    const doc = docWith({
      xml:
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/><w:instrText>QUOTE "n"</w:instrText>' +
        '<w:fldChar w:fldCharType="separate"/><w:t>o</w:t><w:fldChar w:fldCharType="end"/><w:t>!</w:t></w:r></w:p>',
    });
    updateFields(doc);
    expect(doc.text()).toBe('n!');
    expect(listFields(doc)[0]!.instruction).toBe('QUOTE "n"');
  });

  it('skips locked fields and everything inside them', () => {
    const doc = docWith(
      para(
        cf(['QUOTE "', cf('MERGEFIELD a', 'old', { nested: true }), '"'], 'old', { lock: true }),
      ),
      para(sf('MERGEFIELD a', 'old', { lock: true })),
    );
    const r = updateFields(doc, { values: { a: 'new' } });
    expect(r.updated).toBe(0);
    expect(r.skipped.map((f) => f.locked)).toEqual([true, true]);
    expect(results(doc)).toEqual(['old', 'old', 'old']);
  });

  it('sets w:updateFields in settings, creating the settings part', () => {
    const doc = docWith(para(cf('PAGE', '1')));
    updateFields(doc, { markForWord: true });
    const settings = doc.mainPart.relatedByType('settings')[0]!.xml().root;
    const el = settings.element(NS.w, 'updateFields');
    expect(el?.attrNS(NS.w, 'val')).toBe('true');
    // Schema order: before decimalSymbol, after characterSpacingControl.
    const names = settings.elements().map((e) => e.local);
    expect(names.indexOf('updateFields')).toBeGreaterThan(names.indexOf('characterSpacingControl'));
    expect(names.indexOf('updateFields')).toBeLessThan(names.indexOf('decimalSymbol'));
    setUpdateFieldsOnOpen(doc, false);
    expect(settings.element(NS.w, 'updateFields')).toBeUndefined();

    const bare = docWith(para('x'));
    const rel = bare.mainPart.rels.byType('settings')[0]!;
    bare.mainPart.rels.remove(rel.id);
    bare.package.removePart('word/settings.xml');
    setUpdateFieldsOnOpen(bare);
    const created = bare.mainPart.relatedByType('settings')[0]!;
    expect(created.name).toBe('word/settings.xml');
    expect(created.xml().root.element(NS.w, 'updateFields')).toBeDefined();
    const again = readDocx(writeDocx(bare));
    expect(
      again.mainPart.relatedByType('settings')[0]!.xml().root.element(NS.w, 'updateFields'),
    ).toBeDefined();
  });
});

describe('MERGEFIELD, IF, COMPARE', () => {
  it('merges values with \\b, \\f and formats', () => {
    const doc = docWith(
      para(cf('MERGEFIELD first \\b "Dear " \\f ","', '')),
      para(cf('MERGEFIELD missing', 'keep')),
      para(cf('MERGEFIELD empty \\b "x" \\f "y"', 'z')),
      para(cf('MERGEFIELD amount \\# "$#,##0.00"', '')),
      para(cf('MERGEFIELD when \\@ "d MMMM yyyy"', '')),
      para(cf('MERGEFIELD person.name \\* Upper', '')),
      para(cf('MERGEFIELD "Two Words" \\* Lower', '')),
      para(cf('MERGEFIELD CaseLess', '')),
    );
    const r = updateFields(doc, {
      timeZone: 'UTC',
      values: {
        first: 'Ada',
        empty: '',
        amount: 1234.5,
        when: new Date(Date.UTC(2024, 0, 2)),
        person: { name: 'grace' },
        'Two Words': 'ABC',
        caseless: 'ci',
      },
    });
    expect(results(doc)).toEqual([
      'Dear Ada,',
      'keep',
      '',
      '$1,234.50',
      '2 January 2024',
      'GRACE',
      'abc',
      'ci',
    ]);
    expect(r.updated).toBe(7);
    expect(r.skipped.map((f) => f.args[0])).toEqual(['missing']);
  });

  it('evaluates IF comparisons, wildcards and branches', () => {
    const cases: [string, string][] = [
      ['IF 1 < 2 "yes" "no"', 'yes'],
      ['IF 10 > 9 "yes" "no"', 'yes'],
      ['IF "10" > "9" "yes" "no"', 'yes'],
      ['IF "b" > "a" "yes" "no"', 'yes'],
      ['IF abc = "abc" "yes" "no"', 'yes'],
      ['IF abc = "ABC" "yes" "no"', 'no'],
      ['IF "Seattle" = "S*e" "yes" "no"', 'yes'],
      ['IF "Seattle" = "S?attle" "yes" "no"', 'yes'],
      ['IF "Seattle" <> "Port*" "yes" "no"', 'yes'],
      ['IF 3 <= 3 "yes" "no"', 'yes'],
      ['IF 3 >= 4 "yes" "no"', 'no'],
      ['IF 1.0 = 1 "yes" "no"', 'yes'],
      ['IF 2 * 3 = 6 "yes" "no"', 'yes'],
      ['IF 1 = 2 "yes"', ''],
      ['IF 1 = 1 "Dear Sir" "Dear Madam" \\* Upper', 'DEAR SIR'],
      ['IF 1 "on" "off"', 'on'],
      ['COMPARE 5 > 3', '1'],
      ['COMPARE "a" = "b"', '0'],
    ];
    const doc = docWith(...cases.map(([instr]) => para(cf(instr, '?'))));
    updateFields(doc);
    expect(results(doc)).toEqual(cases.map(([, want]) => want));
  });

  it('evaluates nested IF and MERGEFIELD, inner fields first', () => {
    const inner = (v: string): ReturnType<typeof cf> =>
      cf('MERGEFIELD gender', v, { nested: true });
    const doc = docWith(
      para(
        cf(
          [
            'IF ',
            inner('?'),
            ' = "f" "Dear Ms. ',
            cf('MERGEFIELD last', '?', { nested: true }),
            '" "',
            cf(['IF ', inner('?'), ' = "m" "Dear Mr. " "Dear "'], '?', { nested: true }),
            cf('MERGEFIELD last', '?', { nested: true }),
            '"',
          ],
          '?',
        ),
      ),
    );
    const merge = (values: Record<string, string>): string => {
      const d = readDocx(writeDocx(doc));
      updateFields(d, { values });
      return listFields(d)[0]!.result;
    };
    expect(merge({ gender: 'f', last: 'Lovelace' })).toBe('Dear Ms. Lovelace');
    expect(merge({ gender: 'm', last: 'Babbage' })).toBe('Dear Mr. Babbage');
    expect(merge({ gender: 'x', last: 'Hopper' })).toBe('Dear Hopper');
    // Without data the IF is left alone (and so are the fields it holds).
    const d = readDocx(writeDocx(doc));
    const r = updateFields(d, { values: {} });
    expect(r.updated).toBe(0);
    expect(listFields(d)[0]!.result).toBe('?');
  });
});

describe('= formulas', () => {
  const formula = (expr: string, more = ''): ReturnType<typeof cf> => cf(`= ${expr} ${more}`, '?');

  it('computes operators, functions and defaults to two decimals', () => {
    const cases: [string, string][] = [
      ['1 + 2 * 3', '7'],
      ['(1 + 2) * 3', '9'],
      ['2 ^ 3 ^ 2', '64'],
      ['-2 ^ 2', '4'],
      ['10 / 4', '2.5'],
      ['10 / 3', '3.33'],
      ['50%', '0.5'],
      ['200 * 15%', '30'],
      ['3 > 2', '1'],
      ['3 <> 3', '0'],
      ['2 <= 1', '0'],
      ['SUM(1, 2, 3)', '6'],
      ['AVERAGE(1, 2, 3, 4)', '2.5'],
      ['MIN(4, -1, 3)', '-1'],
      ['MAX(4, -1, 3)', '4'],
      ['ABS(-5)', '5'],
      ['INT(-4.7)', '-4'],
      ['ROUND(2.345, 2)', '2.35'],
      ['ROUND(1234, -2)', '1200'],
      ['MOD(7, 3)', '1'],
      ['MOD(-7, 3)', '2'],
      ['AND(1, 0)', '0'],
      ['OR(1, 0)', '1'],
      ['NOT(0)', '1'],
      ['IF(2 > 1, 10, 20)', '10'],
      ['PRODUCT(2, 3, 4)', '24'],
      ['COUNT(5, 6, 7)', '3'],
      ['SIGN(-3)', '-1'],
      ['TRUE + TRUE()', '2'],
      ['FALSE', '0'],
      ['DEFINED(nope)', '0'],
      ['DEFINED(price)', '1'],
      ['price * 2', '25'],
      ['$1,234.50 + 1', '1235.5'],
      ['1/3 \\# "0.0000"', '0.3333'],
      ['1234.5 \\# "#,##0.00"', '1,234.50'],
      ['21 \\* CardText', 'twenty-one'],
    ];
    const doc = docWith(
      para(bookmark('price', 1, '12.5')),
      ...cases.map(([e]) => para(formula(e))),
    );
    updateFields(doc);
    expect(results(doc)).toEqual(cases.map(([, want]) => want));
  });

  it('leaves the field on errors', () => {
    const doc = docWith(
      para(formula('1 / 0')),
      para(formula('1 +')),
      para(formula('unknown + 1')),
      para(formula('FOO(1)')),
    );
    const r = updateFields(doc);
    expect(r.updated).toBe(0);
    expect(results(doc)).toEqual(['?', '?', '?', '?']);
  });

  it('reads table cells: ABOVE, LEFT, RIGHT, BELOW, A1 refs, ranges, R1C1', () => {
    const cell = (expr: string): ReturnType<typeof para>[] => [para(formula(expr))];
    const doc = docWith(
      table([
        ['Item', 'Q1', 'Q2', 'Total'],
        ['a', '1', '2', cell('SUM(LEFT)')],
        ['b', '$3.50', '4', cell('SUM(LEFT)')],
        ['c', 'n/a', '6', cell('SUM(B2:C4)')],
        ['Sum', cell('SUM(ABOVE)'), cell('SUM(ABOVE)'), cell('B2 + R3C3')],
        [cell('SUM(RIGHT)'), '1', '2', '3'],
        [cell('AVERAGE(B2,C2)'), cell('MAX(BELOW)'), '', ''],
        ['', '9', '', ''],
      ]),
    );
    updateFields(doc);
    // SUM(ABOVE) stops at the first non-numeric cell ("n/a" for Q1).
    expect(results(doc)).toEqual(['3', '7.5', '16.5', '0', '12', '5', '6', '1.5', '9']);
  });
});

describe('dates and document information', () => {
  const now = new Date(Date.UTC(2024, 2, 9, 17, 7, 8));

  it('updates DATE and TIME only with a clock', () => {
    const doc = docWith(
      para(cf('DATE', 'old')),
      para(cf('TIME', 'old')),
      para(cf('DATE \\@ "yyyy-MM-dd"', 'old')),
    );
    expect(updateFields(doc).updated).toBe(0);
    expect(results(doc)).toEqual(['old', 'old', 'old']);
    updateFields(doc, { now, timeZone: 'UTC' });
    expect(results(doc)).toEqual(['3/9/2024', '5:07 PM', '2024-03-09']);
    updateFields(doc, { now, timeZone: 'UTC', locale: 'de-DE' });
    expect(results(doc)).toEqual(['09.03.2024', '17:07', '2024-03-09']);
  });

  it('reads core, app and custom properties', () => {
    const doc = docWith(
      para(cf('CREATEDATE \\@ "d MMM yyyy"', '?')),
      para(cf('SAVEDATE', '?')),
      para(cf('PRINTDATE', '?')),
      para(cf('AUTHOR', '?')),
      para(cf('TITLE', '?')),
      para(cf('SUBJECT \\* Upper', '?')),
      para(cf('KEYWORDS', '?')),
      para(cf('COMMENTS', '?')),
      para(cf('LASTSAVEDBY', '?')),
      para(cf('TITLE "Given"', '?')),
      para(cf('DOCPROPERTY Title', '?')),
      para(cf('DOCPROPERTY Company', '?')),
      para(cf('DOCPROPERTY Pages', '?')),
      para(cf('DOCPROPERTY "Project Code"', '?')),
      para(cf('DOCPROPERTY approved', '?')),
      para(cf('DOCPROPERTY Budget \\# "#,##0"', '?')),
      para(cf('DOCPROPERTY NoSuchThing', '?')),
      para(cf('REVNUM', '?')),
      para(cf('FILENAME', 'report.docx')),
    );
    const props = doc.properties;
    props.created = new Date(Date.UTC(2023, 4, 6, 8, 0, 0));
    props.modified = new Date(Date.UTC(2024, 0, 2, 0, 0, 0));
    props.creator = 'Ada';
    props.title = 'Engines';
    props.subject = 'math';
    props.keywords = 'k1, k2';
    props.description = 'notes';
    props.lastModifiedBy = 'Grace';
    props.revision = '7';
    props.company = 'ACME';
    props.setCustom('Project Code', 'PX-1');
    props.setCustom('Approved', true);
    props.setCustom('Budget', 1250000);
    const r = updateFields(doc, { timeZone: 'UTC' });
    expect(results(doc)).toEqual([
      '6 May 2023',
      '1/2/2024 12:00:00 AM',
      '?',
      'Ada',
      'Engines',
      'MATH',
      'k1, k2',
      'notes',
      'Grace',
      'Given',
      'Engines',
      'ACME',
      '?',
      'PX-1',
      'Y',
      '1,250,000',
      '?',
      '7',
      'report.docx',
    ]);
    expect(r.skipped.map((f) => f.type)).toEqual([
      'PRINTDATE',
      'DOCPROPERTY',
      'DOCPROPERTY',
      'FILENAME',
    ]);
  });

  it('reads TEMPLATE and EDITTIME from the extended properties', () => {
    const doc = docWith(para(cf('TEMPLATE', '?')), para(cf('EDITTIME', '?')));
    const app = doc.package.part('docProps/app.xml')!.xml();
    app.root.append(app.createElement(NS.ep, 'Template', {}));
    app.root.lastElement()!.text = 'Normal.dotm';
    app.root.append(app.createElement(NS.ep, 'TotalTime', {}));
    app.root.lastElement()!.text = '42';
    updateFields(doc);
    expect(results(doc)).toEqual(['Normal.dotm', '42']);
  });

  it('reads document variables and counts words and characters', () => {
    const doc = docWith(
      para('Four words in here.'),
      para(cf('DOCVARIABLE client', '?')),
      para(cf('DOCVARIABLE missing', '?')),
      para(cf('NUMWORDS', '0')),
      para(cf('NUMCHARS', '0')),
    );
    const settings = doc.mainPart.relatedByType('settings')[0]!.xml();
    const vars = settings.createElement(NS.w, 'docVars');
    vars.append(settings.createElement(NS.w, 'docVar', { 'w:name': 'Client', 'w:val': 'Contoso' }));
    settings.root.append(vars);
    updateFields(doc);
    // Counted when first needed, after the fields before it were updated:
    // "Four words in here." "Contoso" "?" "0" "0" = 8 words, 26 characters without spaces.
    expect(results(doc)).toEqual(['Contoso', '?', '8', '26']);
  });
});

describe('references, numbering and pages', () => {
  const pages: PageInfoProvider = {
    pageOf: (node: XElement) => {
      const p = node.is(NS.w, 'p') ? node : node.closest(NS.w, 'p');
      const t = p?.text ?? '';
      return t.includes('second page') ? 2 : t.includes('nowhere') ? undefined : 1;
    },
    pageNumberText: (node: XElement) => {
      const p = node.is(NS.w, 'p') ? node : node.closest(NS.w, 'p');
      return p?.text.includes('roman') ? 'ii' : undefined;
    },
    pageCount: () => 5,
    sectionPageCount: () => 3,
  };

  it('evaluates REF, implicit REF, SET, ASK, FILLIN and QUOTE', () => {
    const doc = docWith(
      para(bookmark('Target', 1, 'Chapter One')),
      para(cf('REF Target', '?')),
      para(cf('REF target \\* Upper', '?')),
      para(cf('Target', '?')),
      para(cf('REF Target \\p', '?')),
      para(cf('REF nope', '?')),
      para(cf('SET Greeting "Hello there"', '')),
      para(cf('REF Greeting', '?')),
      para(cf('ASK Client "Client name?" \\d "Anonymous"', '')),
      para(cf('REF Client', '?')),
      para(cf('ASK Asked "Prompt?"', '')),
      para(cf('FILLIN "Your name?"', '?')),
      para(cf('FILLIN "Unknown?" \\d "fallback"', '?')),
      para(cf('QUOTE "quoted text" \\* Upper', '?')),
      para(cf('QUOTE 13 \\* ALPHABETIC', '?')),
      para(cf('REF Later \\p', '?')),
      para(bookmark('Later', 2, 'below me')),
    );
    const r = updateFields(doc, { values: { 'Your name?': 'Ada', Asked: 'yes' } });
    expect(results(doc)).toEqual([
      'Chapter One',
      'CHAPTER ONE',
      'Chapter One',
      'above',
      '?',
      '',
      'Hello there',
      '',
      'Anonymous',
      '',
      'Ada',
      'fallback',
      'QUOTED TEXT',
      'M',
      'below',
    ]);
    expect(r.skipped.map((f) => f.instruction)).toEqual(['REF nope']);
  });

  it('reads bookmarks as updated earlier in the same run', () => {
    const doc = docWith(
      para(bookmark('total', 1, cf('= 2 + 3', '0'))),
      para(cf('REF total', '?')),
      para(cf('= total * 2', '?')),
      para(cf('IF total > 4 "big" "small"', '?')),
    );
    updateFields(doc);
    expect(results(doc)).toEqual(['5', '5', '10', 'big']);
  });

  it('leaves REF \\n alone when the bookmarked paragraph has no number', () => {
    const doc = docWith(para(bookmark('plain', 1, 'Plain')), para(cf('REF plain \\n', '?')));
    expect(updateFields(doc).updated).toBe(0);
    expect(results(doc)).toEqual(['?']);
  });

  it('evaluates PAGEREF, PAGE, NUMPAGES and SECTIONPAGES with a page provider only', () => {
    const doc = docWith(
      para(bookmark('far', 1, 'on the second page')),
      para('roman page here ', cf('PAGE', '?')),
      para(cf('PAGE \\* ALPHABETIC', '?')),
      para(cf('NUMPAGES', '?')),
      para(cf('SECTIONPAGES', '?')),
      para(cf('PAGEREF far', '?')),
      para(cf('PAGEREF far \\p', '?')),
      para(cf('PAGEREF nope', '?')),
      para('nowhere ', cf('PAGE', '?')),
    );
    expect(updateFields(doc).updated).toBe(0);
    updateFields(doc, { pages });
    expect(results(doc)).toEqual(['ii', 'A', '5', '3', '2', 'on page 2', '?', '?']);
  });

  it('numbers SEQ fields with \\c, \\r, \\h, \\n and heading resets', () => {
    const doc = docWith(
      heading(1, 'One'),
      para('Figure ', cf('SEQ Figure \\s 1', '?')),
      para('Figure ', cf('SEQ Figure \\s 1', '?')),
      para('Again ', cf('SEQ Figure \\c', '?')),
      para('Table ', cf('SEQ Table \\* ROMAN', '?')),
      heading(2, 'One.One'),
      para('Figure ', cf('SEQ Figure \\s 1', '?')),
      heading(1, 'Two'),
      para('Figure ', cf('SEQ Figure \\s 1', '?')),
      para('Hidden ', cf('SEQ Figure \\h', '?')),
      para('Next ', cf('SEQ Figure \\n', '?')),
      para('Restart ', cf('SEQ Figure \\r 10', '?')),
      para('Table ', cf('SEQ Table', '?')),
    );
    updateFields(doc);
    expect(results(doc)).toEqual(['1', '2', '2', 'I', '3', '1', '', '3', '10', '2']);
  });

  it('evaluates STYLEREF in the body', () => {
    const doc = docWith(
      heading(1, 'Intro'),
      para(cf('STYLEREF "Heading 1"', '?')),
      heading(1, 'Methods'),
      para(cf('STYLEREF 1', '?')),
      para(cf('STYLEREF "heading 2"', '?')),
      heading(2, 'Later'),
      para(cf('STYLEREF Nonexistent', '?')),
    );
    updateFields(doc);
    expect(results(doc)).toEqual(['Intro', 'Methods', 'Later', '?']);
  });

  it('writes SYMBOL with its font and size', () => {
    const doc = docWith(
      para(cf('SYMBOL 183 \\f Symbol \\s 12', '?')),
      para(cf('SYMBOL 0x263A \\u', '?')),
    );
    updateFields(doc);
    const [a, b] = listFields(doc);
    expect(a!.result).toBe('');
    const rPr = a!.separate!.nextElement()!.element(NS.w, 'rPr')!;
    expect(rPr.element(NS.w, 'rFonts')?.attrNS(NS.w, 'ascii')).toBe('Symbol');
    expect(rPr.element(NS.w, 'sz')?.attrNS(NS.w, 'val')).toBe('24');
    expect(b!.result).toBe('☺');
  });

  it('numbers SECTION by the section breaks before the field', () => {
    const doc = docWith(
      para(cf('SECTION', '?')),
      ppara('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>', 'end of one'),
      para(cf('SECTION', '?')),
    );
    updateFields(doc);
    expect(results(doc)).toEqual(['1', '2']);
  });
});

describe('custom evaluators and options', () => {
  it('lets custom evaluators override built-ins and evaluate nested instructions', () => {
    const doc = docWith(
      para(cf('MERGEFIELD name', '?')),
      para(cf('MYFIELD a b \\x', '?')),
      para(cf('USERNAME', '?')),
    );
    const seen: string[] = [];
    const evaluators: Record<string, FieldEvaluator> = {
      MERGEFIELD: (f) => `custom:${f.args[0]}`,
      myfield: (f, ctx) => {
        seen.push(f.code ?? '');
        return `${f.args.join('+')}=${ctx.evaluate('= 2 + 2')}/${ctx.locale}/${Number.isNaN(ctx.now.getTime())}`;
      },
      USERNAME: () => undefined,
    };
    const r = updateFields(doc, { values: { name: 'Ada' }, evaluators });
    expect(results(doc)).toEqual(['custom:name', 'a+b=4/en-US/true', '?']);
    expect(seen).toEqual(['MYFIELD a b \\x']);
    expect(r.skipped.map((f) => f.type)).toEqual(['USERNAME']);
  });

  it('updates only the requested types', () => {
    const doc = docWith(para(cf('QUOTE "a"', '?')), para(cf('= 1 + 1', '?')));
    const r = updateFields(doc, { types: ['='] });
    expect(results(doc)).toEqual(['?', '2']);
    expect(r.updated).toBe(1);
    expect(r.skipped.map((f) => f.type)).toEqual(['QUOTE']);
  });

  it('updates headers, footers, footnotes and both branches of a text box', () => {
    const tb = (branch: string): string =>
      `<w:txbxContent><w:p><w:r><w:t>${branch} </w:t></w:r>${cf('QUOTE "in box"', 'old').xml}</w:p></w:txbxContent>`;
    const doc = docWith({
      xml:
        '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
        `<mc:Choice Requires="wps"><w:drawing><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:txbx>${tb('choice')}</wps:txbx></wps:wsp></w:drawing></mc:Choice>` +
        `<mc:Fallback><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox>${tb('fallback')}</v:textbox></v:shape></w:pict></mc:Fallback>` +
        '</mc:AlternateContent></w:r></w:p>',
    });
    expect(listFields(doc)).toHaveLength(1);
    const r = updateFields(doc);
    expect(r.updated).toBe(1);
    const x = xmlOf(doc);
    expect(x.match(/<w:t>in box<\/w:t>/g)).toHaveLength(2);
    expect(x).not.toContain('old');
  });
});
