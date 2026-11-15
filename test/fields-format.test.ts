// formatFieldValue: the \* general formats, \# numeric pictures and \@ date-time pictures,
// checked against the examples of Word's field-switch documentation where it has them.

import { describe, expect, it } from 'vitest';
import { formatFieldValue } from '../src/fields/index.js';
import { parseInstruction, tokenize } from '../src/fields/tokenize.js';
import type { FieldSwitch } from '../src/fields/types.js';

const star = (arg: string): FieldSwitch => ({ name: '\\*', arg });
const num = (arg: string): FieldSwitch => ({ name: '\\#', arg });
const at = (arg: string): FieldSwitch => ({ name: '\\@', arg });
const f = (v: unknown, s: FieldSwitch[], locale = 'en-US'): string =>
  formatFieldValue(v, s, { locale, timeZone: 'UTC' });

describe('\\* general formats', () => {
  it('changes case', () => {
    expect(f('hello wORLD', [star('Upper')])).toBe('HELLO WORLD');
    expect(f('Hello World', [star('Lower')])).toBe('hello world');
    expect(f('weekly REPORT on sales', [star('FirstCap')])).toBe('Weekly report on sales');
    expect(f('luis alverca', [star('Caps')])).toBe('Luis Alverca');
    expect(f('straße', [star('Upper')], 'de-DE')).toBe('STRASSE');
  });

  it('formats numbers as numerals and words', () => {
    expect(f(11, [star('Roman')])).toBe('XI');
    expect(f(11, [star('ROMAN')])).toBe('XI');
    expect(f(1994, [star('roman')])).toBe('mcmxciv');
    expect(f(3, [star('ALPHABETIC')])).toBe('C');
    expect(f(27, [star('alphabetic')])).toBe('aa');
    expect(f(53, [star('ALPHABETIC')])).toBe('AAA');
    expect(f('4.0', [star('Arabic')])).toBe('4');
    expect(f(7, [star('ArabicDash')])).toBe('- 7 -');
    expect([1, 2, 3, 4, 11, 12, 13, 21, 22, 101, 111].map((n) => f(n, [star('Ordinal')]))).toEqual([
      '1st',
      '2nd',
      '3rd',
      '4th',
      '11th',
      '12th',
      '13th',
      '21st',
      '22nd',
      '101st',
      '111th',
    ]);
    expect(f(0, [star('CardText')])).toBe('zero');
    expect(f(780, [star('CardText')])).toBe('seven hundred eighty');
    expect(f(21, [star('CardText')])).toBe('twenty-one');
    expect(f(1234567, [star('CardText')])).toBe(
      'one million two hundred thirty-four thousand five hundred sixty-seven',
    );
    expect(f(1, [star('OrdText')])).toBe('first');
    expect(f(12, [star('OrdText')])).toBe('twelfth');
    expect(f(21, [star('OrdText')])).toBe('twenty-first');
    expect(f(40, [star('OrdText')])).toBe('fortieth');
    expect(f(100, [star('OrdText')])).toBe('one hundredth');
    expect(f(14.55, [star('DollarText'), star('Upper')])).toBe('FOURTEEN AND 55/100');
    expect(f(458, [star('Hex')])).toBe('1CA');
    expect(f(21, [star('CardText'), star('Caps')])).toBe('Twenty-one');
  });

  it('leaves the value alone for MERGEFORMAT, CHARFORMAT, unknown formats and text', () => {
    expect(f('Abc', [star('MERGEFORMAT')])).toBe('Abc');
    expect(f('Abc', [star('CHARFORMAT')])).toBe('Abc');
    expect(f('Abc', [star('Nonsense')])).toBe('Abc');
    expect(f('Abc', [star('Roman')])).toBe('Abc');
    expect(f(null, [star('Upper')])).toBe('');
  });
});

describe('\\# numeric pictures', () => {
  it('follows the documented examples', () => {
    expect(f(9, [num('00.00')])).toBe('09.00');
    expect(f(15, [num('$###')])).toBe('$ 15');
    expect(f(222492, [num('x##')])).toBe('492');
    expect(f(0.125, [num('0.00x')])).toBe('0.125');
    expect(f(0.75, [num('.x')])).toBe('.8');
    expect(f(495.47, [num('$###.00')])).toBe('$495.47');
    expect(f(2456800, [num('$#,###,###')])).toBe('$2,456,800');
    expect(f(-80, [num('-##')])).toBe('-80');
    expect(f(10, [num('+##')])).toBe('+10');
    expect(f(-10, [num('+##')])).toBe('-10');
    expect(f(33, [num('##%')])).toBe('33%');
    expect(f(1245.65, [num('$#,##0.00;-$#,##0.00')])).toBe('$1,245.65');
    expect(f(-345.56, [num('$#,##0.00;-$#,##0.00')])).toBe('-$ 345.56');
    expect(f(-345.56, [num('$#,##0.00;($#,##0.00);$0')])).toBe('($ 345.56)');
    expect(f(0, [num('$#,##0.00;($#,##0.00);$0')])).toBe('$0');
    expect(f(347.44, [num("$##0.00 'is sales tax'")])).toBe('$347.44 is sales tax');
  });

  it('groups, pads, rounds half away from zero and signs', () => {
    expect(f(1234.5, [num('$,0.00')])).toBe('$1,234.50');
    expect(f(1.005, [num('0.00')])).toBe('1.01');
    expect(f(2.5, [num('0')])).toBe('3');
    expect(f(-2.5, [num('0')])).toBe('-3');
    expect(f(-5, [num('0.00')])).toBe('-5.00');
    expect(f(-0.001, [num('0.00')])).toBe('0.00');
    expect(f(5, [num('#,##0')])).toBe('   5');
    expect(f(5551234, [num("000'-'0000")])).toBe('555-1234');
    expect(f(1234567.891, [num('#,##0.00')])).toBe('1,234,567.89');
  });

  it('reads numbers from text and uses the locale separators', () => {
    expect(f('1,234.5', [num('0.00')])).toBe('1234.50');
    expect(f('$1,234.50', [num('#,##0')])).toBe('1,235');
    expect(f('(12)', [num('0')])).toBe('-12');
    expect(f(1234.5, [num('#.##0,00')], 'de-DE')).toBe('1.234,50');
    expect(f('1.234,5', [num('0,00')], 'de-DE')).toBe('1234,50');
    expect(f('n/a', [num('0.00')])).toBe('n/a');
  });
});

describe('\\@ date-time pictures', () => {
  const d = new Date(Date.UTC(2024, 2, 9, 17, 7, 8));

  it('formats every token', () => {
    expect(f(d, [at('dddd, MMMM d, yyyy')])).toBe('Saturday, March 9, 2024');
    expect(f(d, [at('ddd MMM dd yy')])).toBe('Sat Mar 09 24');
    expect(f(d, [at('M/d/yyyy')])).toBe('3/9/2024');
    expect(f(d, [at('MM-dd')])).toBe('03-09');
    expect(f(d, [at('h:mm AM/PM')])).toBe('5:07 PM');
    expect(f(d, [at('hh:mm:ss am/pm')])).toBe('05:07:08 pm');
    expect(f(d, [at('H:m:s')])).toBe('17:7:8');
    expect(f(d, [at('HH:mm A/P')])).toBe('17:07 P');
    expect(f(d, [at("d 'of' MMMM")])).toBe('9 of March');
    expect(f(d, [at('MMMM'), star('Upper')])).toBe('MARCH');
  });

  it('uses the locale for names', () => {
    expect(f(d, [at('dddd, d. MMMM yyyy')], 'de-DE')).toBe('Samstag, 9. März 2024');
    expect(f(d, [at("d 'de' MMMM 'de' yyyy")], 'es-ES')).toBe('9 de marzo de 2024');
    expect(f(d, [at('dddd d MMMM')], 'fr-FR')).toBe('samedi 9 mars');
  });

  it('honours the time zone and parses date strings', () => {
    expect(formatFieldValue(d, [at('HH:mm')], { timeZone: 'Asia/Tokyo' })).toBe('02:07');
    expect(f('2024-03-09', [at('dddd')])).toBe('Saturday');
    expect(f('2024-03-09T10:30:00Z', [at('HH:mm')])).toBe('10:30');
    expect(f('3/9/2024', [at('MMMM d')])).toBe('March 9');
    expect(f('9.3.2024', [at('MMMM d')], 'de-DE')).toBe('März 9');
    expect(f('not a date', [at('yyyy')])).toBe('not a date');
  });

  it('writes dates and numbers by default in the locale', () => {
    expect(f(d, [])).toBe('3/9/2024 5:07:08 PM');
    expect(f(new Date(Date.UTC(2024, 2, 9)), [])).toBe('3/9/2024');
    expect(f(new Date(Date.UTC(2024, 2, 9)), [], 'de-DE')).toBe('09.03.2024');
    expect(f(new Date(Date.UTC(2024, 2, 9)), [], 'en-GB')).toBe('09/03/2024');
    expect(f(1.5, [], 'de-DE')).toBe('1,5');
    expect(f(0.1 + 0.2, [])).toBe('0.3');
    expect(f(true, [])).toBe('true');
  });
});

describe('instruction syntax', () => {
  it('tokenizes with double quotes, escapes and concatenation', () => {
    expect(tokenize(' MERGEFIELD "Hello world" \\b "say \\"hi\\" " ')).toEqual([
      { text: 'MERGEFIELD', quoted: false },
      { text: 'Hello world', quoted: true },
      { text: '\\b', quoted: false },
      { text: 'say "hi" ', quoted: true },
    ]);
    expect(tokenize('a"b c"d "" \'x y\'').map((t) => t.text)).toEqual(['ab cd', '', "'x", "y'"]);
    expect(tokenize('INCLUDETEXT "C:\\\\Docs\\\\a.docx"')[1]?.text).toBe('C:\\Docs\\a.docx');
    expect(tokenize('QUOTE "open')[1]).toEqual({ text: 'open', quoted: true });
  });

  it('parses type, arguments and switches', () => {
    const p = parseInstruction(' MERGEFIELD Name \\b "Dear " \\f "," \\* Caps \\m');
    expect(p.type).toBe('MERGEFIELD');
    expect(p.args.map((a) => a.text)).toEqual(['Name']);
    expect(p.switches).toEqual([
      { name: '\\b', arg: 'Dear ' },
      { name: '\\f', arg: ',' },
      { name: '\\*', arg: 'Caps' },
      { name: '\\m' },
    ]);
    expect(parseInstruction('TOC \\o "1-3" \\h \\z \\u').switches).toEqual([
      { name: '\\o', arg: '1-3' },
      { name: '\\h' },
      { name: '\\z' },
      { name: '\\u' },
    ]);
    expect(parseInstruction('TOC \\o \\h').switches).toEqual([{ name: '\\o' }, { name: '\\h' }]);
    expect(parseInstruction('page \\*roman').switches).toEqual([{ name: '\\*', arg: 'roman' }]);
    expect(parseInstruction('DATE \\@"d MMM"').switches).toEqual([{ name: '\\@', arg: 'd MMM' }]);
    expect(parseInstruction('REF _Ref1 \\h \\n').args.map((a) => a.text)).toEqual(['_Ref1']);
    expect(parseInstruction('SEQ Figure \\r 3 \\* ARABIC').switches).toEqual([
      { name: '\\r', arg: '3' },
      { name: '\\*', arg: 'ARABIC' },
    ]);
    const formula = parseInstruction('=SUM(ABOVE) * 1.5 \\# "0.00"');
    expect(formula.type).toBe('=');
    expect(formula.args[0]?.text).toBe('SUM(ABOVE) * 1.5');
    expect(formula.switches).toEqual([{ name: '\\#', arg: '0.00' }]);
    expect(parseInstruction('').type).toBe('');
  });
});
