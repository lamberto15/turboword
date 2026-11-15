// Number formats of list labels (ST_NumberFormat, ECMA-376 §17.18.59).
//
// Word's conventions where they differ from the obvious: letters repeat instead of carrying
// (27 → AA, 28 → BB, 53 → AAA), `decimalZero` pads to two digits, enclosed numbers exist only
// for 1-20 (larger values fall back to plain or parenthesized decimals), English text formats
// capitalize the first letter only ("Twenty-one"). Formats that are not implemented fall back
// to decimal, as LibreOffice does, so a label is never empty by accident.

const ROMAN: readonly (readonly [number, string])[] = [
  [1000, 'm'],
  [900, 'cm'],
  [500, 'd'],
  [400, 'cd'],
  [100, 'c'],
  [90, 'xc'],
  [50, 'l'],
  [40, 'xl'],
  [10, 'x'],
  [9, 'ix'],
  [5, 'v'],
  [4, 'iv'],
  [1, 'i'],
];

function roman(n: number): string {
  if (n <= 0) return String(n);
  let out = '';
  let v = n;
  for (const [value, s] of ROMAN) {
    while (v >= value) {
      out += s;
      v -= value;
    }
  }
  return out;
}

/** Word's repeating alphabet: 1..k → single letters, k+1 → doubled first letter, ... */
function repeatAlphabet(n: number, alphabet: readonly string[]): string {
  if (n <= 0) return String(n);
  const k = alphabet.length;
  return alphabet[(n - 1) % k]!.repeat(Math.floor((n - 1) / k) + 1);
}

const LATIN = 'abcdefghijklmnopqrstuvwxyz'.split('');
const RUSSIAN = 'абвгдежзиклмнопрстуфхцчшщэюя'.split('');

const ONES = [
  '',
  'one',
  'two',
  'three',
  'four',
  'five',
  'six',
  'seven',
  'eight',
  'nine',
  'ten',
  'eleven',
  'twelve',
  'thirteen',
  'fourteen',
  'fifteen',
  'sixteen',
  'seventeen',
  'eighteen',
  'nineteen',
];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];

function cardinalWords(n: number): string {
  if (n === 0) return 'zero';
  if (n < 0) return `minus ${cardinalWords(-n)}`;
  const parts: string[] = [];
  let v = n;
  const scales: readonly (readonly [number, string])[] = [
    [1_000_000_000, 'billion'],
    [1_000_000, 'million'],
    [1000, 'thousand'],
    [100, 'hundred'],
  ];
  for (const [value, word] of scales) {
    if (v >= value) {
      parts.push(`${cardinalWords(Math.floor(v / value))} ${word}`);
      v %= value;
    }
  }
  if (v > 0) {
    if (v < 20) parts.push(ONES[v]!);
    else parts.push(TENS[Math.floor(v / 10)]! + (v % 10 > 0 ? `-${ONES[v % 10]}` : ''));
  }
  return parts.join(' ');
}

const ORDINAL_WORDS: Readonly<Record<string, string>> = {
  one: 'first',
  two: 'second',
  three: 'third',
  five: 'fifth',
  eight: 'eighth',
  nine: 'ninth',
  twelve: 'twelfth',
};

function ordinalWords(n: number): string {
  const words = cardinalWords(n);
  const m = /([a-z]+)$/.exec(words);
  if (m === null) return words;
  const last = m[1]!;
  let ord = ORDINAL_WORDS[last];
  if (ord === undefined) ord = last.endsWith('y') ? `${last.slice(0, -1)}ieth` : `${last}th`;
  return words.slice(0, m.index) + ord;
}

function capitalize(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function ordinalSuffix(n: number): string {
  const mod100 = Math.abs(n) % 100;
  if (mod100 >= 11 && mod100 <= 13) return 'th';
  switch (Math.abs(n) % 10) {
    case 1:
      return 'st';
    case 2:
      return 'nd';
    case 3:
      return 'rd';
    default:
      return 'th';
  }
}

/** Replaces ASCII digits with the ten characters of `digits`. */
function mapDigits(n: number, digits: string): string {
  const chars = [...digits];
  return String(n).replace(/[0-9]/g, (d) => chars[Number(d)]!);
}

const CJK_DIGITS = '〇一二三四五六七八九';

/** CJK counting numerals (一, 十, 十一, 二十, 百, 千, 万). */
function cjkCounting(n: number): string {
  if (n <= 0) return String(n);
  const d = (k: number): string => CJK_DIGITS.charAt(k);
  const under10k = (v: number): string => {
    let out = '';
    const units: readonly (readonly [number, string])[] = [
      [1000, '千'],
      [100, '百'],
      [10, '十'],
    ];
    let rest = v;
    let zero = false;
    for (const [value, u] of units) {
      const q = Math.floor(rest / value);
      rest %= value;
      if (q > 0) {
        if (zero) out += d(0);
        out += (q === 1 && out === '' && value === 10 ? '' : d(q)) + u;
        zero = false;
      } else if (out !== '') zero = true;
    }
    if (rest > 0) out += (zero ? d(0) : '') + d(rest);
    return out;
  };
  if (n < 10000) return under10k(n);
  const high = Math.floor(n / 10000);
  const low = n % 10000;
  return `${cjkCounting(high)}万${low === 0 ? '' : (low < 1000 ? d(0) : '') + under10k(low)}`;
}

const CHICAGO = ['*', '†', '‡', '§'];
const IDEOGRAPH_TRADITIONAL = [...'甲乙丙丁戊己庚辛壬癸'];
const IDEOGRAPH_ZODIAC = [...'子丑寅卯辰巳午未申酉戌亥'];

function enclosed(n: number, first: number, fallback: (n: number) => string): string {
  return n >= 1 && n <= 20 ? String.fromCodePoint(first + n - 1) : fallback(n);
}

/**
 * `w:numFmt/@w:format` of a `custom` format (Word 2013+), e.g. `001, 002, 003, ...`: the first
 * sample's width and digits decide the padding and the digit set.
 */
const GREEK_LOWER = [...'αβγδεζηθικλμνξοπρστυφχψω'];
const GREEK_UPPER = [...'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ'];
const ALPHABETS: readonly (readonly string[])[] = [
  LATIN,
  LATIN.map((c) => c.toUpperCase()),
  RUSSIAN,
  RUSSIAN.map((c) => c.toUpperCase()),
  GREEK_LOWER,
  GREEK_UPPER,
];

function customFormat(n: number, pattern: string): string {
  const samples = pattern
    .split(',')
    .map((x) => x.trim())
    .filter((x) => x !== '' && x !== '...' && x !== '…');
  const sample = samples[0] ?? '';
  const chars = [...sample];
  if (chars.length === 0) return String(n);
  const last = chars[chars.length - 1]!;
  const cp = last.codePointAt(0)!;
  if (/\p{Nd}/u.test(last)) {
    // The sample ends in the "1" of some decimal digit set: derive that set's zero.
    const zeroChar = String.fromCodePoint(cp - 1);
    if (!/\p{Nd}/u.test(zeroChar)) return String(n);
    const digits = Array.from({ length: 10 }, (_, k) => String.fromCodePoint(cp - 1 + k)).join('');
    let s = mapDigits(n, digits);
    while ([...s].length < chars.length) s = zeroChar + s;
    return s;
  }
  if (chars.length !== 1) return String(n);
  // An alphabetic sequence ("α, β, γ, ..."): a known alphabet, else 26 consecutive letters.
  const known = ALPHABETS.find((a) => a[0] === sample);
  if (known !== undefined) return repeatAlphabet(n, known);
  const next = samples[1];
  if (next !== undefined && next.codePointAt(0) === cp + 1) {
    return repeatAlphabet(
      n,
      Array.from({ length: 26 }, (_, k) => String.fromCodePoint(cp + k)),
    );
  }
  return String(n);
}

/**
 * Formats a counter value with a `w:numFmt` value. `custom` is the `w:format` attribute of a
 * `custom` format. Unknown formats fall back to decimal; `bullet` and `none` give ''.
 */
export function formatNumberValue(value: number, format: string, custom?: string): string {
  const n = Math.trunc(value);
  switch (format) {
    case 'decimal':
    case 'decimalHalfWidth':
      return String(n);
    case 'decimalZero':
      return n >= 0 && n < 10 ? `0${n}` : String(n);
    case 'upperRoman':
      return roman(n).toUpperCase();
    case 'lowerRoman':
      return roman(n);
    case 'upperLetter':
      return repeatAlphabet(n, LATIN).toUpperCase();
    case 'lowerLetter':
      return repeatAlphabet(n, LATIN);
    case 'russianUpper':
      return repeatAlphabet(n, RUSSIAN).toUpperCase();
    case 'russianLower':
      return repeatAlphabet(n, RUSSIAN);
    case 'ordinal':
      return `${n}${ordinalSuffix(n)}`;
    case 'cardinalText':
      return capitalize(cardinalWords(n));
    case 'ordinalText':
      return capitalize(ordinalWords(n));
    case 'hex':
      return n >= 0 ? n.toString(16).toUpperCase() : String(n);
    case 'chicago': {
      if (n <= 0) return String(n);
      return CHICAGO[(n - 1) % 4]!.repeat(Math.floor((n - 1) / 4) + 1);
    }
    case 'bullet':
    case 'none':
      return '';
    case 'decimalEnclosedCircle':
    case 'decimalEnclosedCircleChinese':
      return enclosed(n, 0x2460, String);
    case 'decimalEnclosedParen':
      return enclosed(n, 0x2474, (v) => `(${v})`);
    case 'decimalEnclosedFullstop':
      return enclosed(n, 0x2488, (v) => `${v}.`);
    case 'decimalFullWidth':
    case 'decimalFullWidth2':
      return mapDigits(n, '０１２３４５６７８９');
    case 'ideographDigital':
      return mapDigits(n, CJK_DIGITS);
    case 'japaneseCounting':
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'taiwaneseCounting':
    case 'taiwaneseCountingThousand':
    case 'japaneseDigitalTenThousand':
      return cjkCounting(n);
    case 'ideographTraditional':
      return n >= 1 ? IDEOGRAPH_TRADITIONAL[(n - 1) % 10]! : String(n);
    case 'ideographZodiac':
      return n >= 1 ? IDEOGRAPH_ZODIAC[(n - 1) % 12]! : String(n);
    case 'hindiNumbers':
    case 'arabicIndic':
      return mapDigits(n, '٠١٢٣٤٥٦٧٨٩');
    case 'thaiNumbers':
      return mapDigits(n, '๐๑๒๓๔๕๖๗๘๙');
    case 'custom':
      return custom === undefined ? String(n) : customFormat(n, custom);
    default:
      return String(n);
  }
}
