// Number names for the `\*` general-format switch: Roman numerals, alphabetic labels, ordinal
// suffixes and English number words (CardText, OrdText, DollarText).
//
// Word produces the words in the document's language; turboword implements English, which is
// what Word writes for en-* documents. Other languages get the English words too.

const ONES = [
  'zero',
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
const SCALES = ['', 'thousand', 'million', 'billion', 'trillion'];

/** Words for 0 < n < 1000. */
function hundreds(n: number): string {
  const parts: string[] = [];
  const h = Math.floor(n / 100);
  const rest = n % 100;
  if (h > 0) parts.push(`${ONES[h]} hundred`);
  if (rest >= 20) {
    const t = TENS[Math.floor(rest / 10)]!;
    parts.push(rest % 10 === 0 ? t : `${t}-${ONES[rest % 10]}`);
  } else if (rest > 0) {
    parts.push(ONES[rest]!);
  }
  return parts.join(' ');
}

/** Cardinal English words ("one hundred twenty-three"); the integer part of |n|. */
export function cardText(n: number): string {
  let v = Math.floor(Math.abs(n));
  if (v === 0) return 'zero';
  if (v >= 1e15) return String(v);
  const groups: string[] = [];
  for (let s = 0; v > 0; s++) {
    const g = v % 1000;
    if (g > 0) groups.unshift(SCALES[s] ? `${hundreds(g)} ${SCALES[s]}` : hundreds(g));
    v = Math.floor(v / 1000);
  }
  return groups.join(' ');
}

const ORD_IRREGULAR: Readonly<Record<string, string>> = {
  one: 'first',
  two: 'second',
  three: 'third',
  five: 'fifth',
  eight: 'eighth',
  nine: 'ninth',
  twelve: 'twelfth',
};

/** Ordinal English words ("twenty-first"). */
export function ordText(n: number): string {
  const words = cardText(n);
  const m = /([a-z]+)$/.exec(words);
  if (!m) return words;
  const last = m[1]!;
  let ord: string;
  if (ORD_IRREGULAR[last] !== undefined) ord = ORD_IRREGULAR[last];
  else if (last.endsWith('y')) ord = `${last.slice(0, -1)}ieth`;
  else ord = `${last}th`;
  return words.slice(0, words.length - last.length) + ord;
}

/** Ordinal suffix form ("1st", "22nd", "113th"). */
export function ordinal(n: number): string {
  const v = Math.trunc(n);
  const a = Math.abs(v);
  const t = a % 100;
  const suffix =
    t >= 11 && t <= 13
      ? 'th'
      : a % 10 === 1
        ? 'st'
        : a % 10 === 2
          ? 'nd'
          : a % 10 === 3
            ? 'rd'
            : 'th';
  return `${v}${suffix}`;
}

/** Cheque text: "one hundred twenty-three and 45/100". */
export function dollarText(n: number): string {
  const a = Math.abs(n);
  let whole = Math.floor(a);
  let cents = Math.round((a - whole) * 100);
  if (cents === 100) {
    whole += 1;
    cents = 0;
  }
  return `${cardText(whole)} and ${String(cents).padStart(2, '0')}/100`;
}

const ROMAN: readonly (readonly [number, string])[] = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];

/** Upper-case Roman numerals (1-32767, as Word; outside that range the decimal number). */
export function roman(n: number): string {
  let v = Math.trunc(n);
  if (v < 1 || v > 32767) return String(v);
  let out = '';
  for (const [value, digits] of ROMAN) {
    while (v >= value) {
      out += digits;
      v -= value;
    }
  }
  return out;
}

/** Upper-case alphabetic label as Word writes it: A..Z, then AA..ZZ, AAA..: 27 = "AA". */
export function alphabetic(n: number): string {
  const v = Math.trunc(n);
  if (v < 1 || v > 780) return String(v);
  const letter = String.fromCharCode(65 + ((v - 1) % 26));
  return letter.repeat(Math.floor((v - 1) / 26) + 1);
}
