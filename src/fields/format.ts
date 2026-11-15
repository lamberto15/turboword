// Field result formatting: the general (`\*`), numeric picture (`\#`) and date-time picture
// (`\@`) switches, ECMA-376 Part 1 §17.16.4.
//
// Started as a port of turbomerge `src/format.ts` (MIT, same author), then aligned with Word:
// - `\#`: `0` shows a digit or 0, `#` a digit or a space, `x` drops the digits to its left (or
//   sets the rounding place after the decimal point); `-`/`+` are sign placeholders; `;`
//   separates positive;negative;zero sections; 'text' is literal; every other character
//   (`$`, `%`, letters) is literal too (`%` does not scale, as in Word). The decimal and grouping
//   symbols of the picture and of the result are the locale's, as Word uses the system's;
// - `\@`: d dd ddd dddd, M MM MMM MMMM, yy yyyy, h hh H HH, m mm, s ss, AM/PM am/pm A/P a/p and
//   'literal text'; month, day and AM/PM names come from `Intl` for the locale;
// - `\*`: MERGEFORMAT/CHARFORMAT (formatting directives, no effect on the text), Upper, Lower,
//   FirstCap, Caps, Arabic, ArabicDash, Roman/roman, ALPHABETIC/alphabetic, Ordinal, CardText,
//   OrdText, DollarText, Hex. Unknown formats leave the value as it is.
// Rounding is half away from zero on the decimal representation (so 1.005 → 1.01).

import type { FieldSwitch } from './types.js';
import { alphabetic, cardText, dollarText, ordinal, ordText, roman } from './words.js';

/** Locale and time zone used by {@link formatValue}. */
export interface FormatContext {
  locale: string;
  timeZone?: string | undefined;
}

// ---- Intl helpers --------------------------------------------------------------------------

const dtfCache = new Map<string, Intl.DateTimeFormat>();

function dtf(
  locale: string,
  timeZone: string | undefined,
  opts: Intl.DateTimeFormatOptions,
): Intl.DateTimeFormat {
  const key = `${locale}|${timeZone ?? ''}|${JSON.stringify(opts)}`;
  let f = dtfCache.get(key);
  if (f === undefined) {
    const o: Intl.DateTimeFormatOptions = { ...opts };
    if (timeZone !== undefined) o.timeZone = timeZone;
    try {
      f = new Intl.DateTimeFormat(locale, o);
    } catch {
      f = new Intl.DateTimeFormat('en-US', o);
    }
    dtfCache.set(key, f);
  }
  return f;
}

const sepCache = new Map<string, { decimal: string; group: string }>();

/** Decimal and grouping symbols of `locale`. */
export function separators(locale: string): { decimal: string; group: string } {
  let s = sepCache.get(locale);
  if (s === undefined) {
    let decimal = '.';
    let group = ',';
    try {
      for (const p of new Intl.NumberFormat(locale).formatToParts(12345.6)) {
        if (p.type === 'decimal') decimal = p.value;
        else if (p.type === 'group') group = p.value;
      }
    } catch {
      // unknown locale: keep the defaults
    }
    s = { decimal, group };
    sepCache.set(locale, s);
  }
  return s;
}

interface DateParts {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  weekday: number;
}

const WEEKDAYS: Readonly<Record<string, number>> = {
  Sun: 0,
  Mon: 1,
  Tue: 2,
  Wed: 3,
  Thu: 4,
  Fri: 5,
  Sat: 6,
};

/** Calendar fields of `date` in `timeZone` (the host's zone when undefined). */
export function dateParts(date: Date, timeZone: string | undefined): DateParts {
  const parts = dtf('en-US', timeZone, {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: 'numeric',
    second: 'numeric',
    weekday: 'short',
    hourCycle: 'h23',
  }).formatToParts(date);
  const get = (type: string): string => parts.find((p) => p.type === type)?.value ?? '0';
  return {
    year: Number(get('year')),
    month: Number(get('month')),
    day: Number(get('day')),
    hour: Number(get('hour')) % 24,
    minute: Number(get('minute')),
    second: Number(get('second')),
    weekday: WEEKDAYS[get('weekday')] ?? 0,
  };
}

/** The instant of a wall-clock time in `timeZone` (host local time when undefined). */
export function zonedDate(
  y: number,
  mo: number,
  d: number,
  h: number,
  mi: number,
  s: number,
  timeZone: string | undefined,
): Date {
  if (timeZone === undefined) return new Date(y, mo - 1, d, h, mi, s);
  const want = Date.UTC(y, mo - 1, d, h, mi, s);
  let t = want;
  for (let i = 0; i < 2; i++) {
    const p = dateParts(new Date(t), timeZone);
    const shown = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
    t += want - shown;
  }
  return new Date(t);
}

// ---- values --------------------------------------------------------------------------------

const ISO_DATE =
  /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/;
const NUMERIC_DATE =
  /^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})(?:[ ,]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?$/;

const orderCache = new Map<string, string>();

/** Order of day/month/year in the locale's numeric dates ('mdy', 'dmy' or 'ymd'). */
function dateOrder(locale: string): string {
  let o = orderCache.get(locale);
  if (o === undefined) {
    o = '';
    for (const p of dtf(locale, 'UTC', {
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
    }).formatToParts(new Date(Date.UTC(2001, 1, 3)))) {
      if (p.type === 'year') o += 'y';
      else if (p.type === 'month') o += 'm';
      else if (p.type === 'day') o += 'd';
    }
    if (o.length !== 3) o = 'mdy';
    orderCache.set(locale, o);
  }
  return o;
}

/**
 * A Date from a value: a Date, an ISO 8601 string, or a numeric date in the locale's order
 * ("3/9/2024" in en-US, "9.3.2024" in de-DE), optionally with a time. Undefined otherwise.
 */
export function toDate(value: unknown, ctx: FormatContext): Date | undefined {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? undefined : value;
  if (typeof value !== 'string') return undefined;
  const s = value.trim();
  const iso = ISO_DATE.exec(s);
  if (iso) {
    const [, y, mo, d, h, mi, sec, zone] = iso;
    if (zone !== undefined) {
      const date = new Date(s.replace(' ', 'T'));
      return Number.isNaN(date.getTime()) ? undefined : date;
    }
    return zonedDate(+y!, +mo!, +d!, +(h ?? 0), +(mi ?? 0), +(sec ?? 0), ctx.timeZone);
  }
  const num = NUMERIC_DATE.exec(s);
  if (num) {
    const order = dateOrder(ctx.locale);
    const fields = [num[1]!, num[2]!, num[3]!].map(Number);
    let y: number;
    let mo: number;
    let d: number;
    if (num[1]!.length === 4) [y, mo, d] = fields as [number, number, number];
    else if (order === 'dmy') [d, mo, y] = fields as [number, number, number];
    else [mo, d, y] = fields as [number, number, number];
    if (y < 100) y += 2000;
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return undefined;
    let h = Number(num[4] ?? 0);
    const ampm = num[7]?.toLowerCase();
    if (ampm === 'pm' && h < 12) h += 12;
    if (ampm === 'am' && h === 12) h = 0;
    return zonedDate(y, mo, d, h, Number(num[5] ?? 0), Number(num[6] ?? 0), ctx.timeZone);
  }
  return undefined;
}

const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i;

/**
 * A number from a value: numbers, bigints and booleans as such; strings leniently, with the
 * locale's grouping and decimal symbols, currency symbols, a trailing % and accounting
 * parentheses for negatives ("(1,234.50)" → -1234.5). Undefined when not numeric.
 */
export function toNumber(value: unknown, locale: string): number | undefined {
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined;
  if (typeof value === 'bigint') return Number(value);
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value !== 'string') return undefined;
  let s = value.trim().replace(/[\s\u00a0\u202f$€£¥¢₹%]/g, '');
  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.slice(1, -1);
  }
  const { decimal, group } = separators(locale);
  if (group !== '' && /^[+-]?\d{1,3}(?:[.,'’]\d{3})+(?:[.,]\d+)?$/.test(s)) {
    s = s.split(group).join('');
  }
  if (decimal !== '.') s = s.replace(decimal, '.');
  if (!NUMBER.test(s)) return undefined;
  const n = Number(s);
  return negative ? -n : n;
}

/** Rounds |n| to `places` decimals, half away from zero, on the decimal representation. */
function roundAbs(n: number, places: number): string {
  const a = Math.abs(n);
  const p = Math.min(Math.max(places, 0), 20);
  const shifted = Number(`${a}e${p}`);
  if (!Number.isFinite(shifted)) return a.toFixed(p);
  const r = Math.round(shifted);
  const s = Number(`${r}e-${p}`).toFixed(p);
  return s;
}

/** Number as text with the locale's decimal symbol, no grouping, at most 15 significant digits. */
export function numberText(n: number, locale: string): string {
  const s = String(Number(n.toPrecision(15)));
  if (/e/i.test(s)) return s;
  const { decimal } = separators(locale);
  return decimal === '.' ? s : s.replace('.', decimal);
}

/**
 * Languages whose usual short date pads day and month ("11.06.2019" in de-DE) where `Intl`'s
 * numeric style does not.
 */
const PADDED_DATES: ReadonlySet<string> = new Set([
  'de',
  'es',
  'it',
  'ja',
  'pl',
  'cs',
  'da',
  'fi',
  'nb',
]);

/** Default text of a date: the locale's numeric date, plus the time when it is not midnight. */
export function dateText(date: Date, ctx: FormatContext, withTime?: boolean): string {
  const lang = ctx.locale.toLowerCase().split(/[-_]/)[0] ?? '';
  const pad = PADDED_DATES.has(lang) ? '2-digit' : 'numeric';
  const day = dtf(ctx.locale, ctx.timeZone, {
    year: 'numeric',
    month: pad,
    day: pad,
  }).format(date);
  const p = dateParts(date, ctx.timeZone);
  const timed = withTime ?? !(p.hour === 0 && p.minute === 0 && p.second === 0);
  if (!timed) return day;
  const time = dtf(ctx.locale, ctx.timeZone, {
    hour: 'numeric',
    minute: '2-digit',
    second: '2-digit',
  }).format(date);
  return `${day} ${time}`;
}

/** Final text of a value (no switches). */
export function toText(value: unknown, ctx: FormatContext): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return dateText(value, ctx);
  if (typeof value === 'number') return numberText(value, ctx.locale);
  if (typeof value === 'string') return value;
  return String(value);
}

// ---- \# numeric pictures ---------------------------------------------------------------------

type PicItem =
  | { k: 'digit'; c: '0' | '#' | 'x' }
  | { k: 'dec' }
  | { k: 'group' }
  | { k: 'minus' }
  | { k: 'plus' }
  | { k: 'lit'; t: string };

/** Splits a picture into its `;` sections (outside 'quoted' text). */
function sections(picture: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quote = false;
  for (const ch of picture) {
    if (ch === "'") quote = !quote;
    if (ch === ';' && !quote) {
      out.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

function pictureItems(section: string, decimal: string, group: string): PicItem[] {
  const items: PicItem[] = [];
  let hasDec = false;
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]!;
    if (ch === "'" || ch === '`') {
      const close = section.indexOf(ch, i + 1);
      const end = close < 0 ? section.length : close;
      if (ch === "'") items.push({ k: 'lit', t: section.slice(i + 1, end) });
      i = end;
    } else if (ch === '0' || ch === '#' || ch === 'x') {
      items.push({ k: 'digit', c: ch });
    } else if (ch === decimal && !hasDec) {
      hasDec = true;
      items.push({ k: 'dec' });
    } else if (ch === group) {
      items.push({ k: 'group' });
    } else if (ch === '-') {
      items.push({ k: 'minus' });
    } else if (ch === '+') {
      items.push({ k: 'plus' });
    } else {
      items.push({ k: 'lit', t: ch });
    }
  }
  return items;
}

function groupDigits(digits: string, sep: string): string {
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, sep);
}

/** Applies a `\#` numeric picture. */
export function formatNumberPicture(n: number, picture: string, locale: string): string {
  const { decimal, group } = separators(locale);
  const secs = sections(picture);
  let section = secs[0]!;
  let v = n;
  let ownSign = false;
  if (n === 0 && secs.length >= 3) section = secs[2]!;
  else if (n < 0 && secs.length >= 2) {
    section = secs[1]!;
    v = -n;
    ownSign = true;
  }
  const items = pictureItems(section, decimal, group);
  const isPh = (it: PicItem): boolean => it.k === 'digit' || it.k === 'dec';
  const first = items.findIndex(isPh);
  if (first < 0) return items.map((it) => (it.k === 'lit' ? it.t : '')).join('');
  let last = items.length - 1;
  while (!isPh(items[last]!)) last--;
  const body = items.slice(first, last + 1);
  const decAt = body.findIndex((it) => it.k === 'dec');
  const intBody = decAt < 0 ? body : body.slice(0, decAt);
  const fracBody = decAt < 0 ? [] : body.slice(decAt + 1);
  const intPh = intBody.filter((it) => it.k === 'digit') as { k: 'digit'; c: string }[];
  const fracPh = fracBody.filter((it) => it.k === 'digit');
  const grouping = items.some((it) => it.k === 'group');

  const rounded = roundAbs(v, fracPh.length);
  const [intRaw = '0', fracRaw = ''] = rounded.split('.');
  let digits = intRaw;
  const xAt = intPh.findIndex((p) => p.c === 'x');
  if (xAt >= 0) digits = digits.slice(-(intPh.length - xAt));
  const zeroAt = intPh.findIndex((p) => p.c === '0');
  const minDigits = zeroAt < 0 ? 0 : intPh.length - zeroAt;
  if (/^0+$/.test(digits) && minDigits === 0) digits = '';
  digits = digits.padStart(minDigits, '0');

  let intText = '';
  if (grouping || intPh.length === 0) {
    intText = ' '.repeat(Math.max(0, intPh.length - digits.length)) + groupDigits(digits, group);
    if (!grouping) intText = digits;
    if (grouping) {
      const lits = intBody.filter((it) => it.k === 'lit').map((it) => (it as { t: string }).t);
      intText += lits.join('');
    }
  } else {
    // Positional: digits fill the placeholders from the right; extra digits go to the first one.
    let di = digits.length - 1;
    let phLeft = intPh.length;
    const parts: string[] = [];
    for (let i = intBody.length - 1; i >= 0; i--) {
      const it = intBody[i]!;
      if (it.k === 'lit') parts.push(it.t);
      else if (it.k === 'digit') {
        phLeft--;
        if (phLeft === 0 && di >= 0) {
          parts.push(digits.slice(0, di + 1));
          di = -1;
        } else if (di >= 0) {
          parts.push(digits[di]!);
          di--;
        } else {
          parts.push(it.c === '0' ? '0' : ' ');
        }
      }
    }
    intText = parts.reverse().join('');
  }
  let fracText = '';
  let fi = 0;
  for (const it of fracBody) {
    if (it.k === 'digit') fracText += fracRaw[fi++] ?? '0';
    else if (it.k === 'lit') fracText += it.t;
  }
  const numberPart = intText + (decAt >= 0 ? decimal + fracText : '');

  const hasSignPh = items.some((it) => it.k === 'minus' || it.k === 'plus');
  const negative = n < 0 && !/^[0.]*$/.test(rounded);
  let out = '';
  for (let i = 0; i < first; i++) out += renderSign(items[i]!, n, negative);
  out += numberPart;
  for (let i = last + 1; i < items.length; i++) out += renderSign(items[i]!, n, negative);
  if (negative && !ownSign && !hasSignPh) out = `-${out}`;
  return out;
}

function renderSign(it: PicItem, n: number, negative: boolean): string {
  switch (it.k) {
    case 'lit':
      return it.t;
    case 'minus':
      return negative ? '-' : ' ';
    case 'plus':
      return negative ? '-' : n > 0 ? '+' : ' ';
    case 'group':
      return '';
    default:
      return '';
  }
}

// ---- \@ date-time pictures -----------------------------------------------------------------

const DATE_TOKEN = /'[^']*'?|`[^`]*`?|AM\/PM|am\/pm|A\/P|a\/p|M+|m+|d+|D+|y+|Y+|h+|H+|s+|S+/g;

const pad2 = (n: number): string => String(n).padStart(2, '0');

function dayPeriod(date: Date, ctx: FormatContext): string {
  const p = dtf(ctx.locale, ctx.timeZone, { hour: 'numeric', hour12: true })
    .formatToParts(date)
    .find((x) => x.type === 'dayPeriod')?.value;
  if (p !== undefined && p !== '') return p;
  return dateParts(date, ctx.timeZone).hour < 12 ? 'AM' : 'PM';
}

/** Applies a `\@` date-time picture. */
export function formatDatePicture(date: Date, picture: string, ctx: FormatContext): string {
  const p = dateParts(date, ctx.timeZone);
  const name = (o: Intl.DateTimeFormatOptions): string =>
    dtf(ctx.locale, ctx.timeZone, o).format(date);
  return picture.replace(DATE_TOKEN, (tok) => {
    const c = tok[0]!;
    if (c === "'") return tok.slice(1, tok.endsWith("'") && tok.length > 1 ? -1 : undefined);
    if (c === '`') return '';
    if (tok === 'AM/PM') return dayPeriod(date, ctx);
    if (tok === 'am/pm') return dayPeriod(date, ctx).toLocaleLowerCase(ctx.locale);
    if (tok === 'A/P') return dayPeriod(date, ctx).charAt(0).toLocaleUpperCase(ctx.locale);
    if (tok === 'a/p') return dayPeriod(date, ctx).charAt(0).toLocaleLowerCase(ctx.locale);
    const n = tok.length;
    switch (c) {
      case 'M':
        if (n === 1) return String(p.month);
        if (n === 2) return pad2(p.month);
        return name({ month: n === 3 ? 'short' : 'long' });
      case 'd':
      case 'D':
        if (n === 1) return String(p.day);
        if (n === 2) return pad2(p.day);
        return name({ weekday: n === 3 ? 'short' : 'long' });
      case 'y':
      case 'Y':
        return n <= 2 ? pad2(p.year % 100) : String(p.year);
      case 'h': {
        const h12 = p.hour % 12 || 12;
        return n === 1 ? String(h12) : pad2(h12);
      }
      case 'H':
        return n === 1 ? String(p.hour) : pad2(p.hour);
      case 'm':
        return n === 1 ? String(p.minute) : pad2(p.minute);
      default:
        return n === 1 ? String(p.second) : pad2(p.second);
    }
  });
}

// ---- \* general formats --------------------------------------------------------------------

function generalFormat(value: unknown, format: string, ctx: FormatContext): unknown {
  const key = format.toLowerCase();
  const text = (): string => toText(value, ctx);
  const loc = ctx.locale;
  switch (key) {
    case 'upper':
      return text().toLocaleUpperCase(loc);
    case 'lower':
      return text().toLocaleLowerCase(loc);
    case 'firstcap': {
      const t = text();
      return t.charAt(0).toLocaleUpperCase(loc) + t.slice(1).toLocaleLowerCase(loc);
    }
    case 'caps':
      return text()
        .toLocaleLowerCase(loc)
        .replace(/(^|\s)(\p{L})/gu, (_m, sp: string, ch: string) => sp + ch.toLocaleUpperCase(loc));
    default:
      break;
  }
  const n = value instanceof Date ? undefined : toNumber(value, loc);
  if (n === undefined) return value;
  const int = Math.round(n);
  const upper = format[0] !== undefined && format[0] === format[0].toUpperCase();
  switch (key) {
    case 'arabic':
      return String(int);
    case 'arabicdash':
      return `- ${int} -`;
    case 'roman':
      return upper ? roman(int) : roman(int).toLowerCase();
    case 'alphabetic':
      return upper ? alphabetic(int) : alphabetic(int).toLowerCase();
    case 'ordinal':
      return ordinal(int);
    case 'cardtext':
      return cardText(int);
    case 'ordtext':
      return ordText(int);
    case 'dollartext':
      return dollarText(n);
    case 'hex':
      return Math.trunc(Math.abs(n)).toString(16).toUpperCase();
    default:
      return value;
  }
}

/** Applies the `\*`, `\#` and `\@` switches in order, leaving others alone; returns raw value. */
export function applySwitches(
  value: unknown,
  switches: readonly FieldSwitch[],
  ctx: FormatContext,
): unknown {
  let v = value;
  for (const s of switches) {
    if (s.arg === undefined) continue;
    if (s.name === '\\#') {
      const n = v instanceof Date ? undefined : toNumber(v, ctx.locale);
      if (n !== undefined) v = formatNumberPicture(n, s.arg, ctx.locale);
    } else if (s.name === '\\@') {
      const d = toDate(v, ctx);
      if (d !== undefined) v = formatDatePicture(d, s.arg, ctx);
    } else if (s.name === '\\*') {
      v = generalFormat(v, s.arg, ctx);
    }
  }
  return v;
}

/** True when the switches include a `\*` format other than MERGEFORMAT/CHARFORMAT. */
export function hasTextFormat(switches: readonly FieldSwitch[]): boolean {
  return switches.some(
    (s) => s.name === '\\*' && s.arg !== undefined && !/^(merge|char)format$/i.test(s.arg),
  );
}

/** {@link applySwitches} then {@link toText}. */
export function formatValue(
  value: unknown,
  switches: readonly FieldSwitch[],
  ctx: FormatContext,
): string {
  return toText(applySwitches(value, switches, ctx), ctx);
}
