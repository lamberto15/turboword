// Word date pictures (`w:dateFormat` of date content controls, `\@` of DATE fields, legacy
// date form fields): `d dd ddd dddd M MM MMM MMMM yy yyyy H HH h hh m mm s ss am/pm`, quoted
// literals ('text'), everything else copied as is.
//
// Month and day names come from `Intl.DateTimeFormat` for the control's language (`w:lid`), so
// "d MMMM yyyy" in de-DE reads "15 Januar 2024". When the picture also shows the day of the
// month, the month name is taken in its format context, which gives the genitive in languages
// that have one (ru "15 января 2024"), as Word writes it.
//
// Dates are read in UTC: a content control stores a calendar date (`2024-01-15T00:00:00Z`), and
// UTC keeps the output the same on every machine.

/** Formatter cache: `lid|kind` → formatter (creating Intl formatters is slow). */
const FORMATTERS = new Map<string, Intl.DateTimeFormat>();

function formatter(lid: string, kind: string): Intl.DateTimeFormat {
  const key = `${lid}|${kind}`;
  let f = FORMATTERS.get(key);
  if (f === undefined) {
    const o: Intl.DateTimeFormatOptions = { timeZone: 'UTC' };
    if (kind === 'dddd' || kind === 'ddd') o.weekday = kind === 'dddd' ? 'long' : 'short';
    else {
      o.month = kind.startsWith('MMMM') ? 'long' : 'short';
      if (kind.endsWith('+d')) o.day = 'numeric';
    }
    try {
      f = new Intl.DateTimeFormat(lid, o);
    } catch {
      // An invalid or unsupported language tag: Word falls back to its UI language.
      f = new Intl.DateTimeFormat('en-US', o);
    }
    FORMATTERS.set(key, f);
  }
  return f;
}

/** A name part (`month` or `weekday`) of `d` in language `lid`. */
function namePart(d: Date, lid: string, kind: string): string {
  const type = kind.startsWith('ddd') ? 'weekday' : 'month';
  for (const p of formatter(lid, kind).formatToParts(d)) if (p.type === type) return p.value;
  return '';
}

function pad(n: number, width: number): string {
  return String(n).padStart(width, '0');
}

/**
 * Formats `d` (UTC) with Word date picture `picture` in language `lid` (a BCP 47 tag such as
 * `en-US`, as found in `w:lid`).
 */
export function formatWordDate(d: Date, picture: string, lid = 'en-US'): string {
  const withDay = /(^|[^d])d{1,2}(?!d)/.test(picture.replace(/'[^']*'/g, ''));
  let out = '';
  let i = 0;
  while (i < picture.length) {
    const c = picture[i]!;
    if (c === "'") {
      // A quoted literal; `''` stands for one apostrophe, inside or outside quotes.
      i++;
      if (picture[i] === "'") {
        out += "'";
        i++;
        continue;
      }
      while (i < picture.length) {
        if (picture[i] === "'") {
          if (picture[i + 1] !== "'") break;
          i++;
        }
        out += picture[i];
        i++;
      }
      i++;
      continue;
    }
    if (/^am\/pm/i.test(picture.slice(i, i + 5))) {
      const pm = d.getUTCHours() >= 12;
      const s = pm ? 'PM' : 'AM';
      out += picture[i] === 'a' ? s.toLowerCase() : s;
      i += 5;
      continue;
    }
    let n = 1;
    while (picture[i + n] === c) n++;
    switch (c) {
      case 'd':
        if (n <= 2) out += pad(d.getUTCDate(), n);
        else out += namePart(d, lid, n === 3 ? 'ddd' : 'dddd');
        break;
      case 'M':
        if (n <= 2) out += pad(d.getUTCMonth() + 1, n);
        else out += namePart(d, lid, (n === 3 ? 'MMM' : 'MMMM') + (withDay ? '+d' : ''));
        break;
      case 'y':
      case 'Y':
        out += n <= 2 ? pad(d.getUTCFullYear() % 100, 2) : pad(d.getUTCFullYear(), 4);
        break;
      case 'H':
        out += pad(d.getUTCHours(), Math.min(n, 2));
        break;
      case 'h':
        out += pad(d.getUTCHours() % 12 || 12, Math.min(n, 2));
        break;
      case 'm':
        out += pad(d.getUTCMinutes(), Math.min(n, 2));
        break;
      case 's':
        out += pad(d.getUTCSeconds(), Math.min(n, 2));
        break;
      default:
        out += picture.slice(i, i + n);
    }
    i += n;
  }
  return out;
}

/** `yyyy-mm-ddTHH:MM:SSZ` of `d` (UTC), the form of `w:fullDate`. */
export function isoDateTime(d: Date): string {
  return `${d.toISOString().slice(0, 19)}Z`;
}

/**
 * Parses an ISO 8601 date (`2024-01-15`) or date-time (`2024-01-15T10:30:00`, with optional
 * fraction and `Z`/offset). A date or date-time without a zone is taken as UTC. Undefined when
 * the string is not one.
 */
export function parseIsoDate(s: string): Date | undefined {
  const m =
    /^\s*(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?\s*$/.exec(
      s,
    );
  if (m === null) return undefined;
  const [, y, mo, da, hh, mi, ss, frac, zone] = m;
  const ms = Date.UTC(
    Number(y),
    Number(mo) - 1,
    Number(da),
    Number(hh ?? 0),
    Number(mi ?? 0),
    Number(ss ?? 0),
    frac === undefined ? 0 : Math.round(Number(frac) * 1000),
  );
  const check = new Date(ms);
  if (check.getUTCMonth() !== Number(mo) - 1 || check.getUTCDate() !== Number(da)) return undefined;
  let offset = 0;
  if (zone !== undefined && zone !== 'Z') {
    const sign = zone[0] === '-' ? -1 : 1;
    const digits = zone.slice(1).replace(':', '');
    offset = sign * (Number(digits.slice(0, 2)) * 60 + Number(digits.slice(2))) * 60000;
  }
  const d = new Date(ms - offset);
  return Number.isNaN(d.getTime()) ? undefined : d;
}
