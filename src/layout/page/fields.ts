// Page-dependent fields (PAGE, NUMPAGES, SECTIONPAGES, PAGEREF) and page number text.
//
// Page numbers are formatted with the section's `w:pgNumType/@w:fmt` through resolve's
// formatNumber; a field's own `\*` format switch (roman, ALPHABETIC, ...) overrides it.

import { formatNumber } from '../../resolve/index.js';
import type { FieldInline } from './strip.js';

/** `\*` switch names → w:numFmt values. */
const SWITCH_FORMATS: Readonly<Record<string, string>> = {
  arabic: 'decimal',
  roman: 'lowerRoman',
  ROMAN: 'upperRoman',
  alphabetic: 'lowerLetter',
  ALPHABETIC: 'upperLetter',
  ordinal: 'ordinal',
  cardtext: 'cardinalText',
  ordtext: 'ordinalText',
  hex: 'hex',
  arabicdash: 'numberInDash',
};

/** The w:numFmt a field's `\*` switch asks for, or undefined (MERGEFORMAT, Upper, ... or none). */
export function switchFormat(format: string | undefined): string | undefined {
  if (format === undefined) return undefined;
  const f = format.trim();
  return SWITCH_FORMATS[f] ?? SWITCH_FORMATS[f.toLowerCase()];
}

/** Page number text; unsupported formats degrade to decimal. */
export function pageNumberText(value: number, fmt: string | undefined): string {
  if (fmt === undefined || fmt === '' || fmt === 'decimal') return String(value);
  try {
    const s = formatNumber(value, fmt);
    return s === '' ? String(value) : s;
  } catch {
    return String(value);
  }
}

/** What a field needs to know about pagination. */
export interface PageFacts {
  /** Displayed number and its format, section index of each physical page. */
  pages: readonly { value: number; format: string; sectionIndex: number }[];
  /** Bookmark → 0-based page index. */
  bookmarks: ReadonlyMap<string, number>;
}

/** The text of a page-dependent field on page `index`, or undefined to keep its current result. */
export function fieldValue(f: FieldInline, index: number, facts: PageFacts): string | undefined {
  const page = facts.pages[index];
  if (page === undefined) return undefined;
  const own = switchFormat(f.format);
  switch (f.type.toUpperCase()) {
    case 'PAGE':
      return pageNumberText(page.value, own ?? page.format);
    case 'NUMPAGES':
      return pageNumberText(facts.pages.length, own);
    case 'SECTIONPAGES': {
      let n = 0;
      for (const p of facts.pages) if (p.sectionIndex === page.sectionIndex) n++;
      return pageNumberText(n, own);
    }
    case 'PAGEREF': {
      const target = f.arg === undefined ? undefined : facts.bookmarks.get(f.arg);
      const tp = target === undefined ? undefined : facts.pages[target];
      return tp === undefined ? undefined : pageNumberText(tp.value, own ?? tp.format);
    }
    default:
      return undefined;
  }
}
