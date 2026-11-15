// turboword/fields: the field index, result formatting, field updates and the TOC.
//
// Everything here is non-destructive: field codes stay exactly as they are, and only the
// result (the runs between `separate` and `end`, or the content of `w:fldSimple`) is
// rewritten. See scan.ts (index), format.ts (switches), evaluate.ts (evaluators), write.ts
// (result writer), update.ts (order of evaluation) and toc.ts (TOC).

import type { WordDocument } from '../wml/document.js';
import type { Story } from '../wml/story.js';
import { formatValue } from './format.js';
import type {
  Field,
  FieldSwitch,
  TocOptions,
  UpdateFieldsOptions,
  UpdateFieldsResult,
} from './types.js';
import { listScans, updateFieldsImpl, updateTocImpl } from './update.js';

export { setUpdateFieldsOnOpen } from './settings.js';
export type * from './types.js';

/** Every field in the target (all stories of a document by default), in document order. */
export function listFields(
  target: WordDocument | Story,
  filter?: { type?: string } | ((f: Field) => boolean),
): Field[] {
  const stories = 'stories' in target ? target.stories() : [target];
  const all = listScans(stories);
  if (filter === undefined) return all;
  if (typeof filter === 'function') return all.filter(filter);
  const type = filter.type?.toUpperCase();
  return type === undefined ? all : all.filter((f) => f.type === type);
}

/** Recomputes field results in place (non-destructive: field codes stay). */
export function updateFields(doc: WordDocument, o?: UpdateFieldsOptions): UpdateFieldsResult {
  return updateFieldsImpl(doc, o);
}

/** Rebuilds every TOC field's result from the headings (and page numbers when `pages` is given). Returns the count. */
export function updateToc(doc: WordDocument, o?: TocOptions): number {
  return updateTocImpl(doc, o);
}

/** Applies \* \# \@ switches to a value (Word's general, numeric and date-time pictures). */
export function formatFieldValue(
  value: unknown,
  switches: readonly FieldSwitch[],
  o?: { locale?: string; timeZone?: string },
): string {
  return formatValue(value, switches, { locale: o?.locale ?? 'en-US', timeZone: o?.timeZone });
}
