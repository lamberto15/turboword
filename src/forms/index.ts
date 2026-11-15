// turboword/forms: content controls (structured document tags), data binding and legacy form
// fields.
//
//   contentControls(doc)              every `w:sdt`, typed, with a setValue that writes like Word
//   fillContentControls(doc, values)  by tag or alias; arrays of records fill repeating sections
//   bindContentControls(doc, xml)     write the custom XML store, refresh every bound control
//   legacyFormFields / fillLegacyFormFields   FORMTEXT, FORMCHECKBOX, FORMDROPDOWN
//
// The modules: sdt.ts (discovery, reading, setValue), content.ts (rewriting content the way Word
// does), fill.ts, bind.ts with xpath.ts and customxml.ts, legacy.ts, date-format.ts (Word date
// pictures), ids.ts (fresh ids for cloned items).

import type { WordDocument } from '../wml/document.js';
import { bind } from './bind.js';
import { fill } from './fill.js';
import { fillLegacy, legacyFields } from './legacy.js';
import { allControls, FormsContext } from './sdt.js';
import type { ContentControl, FillResult, FillValue, LegacyFormField } from './types.js';

export { formatWordDate } from './date-format.js';
export type {
  ContentControl,
  ContentControlType,
  ContentControlValue,
  FillRecord,
  FillResult,
  FillValue,
  LegacyFormField,
} from './types.js';

/**
 * Every content control in the document (all stories), in document order, nested included.
 *
 * Stories are visited body first, then headers and footers, footnotes, endnotes and comments;
 * controls in text boxes are included. Each control is a live view of its `w:sdt`.
 */
export function contentControls(doc: WordDocument): ContentControl[] {
  return allControls(new FormsContext(doc));
}

/**
 * Fills controls by tag (default) or alias. Repeating sections accept arrays of records.
 *
 * - A scalar sets every control with that key ({@link ContentControl.setValue} rules); a
 *   content-locked control throws a `UsageError`.
 * - An array fills a repeating section: the first item is cloned once per record (fresh ids)
 *   and the clone's nested controls are filled from the record by the same key kind (nested
 *   arrays fill nested sections). The section is unbound from custom XML; an empty array
 *   removes it.
 * - `removeControls`: every control that was filled (sections, their items and the nested
 *   controls included) is replaced by its content.
 *
 * `filled` counts the controls set (a repeating section counts once, plus its nested controls);
 * `missing` lists the keys that matched no control.
 */
export function fillContentControls(
  doc: WordDocument,
  values: Readonly<Record<string, FillValue>>,
  o?: { by?: 'tag' | 'alias'; removeControls?: boolean },
): FillResult {
  return fill(doc, values, o);
}

/**
 * Writes `xml` as (or into) the custom XML part bound by the controls' w:dataBinding storeItemID
 * (created if missing) and refreshes every bound control from it. Returns the controls updated.
 *
 * The store is `o.storeItemId`, else the one the bound controls name (with several, the one whose
 * current XML has the same root element), else a new store whose GUID is derived from the XML
 * content (deterministic). Repeating sections bound to the store are expanded to one item per
 * record first. Paths are evaluated with Word's XPath subset (absolute child steps, `[n]`,
 * `@attr`, `text()`, prefixes from `w:prefixMappings`).
 */
export function bindContentControls(
  doc: WordDocument,
  xml: string,
  o?: { storeItemId?: string },
): number {
  return bind(doc, xml, o);
}

/** Legacy form fields (FORMTEXT, FORMCHECKBOX, FORMDROPDOWN), all stories, document order. */
export function legacyFormFields(doc: WordDocument): LegacyFormField[] {
  return legacyFields(doc);
}

/**
 * Fills legacy form fields by name (`w:ffData/w:name`, which Word also gives the field's
 * bookmark). Text fields honour `w:maxLength` (longer text throws a `UsageError`) and the
 * case/date `w:format`; checkboxes take booleans; dropdowns take one of their entries.
 */
export function fillLegacyFormFields(
  doc: WordDocument,
  values: Readonly<Record<string, string | boolean>>,
): FillResult {
  return fillLegacy(doc, values);
}
