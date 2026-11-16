// turboword/mailmerge: Word's mail merge ("Finish & Merge" → "Edit individual documents")
// without Word. A template with merge fields plus records gives one document per record (or one
// document holding every record).
//
//   mergeRecord(template, record)   → one document (undefined when SKIPIF drops the record)
//   mergeEach(template, records)    → a lazy generator, one document per output
//   mergeAll(template, records)     → one document: every output appended in order
//
// The template is never modified: each output starts as `template.fork()` (copy-on-write, so a
// template parsed once merges thousands of records). Per output document:
//   1. Array values repeat table rows: a row holding `MERGEFIELD <key>` is cloned once per item
//      (fresh ids, see Table.cloneRow), and the fields in each copy read the item over the record.
//      Items that hold arrays expand nested tables the same way. An empty array removes the row.
//   2. The field engine (turboword/fields) evaluates every field in document order, body first:
//      MERGEFIELD (with its \b \f \* \# \@ switches), IF, COMPARE, formulas, ASK, FILLIN, SET,
//      REF... NEXT moves on to the next record, NEXTIF does when its condition holds, SKIPIF drops
//      the output document and starts the next one at the following record, MERGEREC and
//      MERGESEQ give the record and sequence numbers. Headers, footers, notes and comments read
//      the output's first record. Once NEXT runs past the last record, merge fields are empty
//      (as in Word).
//   3. With `fields: 'unlink'` (default) the evaluated merge fields become plain text, as in
//      Word's merged documents: MERGEFIELD, IF, NEXT, NEXTIF, SKIPIF, MERGEREC, MERGESEQ and
//      FILLIN. Other fields (PAGE, DATE, TOC, REF, ...) stay live fields with updated results.
//   4. The mail-merge setup (`w:mailMerge`: data source and query) is removed, so Word opens the
//      output as a plain document instead of asking to run the query.
//
// mergeAll appends the outputs with turboword/compose: each record keeps its own sections,
// headers and footers (the default separator 'section'), lists restart per record, and ids,
// bookmarks and notes are made unique. 'pageBreak' and 'none' put the records in one section
// (the default for Word "Directory"/catalog documents is 'none').

import { appendDocument, importBlocks } from '../compose/index.js';
import { TurbowordError, UsageError } from '../core/errors.js';
import type { MergeState, UpdateState } from '../fields/evaluate.js';
import { lookup } from '../fields/evaluate.js';
import { attached, type Scan, scanStory } from '../fields/scan.js';
import { mailMergeType, removeMailMerge } from '../fields/settings.js';
import type { UpdateFieldsOptions } from '../fields/types.js';
import { runUpdate } from '../fields/update.js';
import { unlinkField } from '../fields/write.js';
import type { WordDocument } from '../wml/document.js';
import { DocumentIds } from '../wml/ids.js';
import { wElement } from '../wml/order.js';
import type { SectionType } from '../wml/section.js';
import type { Story } from '../wml/story.js';
import { Table } from '../wml/table.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';

/** One record of merge data: field name → value. Arrays of records repeat table rows. */
export type MergeRecord = Readonly<Record<string, unknown>>;

/** How mergeAll separates the records. */
export type MergeSeparator = 'section' | SectionType | 'pageBreak' | 'none';

export interface MergeOptions
  extends Pick<UpdateFieldsOptions, 'now' | 'locale' | 'timeZone' | 'pages' | 'evaluators'> {
  /**
   * 'unlink' (default): merged fields become plain text, as in Word's merged documents.
   * 'keep': they stay live fields showing the merged values.
   */
  fields?: 'unlink' | 'keep';
  /**
   * A MERGEFIELD the record lacks: 'empty' (default) merges it as '' (with its \b/\f text
   * dropped, as Word does for an empty value); 'keep' leaves the field unmerged; 'error' throws
   * a TurbowordError with code 'MERGE_MISSING_FIELD' naming the fields.
   */
  missing?: 'empty' | 'keep' | 'error';
  /**
   * mergeAll only. 'section' (default; 'none' for a catalog/directory template) starts each record
   * on a new page with a section break, as Word's merged letters do, so every record keeps its
   * headers and footers (an 'oddPage' or 'evenPage' start of the template's first section is
   * kept, for duplex printing); a section type ('nextPage', 'oddPage', 'evenPage', 'continuous',
   * 'nextColumn') forces that break; 'pageBreak' puts a page break between records in one
   * section; 'none' runs them on.
   */
  separator?: MergeSeparator;
}

/** One output document of mergeEach. */
export interface MergedDocument {
  doc: WordDocument;
  /** 0-based indices of the records merged into it (more than one with NEXT/NEXTIF). */
  records: number[];
  /** Its MERGESEQ: 1-based position among the documents merged so far. */
  sequence: number;
  /** MERGEFIELD names the records lacked (sorted). */
  missing: string[];
}

/** Field types that become plain text in a merged document (fields: 'unlink'). */
const UNLINK: ReadonlySet<string> = new Set([
  'MERGEFIELD',
  'IF',
  'NEXT',
  'NEXTIF',
  'SKIPIF',
  'MERGEREC',
  'MERGESEQ',
  'FILLIN',
]);

const SECTION_TYPES: ReadonlySet<string> = new Set([
  'nextPage',
  'oddPage',
  'evenPage',
  'continuous',
  'nextColumn',
]);

/** Thrown by SKIPIF through the field engine; caught here. */
class SkipRecord extends Error {}

/** The template merged with one record: the document, or undefined when SKIPIF dropped it. */
export function mergeRecord(
  template: WordDocument,
  record: MergeRecord,
  o?: MergeOptions,
): WordDocument | undefined {
  for (const m of mergeEach(template, [record], o)) return m.doc;
  return undefined;
}

/**
 * One merged document per output, lazily: records are read from `records` only as they are
 * merged, so a generator or a database cursor streams with flat memory.
 */
export function* mergeEach(
  template: WordDocument,
  records: Iterable<MergeRecord>,
  o: MergeOptions = {},
): Generator<MergedDocument, void, undefined> {
  checkOptions(o);
  const it = records[Symbol.iterator]();
  let read = 0;
  const pull = (): { index: number; value: MergeRecord } | undefined => {
    const r = it.next();
    if (r.done === true) return undefined;
    const value: unknown = r.value;
    if (value === null || typeof value !== 'object' || Array.isArray(value)) {
      throw new UsageError(`mailmerge: record ${read} is not an object`);
    }
    return { index: read++, value: value as MergeRecord };
  };
  const opts: UpdateFieldsOptions = {};
  if (o.now !== undefined) opts.now = o.now;
  if (o.locale !== undefined) opts.locale = o.locale;
  if (o.timeZone !== undefined) opts.timeZone = o.timeZone;
  if (o.pages !== undefined) opts.pages = o.pages;
  if (o.evaluators !== undefined) opts.evaluators = o.evaluators;
  let sequence = 0;
  let pending = pull();
  while (pending !== undefined) {
    const first = pending;
    let current = first;
    const consumed = [first.index];
    const doc = template.fork();
    const scopes = new Map<XElement, MergeRecord>();
    expandRows(doc, first.value, scopes);
    let inBody = true;
    const merge: MergeState = {
      scope: (s) => (scopes.size === 0 ? undefined : scopeOf(s, scopes)),
      missing: o.missing === 'keep' ? 'keep' : 'empty',
      missingNames: new Set(),
      exhausted: false,
      next(st: UpdateState) {
        if (!inBody || merge.exhausted) return;
        const n = pull();
        if (n === undefined) {
          merge.exhausted = true;
          st.values = {};
          return;
        }
        current = n;
        consumed.push(n.index);
        st.values = n.value;
      },
      skip(): never {
        throw new SkipRecord();
      },
      record: () => current.index + 1,
      sequence: () => sequence + 1,
    };
    let run: ReturnType<typeof runUpdate>;
    try {
      run = runUpdate(
        doc,
        { ...opts, values: first.value },
        {
          merge,
          enterStory(story: Story, st: UpdateState) {
            if (story.kind === 'body') return;
            inBody = false;
            merge.exhausted = false;
            st.values = first.value;
          },
        },
      );
    } catch (e) {
      if (!(e instanceof SkipRecord)) throw e;
      pending = merge.exhausted ? undefined : pull();
      continue;
    }
    const missing = [...merge.missingNames].sort();
    if (o.missing === 'error' && missing.length > 0) {
      throw new TurbowordError(
        'MERGE_MISSING_FIELD',
        `mailmerge: record ${first.index} has no value for ${missing.map((n) => `"${n}"`).join(', ')}`,
      );
    }
    if (o.fields !== 'keep') unlinkMerged(run.scans, run.done);
    removeMailMerge(doc);
    sequence++;
    pending = merge.exhausted ? undefined : pull();
    yield { doc, records: consumed, sequence, missing };
  }
}

/**
 * Every output in one document, in record order (see `separator`). Throws a UsageError when no
 * record was merged (no records, or SKIPIF dropped them all).
 */
export function mergeAll(
  template: WordDocument,
  records: Iterable<MergeRecord>,
  o: MergeOptions = {},
): WordDocument {
  checkOptions(o);
  const sep: MergeSeparator =
    o.separator ?? (mailMergeType(template) === 'catalog' ? 'none' : 'section');
  const start: SectionType | undefined =
    sep === 'section'
      ? recordStart(template)
      : SECTION_TYPES.has(sep)
        ? (sep as SectionType)
        : undefined;
  /** Appends `right` (one or more records) to `left`. */
  const join = (left: WordDocument, right: WordDocument): void => {
    if (start !== undefined) {
      appendDocument(left, right, { sectionStart: start, restartListsAtImport: true });
      return;
    }
    const blocks = right.body.element.elements().filter((e) => !e.is(NS.w, 'sectPr'));
    if (blocks.length === 0) return;
    if (sep === 'pageBreak') blocks.splice(0, 1, ...pageBreakBefore(blocks[0]!));
    importBlocks(left, right, blocks, { restartListsAtImport: true });
  };
  // Every append scans the whole target, so appending record after record to one growing
  // document costs O(n²). Runs of equal size are joined pairwise instead (a binary counter, as in
  // merge sort): each record is copied O(log n) times and at most log2(n) runs are held.
  const runs: { doc: WordDocument; size: number }[] = [];
  for (const m of mergeEach(template, records, o)) {
    let top = { doc: m.doc, size: 1 };
    while (runs.length > 0 && runs[runs.length - 1]!.size === top.size) {
      const left = runs.pop()!;
      join(left.doc, top.doc);
      top = { doc: left.doc, size: left.size + top.size };
    }
    runs.push(top);
  }
  const first = runs[0];
  if (first === undefined) {
    throw new UsageError(
      'mergeAll(): no record was merged (no records, or SKIPIF dropped them all)',
    );
  }
  // Remaining runs are in record order with decreasing sizes: join from the right.
  for (let i = runs.length - 1; i > 0; i--) join(runs[i - 1]!.doc, runs[i]!.doc);
  return first.doc;
}

// ---- helpers ------------------------------------------------------------------------------

/**
 * The break before each record for separator 'section': a new page, or the odd/even page the
 * template's first section asks for. (A first section's own 'continuous' or 'nextColumn' start
 * has no effect in the template, where nothing precedes it; between records it would run the
 * letters together.)
 */
function recordStart(template: WordDocument): SectionType {
  const type = template.sections()[0]?.type;
  return type === 'oddPage' || type === 'evenPage' ? type : 'nextPage';
}

function checkOptions(o: MergeOptions): void {
  const bad = (name: string, v: unknown, ok: readonly string[]): void => {
    if (v !== undefined && !ok.includes(v as string)) {
      throw new UsageError(`mailmerge: ${name} must be one of ${ok.join(', ')}`);
    }
  };
  bad('fields', o.fields, ['unlink', 'keep']);
  bad('missing', o.missing, ['empty', 'keep', 'error']);
  bad('separator', o.separator, ['section', ...SECTION_TYPES, 'pageBreak', 'none']);
}

/** The scope of the nearest expanded table row around field `s`. */
function scopeOf(s: Scan, scopes: ReadonlyMap<XElement, MergeRecord>): MergeRecord | undefined {
  let n = s.begin.parent;
  while (n !== null && n.nodeType === 1) {
    if (n.local === 'tr') {
      const v = scopes.get(n);
      if (v !== undefined) return v;
    }
    n = n.parent;
  }
  return undefined;
}

/** Repeats the table rows that array values of `record` anchor (see the file header). */
function expandRows(
  doc: WordDocument,
  record: MergeRecord,
  scopes: Map<XElement, MergeRecord>,
): void {
  if (!Object.values(record).some(Array.isArray)) return;
  const release = DocumentIds.of(doc).hold();
  try {
    for (const story of doc.stories()) expandIn(story, story.element, record, record, scopes);
  } finally {
    release();
  }
}

function expandIn(
  story: Story,
  root: XElement,
  own: MergeRecord,
  base: MergeRecord,
  scopes: Map<XElement, MergeRecord>,
): void {
  for (const [key, value] of Object.entries(own)) {
    if (!Array.isArray(value)) continue;
    for (const tr of anchorRows(story, root, key)) {
      if (!attached(tr, story.element)) continue;
      const tbl = tr.closest(NS.w, 'tbl');
      if (tbl === undefined) continue;
      const table = new Table(tbl, story);
      const index = table.rows().findIndex((r) => r.element === tr);
      if (index < 0) continue;
      if (value.length === 0) {
        table.deleteRow(index);
        continue;
      }
      const rows = [tr];
      for (let i = 1; i < value.length; i++) {
        rows.push(table.cloneRow(index, { after: index + i - 1 }).element);
      }
      rows.forEach((row, i) => {
        const item: unknown = value[i];
        const mine: MergeRecord =
          item !== null && typeof item === 'object' && !Array.isArray(item)
            ? (item as MergeRecord)
            : { [key]: item };
        const scope = { ...base, ...mine };
        scopes.set(row, scope);
        expandIn(story, row, mine, scope, scopes);
      });
    }
  }
}

/** Distinct table rows strictly inside `root` that hold `MERGEFIELD <key>` (nearest row). */
function anchorRows(story: Story, root: XElement, key: string): XElement[] {
  const out: XElement[] = [];
  for (const s of scanStory(story, true)) {
    if (s.parsed.type !== 'MERGEFIELD') continue;
    const name = s.parsed.args[0]?.text;
    if (name === undefined || lookup({ [key]: true }, name) !== true) continue;
    const tr = s.begin.closest(NS.w, 'tr');
    if (tr === undefined || tr === root || !inside(tr, root) || out.includes(tr)) continue;
    out.push(tr);
  }
  return out;
}

function inside(el: XElement, root: XElement): boolean {
  for (let n = el.parent; n !== null && n.nodeType === 1; n = n.parent) if (n === root) return true;
  return false;
}

/** Unlinks the evaluated merge fields, last first (see UNLINK). */
function unlinkMerged(scans: ReadonlyMap<Story, Scan[]>, done: ReadonlySet<Scan>): void {
  for (const [story, list] of scans) {
    for (let i = list.length - 1; i >= 0; i--) {
      const s = list[i]!;
      if (!done.has(s) || !UNLINK.has(s.parsed.type)) continue;
      if (s.parent !== undefined && s.inInstr) continue;
      if (!attached(s.begin, story.element)) continue;
      unlinkField(s);
    }
  }
}

/** `block` with a page break in front: in its first paragraph, or a new paragraph before a table. */
function pageBreakBefore(block: XElement): XElement[] {
  const run = wElement(block, 'r');
  run.append(wElement(block, 'br', { type: 'page' }));
  if (block.is(NS.w, 'p')) {
    const pPr = block.element(NS.w, 'pPr');
    const first = block.children[0];
    if (pPr !== undefined) pPr.after(run);
    else if (first !== undefined) first.before(run);
    else block.append(run);
    return [block];
  }
  const p = wElement(block, 'p');
  p.append(run);
  block.before(p);
  return [p, block];
}
