// CONTRACT (fields agent): the field model and the page-information hook layout implements.
import type { Story } from '../wml/story.js';
import type { XElement } from '../xml/dom.js';

/**
 * Page information for PAGE / NUMPAGES / SECTIONPAGES / PAGEREF / TOC page numbers.
 * `turboword/layout` provides one from a pagination (`layout.pages`); callers may supply their own.
 */
export interface PageInfoProvider {
  /** 1-based physical page index where `node` (an element in the document) is laid out. */
  pageOf(node: XElement): number | undefined;
  /** Displayed page number text at `node` (honours pgNumType start/format). */
  pageNumberText?(node: XElement): string | undefined;
  /** Total pages (NUMPAGES). */
  pageCount(): number;
  /** Pages in the section containing `node` (SECTIONPAGES). */
  sectionPageCount?(node: XElement): number | undefined;
}

export interface FieldSwitch {
  /** Switch letter(s) with the backslash, e.g. '\\*', '\\#', '\\@', '\\h', '\\o'. */
  name: string;
  /** Argument, unquoted, when the switch takes one. */
  arg?: string;
}

/** One field occurrence (complex fldChar/instrText fields and w:fldSimple). */
export interface Field {
  readonly kind: 'complex' | 'simple';
  /** Upper-case field type: MERGEFIELD, IF, DATE, TOC, PAGE, REF, SEQ, '=' (formula), ... */
  readonly type: string;
  /** Full instruction text (nested fields' results substituted as they currently read). */
  readonly instruction: string;
  /** Positional arguments after the type (unquoted). */
  readonly args: readonly string[];
  readonly switches: readonly FieldSwitch[];
  readonly story: Story;
  /** fldChar begin element's run (complex) or the w:fldSimple element (simple). */
  readonly start: XElement;
  /** fldChar separate run (complex; undefined when there is no result section). */
  readonly separate: XElement | undefined;
  /** fldChar end run (complex) or the w:fldSimple element. */
  readonly end: XElement;
  /** Current result text. */
  readonly result: string;
  /** Fields nested in this field's instruction. */
  readonly nested: readonly Field[];
  /** `w:fldLock` set. */
  readonly locked: boolean;
  /** `w:dirty` set. */
  readonly dirty: boolean;
  /**
   * The instruction with nested fields shown as `{code}` (Word's Alt+F9 view), e.g.
   * `IF {MERGEFIELD Qty} > 1 "s" ""`. Unlike `instruction` it does not change when nested
   * results do. (Additive; always set by `listFields`.)
   */
  readonly code?: string;
  /** The field this one is nested in (its instruction or its result). (Additive.) */
  readonly parent?: Field;
}

export interface EvalContext {
  /** MERGEFIELD / ASK / FILLIN data. */
  values: Readonly<Record<string, unknown>>;
  now: Date;
  locale: string;
  timeZone?: string;
  pages?: PageInfoProvider;
  /** Evaluates a nested instruction (for IF, = and nested fields). */
  evaluate(instruction: string): string | undefined;
}

/** Custom evaluator: return the result text, or undefined to leave the field unchanged. */
export type FieldEvaluator = (field: Field, ctx: EvalContext) => string | undefined;

export interface UpdateFieldsOptions {
  values?: Readonly<Record<string, unknown>>;
  /** Clock for DATE/TIME (default: no update of DATE/TIME without it — determinism). */
  now?: Date;
  /** BCP 47 locale for number/date formatting. Default 'en-US'. */
  locale?: string;
  timeZone?: string;
  pages?: PageInfoProvider;
  /** Also set w:updateFields in settings so Word refreshes on open. Default false. */
  markForWord?: boolean;
  /** Field types to update (default: all known). */
  types?: readonly string[];
  /** Extra or overriding evaluators by field type. */
  evaluators?: Readonly<Record<string, FieldEvaluator>>;
}

export interface UpdateFieldsResult {
  updated: number;
  /** Fields left as they were (unknown type, locked, no data, no page info). */
  skipped: Field[];
}

export interface TocOptions {
  pages?: PageInfoProvider;
  /** Heading levels, default from the TOC's \o switch or 1-3. */
  levels?: [number, number];
}
