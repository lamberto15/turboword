// Resolved formatting: the public entry points of the style cascade (ECMA-376 §17.7).
//
// Each function takes a wml view (Paragraph, Run, Table, TableCell) and returns plain data
// with everything resolved: document defaults, table style conditional formatting, style
// chains, numbering, direct formatting, theme fonts and colours. Results are cached per
// document and invalidated through `XDocument.version` of the parts involved (styles,
// numbering, theme, settings, the main part's relationships and the story part); see
// context.ts. Lengths are twips, font sizes points (see types.ts).

import type { WordDocument } from '../wml/document.js';
import type { Paragraph, Run } from '../wml/paragraph.js';
import type { Table, TableCell } from '../wml/table.js';
import { NS } from '../xml/ns.js';
import { contextOf } from './context.js';
import { formatNumberValue } from './format.js';
import { flowCounters } from './labels.js';
import { labelRunProps, paragraphResult, resultForProps, runProps } from './paragraph.js';
import { cellProps, tableInfo, tableProps } from './table.js';
import { themeInfo } from './theme.js';
import type {
  ListInfo,
  ResolvedCellProps,
  ResolvedParagraphProps,
  ResolvedRunProps,
  ResolvedTableProps,
  ThemeInfo,
} from './types.js';

export type * from './types.js';

/**
 * Fully resolved paragraph properties (inside a table: including the table style). The object
 * is cached and shared until the document changes: treat it as read-only.
 */
export function resolveParagraph(p: Paragraph): ResolvedParagraphProps {
  return paragraphResult(contextOf(p.story.doc), p.element, p.story).props;
}

/** Fully resolved properties of a run (`para` = its paragraph's resolved props, computed if omitted). */
export function resolveRun(r: Run, para?: ResolvedParagraphProps): ResolvedRunProps {
  const story = r.paragraph.story;
  const ctx = contextOf(story.doc);
  const known = para === undefined ? undefined : resultForProps(para);
  const result = known ?? paragraphResult(ctx, r.paragraph.element, story);
  return runProps(ctx, r.element, result);
}

/** Resolved table properties. */
export function resolveTable(t: Table): ResolvedTableProps {
  return tableProps(tableInfo(contextOf(t.story.doc), t.element, t.story));
}

/** Resolved cell properties incl. table-style conditional formatting and border conflicts. */
export function resolveCell(cell: TableCell): ResolvedCellProps {
  const ctx = contextOf(cell.story.doc);
  const tbl = cell.element.closest(NS.w, 'tbl');
  if (tbl === undefined) {
    return {
      margins: { top: 0, left: 0, bottom: 0, right: 0 },
      borders: {},
      vAlign: 'top',
      noWrap: false,
      conditional: [],
    };
  }
  return cellProps(ctx, tableInfo(ctx, tbl, cell.story), cell.element);
}

/** List membership and rendered label of a paragraph (counters computed over its story). */
export function listInfo(p: Paragraph): ListInfo | undefined {
  const story = p.story;
  const ctx = contextOf(story.doc);
  const result = paragraphResult(ctx, p.element, story);
  if (result.list === undefined) return undefined;
  const entry = flowCounters(ctx, p.element, story).entries.get(p.element);
  if (entry === undefined) return undefined;
  const level = entry.eff.level;
  return {
    numId: entry.eff.numId,
    abstractNumId: entry.eff.abstractId,
    level: entry.ilvl,
    label: entry.label,
    format: level.numFmt,
    isBullet: level.numFmt === 'bullet',
    suffix: level.suff,
    alignment: level.lvlJc,
    runProps: labelRunProps(ctx, result, entry.eff),
    counters: [...entry.counters],
  };
}

/** Shorthand for `listInfo(p)?.label`. */
export function listLabel(p: Paragraph): string | undefined {
  const story = p.story;
  const ctx = contextOf(story.doc);
  if (paragraphResult(ctx, p.element, story).list === undefined) return undefined;
  return flowCounters(ctx, p.element, story).entries.get(p.element)?.label;
}

/** Theme fonts and colors of the document (empty when it has no theme part). */
export function themeOf(doc: WordDocument): ThemeInfo {
  const ctx = contextOf(doc);
  return themeInfo(ctx.theme, ctx.settings);
}

/** Formats a counter value with a `w:numFmt` (decimal, upperRoman, lowerLetter, ordinal, ...). */
export function formatNumber(value: number, format: string): string {
  return formatNumberValue(value, format);
}
