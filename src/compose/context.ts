// The ImportContext: the mutable state one import (appendDocument, insertDocument, importBlocks)
// shares between its pipeline steps (see types.ts). Built once per import by the orchestrator.
//
// Options are resolved to their documented defaults here, so steps never deal with `undefined`.
// The id registry is the target's DocumentIds; the orchestrator holds it (no re-scans) for the
// duration of the import, after scanning the target BEFORE the source blocks are inserted, so
// `ctx.ids.has()` answers "is this id used by the target (or by something imported already)?"
// and never sees the not-yet-renumbered source ids. Steps that keep a source id must reserve it.

import { type WarningHandler, warning } from '../core/warnings.js';
import type { WordDocument } from '../wml/document.js';
import { DocumentIds } from '../wml/ids.js';
import type { XElement } from '../xml/dom.js';
import { PackageCopier } from './part-copier.js';
import type { ImportContext, ImportOptions } from './types.js';

/** Resolved options (every field set). */
export type ResolvedImportOptions = Readonly<Required<Omit<ImportOptions, 'onWarning'>>>;

/** The documented defaults of {@link ImportOptions}. */
export const DEFAULT_IMPORT_OPTIONS: ResolvedImportOptions = {
  styles: 'useDestinationStyles',
  smartStyleBehavior: true,
  lists: 'keepSource',
  restartListsAtImport: false,
  sectionStart: 'nextPage',
  headersFooters: 'keepSource',
  pageNumbering: 'keepSource',
  keepSourceSections: true,
  theme: 'destination',
  bookmarks: 'rename',
  comments: 'keep',
  revisions: 'keep',
  validate: false,
};

/** `o` with every unset option replaced by its default. */
export function resolveImportOptions(o: ImportOptions | undefined): ResolvedImportOptions {
  const out: Record<string, unknown> = { ...DEFAULT_IMPORT_OPTIONS };
  if (o !== undefined) {
    for (const [k, v] of Object.entries(o)) {
      if (k !== 'onWarning' && v !== undefined) out[k] = v;
    }
  }
  return out as unknown as ResolvedImportOptions;
}

/** Builds the context of one import of `source` into `target`. */
export function createImportContext(
  target: WordDocument,
  source: WordDocument,
  options: ResolvedImportOptions,
  onWarning: WarningHandler | undefined,
): ImportContext {
  const addedParts: string[] = [];
  const warn = (code: string, message: string, part?: string): void => {
    onWarning?.(warning(code, message, part));
  };
  const copier = new PackageCopier(source.package, target.package, {
    added: addedParts,
    warn,
    sourceMain: source.mainPart,
    targetMain: target.mainPart,
  });
  return {
    target,
    source,
    options,
    ids: DocumentIds.of(target),
    copier,
    warn,
    styleMap: new Map<string, string>(),
    numIdMap: new Map<number, number>(),
    bookmarkRenames: new Map<string, string>(),
    commentIdMap: new Map<number, number>(),
    footnoteIdMap: new Map<number, number>(),
    endnoteIdMap: new Map<number, number>(),
    addedParts,
    importedStyles: [] as XElement[],
  };
}
