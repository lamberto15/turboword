// Document composition: appendDocument, insertDocument, composeDocuments, importBlocks.
//
// Every entry point runs the same pipeline (types.ts documents the steps and their modules):
//   1. clone the selected source blocks (body children, paragraph sectPrs included) into the
//      target at the insertion point;
//   2-11. namespaces, styles, numbering, notes, comments, bookmarks, relationships, ids,
//      sections, settings; then customXml items used by data bindings are brought along.
// Before anything is inserted, the target's id registry is scanned and held (no re-scans for the
// duration of the import): steps compare source ids against what the target already uses, and
// every id they write comes from (or is reserved in) the registry.
//
// The source is never modified: it is forked first when the import must change it (tracked
// changes accepted/rejected) or when it is the target's own package (importing a document into
// itself).
//
// `validate: true` runs validatePackage on both inputs before and on the result after, and
// throws a TurbowordError('VALIDATION') when the result has errors of a kind the inputs did not
// (by count per code). The target is modified either way.

import { TurbowordError, UsageError } from '../core/errors.js';
import type { Part } from '../opc/part.js';
import { acceptAllRevisions, rejectAllRevisions } from '../text/index.js';
import type { WordDocument } from '../wml/document.js';
import { Paragraph } from '../wml/paragraph.js';
import { Table } from '../wml/table.js';
import { type ValidationIssue, validatePackage } from '../wml/validate.js';
import type { XDocument, XElement, XNode } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { importBookmarks } from './bookmarks.js';
import { importComments } from './comments.js';
import {
  createImportContext,
  type ResolvedImportOptions,
  resolveImportOptions,
} from './context.js';
import {
  dropEmptyHalves,
  ensureCellsEndWithParagraph,
  type InsertionPoint,
  insertBlocks,
  resolveInsertionPoint,
  stripSectionProperties,
} from './core-insert.js';
import { reconcileNamespaces } from './namespaces.js';
import { importNotes } from './notes.js';
import { importNumbering } from './numbering.js';
import { type PackageCopier, remapRelationships } from './part-copier.js';
import { renumberIds } from './renumber.js';
import { importSections } from './sections.js';
import { mergeSettings } from './settings.js';
import { importStyles } from './styles.js';
import type { ImportContext, ImportOptions, ImportResult, InsertTarget } from './types.js';

export type * from './types.js';

/** What to import and where (the parts that differ between the entry points). */
interface ImportPlan {
  /** The source elements to clone (in the possibly forked source). */
  select(source: WordDocument): { part: Part; elements: XElement[] };
  /** Where they go (called once, before anything is inserted). */
  locate(ctx: ImportContext): InsertionPoint;
  /** Run importSections (false: paragraph sectPrs are stripped instead). */
  sections: boolean;
  /** SectionPlacement mode for importSections. */
  mode: 'append' | 'insert';
}

/** Appends `source` (all sections, headers/footers, styles, numbering, notes, comments, ...) to `target`. */
export function appendDocument(
  target: WordDocument,
  source: WordDocument,
  o?: ImportOptions,
): ImportResult {
  return runImport(target, source, o, {
    select: bodyBlocks,
    locate: (ctx) => ({
      container: ctx.target.body.element,
      before: null,
      part: ctx.target.mainPart,
      halves: [],
      bodyLevel: true,
    }),
    sections: true,
    mode: 'append',
  });
}

/** Inserts `source` at a placeholder, bookmark, content control, paragraph or story index. */
export function insertDocument(
  target: WordDocument,
  at: InsertTarget,
  source: WordDocument,
  o?: ImportOptions,
): ImportResult {
  return runImport(target, source, o, {
    select: bodyBlocks,
    locate: (ctx) => resolveInsertionPoint(ctx.target, at, ctx.ids),
    sections: true,
    mode: 'insert',
  });
}

/** A new document: the first document with the others appended in order. Inputs are not modified. */
export function composeDocuments(docs: readonly WordDocument[], o?: ImportOptions): WordDocument {
  const first = docs[0];
  if (first === undefined) throw new UsageError('composeDocuments(): no documents');
  const out = first.fork();
  for (let i = 1; i < docs.length; i++) appendDocument(out, docs[i]!, o);
  return out;
}

/**
 * Imports selected blocks of `source` at the end of the target body (or `at`). The blocks keep
 * their styles, lists, notes, comments, bookmarks, images and links; section properties inside
 * them are not imported (the blocks adopt the target's section).
 */
export function importBlocks(
  target: WordDocument,
  source: WordDocument,
  blocks: readonly (Paragraph | Table | XElement)[],
  o?: ImportOptions & { at?: InsertTarget },
): ImportResult {
  const wanted = blocks.map((b) => (b instanceof Paragraph || b instanceof Table ? b.element : b));
  const located = wanted.map((el) => locateInSource(source, el));
  const part = located[0]?.part ?? source.mainPart;
  if (located.some((l) => l.part !== part)) {
    throw new UsageError('importBlocks(): all blocks must come from the same part of the source');
  }
  const at = o?.at;
  return runImport(target, source, o, {
    select: (src) => {
      const p = src.package.part(part.name) ?? src.mainPart;
      const doc = p.xml();
      const elements: XElement[] = [];
      for (const l of located) {
        const el = resolvePath(doc, l.path);
        if (el !== undefined && el.parent !== null) elements.push(el);
      }
      return { part: p, elements };
    },
    locate: (ctx) =>
      at === undefined
        ? {
            container: ctx.target.body.element,
            before: null,
            part: ctx.target.mainPart,
            halves: [],
            bodyLevel: true,
          }
        : resolveInsertionPoint(ctx.target, at, ctx.ids),
    sections: false,
    mode: 'insert',
  });
}

// ---- pipeline -----------------------------------------------------------------------------

/** The source body's block children (everything but the final sectPr). */
function bodyBlocks(src: WordDocument): { part: Part; elements: XElement[] } {
  const body = src.body.element;
  const elements: XElement[] = [];
  for (const k of body.children) {
    if (k.nodeType !== 1) continue;
    if (k.is(NS.w, 'sectPr') && k === body.lastElement()) continue;
    elements.push(k);
  }
  return { part: src.mainPart, elements };
}

function runImport(
  target: WordDocument,
  original: WordDocument,
  o: ImportOptions | undefined,
  plan: ImportPlan,
): ImportResult {
  const options = resolveImportOptions(o);
  const baseline = options.validate
    ? errorCounts([...validatePackage(target), ...validatePackage(original)])
    : undefined;
  const source = prepareSource(target, original, options);
  const ctx = createImportContext(target, source, options, o?.onWarning);
  // Scan the target now, before the source ids arrive, and keep that view for the import.
  const release = ctx.ids.hold();
  let sections = 0;
  let blocks: XElement[];
  try {
    const point = plan.locate(ctx);
    const selected = plan.select(source);
    blocks = selected.elements.map((e) => e.clone());
    if (!plan.sections) stripSectionProperties(blocks);
    insertBlocks(point, blocks);
    const sourcePart = selected.part;

    reconcileNamespaces(ctx, blocks, point.part, sourcePart);
    importStyles(ctx, blocks);
    importNumbering(ctx, blocks);
    importNotes(ctx, blocks);
    importComments(ctx, blocks);
    importBookmarks(ctx, blocks);
    remapRelationships(ctx, sourcePart, point.part, blocks);
    renumberIds(ctx, blocks);
    if (plan.sections) {
      sections = importSections(ctx, blocks, { mode: plan.mode, part: point.part });
    }
    mergeSettings(ctx);
    importBoundCustomXml(ctx, blocks);

    dropEmptyHalves(point.halves);
    ensureCellsEndWithParagraph(point.container);
  } finally {
    release();
  }
  if (baseline !== undefined) checkNoNewErrors(baseline, validatePackage(target));
  return {
    blocks: blocks.filter((b) => b.parent !== null),
    sections,
    styleMap: ctx.styleMap,
    numIdMap: ctx.numIdMap,
    bookmarkRenames: ctx.bookmarkRenames,
    addedParts: ctx.addedParts,
  };
}

/** The source to import from: a fork when it must be changed or is the target itself. */
function prepareSource(
  target: WordDocument,
  original: WordDocument,
  options: ResolvedImportOptions,
): WordDocument {
  const mustFork = options.revisions !== 'keep' || original.package === target.package;
  if (!mustFork) return original;
  const src = original.fork();
  if (options.revisions === 'accept') acceptAllRevisions(src);
  else if (options.revisions === 'reject') rejectAllRevisions(src);
  return src;
}

/** Brings the customXml items that imported data bindings (`w:dataBinding`) point at. */
function importBoundCustomXml(ctx: ImportContext, blocks: readonly XElement[]): void {
  const ids = new Set<string>();
  const scan = (root: XElement): void => {
    for (const b of root.descendants(NS.w, 'dataBinding')) {
      const id = b.attrNS(NS.w, 'storeItemID');
      if (id !== undefined && id !== '') ids.add(id.toUpperCase());
    }
  };
  for (const b of blocks) scan(b);
  for (const name of ctx.addedParts) {
    const part = ctx.target.package.part(name);
    const ct = part?.contentType ?? '';
    if (part === undefined || !(ct.endsWith('header+xml') || ct.endsWith('footer+xml'))) continue;
    scan(part.xml().root);
  }
  const copier = ctx.copier as PackageCopier;
  for (const id of ids) copier.copyCustomXmlItem(id);
}

// ---- importBlocks helpers -----------------------------------------------------------------

/** The part of `doc` whose DOM holds `el`, and the child-index path from the document to it. */
function locateInSource(doc: WordDocument, el: XElement): { part: Part; path: number[] } {
  const owner = el.ownerDocument();
  if (owner === undefined)
    throw new UsageError('importBlocks(): a block is not attached to a document');
  const part = doc.package.parts().find((p) => p._doc === owner);
  if (part === undefined)
    throw new UsageError('importBlocks(): a block does not belong to the source');
  const path: number[] = [];
  let n: XNode = el;
  while (n.parent !== null) {
    path.unshift(n.parent._children.indexOf(n));
    if (n.parent.nodeType !== 1) break;
    n = n.parent;
  }
  return { part, path };
}

function resolvePath(doc: XDocument, path: readonly number[]): XElement | undefined {
  let kids: readonly XNode[] = doc._children;
  let node: XNode | undefined;
  for (const i of path) {
    node = kids[i];
    if (node === undefined || node.nodeType !== 1) return undefined;
    kids = node._children;
  }
  return node?.nodeType === 1 ? node : undefined;
}

// ---- validation ---------------------------------------------------------------------------

function errorCounts(issues: readonly ValidationIssue[]): Map<string, number> {
  const m = new Map<string, number>();
  for (const i of issues) if (i.severity === 'error') m.set(i.code, (m.get(i.code) ?? 0) + 1);
  return m;
}

function checkNoNewErrors(
  baseline: ReadonlyMap<string, number>,
  after: readonly ValidationIssue[],
): void {
  const counts = errorCounts(after);
  const fresh: ValidationIssue[] = [];
  for (const [code, n] of counts) {
    if (n > (baseline.get(code) ?? 0))
      fresh.push(...after.filter((i) => i.code === code && i.severity === 'error'));
  }
  if (fresh.length === 0) return;
  const lines = fresh.map((e) => `  ${e.code}${e.part ? ` (${e.part})` : ''}: ${e.message}`);
  throw new TurbowordError(
    'VALIDATION',
    `The composed package has ${fresh.length} new validation error(s):\n${lines.join('\n')}`,
  );
}
