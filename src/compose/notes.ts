// importNotes: the footnotes and endnotes referenced by the imported blocks.
//
// A `w:footnoteReference`/`w:endnoteReference` points by id at a note in the source's notes part.
// Each referenced note is copied into the target's notes part (created with the separator and
// continuation separator Word requires when the target has none) under a fresh id from the
// document-wide registry, and the reference is rewritten. Separator notes are recognised by
// their `w:type`, never by id: producers number them -1/0, 0/1 or anything else. The copied
// note bodies then go through importContent (styles, numbering, the notes part's own
// relationships, document-wide ids), and are recorded so the bookmark step rewrites them too.

import type { Part } from '../opc/part.js';
import { CT, REL } from '../opc/rel-types.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { importContent } from './content.js';
import {
  addMainPart,
  importedExtras,
  mainRelated,
  newRoot,
  partRoot,
  removeElement,
  removeIfEmptyRun,
  wChild,
} from './parts-util.js';
import type { ImportContext } from './types.js';

/** `w:type` values of notes that are separators, not notes. */
const SEPARATOR_TYPES: ReadonlySet<string> = new Set([
  'separator',
  'continuationSeparator',
  'continuationNotice',
]);

interface NoteKind {
  readonly note: 'footnote' | 'endnote';
  readonly part: 'footnotes' | 'endnotes';
  readonly reference: 'footnoteReference' | 'endnoteReference';
  readonly contentType: string;
  readonly relType: string;
}

const FOOTNOTES: NoteKind = {
  note: 'footnote',
  part: 'footnotes',
  reference: 'footnoteReference',
  contentType: CT.footnotes,
  relType: REL.footnotes,
};

const ENDNOTES: NoteKind = {
  note: 'endnote',
  part: 'endnotes',
  reference: 'endnoteReference',
  contentType: CT.endnotes,
  relType: REL.endnotes,
};

/**
 * Copies the footnotes and endnotes referenced in `blocks` into the target (new ids in
 * `ctx.footnoteIdMap`/`ctx.endnoteIdMap`) and rewrites the references. A reference to a note the
 * source does not have (or to a separator) is removed with a warning.
 */
export function importNotes(ctx: ImportContext, blocks: XElement[]): void {
  importKind(ctx, blocks, FOOTNOTES, ctx.footnoteIdMap);
  importKind(ctx, blocks, ENDNOTES, ctx.endnoteIdMap);
}

function importKind(
  ctx: ImportContext,
  blocks: XElement[],
  kind: NoteKind,
  idMap: Map<number, number>,
): void {
  const refs: XElement[] = [];
  for (const b of blocks) {
    if (b.is(NS.w, kind.reference)) refs.push(b);
    refs.push(...b.descendants(NS.w, kind.reference));
  }
  if (refs.length === 0) return;

  const sourcePart = mainRelated(ctx.source, kind.part);
  const sourceRoot = sourcePart === undefined ? undefined : partRoot(sourcePart);
  const notes = new Map<number, XElement>();
  if (sourceRoot !== undefined) {
    for (const n of sourceRoot.elements(NS.w, kind.note)) {
      const id = noteId(n);
      const type = n.attrNS(NS.w, 'type');
      if (id === undefined || (type !== undefined && SEPARATOR_TYPES.has(type))) continue;
      if (!notes.has(id)) notes.set(id, n);
    }
  }

  let target: { part: Part; root: XElement } | undefined;
  const copies: XElement[] = [];
  for (const ref of refs) {
    const id = noteId(ref);
    const note = id === undefined ? undefined : notes.get(id);
    if (id === undefined || note === undefined) {
      ctx.warn(
        'NOTE_MISSING',
        `${kind.reference} to ${kind.note} ${ref.attrNS(NS.w, 'id') ?? '(no id)'} not found in the source; reference removed`,
        sourcePart?.name,
      );
      const run = ref.parent?.nodeType === 1 ? ref.parent : undefined;
      removeElement(blocks, ref);
      removeIfEmptyRun(run);
      continue;
    }
    let fresh = idMap.get(id);
    if (fresh === undefined) {
      target ??= ensureNotesPart(ctx, kind);
      fresh = ctx.ids.next(kind.note);
      idMap.set(id, fresh);
      const copy = note.clone();
      copy.setAttrNS(NS.w, 'id', String(fresh));
      target.root.append(copy);
      copies.push(copy);
    }
    ref.setAttrNS(NS.w, 'id', String(fresh));
  }
  if (copies.length > 0 && target !== undefined && sourcePart !== undefined) {
    importContent(ctx, copies, sourcePart, target.part);
    importedExtras(ctx).push(...copies);
  }
}

/** Integer `w:id` of a note or reference, or undefined. */
function noteId(el: XElement): number | undefined {
  const v = el.attrNS(NS.w, 'id');
  if (v === undefined || !/^\s*-?\d+\s*$/.test(v)) return undefined;
  return Number.parseInt(v, 10);
}

/** The target notes part and root, created with the two separator notes when missing. */
function ensureNotesPart(ctx: ImportContext, kind: NoteKind): { part: Part; root: XElement } {
  const existing = mainRelated(ctx.target, kind.part);
  const existingRoot = existing === undefined ? undefined : partRoot(existing);
  if (existing !== undefined && existingRoot !== undefined) {
    return { part: existing, root: existingRoot };
  }
  const root = newRoot(ctx.target, NS.w, kind.part);
  root.append(
    separator(root, kind, 'separator', -1),
    separator(root, kind, 'continuationSeparator', 0),
  );
  const part = addMainPart(
    ctx,
    `word/${kind.part}.xml`,
    `word/${kind.part}{n}.xml`,
    kind.contentType,
    kind.relType,
    root,
  );
  ctx.ids.reserve(kind.note, -1);
  ctx.ids.reserve(kind.note, 0);
  return { part, root };
}

/** A separator note as Word writes it (no paragraph ids, so nothing to renumber). */
function separator(
  scope: XElement,
  kind: NoteKind,
  type: 'separator' | 'continuationSeparator',
  id: number,
): XElement {
  const note = wChild(scope, kind.note, { type, id: String(id) });
  const p = wChild(scope, 'p');
  const pPr = wChild(scope, 'pPr');
  pPr.append(wChild(scope, 'spacing', { after: '0', line: '240', lineRule: 'auto' }));
  const r = wChild(scope, 'r');
  r.append(wChild(scope, type));
  p.append(pPr, r);
  note.append(p);
  return note;
}
