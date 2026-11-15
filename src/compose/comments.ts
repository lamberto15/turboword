// importComments: the comments anchored in the imported blocks, across all five comment parts.
//
// Word spreads a comment over up to five parts, all related from the main document:
//   comments.xml            w:comment (w:id, author, the comment's paragraphs)
//   commentsExtended.xml    w15:commentEx   paraId (= the comment's LAST paragraph's w14:paraId),
//                                           paraIdParent (reply threading), done
//   commentsIds.xml         w16cid:commentId   paraId -> durableId
//   commentsExtensible.xml  w16cex:commentExtensible   durableId -> dateUtc, ...
//   people.xml              w15:person   author -> presence info
// They must be remapped together: the comment gets a fresh w:id (rewritten in every
// commentRangeStart/End/commentReference of the blocks), importContent may renumber the
// comment's paragraph ids (so the paraId keys of the extended parts follow), a durableId already
// used in the target gets a fresh one (and the extensible entry follows), and authors are
// deduplicated in people.xml. Replies (paraIdParent chains) come along with their parent.
// With `comments: 'drop'` the markers are removed instead and nothing is copied.

import type { Part } from '../opc/part.js';
import { CT, REL } from '../opc/rel-types.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { importContent } from './content.js';
import {
  addMainPart,
  adoptNamespaces,
  CT_COMMENTS_EXTENDED,
  CT_COMMENTS_EXTENSIBLE,
  CT_COMMENTS_IDS,
  CT_PEOPLE,
  importedExtras,
  newRoot,
  partRoot,
  REL_COMMENTS_EXTENDED,
  REL_COMMENTS_EXTENSIBLE,
  REL_COMMENTS_IDS,
  REL_PEOPLE,
  removeElement,
  removeIfEmptyRun,
} from './parts-util.js';
import type { ImportContext } from './types.js';

const MARKERS: ReadonlySet<string> = new Set([
  'commentRangeStart',
  'commentRangeEnd',
  'commentReference',
]);

/** One comment part, as found in a document or created in the target. */
interface CommentPart {
  readonly part: Part;
  readonly root: XElement;
}

interface ExtPartSpec {
  readonly rel: string;
  readonly contentType: string;
  readonly uri: string;
  readonly local: string;
  readonly file: string;
}

const EXTENDED: ExtPartSpec = {
  rel: REL_COMMENTS_EXTENDED,
  contentType: CT_COMMENTS_EXTENDED,
  uri: NS.w15,
  local: 'commentsEx',
  file: 'commentsExtended',
};
const IDS: ExtPartSpec = {
  rel: REL_COMMENTS_IDS,
  contentType: CT_COMMENTS_IDS,
  uri: NS.w16cid,
  local: 'commentsIds',
  file: 'commentsIds',
};
const EXTENSIBLE: ExtPartSpec = {
  rel: REL_COMMENTS_EXTENSIBLE,
  contentType: CT_COMMENTS_EXTENSIBLE,
  uri: NS.w16cex,
  local: 'commentsExtensible',
  file: 'commentsExtensible',
};
const PEOPLE: ExtPartSpec = {
  rel: REL_PEOPLE,
  contentType: CT_PEOPLE,
  uri: NS.w15,
  local: 'people',
  file: 'people',
};

/**
 * Copies the comments referenced in `blocks` (and their replies) into the target's comment
 * parts, filling `ctx.commentIdMap`, or removes the markers with `options.comments === 'drop'`.
 */
export function importComments(ctx: ImportContext, blocks: XElement[]): void {
  const markers: XElement[] = [];
  for (const b of blocks) {
    if (isMarker(b)) markers.push(b);
    for (const d of b.descendants(NS.w)) if (MARKERS.has(d.local)) markers.push(d);
  }
  if (markers.length === 0) return;
  if (ctx.options.comments === 'drop') {
    for (const m of markers) dropMarker(blocks, m);
    return;
  }

  const source = existingPart(ctx.source.mainPart, 'comments');
  const byId = new Map<string, XElement>();
  if (source !== undefined) {
    for (const c of source.root.elements(NS.w, 'comment')) {
      const id = c.attrNS(NS.w, 'id');
      if (id !== undefined && !byId.has(id)) byId.set(id, c);
    }
  }

  // Which source comments to copy: the referenced ones plus replies to them (any depth).
  const wanted = new Set<XElement>();
  for (const m of markers) {
    const c = byId.get(m.attrNS(NS.w, 'id') ?? '');
    if (c !== undefined) wanted.add(c);
  }
  const sourceExtended = existingPart(ctx.source.mainPart, 'commentsExtended');
  if (source !== undefined && sourceExtended !== undefined)
    addReplies(source.root, sourceExtended.root, wanted);

  if (source !== undefined && wanted.size > 0) {
    copyComments(ctx, source, sourceExtended, wanted);
  }

  let missing = 0;
  for (const m of markers) {
    const id = Number.parseInt(m.attrNS(NS.w, 'id') ?? '', 10);
    const fresh = Number.isInteger(id) ? ctx.commentIdMap.get(id) : undefined;
    if (fresh === undefined) {
      missing++;
      dropMarker(blocks, m);
    } else {
      m.setAttrNS(NS.w, 'id', String(fresh));
    }
  }
  if (missing > 0) {
    ctx.warn(
      'COMMENT_MISSING',
      `${missing} comment marker(s) refer to comments the source does not have; removed`,
      source?.part.name,
    );
  }
}

function isMarker(el: XElement): boolean {
  return el.uri === NS.w && MARKERS.has(el.local);
}

/** Removes a comment marker (and the run a commentReference leaves empty). */
function dropMarker(blocks: XElement[], m: XElement): void {
  const run = m.parent?.nodeType === 1 ? m.parent : undefined;
  removeElement(blocks, m);
  if (m.local === 'commentReference') removeIfEmptyRun(run);
}

function existingPart(main: Part, kind: string): CommentPart | undefined {
  const part = main.relatedByType(kind)[0];
  const root = part === undefined ? undefined : partRoot(part);
  return part === undefined || root === undefined ? undefined : { part, root };
}

/** `w14:paraId` of the last paragraph of a comment (the key of its extended entries). */
function lastParaId(comment: XElement): string | undefined {
  const ps = comment.descendants(NS.w, 'p');
  return ps.at(-1)?.attrNS(NS.w14, 'paraId');
}

/** Adds to `wanted` every comment whose paraIdParent chain reaches a wanted comment. */
function addReplies(comments: XElement, extended: XElement, wanted: Set<XElement>): void {
  const byParaId = new Map<string, XElement>();
  for (const c of comments.elements(NS.w, 'comment')) {
    const p = lastParaId(c);
    if (p !== undefined) byParaId.set(p.toUpperCase(), c);
  }
  const parentOf = new Map<XElement, XElement>();
  for (const ex of extended.elements(NS.w15, 'commentEx')) {
    const child = byParaId.get((ex.attrNS(NS.w15, 'paraId') ?? '').toUpperCase());
    const parent = byParaId.get((ex.attrNS(NS.w15, 'paraIdParent') ?? '').toUpperCase());
    if (child !== undefined && parent !== undefined) parentOf.set(child, parent);
  }
  for (let changed = true; changed; ) {
    changed = false;
    for (const [child, parent] of parentOf) {
      if (!wanted.has(child) && wanted.has(parent)) {
        wanted.add(child);
        changed = true;
      }
    }
  }
}

function copyComments(
  ctx: ImportContext,
  source: CommentPart,
  sourceExtended: CommentPart | undefined,
  wanted: Set<XElement>,
): void {
  const target = ensurePart(ctx, 'comments', {
    rel: REL.comments,
    contentType: CT.comments,
    uri: NS.w,
    local: 'comments',
    file: 'comments',
  });

  // Copy in source order; remember each copied paragraph's paraId before importContent.
  const copies: XElement[] = [];
  const paras: [XElement, string][] = [];
  const authors = new Set<string>();
  for (const c of source.root.elements(NS.w, 'comment')) {
    if (!wanted.has(c)) continue;
    const id = Number.parseInt(c.attrNS(NS.w, 'id') ?? '', 10);
    if (!Number.isInteger(id)) continue;
    const fresh = ctx.ids.next('comment');
    ctx.commentIdMap.set(id, fresh);
    const copy = c.clone();
    copy.setAttrNS(NS.w, 'id', String(fresh));
    const author = c.attrNS(NS.w, 'author');
    if (author !== undefined) authors.add(author);
    for (const p of copy.descendants(NS.w, 'p')) {
      const paraId = p.attrNS(NS.w14, 'paraId');
      if (paraId !== undefined) paras.push([p, paraId]);
    }
    target.root.append(copy);
    copies.push(copy);
  }
  if (copies.length === 0) return;
  importContent(ctx, copies, source.part, target.part);
  importedExtras(ctx).push(...copies);

  const paraMap = new Map<string, string>();
  for (const [p, old] of paras) {
    paraMap.set(old.toUpperCase(), p.attrNS(NS.w14, 'paraId') ?? old);
  }

  if (sourceExtended !== undefined) copyExtended(ctx, sourceExtended, paraMap);
  const durableMap = copyIds(ctx, paraMap);
  if (durableMap.size > 0) copyExtensible(ctx, durableMap);
  copyPeople(ctx, authors);
}

/** commentsExtended: entries keyed by (remapped) paraId; reply links follow the map. */
function copyExtended(
  ctx: ImportContext,
  source: CommentPart,
  paraMap: ReadonlyMap<string, string>,
): void {
  let target: CommentPart | undefined;
  for (const ex of source.root.elements(NS.w15, 'commentEx')) {
    const paraId = paraMap.get((ex.attrNS(NS.w15, 'paraId') ?? '').toUpperCase());
    if (paraId === undefined) continue;
    target ??= ensurePart(ctx, 'commentsExtended', EXTENDED);
    const copy = ex.clone();
    copy.setAttrNS(NS.w15, 'paraId', paraId);
    const parent = ex.attrNS(NS.w15, 'paraIdParent');
    if (parent !== undefined) {
      const mapped = paraMap.get(parent.toUpperCase());
      if (mapped === undefined) copy.removeAttrNS(NS.w15, 'paraIdParent');
      else copy.setAttrNS(NS.w15, 'paraIdParent', mapped);
    }
    target.root.append(copy);
    adoptNamespaces(copy, target.root);
  }
}

/**
 * commentsIds: entries keyed by (remapped) paraId; a durableId already used in the target gets
 * a fresh one. Returns source durableId → target durableId (upper-case keys).
 */
function copyIds(ctx: ImportContext, paraMap: ReadonlyMap<string, string>): Map<string, string> {
  const durableMap = new Map<string, string>();
  const source = existingPart(ctx.source.mainPart, 'commentsIds');
  if (source === undefined) return durableMap;
  let target: CommentPart | undefined;
  let used: Set<string> | undefined;
  for (const entry of source.root.elements(NS.w16cid, 'commentId')) {
    const paraId = paraMap.get((entry.attrNS(NS.w16cid, 'paraId') ?? '').toUpperCase());
    if (paraId === undefined) continue;
    target ??= ensurePart(ctx, 'commentsIds', IDS);
    used ??= targetDurableIds(ctx);
    const copy = entry.clone();
    copy.setAttrNS(NS.w16cid, 'paraId', paraId);
    const durable = entry.attrNS(NS.w16cid, 'durableId');
    if (durable !== undefined) {
      const key = durable.toUpperCase();
      let fresh = durableMap.get(key);
      if (fresh === undefined) {
        if (used.has(key) || ctx.ids.has('durableId', durable)) {
          fresh = ctx.ids.nextHex('durableId');
        } else {
          fresh = durable;
          ctx.ids.reserve('durableId', durable);
        }
        used.add(fresh.toUpperCase());
        durableMap.set(key, fresh);
      }
      copy.setAttrNS(NS.w16cid, 'durableId', fresh);
    }
    target.root.append(copy);
    adoptNamespaces(copy, target.root);
  }
  return durableMap;
}

/** durableIds already present in the target's commentsIds/commentsExtensible parts. */
function targetDurableIds(ctx: ImportContext): Set<string> {
  const out = new Set<string>();
  const ids = existingPart(ctx.target.mainPart, 'commentsIds');
  if (ids !== undefined) {
    for (const e of ids.root.elements(NS.w16cid, 'commentId')) {
      const v = e.attrNS(NS.w16cid, 'durableId');
      if (v !== undefined) out.add(v.toUpperCase());
    }
  }
  const ext = existingPart(ctx.target.mainPart, 'commentsExtensible');
  if (ext !== undefined) {
    for (const e of ext.root.elements(NS.w16cex, 'commentExtensible')) {
      const v = e.attrNS(NS.w16cex, 'durableId');
      if (v !== undefined) out.add(v.toUpperCase());
    }
  }
  return out;
}

/** commentsExtensible: entries keyed by durableId, renamed through `durableMap`. */
function copyExtensible(ctx: ImportContext, durableMap: ReadonlyMap<string, string>): void {
  const source = existingPart(ctx.source.mainPart, 'commentsExtensible');
  if (source === undefined) return;
  let target: CommentPart | undefined;
  for (const entry of source.root.elements(NS.w16cex, 'commentExtensible')) {
    const fresh = durableMap.get((entry.attrNS(NS.w16cex, 'durableId') ?? '').toUpperCase());
    if (fresh === undefined) continue;
    target ??= ensurePart(ctx, 'commentsExtensible', EXTENSIBLE);
    const copy = entry.clone();
    copy.setAttrNS(NS.w16cex, 'durableId', fresh);
    target.root.append(copy);
    adoptNamespaces(copy, target.root);
  }
}

/** people.xml: the copied comments' authors the target does not list yet. */
function copyPeople(ctx: ImportContext, authors: ReadonlySet<string>): void {
  const source = existingPart(ctx.source.mainPart, 'people');
  if (source === undefined || authors.size === 0) return;
  const existing = existingPart(ctx.target.mainPart, 'people');
  const known = new Set<string>();
  if (existing !== undefined) {
    for (const p of existing.root.elements(NS.w15, 'person')) {
      const a = p.attrNS(NS.w15, 'author');
      if (a !== undefined) known.add(a);
    }
  }
  let target = existing;
  for (const person of source.root.elements(NS.w15, 'person')) {
    const author = person.attrNS(NS.w15, 'author');
    if (author === undefined || !authors.has(author) || known.has(author)) continue;
    known.add(author);
    target ??= ensurePart(ctx, 'people', PEOPLE);
    const copy = person.clone();
    target.root.append(copy);
    adoptNamespaces(copy, target.root);
  }
}

/** The target's comment part of `kind`, created (empty root) when missing. */
function ensurePart(ctx: ImportContext, kind: string, spec: ExtPartSpec): CommentPart {
  const existing = existingPart(ctx.target.mainPart, kind);
  if (existing !== undefined) return existing;
  const root = newRoot(ctx.target, spec.uri, spec.local);
  const part = addMainPart(
    ctx,
    `word/${spec.file}.xml`,
    `word/${spec.file}{n}.xml`,
    spec.contentType,
    spec.rel,
    root,
  );
  return { part, root };
}
