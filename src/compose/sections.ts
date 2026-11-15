// importSections: section properties and headers/footers of the imported content.
//
// WordprocessingML has no section container: a `w:sectPr` in a paragraph's `w:pPr` ends a
// section and the body's last `w:sectPr` describes the final one. Headers and footers are
// inherited per type (default / first / even) from the previous section when a section does not
// reference one. Appending a document therefore takes more than copying its sectPrs (this
// generalizes turbomerge's `#commitHeaders`):
//   1. The target's body-level sectPr moves onto the target's last paragraph, so the target keeps
//      its final section; the source's final sectPr becomes the body's.
//   2. Every imported sectPr gets its header/footer references remapped: the parts are copied
//      with ctx.copier.copyPart (with their own relationships: images, hyperlinks) and their
//      content made valid through importContent(..., { remapRelationships: false }).
//   3. Inheritance is materialized: for each header/footer type the first imported section does
//      not reference, a reference to a new EMPTY part is added. The source showed nothing there;
//      without it the imported pages would wrongly inherit the target's last header.
//   4. `w:evenAndOddHeaders` is document-wide: when one side needs it and the other did not use
//      it, every section of the other side gets an `even` reference equal to its default one,
//      so its appearance does not change.
//   5. pageNumbering / sectionStart adjust the first imported section; headersFooters
//      'linkToPrevious' drops the references (the target's are inherited), 'none' gives the
//      imported sections empty headers and footers.
// Inserting in the middle of the body splits the host section (its sectPr is copied onto the
// paragraph before the insertion) and pins the host remainder's headers/footers so they do not
// start inheriting from the imported sections. With keepSourceSections false (insertDocument),
// or content landing outside the body (a cell, a header), the source sectPrs are dropped.
//
// Relationship ids: the relationship step (remapRelationships) leaves header/footer references
// alone, so they always carry source ids here; it does remap the other r:* attributes of the
// block sectPrs (`w:printerSettings`). The source's final sectPr is cloned here, so all of its
// relationship ids are the source's and are remapped by this step.

import type { Part } from '../opc/part.js';
import { REL } from '../opc/rel-types.js';
import { createHeaderFooter } from '../wml/header-footer.js';
import { insertOrdered, makeOrder, PPR_ORDER, pruneEmpty } from '../wml/order.js';
import type { HeaderFooterType } from '../wml/section.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { renameBookmarks, rewriteReferences } from './bookmarks.js';
import { importContent } from './content.js';
import {
  adoptNamespaces,
  ensureSettings,
  ensureSettingsChild,
  evenAndOddHeaders,
  firstVisit,
  onOff,
  partRoot,
  relationshipTo,
  relTypeFor,
  removeElement,
  wChild,
} from './parts-util.js';
import type { ImportContext, SectionPlacement } from './types.js';

/** `w:sectPr` children in schema order (CT_SectPr, ECMA-376 Part 1 §17.6.17). */
const SECTPR_ORDER = /* @__PURE__ */ makeOrder([
  'headerReference',
  'footerReference',
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'sectPrChange',
]);

type Kind = 'header' | 'footer';
const KINDS: readonly Kind[] = ['header', 'footer'];
const TYPES: readonly HeaderFooterType[] = ['default', 'first', 'even'];

/** An imported sectPr. */
interface Imported {
  readonly sectPr: XElement;
}

/** Imports section properties and headers/footers; returns the number of sections imported. */
export function importSections(
  ctx: ImportContext,
  blocks: XElement[],
  placement: SectionPlacement,
): number {
  const target = ctx.target;
  const body = bodyOf(ctx);
  // A body-level sectPr among the blocks (one that was not the source body's last child) is
  // misplaced markup: dropped, the source's final sectPr is read from the source below.
  for (const b of [...blocks]) if (b.is(NS.w, 'sectPr')) removeElement(blocks, b);
  const blockSectPrs = paragraphSectPrs(blocks);
  const first0 = blocks[0];
  const inBody =
    body !== undefined &&
    placement.part === target.mainPart &&
    (first0 === undefined ? placement.mode === 'append' : first0.parent === body);

  if (body === undefined || !inBody || !ctx.options.keepSourceSections) {
    for (const s of blockSectPrs) dropSectPr(s);
    if (sourceHasHeadersFooters(ctx)) {
      ctx.warn(
        'SECTIONS_DROPPED',
        inBody
          ? 'keepSourceSections is false: the source section properties and headers/footers were dropped'
          : 'content imported outside the body: the source section properties and headers/footers were dropped',
      );
    }
    return 0;
  }

  const srcFinal = sourceFinalSectPr(ctx);
  const imported: Imported[] = blockSectPrs.map((s) => ({ sectPr: s }));
  const finalImported: Imported = { sectPr: srcFinal ?? wChild(body, 'sectPr') };
  imported.push(finalImported);
  const importedSet = new Set(imported.map((i) => i.sectPr));

  let host: HostSplit | undefined;
  if (placement.mode === 'append') placeAppend(body, blocks, finalImported.sectPr);
  else host = placeInsert(ctx, body, blocks, finalImported.sectPr, importedSet);
  const mainRoot = partRoot(target.mainPart);
  if (mainRoot !== undefined && srcFinal !== undefined) adoptNamespaces(srcFinal, mainRoot);

  const mode = ctx.options.headersFooters;
  const srcEvenOdd = evenAndOddHeaders(ctx.source);
  const tgtEvenOdd = evenAndOddHeaders(target);
  const anyTitlePg = imported.some((i) => onOff(i.sectPr.element(NS.w, 'titlePg')));
  const headerRoots: XElement[] = [];

  if (mode === 'keepSource') remapReferences(ctx, imported, headerRoots);
  else for (const i of imported) removeReferences(i.sectPr);
  if (srcFinal !== undefined) {
    remapOtherRelationships(ctx, srcFinal);
    renumberRevision(ctx, srcFinal);
  }

  // Materialize what the first imported section would otherwise inherit from the target.
  const first = imported[0]!.sectPr;
  if (mode !== 'linkToPrevious') {
    const types: HeaderFooterType[] = ['default'];
    if (anyTitlePg) types.push('first');
    if (mode === 'none' ? tgtEvenOdd : srcEvenOdd) types.push('even');
    for (const kind of KINDS) {
      for (const type of types) {
        if (ownReference(first, kind, type) === undefined) {
          addReference(first, kind, type, emptyPartId(ctx, kind, type));
        }
      }
    }
  }
  if (host !== undefined) pinHost(ctx, host, tgtEvenOdd);
  if (mode === 'keepSource') {
    if (srcEvenOdd && !tgtEvenOdd) {
      const el = ensureSettingsChild(ensureSettings(ctx), 'evenAndOddHeaders');
      el.removeAttrNS(NS.w, 'val');
      mirrorEven(targetSectPrs(ctx).filter((s) => !importedSet.has(s)));
    } else if (tgtEvenOdd && !srcEvenOdd) {
      mirrorEven(imported.map((i) => i.sectPr));
    }
  }

  applyPageNumbering(ctx, first);
  setType(first, ctx.options.sectionStart);

  if (headerRoots.length > 0) {
    renameBookmarks(ctx, headerRoots);
    for (const r of headerRoots) rewriteReferences([r], ctx.bookmarkRenames);
  }
  return imported.length;
}

// ---- placement ------------------------------------------------------------------------------

function bodyOf(ctx: ImportContext): XElement | undefined {
  const el = ctx.target.body.element;
  return el.is(NS.w, 'document') ? el.element(NS.w, 'body') : el;
}

/** sectPrs of body-level paragraphs inside `blocks`, in document order. */
function paragraphSectPrs(blocks: readonly XElement[]): XElement[] {
  const out: XElement[] = [];
  for (const b of blocks) {
    for (const s of b.descendants(NS.w, 'sectPr')) {
      const pPr = s.parent;
      if (pPr?.nodeType === 1 && pPr.is(NS.w, 'pPr') && pPr.parent?.nodeType === 1) {
        if (pPr.parent.is(NS.w, 'p')) out.push(s);
      }
    }
  }
  return out;
}

/** Removes a paragraph's sectPr (and the pPr it leaves empty). */
function dropSectPr(sectPr: XElement): void {
  const pPr = sectPr.parent?.nodeType === 1 ? sectPr.parent : undefined;
  sectPr.remove();
  if (pPr !== undefined) pruneEmpty(pPr);
}

/** A clone of the source body's final sectPr (undefined when the source has an implicit one). */
function sourceFinalSectPr(ctx: ImportContext): XElement | undefined {
  const el = ctx.source.body.element;
  const body = el.is(NS.w, 'document') ? el.element(NS.w, 'body') : el;
  return body?.element(NS.w, 'sectPr')?.clone();
}

/** The paragraph that should end the section closed by a sectPr placed after `before`. */
function holderAfter(body: XElement, before: XElement | undefined): XElement {
  if (before?.is(NS.w, 'p')) {
    const pPr = before.element(NS.w, 'pPr');
    if (pPr?.element(NS.w, 'sectPr') === undefined) return before;
  }
  const p = wChild(body, 'p');
  if (before === undefined) body.prepend(p);
  else body.insertAfter(p, before);
  return p;
}

/** Puts `sectPr` into paragraph `p`'s properties at its schema position. */
function attachSectPr(p: XElement, sectPr: XElement): void {
  let pPr = p.element(NS.w, 'pPr');
  if (pPr === undefined) {
    pPr = wChild(p, 'pPr');
    p.prepend(pPr);
  }
  insertOrdered(pPr, sectPr, PPR_ORDER);
}

/** Previous sibling element of `el`, skipping `skip`. */
function previousElement(el: XElement, skip?: XElement): XElement | undefined {
  let prev = el.previousElement();
  while (prev !== undefined && prev === skip) prev = prev.previousElement();
  return prev;
}

/**
 * Append: the target's body sectPr moves to its last paragraph (a new one when the target ends
 * with a table or a section break), the imported final sectPr becomes the body's.
 */
function placeAppend(body: XElement, blocks: readonly XElement[], finalSectPr: XElement): void {
  const importedSet = new Set(blocks);
  const own = body.elements(NS.w, 'sectPr').filter((s) => !importedSet.has(s));
  const targetFinal = own.at(-1);
  for (const s of own) s.remove();
  const firstBlock = blocks[0];
  const lastTarget =
    firstBlock !== undefined ? previousElement(firstBlock, targetFinal) : body.lastElement();
  if (lastTarget !== undefined) {
    const holder = holderAfter(body, lastTarget);
    attachSectPr(holder, targetFinal ?? wChild(body, 'sectPr'));
  }
  body.append(finalSectPr);
}

/** Where an insertion split the host section. */
interface HostSplit {
  /** The host section's own sectPr, now describing the part after the insertion. */
  readonly remainder: XElement;
  /** Its effective header/footer relationship ids before the insertion (per kind and type). */
  readonly effective: Map<string, string | undefined>;
  /** True when content of the host section precedes the insertion (it was split). */
  readonly split: boolean;
}

/**
 * Insert: the imported content ends with the source's final section (on its last paragraph);
 * the host section is split around it (its sectPr copied onto the paragraph before the
 * insertion) unless the insertion is at a section boundary.
 */
function placeInsert(
  ctx: ImportContext,
  body: XElement,
  blocks: readonly XElement[],
  finalSectPr: XElement,
  imported: ReadonlySet<XElement>,
): HostSplit {
  const sections = targetSectPrs(ctx).filter((s) => !imported.has(s));
  const lastBlock = blocks.at(-1);
  const firstBlock = blocks[0];
  let remainder = hostSectPr(body, blocks, sections);
  const effective = effectiveRefs(sections, remainder);
  remainder ??= ensureBodySectPr(body);
  // Close the imported content with the source's final section.
  attachSectPr(holderAfter(body, lastBlock), finalSectPr);
  // Split the host section when part of it comes before the insertion.
  const prev = firstBlock?.previousElement();
  const prevEndsSection =
    prev !== undefined && prev.element(NS.w, 'pPr')?.element(NS.w, 'sectPr') !== undefined;
  const split = prev !== undefined && !prevEndsSection;
  if (split) attachSectPr(holderAfter(body, prev), remainder.clone());
  return { remainder, effective, split };
}

function ensureBodySectPr(body: XElement): XElement {
  const existing = body.elements(NS.w, 'sectPr').at(-1);
  if (existing !== undefined) return existing;
  const s = wChild(body, 'sectPr');
  body.append(s);
  return s;
}

/** The sectPr of the section containing the insertion (first one after the last block). */
function hostSectPr(
  body: XElement,
  blocks: readonly XElement[],
  sections: readonly XElement[],
): XElement | undefined {
  const lastBlock = blocks.at(-1);
  if (lastBlock === undefined) return body.element(NS.w, 'sectPr');
  const kids = body.elements();
  const after = kids.indexOf(lastBlock);
  for (const s of sections) {
    let top: XElement = s;
    while (top.parent !== null && top.parent !== body && top.parent.nodeType === 1)
      top = top.parent;
    if (top.parent === body && kids.indexOf(top) > after) return s;
  }
  return undefined;
}

/** Every sectPr of the target body in document order (paragraph ones, then the body's). */
function targetSectPrs(ctx: ImportContext): XElement[] {
  return ctx.target
    .sections()
    .map((s) => s.element)
    .filter((e): e is XElement => e !== undefined);
}

/** Effective header/footer relationship ids of `sectPr` (own, else inherited), per kind+type. */
function effectiveRefs(
  sections: readonly XElement[],
  sectPr: XElement | undefined,
): Map<string, string | undefined> {
  const out = new Map<string, string | undefined>();
  const at = sectPr === undefined ? sections.length - 1 : sections.indexOf(sectPr);
  for (const kind of KINDS) {
    for (const type of TYPES) {
      let id: string | undefined;
      for (let i = at; i >= 0; i--) {
        const ref = ownReference(sections[i]!, kind, type);
        if (ref !== undefined) {
          id = ref.attrNS(NS.r, 'id');
          break;
        }
      }
      out.set(`${kind}:${type}`, id);
    }
  }
  return out;
}

/**
 * Pins the host remainder after an insertion: every header/footer slot it inherited gets an
 * explicit reference (an empty part when it inherited nothing), so it does not inherit from the
 * imported sections. A split remainder continues the page numbering and starts like the import.
 */
function pinHost(ctx: ImportContext, host: HostSplit, evenOdd: boolean): void {
  const s = host.remainder;
  const titlePg = onOff(s.element(NS.w, 'titlePg'));
  for (const kind of KINDS) {
    for (const type of TYPES) {
      if (ownReference(s, kind, type) !== undefined) continue;
      if ((type === 'first' && !titlePg) || (type === 'even' && !evenOdd)) continue;
      const id = host.effective.get(`${kind}:${type}`);
      addReference(s, kind, type, id ?? emptyPartId(ctx, kind, type));
    }
  }
  if (host.split) {
    const pg = s.element(NS.w, 'pgNumType');
    if (pg !== undefined) {
      pg.removeAttrNS(NS.w, 'start');
      if (pg._attrs.length === 0 && pg.firstElement() === undefined) pg.remove();
    }
    setType(s, ctx.options.sectionStart);
  }
}

// ---- header/footer references ---------------------------------------------------------------

function kindOf(ref: XElement): Kind | undefined {
  if (ref.uri !== NS.w) return undefined;
  if (ref.local === 'headerReference') return 'header';
  if (ref.local === 'footerReference') return 'footer';
  return undefined;
}

function referencesOf(sectPr: XElement): XElement[] {
  return sectPr.elements(NS.w).filter((e) => kindOf(e) !== undefined);
}

function ownReference(sectPr: XElement, kind: Kind, type: HeaderFooterType): XElement | undefined {
  for (const ref of sectPr.elements(NS.w, `${kind}Reference`)) {
    if ((ref.attrNS(NS.w, 'type') ?? 'default') === type) return ref;
  }
  return undefined;
}

function addReference(sectPr: XElement, kind: Kind, type: HeaderFooterType, rid: string): void {
  const ref = wChild(sectPr, `${kind}Reference`, { type });
  ref.setAttrNS(NS.r, 'id', rid);
  insertOrdered(sectPr, ref, SECTPR_ORDER);
}

function removeReferences(sectPr: XElement): void {
  for (const ref of referencesOf(sectPr)) ref.remove();
}

/**
 * Copies the headers/footers the imported sectPrs reference (source relationship ids) and points
 * the references at the copies; each copied part's content goes through importContent once. All
 * references are resolved before any relationship is added to the target.
 */
function remapReferences(
  ctx: ImportContext,
  imported: readonly Imported[],
  headerRoots: XElement[],
): void {
  const target = ctx.target;
  const resolved: { ref: XElement; kind: Kind; src: Part }[] = [];
  for (const imp of imported) {
    for (const ref of referencesOf(imp.sectPr)) {
      const rid = ref.attrNS(NS.r, 'id');
      const src = rid === undefined ? undefined : ctx.source.mainPart.related(rid);
      if (src === undefined) {
        ctx.warn(
          'HEADER_MISSING',
          `${ref.local} ${rid ?? '(no r:id)'} of an imported section has no part; removed`,
          ctx.source.mainPart.name,
        );
        ref.remove();
        continue;
      }
      resolved.push({ ref, kind: kindOf(ref)!, src });
    }
  }
  for (const { ref, kind, src } of resolved) {
    const copy = ctx.copier.copyPart(src);
    ref.setAttrNS(NS.r, 'id', relationshipTo(target.mainPart, copy, relTypeFor(target, REL[kind])));
    if (!firstVisit(ctx, copy)) continue;
    const root = partRoot(copy);
    if (root === undefined) continue;
    importContent(ctx, root.elements(), src, copy, { remapRelationships: false });
    headerRoots.push(root);
  }
}

/**
 * Other relationship ids of the source's final sectPr (`w:printerSettings/@r:id`), copied from
 * the source; an element whose relationship is missing is dropped. (Block sectPrs had theirs
 * remapped by the relationship step.)
 */
function remapOtherRelationships(ctx: ImportContext, sectPr: XElement): void {
  for (const el of sectPr.descendants()) {
    if (kindOf(el) !== undefined) continue;
    for (const a of el.attributes()) {
      if (a.uri !== NS.r) continue;
      if (ctx.source.mainPart.rels.get(a.value) === undefined) {
        ctx.warn('REL_MISSING', `${el.name} of an imported section has no relationship; removed`);
        el.remove();
        break;
      }
      el.setAttr(
        a.name,
        ctx.copier.copyRelationship(ctx.source.mainPart, a.value, ctx.target.mainPart),
      );
    }
  }
}

/** A fresh revision id for a tracked section change (`w:sectPrChange/@w:id`). */
function renumberRevision(ctx: ImportContext, sectPr: XElement): void {
  for (const ch of sectPr.elements(NS.w, 'sectPrChange')) {
    if (ch.attrNS(NS.w, 'id') !== undefined)
      ch.setAttrNS(NS.w, 'id', String(ctx.ids.next('revision')));
  }
}

/**
 * Even pages show the default header/footer in every section of `sectPrs`: each section's own
 * `even` reference becomes a copy of its own `default` one (or is removed when the section only
 * inherits its default, so it inherits the matching even one too).
 */
function mirrorEven(sectPrs: readonly XElement[]): void {
  for (const s of sectPrs) {
    for (const kind of KINDS) {
      const def = ownReference(s, kind, 'default');
      const even = ownReference(s, kind, 'even');
      const id = def?.attrNS(NS.r, 'id');
      if (def === undefined || id === undefined) {
        even?.remove();
      } else if (even !== undefined) {
        even.setAttrNS(NS.r, 'id', id);
      } else {
        addReference(s, kind, 'even', id);
      }
    }
  }
}

/** Empty header/footer parts created for this import, per kind+type. */
const emptyParts = new WeakMap<ImportContext, Map<string, string>>();

/** Relationship id of a new empty header/footer part (one per kind and type per import). */
function emptyPartId(ctx: ImportContext, kind: Kind, type: HeaderFooterType): string {
  let memo = emptyParts.get(ctx);
  if (memo === undefined) {
    memo = new Map();
    emptyParts.set(ctx, memo);
  }
  const key = `${kind}:${type}`;
  const known = memo.get(key);
  if (known !== undefined) return known;
  const part = createHeaderFooter(ctx.target, kind).part;
  ctx.addedParts.push(part.name);
  firstVisit(ctx, part);
  const id = relationshipTo(ctx.target.mainPart, part, relTypeFor(ctx.target, REL[kind]));
  memo.set(key, id);
  return id;
}

function sourceHasHeadersFooters(ctx: ImportContext): boolean {
  return ctx.source
    .sections()
    .some((s) => s.element !== undefined && referencesOf(s.element).length > 0);
}

// ---- numbering and section start ------------------------------------------------------------

function applyPageNumbering(ctx: ImportContext, sectPr: XElement): void {
  const mode = ctx.options.pageNumbering;
  if (mode === 'keepSource') return;
  let pg = sectPr.element(NS.w, 'pgNumType');
  if (mode === 'continue') {
    if (pg === undefined) return;
    pg.removeAttrNS(NS.w, 'start');
    if (pg._attrs.length === 0 && pg.firstElement() === undefined) pg.remove();
    return;
  }
  if (pg === undefined) pg = insertOrdered(sectPr, wChild(sectPr, 'pgNumType'), SECTPR_ORDER);
  pg.setAttrNS(NS.w, 'start', '1');
}

function setType(sectPr: XElement, type: string): void {
  let el = sectPr.element(NS.w, 'type');
  if (el === undefined) el = insertOrdered(sectPr, wChild(sectPr, 'type'), SECTPR_ORDER);
  el.setAttrNS(NS.w, 'val', type);
}
