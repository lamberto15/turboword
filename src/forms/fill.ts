// fillContentControls: values by tag or alias, records for repeating sections.
//
// Scalars fill every control with the key that existed when the call started; an array fills a
// repeating section: its first item is the template, cloned once per record (fresh content
// control, paragraph and drawing ids; bookmarks and comment anchors are not duplicated), and
// each clone's nested controls are filled from the record by the same key kind. Arrays are
// applied first, so a scalar never overwrites a value a record put in a clone.
//
// A filled repeating section is unbound (its `w15:dataBinding`, and every binding inside the
// clones, is removed): the records no longer mirror the custom XML, and Word would otherwise
// rebuild the section from the XML when it opens the file. An empty array removes the section.

import { UsageError } from '../core/errors.js';
import type { WordDocument } from '../wml/document.js';
import { REVISION_ELEMENTS } from '../wml/ids.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { ensureValid, unwrap } from './content.js';
import type { IdSource } from './ids.js';
import { allControls, controlsUnder, FormsContext, type Sdt } from './sdt.js';
import type { ContentControlValue, FillRecord, FillResult, FillValue } from './types.js';

type Records = readonly FillRecord[];

/** True when `el` is still in the tree under `root`. */
function attached(el: XElement, root: XElement): boolean {
  for (let p: XElement | null = el; p !== null; p = p.parent?.nodeType === 1 ? p.parent : null) {
    if (p === root) return true;
  }
  return false;
}

/**
 * Gives a cloned subtree fresh ids: content-control, paragraph/text, drawing and revision ids
 * (range markers keep theirs: they pair by id). Bookmarks and comment anchors are dropped
 * (their names and ids must stay unique).
 */
export function freshenClone(root: XElement, ids: IdSource): void {
  for (const el of [root, ...root.descendants()]) {
    if (el.uri === NS.w) {
      switch (el.local) {
        case 'sdtPr': {
          const id = el.element(NS.w, 'id');
          if (id !== undefined) id.setAttrNS(NS.w, 'val', String(ids.sdt()));
          break;
        }
        case 'bookmarkStart':
        case 'bookmarkEnd':
        case 'commentRangeStart':
        case 'commentRangeEnd':
        case 'commentReference':
          el.remove();
          break;
        default:
          if (
            REVISION_ELEMENTS.has(el.local) &&
            !el.local.includes('Range') &&
            el.attrNS(NS.w, 'id') !== undefined
          ) {
            el.setAttrNS(NS.w, 'id', String(ids.revision()));
          }
      }
    } else if (el.uri === NS.wp && el.local === 'docPr') {
      el.setAttr('id', String(ids.docPr()));
    }
    if (el.attrNS(NS.w14, 'paraId') !== undefined) el.setAttrNS(NS.w14, 'paraId', ids.paraId());
    if (el.attrNS(NS.w14, 'textId') !== undefined) el.setAttrNS(NS.w14, 'textId', ids.textId());
  }
}

/** Removes every data binding in and under `root`. */
function unbind(root: XElement): void {
  for (const el of [root, ...root.descendants()]) {
    if (el.local === 'dataBinding' && (el.uri === NS.w || el.uri === NS.w15)) el.remove();
  }
}

/** The items of repeating section `sec` (its direct `w:sdt` children), in order. */
export function sectionItems(ctx: FormsContext, sec: Sdt): Sdt[] {
  const c = sec.content;
  if (c === undefined) return [];
  return controlsUnder(ctx, sec.story, c, false).filter(
    (s) => s.type === 'repeatingSectionItem' && s.element.parent === c,
  );
}

class Filler {
  readonly ctx: FormsContext;
  readonly by: 'tag' | 'alias';
  readonly touched: XElement[] = [];
  filled = 0;

  constructor(ctx: FormsContext, by: 'tag' | 'alias') {
    this.ctx = ctx;
    this.by = by;
  }

  key(c: Sdt): string | undefined {
    return this.by === 'tag' ? c.tag : c.alias;
  }

  /** Applies `values` to `controls` (a snapshot); returns the keys nothing matched. */
  apply(controls: readonly Sdt[], values: Readonly<Record<string, FillValue>>): string[] {
    const byKey = new Map<string, Sdt[]>();
    for (const c of controls) {
      const k = this.key(c);
      if (k === undefined) continue;
      const list = byKey.get(k);
      if (list === undefined) byKey.set(k, [c]);
      else list.push(c);
    }
    const missing: string[] = [];
    const entries = Object.entries(values);
    for (const [k, v] of entries) {
      if (!byKey.has(k)) missing.push(k);
      else if (Array.isArray(v)) {
        for (const c of byKey.get(k)!)
          if (attached(c.element, c.story.element)) this.section(c, v as Records);
      }
    }
    for (const [k, v] of entries) {
      if (Array.isArray(v)) continue;
      for (const c of byKey.get(k) ?? []) {
        if (!attached(c.element, c.story.element)) continue;
        if (c.type === 'repeatingSection') {
          throw new UsageError(`Repeating section ${JSON.stringify(k)} takes an array of records`);
        }
        c.setValue(v as ContentControlValue);
        this.touched.push(c.element);
        this.filled++;
      }
    }
    return missing;
  }

  /** Fills repeating section `sec` with one item per record. */
  section(sec: Sdt, records: Records): void {
    if (sec.type !== 'repeatingSection') {
      throw new UsageError(
        `Content control ${JSON.stringify(this.key(sec))} is a ${sec.type} control; arrays fill repeating sections`,
      );
    }
    if (sec.locked.content) {
      throw new UsageError(`Repeating section ${JSON.stringify(this.key(sec))} is locked`);
    }
    const items = sectionItems(this.ctx, sec);
    const template = items[0];
    if (template === undefined) {
      throw new UsageError(`Repeating section ${JSON.stringify(this.key(sec))} has no item`);
    }
    unbind(sec.pr ?? sec.element);
    this.filled++;
    if (records.length === 0) {
      const parent = sec.element.parent;
      sec.element.remove();
      if (parent !== null && parent.nodeType === 1) ensureValid(parent);
      return;
    }
    this.touched.push(sec.element);
    const ids = this.ctx.ids();
    const clones: XElement[] = [];
    const release = ids.hold();
    try {
      for (let i = 0; i < records.length; i++) {
        const clone = template.element.clone();
        freshenClone(clone, ids);
        unbind(clone);
        template.element.before(clone);
        clones.push(clone);
      }
    } finally {
      release();
    }
    for (const it of items) it.element.remove();
    for (let i = 0; i < clones.length; i++) {
      this.touched.push(clones[i]!);
      this.apply(controlsUnder(this.ctx, sec.story, clones[i]!, false), records[i]!);
    }
  }
}

/** See index.ts. */
export function fill(
  doc: WordDocument,
  values: Readonly<Record<string, FillValue>>,
  o?: { by?: 'tag' | 'alias'; removeControls?: boolean },
): FillResult {
  const f = new Filler(new FormsContext(doc), o?.by ?? 'tag');
  const missing = f.apply(allControls(f.ctx), values);
  if (o?.removeControls === true) {
    // Innermost first, so each unwrap sees its own content.
    for (let i = f.touched.length - 1; i >= 0; i--) unwrap(f.touched[i]!);
  }
  return { filled: f.filled, missing };
}
