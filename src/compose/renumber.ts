// renumberIds: makes the document-wide ids of imported markup unique in the target.
//
// Kinds handled here (bookmarks, comments and notes have their own steps):
// - drawings: `wp:docPr/@id`, and `pic:cNvPr/@id` when it repeats its drawing's docPr id;
// - paragraphs and rows: `w14:paraId`, `w14:textId` (Word's shared "77777777" textId is kept);
// - content controls: `w:sdtPr/w:id/@w:val`;
// - revisions: `@w:id` of `w:ins`, `w:del`, `w:moveFrom`/`To` (+ range starts/ends), `w:*Change`,
//   `w:cell*` (ids.ts REVISION_ELEMENTS);
// - VML: `o:spid` and `_x0000_?<n>` shape ids, with the references to them (`o:OLEObject
//   @ShapeID`, `w:control @w:shapeid`) following.
// An id is KEPT when the target does not use it yet (it is then reserved), otherwise replaced by
// a fresh one from ctx.ids. Within one call the same old id always maps to the same new id, so
// pairs that share an id on purpose stay paired: range starts and ends, and the drawing repeated
// in an `mc:AlternateContent` Fallback.
// New ids are all computed first and written afterwards (the registry is never asked while the
// DOM changes under it).

import { REVISION_ELEMENTS } from '../wml/ids.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { ImportContext } from './types.js';

const VML_ID = /^_x0000_([a-z])(\d{1,9})$/;
const KEEP_TEXT_ID = '77777777';

/** Source paraId → target paraId, per import (for steps that reference paragraphs by paraId). */
const paraIdMaps = new WeakMap<ImportContext, Map<string, string>>();

/**
 * The `w14:paraId` values renumbered so far in this import (source value → target value; ids
 * that were kept do not appear). commentsExtended/commentsIds reference comment paragraphs by
 * paraId and must follow.
 */
export function renumberedParaIds(ctx: ImportContext): ReadonlyMap<string, string> {
  return paraIdMaps.get(ctx) ?? new Map();
}

type Kind = 'docPr' | 'sdt' | 'revision' | 'paraId' | 'textId' | 'vml';

/** A pending attribute write. */
interface Write {
  el: XElement;
  name: string;
  value: string;
}

/** Renumbers the ids listed in the file header inside `elements` and their descendants. */
export function renumberIds(ctx: ImportContext, elements: XElement[]): void {
  const ids = ctx.ids;
  const maps: Record<Kind, Map<string, string>> = {
    docPr: new Map(),
    sdt: new Map(),
    revision: new Map(),
    paraId: new Map(),
    textId: new Map(),
    vml: new Map(),
  };
  const map = (kind: Kind, old: string): string => {
    const m = maps[kind];
    const hit = m.get(old);
    if (hit !== undefined) return hit;
    let out: string;
    switch (kind) {
      case 'paraId':
      case 'textId':
        if (kind === 'textId' && old === KEEP_TEXT_ID) out = old;
        else if (/^[0-9A-Fa-f]{1,8}$/.test(old) && !ids.has(kind, old)) {
          ids.reserve(kind, old);
          out = old;
        } else out = ids.nextHex(kind);
        break;
      case 'vml': {
        const mm = VML_ID.exec(old);
        if (mm === null) out = old;
        else if (!ids.hasVmlShapeId(old)) {
          ids.reserveVmlShapeId(old);
          out = old;
        } else out = `_x0000_${mm[1]}${ids.nextVmlShapeId().slice(8)}`;
        break;
      }
      default:
        if (/^-?\d{1,10}$/.test(old) && !ids.has(kind, old)) {
          ids.reserve(kind, old);
          out = old;
        } else out = String(ids.next(kind));
    }
    m.set(old, out);
    return out;
  };

  const writes: Write[] = [];
  const set = (el: XElement, name: string, old: string, value: string): void => {
    if (value !== old) writes.push({ el, name, value });
  };
  /** VML references resolved after the walk (the shape may come later in document order). */
  const vmlRefs: { el: XElement; name: string; old: string }[] = [];

  const stack: XElement[] = [];
  for (let i = elements.length - 1; i >= 0; i--) stack.push(elements[i]!);
  while (stack.length > 0) {
    const el = stack.pop()!;
    const kids = el._children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
    const a = el._attrs;
    if (a.length === 0) continue;
    const uri = el.uri;
    if (uri === NS.w) {
      if (REVISION_ELEMENTS.has(el.local)) {
        attr(el, NS.w, 'id', (name, old) => set(el, name, old, map('revision', old)));
      } else if (el.local === 'id' && (el.parent as XElement | null)?.is(NS.w, 'sdtPr')) {
        attr(el, NS.w, 'val', (name, old) => set(el, name, old, map('sdt', old)));
      } else if (el.local === 'control') {
        attr(el, NS.w, 'shapeid', (name, old) => vmlRefs.push({ el, name, old }));
      }
    } else if (uri === NS.wp) {
      if (el.local === 'docPr')
        attr(el, '', 'id', (name, old) => set(el, name, old, map('docPr', old)));
    } else if (uri === NS.pic) {
      if (el.local === 'cNvPr') {
        // Keep a pic:cNvPr that repeats its drawing's docPr id in step with it.
        const drawing = el.closest(NS.wp, 'inline') ?? el.closest(NS.wp, 'anchor');
        const docPrId = drawing?.element(NS.wp, 'docPr')?.attrNS('', 'id');
        attr(el, '', 'id', (name, old) => {
          if (docPrId !== undefined && old === docPrId) set(el, name, old, map('docPr', old));
        });
      }
    } else if (uri === NS.v || uri === NS.o) {
      if (el.local === 'OLEObject' && uri === NS.o) {
        attr(el, '', 'ShapeID', (name, old) => vmlRefs.push({ el, name, old }));
      } else if (el.local !== 'shapetype') {
        attr(el, '', 'id', (name, old) => {
          if (VML_ID.test(old)) set(el, name, old, map('vml', old));
        });
      }
    }
    for (let i = 0; i < a.length; i += 3) {
      const au = a[i + 1];
      const name = a[i]!;
      const old = a[i + 2]!;
      if (au === NS.w14) {
        if (name.endsWith(':paraId')) set(el, name, old, map('paraId', old));
        else if (name.endsWith(':textId')) set(el, name, old, map('textId', old));
      } else if (au === NS.o && name.endsWith(':spid')) {
        if (VML_ID.test(old)) set(el, name, old, map('vml', old));
      }
    }
  }
  for (const r of vmlRefs) {
    const v = maps.vml.get(r.old);
    if (v !== undefined) set(r.el, r.name, r.old, v);
  }
  for (const w of writes) w.el.setAttr(w.name, w.value);

  if (maps.paraId.size > 0) {
    let acc = paraIdMaps.get(ctx);
    if (acc === undefined) {
      acc = new Map();
      paraIdMaps.set(ctx, acc);
    }
    for (const [from, to] of maps.paraId) if (from !== to) acc.set(from, to);
  }
}

/** Calls `fn(qname, value)` for `el`'s attribute `{uri}local`, when present. */
function attr(
  el: XElement,
  uri: string,
  local: string,
  fn: (name: string, value: string) => void,
): void {
  const a = el._attrs;
  for (let i = 0; i < a.length; i += 3) {
    if (a[i + 1] !== uri) continue;
    const name = a[i]!;
    if (name === local || name.endsWith(`:${local}`)) {
      fn(name, a[i + 2]!);
      return;
    }
  }
}
