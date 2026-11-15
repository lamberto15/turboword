// bindContentControls: write a custom XML store and refresh the controls bound to it.
//
// This is what Word itself does when it opens a file with data-bound controls: the XML is the
// source of truth and the document content a cache. Refreshing the cache here means every other
// consumer (LibreOffice, PDF converters, `text()`) shows the data too.
//
// Repeating sections bound to the store are expanded first, outermost first: the section's path
// (`/invoice[1]/lines[1]/line[1]`) without its last position selects the records, the first item
// is cloned once per record, and the bindings inside clone k are re-pointed at `line[k]`.

import { UsageError } from '../core/errors.js';
import type { WordDocument } from '../wml/document.js';
import { setWAttr } from '../wml/order.js';
import type { XDocument, XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { parseXml } from '../xml/parse.js';
import {
  addCustomXml,
  findCustomXml,
  guidFromContent,
  isPropertyStore,
  normalizeGuid,
  writeCustomXml,
} from './customxml.js';
import { parseIsoDate } from './date-format.js';
import { freshenClone, sectionItems } from './fill.js';
import { allControls, controlsUnder, FormsContext, type Sdt } from './sdt.js';
import { evaluateXPath, parseXPath, readHit, selectElements } from './xpath.js';

/** Types whose value is their nested controls, never set from XML directly. */
const CONTAINERS: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'repeatingSection',
  'repeatingSectionItem',
  'group',
  'buildingBlockGallery',
  'citation',
  'bibliography',
  'equation',
]);

function attached(el: XElement, root: XElement): boolean {
  for (let p: XElement | null = el; p !== null; p = p.parent?.nodeType === 1 ? p.parent : null) {
    if (p === root) return true;
  }
  return false;
}

/** Bytes of base64 text (a `data:` URL prefix and whitespace allowed); undefined when invalid. */
function fromBase64(s: string): Uint8Array | undefined {
  const b64 = s.replace(/^\s*data:[^,]*,/, '').replace(/\s+/g, '');
  if (b64 === '' || !/^[A-Za-z0-9+/]*={0,2}$/.test(b64)) return undefined;
  try {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  } catch {
    return undefined;
  }
}

/** Expands repeating section `sec` to one item per record of `data`; returns the new items. */
function expand(ctx: FormsContext, sec: Sdt, data: XDocument): XElement[] | undefined {
  const b = sec.binding!;
  const p = parseXPath(b.xpath, b.prefixMappings);
  const items = sectionItems(ctx, sec);
  const template = items[0];
  if (p === undefined || p.attr !== undefined || p.text === true || template === undefined) {
    return undefined;
  }
  const last = p.steps[p.steps.length - 1]!;
  const records = selectElements(data, { steps: [...p.steps.slice(0, -1), { ...last, pos: [] }] });
  const base = b.xpath.trim().replace(/\[\s*\d+\s*\]$/, '');
  const prefix = new RegExp(
    `^${base.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?:\\[\\s*\\d+\\s*\\])?(?=/|$)`,
  );
  const ids = ctx.ids();
  const out: XElement[] = [];
  const release = ids.hold();
  // No record: keep one item, as Word does.
  for (let k = 1; k <= Math.max(records.length, 1); k++) {
    const clone = template.element.clone();
    freshenClone(clone, ids);
    for (const el of clone.descendants()) {
      if (el.local !== 'dataBinding' || (el.uri !== NS.w && el.uri !== NS.w15)) continue;
      const x = el.attrNS(NS.w, 'xpath')?.trim();
      if (x !== undefined && prefix.test(x))
        setWAttr(el, 'xpath', x.replace(prefix, `${base}[${k}]`));
    }
    template.element.before(clone);
    out.push(clone);
  }
  release();
  for (const it of items) it.element.remove();
  return out;
}

/** Sets control `c` from bound value `s`; false when it cannot take it. */
function setFromXml(c: Sdt, s: string): boolean {
  switch (c.type) {
    case 'checkbox':
      c.apply(/^\s*(true|1)\s*$/i.test(s));
      return true;
    case 'date': {
      const d = parseIsoDate(s);
      if (d !== undefined) c.apply(d);
      else c.showText(s);
      return true;
    }
    case 'dropDownList':
    case 'comboBox': {
      const hit = c.items.find((i) => i.value === s);
      if (hit !== undefined) c.apply(hit.value);
      else c.showText(s);
      return true;
    }
    case 'picture': {
      const bytes = fromBase64(s);
      if (bytes === undefined) return false;
      c.apply(bytes);
      return true;
    }
    case 'richText':
      // Word stores rich text bindings as an escaped Flat OPC package, which is not text.
      if (s.includes('<pkg:package')) return false;
      c.apply(s);
      return true;
    default:
      c.apply(s);
      return true;
  }
}

/** See index.ts. */
export function bind(doc: WordDocument, xml: string, o?: { storeItemId?: string }): number {
  const root = parseXml(xml, { part: 'customXml' }).root;
  const ctx = new FormsContext(doc);
  ctx.writeThrough = false;
  let id = o?.storeItemId;
  if (id === undefined) {
    const stores: string[] = [];
    for (const c of allControls(ctx)) {
      const s = c.binding?.storeItemId;
      if (s === undefined || isPropertyStore(s)) continue;
      if (!stores.some((x) => normalizeGuid(x) === normalizeGuid(s))) stores.push(s);
    }
    // Several stores: the one whose current data has the same root element as `xml`.
    id =
      stores.find((s) => {
        try {
          const r = findCustomXml(doc, s)?.part.xml().root;
          return r !== undefined && r.local === root.local && r.uri === root.uri;
        } catch {
          return false;
        }
      }) ??
      stores[0] ??
      guidFromContent(xml);
  }
  if (isPropertyStore(id)) {
    throw new UsageError(
      `${id} is the document properties store: set doc.properties instead of binding XML`,
    );
  }
  const want = normalizeGuid(id);
  let store = findCustomXml(doc, id);
  if (store !== undefined) writeCustomXml(store, xml);
  else store = addCustomXml(doc, xml, id);
  const data = store.part.xml();
  const mine = (c: Sdt): boolean => {
    const s = c.binding?.storeItemId;
    return c.binding !== undefined && (s === undefined || normalizeGuid(s) === want);
  };

  let n = 0;
  const work = allControls(ctx).filter((c) => c.type === 'repeatingSection' && mine(c));
  for (let sec = work.shift(); sec !== undefined; sec = work.shift()) {
    if (!attached(sec.element, sec.story.element)) continue;
    const items = expand(ctx, sec, data);
    if (items === undefined) continue;
    n++;
    for (const it of items) {
      for (const c of controlsUnder(ctx, sec.story, it, false)) {
        if (c.type === 'repeatingSection' && mine(c)) work.push(c);
      }
    }
  }
  for (const c of allControls(ctx)) {
    if (!mine(c) || CONTAINERS.has(c.type)) continue;
    const b = c.binding!;
    const hit = evaluateXPath(data, b.xpath, b.prefixMappings);
    if (hit !== undefined && setFromXml(c, readHit(hit))) n++;
  }
  return n;
}
