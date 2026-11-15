// Custom XML data parts: the stores that bound content controls read from.
//
// Each store is `customXml/itemN.xml` (the data, related from the main document with the
// `customXml` relationship) plus `customXml/itemPropsN.xml` (related from the item with
// `customXmlProps`), whose `ds:datastoreItem/@ds:itemID` GUID is what `w:dataBinding/@w:storeItemID`
// names. Word compares those GUIDs case-insensitively and with or without braces, so we do too.

import type { Part } from '../opc/part.js';
import { relativeTarget } from '../opc/uri.js';
import type { WordDocument } from '../wml/document.js';
import { relTypeFor } from '../wml/media.js';
import { escapeAttr } from '../xml/escape.js';
import { parseXml } from '../xml/parse.js';

const DS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PROPS_CT = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';

/** One custom XML store. */
export interface CustomXmlItem {
  /** The data part (`customXml/itemN.xml`). */
  part: Part;
  /** `ds:itemID` of its properties part, as written. */
  id: string | undefined;
}

/** A GUID in comparable form: no braces, upper case. */
export function normalizeGuid(id: string): string {
  return id
    .trim()
    .replace(/^\{|\}$/g, '')
    .toUpperCase();
}

/** Every custom XML store related from the main document, in relationship order. */
export function customXmlItems(doc: WordDocument): CustomXmlItem[] {
  const out: CustomXmlItem[] = [];
  for (const part of doc.mainPart.relatedByType('customXml')) {
    let id: string | undefined;
    const props = part.relatedByType('customXmlProps')[0];
    if (props !== undefined) {
      try {
        id = props.xml().root.attrNS(DS, 'itemID');
      } catch {
        // A broken properties part only means the store has no usable id.
      }
    }
    out.push({ part, id });
  }
  return out;
}

/** The store whose id is `id`, if any. */
export function findCustomXml(doc: WordDocument, id: string): CustomXmlItem | undefined {
  const want = normalizeGuid(id);
  const kind = PROPERTY_STORES.get(want);
  if (kind !== undefined) {
    const rel = doc.package.rels.byType(kind)[0];
    const part = rel?.partName === undefined ? undefined : doc.package.part(rel.partName);
    return part === undefined ? undefined : { part, id };
  }
  return customXmlItems(doc).find((i) => i.id !== undefined && normalizeGuid(i.id) === want);
}

/**
 * Store ids Word maps to package parts instead of custom XML: the core properties
 * (`docProps/core.xml`, title, subject, author...) and the extended properties
 * (`docProps/app.xml`, company, manager...). Controls bound to them show document properties.
 */
const PROPERTY_STORES: ReadonlyMap<string, string> = /* @__PURE__ */ new Map([
  ['6C3C8BC8-F283-45AE-878A-BAB7291924A1', 'core-properties'],
  ['6668398D-A668-4E3E-A5EB-62B293D839F1', 'extended-properties'],
]);

/** True for the core/extended document properties stores (see {@link findCustomXml}). */
export function isPropertyStore(id: string): boolean {
  return PROPERTY_STORES.has(normalizeGuid(id));
}

/** 32-bit FNV-1a variant with a seed, over UTF-16 code units. */
function hash32(s: string, seed: number): number {
  let h = (0x811c9dc5 ^ seed) >>> 0;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  h ^= h >>> 15;
  h = Math.imul(h, 0x2c1b3c6d) >>> 0;
  h ^= h >>> 12;
  return h >>> 0;
}

/** A deterministic GUID (`{XXXXXXXX-XXXX-4XXX-8XXX-XXXXXXXXXXXX}`) derived from `content`. */
export function guidFromContent(content: string): string {
  let hex = '';
  for (let k = 0; k < 4; k++)
    hex += hash32(content, k * 0x9e3779b9)
      .toString(16)
      .padStart(8, '0');
  hex = `${hex.slice(0, 12)}4${hex.slice(13, 16)}${((Number.parseInt(hex[16]!, 16) & 3) | 8).toString(16)}${hex.slice(17)}`;
  const u = hex.toUpperCase();
  return `{${u.slice(0, 8)}-${u.slice(8, 12)}-${u.slice(12, 16)}-${u.slice(16, 20)}-${u.slice(20)}}`;
}

/** UTF-8 bytes of an XML string whose declaration (if any) says UTF-8, as the bytes will be. */
function xmlBytes(xml: string): Uint8Array {
  const text = xml.replace(
    /^(﻿?<\?xml[^>]*?encoding\s*=\s*)(["'])[^"']*\2/,
    (_m, head: string, q: string) => `${head}${q}UTF-8${q}`,
  );
  return new TextEncoder().encode(text);
}

/** Replaces the data of store `item` with `xml` (checked for well-formedness first). */
export function writeCustomXml(item: CustomXmlItem, xml: string): void {
  parseXml(xml, { part: item.part.name });
  item.part.setBytes(xmlBytes(xml));
}

/**
 * Adds a new store holding `xml` with id `id`: `customXml/itemN.xml`, `customXml/itemPropsN.xml`
 * (same N, the first free one), their relationship and the main document's relationship.
 */
export function addCustomXml(doc: WordDocument, xml: string, id: string): CustomXmlItem {
  const parsed = parseXml(xml, { part: 'customXml' });
  const pkg = doc.package;
  let n = 1;
  while (
    pkg.part(`customXml/item${n}.xml`) !== undefined ||
    pkg.part(`customXml/itemProps${n}.xml`) !== undefined
  ) {
    n++;
  }
  const part = pkg.addPart(`customXml/item${n}.xml`, 'application/xml', xmlBytes(xml));
  const guid = `{${normalizeGuid(id)}}`;
  let ns = '';
  try {
    ns = parsed.root.uri;
  } catch {
    // parseXml guarantees a root; keep the schema list empty otherwise.
  }
  const refs =
    ns === ''
      ? '<ds:schemaRefs/>'
      : `<ds:schemaRefs><ds:schemaRef ds:uri="${escapeAttr(ns)}"/></ds:schemaRefs>`;
  const props =
    '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n' +
    `<ds:datastoreItem ds:itemID="${guid}" xmlns:ds="${DS}">${refs}</ds:datastoreItem>`;
  pkg.addPart(`customXml/itemProps${n}.xml`, PROPS_CT, new TextEncoder().encode(props));
  part.rels.add(relTypeFor(doc, `${REL_BASE}customXmlProps`), `itemProps${n}.xml`);
  doc.mainPart.rels.add(
    relTypeFor(doc, `${REL_BASE}customXml`),
    relativeTarget(doc.mainPart.name, part.name),
  );
  return { part, id: guid };
}
