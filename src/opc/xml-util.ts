// Small DOM helpers shared by the OPC views ([Content_Types].xml and .rels parts).
//
// New children copy the parent's prefix and namespace, so they are valid however the producer
// wrote the root (default namespace, as Word does, or an unusual prefix), and their attributes are
// unprefixed, as OPC requires.

import { XDocument, XElement } from '../xml/dom.js';
import { NS, prefixOf } from '../xml/ns.js';

/** The XML declaration Word writes on package parts. */
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** New detached element `local` in `parent`'s namespace and prefix, with unprefixed `attrs`. */
export function childElement(
  parent: XElement,
  local: string,
  attrs: readonly (readonly [string, string])[],
): XElement {
  const prefix = prefixOf(parent.name);
  const flat: string[] = [];
  for (const [name, value] of attrs) flat.push(name, '', value);
  return new XElement(prefix === '' ? local : `${prefix}:${local}`, parent.uri, flat);
}

/** Child elements of `parent` with local name `local` (any namespace: lenient with odd producers). */
export function childrenNamed(parent: XElement, local: string): XElement[] {
  return parent.elements(undefined, local);
}

/** A new document whose root `{uri}local` declares `uri` as the default namespace. */
export function emptyDocument(uri: string, local: string, children: XElement[] = []): XDocument {
  const root = new XElement(local, uri, ['xmlns', NS.xmlns, uri]);
  for (const c of children) root.append(c);
  return new XDocument(XML_DECLARATION, [root]);
}
