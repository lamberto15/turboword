// The XPath subset Word writes in `w:dataBinding/@w:xpath`.
//
// Word's XML Mapping pane only ever produces absolute location paths of child steps with
// positional predicates, optionally ending in an attribute or `text()`:
//   /ns0:invoice[1]/ns0:lines[1]/ns0:line[2]/@ns1:id
// Prefixes are bound by `w:prefixMappings` (`xmlns:ns0='urn:x' xmlns:ns1="urn:y"`), never by the
// document. An unprefixed step matches an element in no namespace (XPath 1.0). Anything outside
// this subset (`//`, axes, functions, value predicates) is reported as unsupported (undefined),
// never guessed.

import { type XDocument, type XElement, XText } from '../xml/dom.js';
import { canonicalNs } from '../xml/ns.js';

interface Step {
  /** Canonical namespace URI ('' = none); undefined for `*` (any element). */
  uri: string | undefined;
  /** Local name, or '*' for any. */
  local: string;
  /** Positional predicates (1-based), applied in order. */
  pos: number[];
}

/** A parsed binding path. */
export interface ParsedXPath {
  steps: Step[];
  /** Final `@attr` step. */
  attr?: { uri: string; local: string };
  /** Final `text()` step. */
  text?: boolean;
}

/** What a binding path selects: an element, one of its attributes, or its text. */
export interface XPathHit {
  el: XElement;
  attr?: { uri: string; local: string };
  text?: boolean;
}

/** `xmlns:p='uri'` pairs of a `w:prefixMappings` value. */
export function prefixMap(mappings: string | undefined): Map<string, string> {
  const out = new Map<string, string>();
  if (mappings === undefined) return out;
  const re = /xmlns:([^\s=]+)\s*=\s*(?:'([^']*)'|"([^"]*)")/g;
  for (let m = re.exec(mappings); m !== null; m = re.exec(mappings)) {
    out.set(m[1]!, canonicalNs(m[2] ?? m[3] ?? ''));
  }
  return out;
}

const STEP = /^(?:([^\s[\]:@/()*]+):)?([^\s[\]:@/()]+)((?:\[\s*\d+\s*\])*)$/;

/** Parses `xpath` against prefix mappings; undefined when outside the supported subset. */
export function parseXPath(xpath: string, mappings?: string): ParsedXPath | undefined {
  const s = xpath.trim();
  if (!s.startsWith('/') || s.includes('//')) return undefined;
  const ns = prefixMap(mappings);
  const parts = s.slice(1).split('/');
  const out: ParsedXPath = { steps: [] };
  for (let i = 0; i < parts.length; i++) {
    let part = parts[i]!.trim();
    const last = i === parts.length - 1;
    if (last && /^text\(\s*\)$/.test(part)) {
      if (i === 0) return undefined;
      out.text = true;
      break;
    }
    const isAttr = part.startsWith('@');
    if (isAttr) {
      if (!last || i === 0) return undefined;
      part = part.slice(1);
    }
    const m = STEP.exec(part);
    if (m === null) return undefined;
    const [, prefix, local, preds] = m;
    let uri: string | undefined = '';
    if (prefix !== undefined) {
      uri = ns.get(prefix);
      if (uri === undefined) return undefined;
    } else if (local === '*') uri = undefined;
    if (isAttr) {
      if (preds !== '' || uri === undefined) return undefined;
      out.attr = { uri, local: local! };
      break;
    }
    const pos = (preds!.match(/\d+/g) ?? []).map(Number);
    out.steps.push({ uri, local: local!, pos });
  }
  return out.steps.length > 0 ? out : undefined;
}

function matches(el: XElement, step: Step): boolean {
  return (
    (step.local === '*' || el.local === step.local) &&
    (step.uri === undefined || el.uri === step.uri)
  );
}

/** Applies positional predicates to one context node's candidates. */
function filter(cands: XElement[], pos: readonly number[]): XElement[] {
  let out = cands;
  for (const p of pos) {
    const hit = out[p - 1];
    out = hit === undefined ? [] : [hit];
  }
  return out;
}

/** Elements selected by the element steps of `p` (attribute/text steps ignored), in order. */
export function selectElements(doc: XDocument, p: ParsedXPath): XElement[] {
  let root: XElement;
  try {
    root = doc.root;
  } catch {
    return [];
  }
  const [first, ...rest] = p.steps;
  let ctx = first !== undefined && matches(root, first) ? filter([root], first.pos) : [];
  for (const step of rest) {
    const next: XElement[] = [];
    for (const el of ctx)
      next.push(
        ...filter(
          el.elements().filter((k) => matches(k, step)),
          step.pos,
        ),
      );
    ctx = next;
    if (ctx.length === 0) break;
  }
  return ctx;
}

/** The first node `xpath` selects in `doc`, or undefined (no match or unsupported path). */
export function evaluateXPath(
  doc: XDocument,
  xpath: string,
  mappings?: string,
): XPathHit | undefined {
  const p = parseXPath(xpath, mappings);
  if (p === undefined) return undefined;
  const el = selectElements(doc, p)[0];
  if (el === undefined) return undefined;
  if (p.attr !== undefined) {
    return el.attrNS(p.attr.uri, p.attr.local) === undefined ? undefined : { el, attr: p.attr };
  }
  return p.text === true ? { el, text: true } : { el };
}

/** String value of a hit (an element's string value is all its descendant text). */
export function readHit(h: XPathHit): string {
  if (h.attr !== undefined) return h.el.attrNS(h.attr.uri, h.attr.local) ?? '';
  if (h.text === true) {
    let s = '';
    for (const k of h.el.children) if (k.nodeType === 3) s += k.text;
    return s;
  }
  return h.el.text;
}

/** Writes a string value where a hit points (an element's content becomes that text). */
export function writeHit(h: XPathHit, value: string): void {
  if (h.attr !== undefined) {
    if (h.attr.uri === '') h.el.setAttr(h.attr.local, value);
    else h.el.setAttrNS(h.attr.uri, h.attr.local, value);
  } else if (readHit(h) === value) {
    return;
  } else if (h.text === true && h.el.firstElement() !== undefined) {
    // text() of mixed content: replace the text nodes only, keep the child elements.
    for (const k of h.el.children.slice()) if (k.nodeType === 3) k.remove();
    h.el.prepend(new XText(value));
  } else {
    h.el.text = value;
  }
}
