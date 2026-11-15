// Differential test: turboword's parser against @xmldom/xmldom on every XML part of the fixture
// corpus. Element names, canonical namespace URIs, attributes (order, names, URIs, values), text,
// comments and processing instructions of the root subtree must agree.

import { readFileSync } from 'node:fs';
import {
  DOMParser,
  type Element as XmElement,
  type Node as XmNode,
  type ProcessingInstruction as XmPI,
} from '@xmldom/xmldom';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  canonicalNs,
  NS,
  parseXml,
  parseXmlBytes,
  type XElement,
  type XNode,
} from '../src/xml/index.js';
import { fixtures as curatedFixtures } from './helpers/fixtures.js';

function fixtureParts(): { name: string; bytes: Uint8Array }[] {
  // The curated corpus (test/fixtures/expected.json); feature-specific fixtures are excluded.
  const files = curatedFixtures().map((f) => f.path);
  const out: { name: string; bytes: Uint8Array }[] = [];
  for (const f of files) {
    let entries: Record<string, Uint8Array>;
    try {
      entries = unzipSync(readFileSync(f));
    } catch {
      continue;
    }
    for (const [name, bytes] of Object.entries(entries)) {
      if (/\.(xml|rels)$/i.test(name)) out.push({ name: `${f}:${name}`, bytes });
    }
  }
  return out;
}

function xmldom(text: string): XmElement {
  const errors: string[] = [];
  const doc = new DOMParser({
    // XML 1.0 line-ending rules (xmldom's default also folds the XML 1.1 NEL/LS/PS characters).
    normalizeLineEndings: (s: string) => s.replace(/\r\n?/g, '\n'),
    onError: (level: string, msg: string) => {
      if (level !== 'warning') errors.push(msg);
    },
  }).parseFromString(text, 'text/xml');
  if (errors.length > 0) throw new Error(`xmldom: ${errors.join('; ')}`);
  return doc.documentElement!;
}

/** Children with adjacent text/CDATA merged into one string. */
type Item = XmNode | string;
function xmKids(n: XmNode): Item[] {
  const out: Item[] = [];
  for (let k = n.firstChild; k !== null; k = k.nextSibling) {
    if (k.nodeType === 3 || k.nodeType === 4) {
      const t = k.nodeValue ?? '';
      if (typeof out[out.length - 1] === 'string') out[out.length - 1] += t;
      else out.push(t);
    } else out.push(k);
  }
  return out;
}

/** First difference between the two trees, or undefined when they agree. */
function compare(ours: XElement, theirs: XmElement): string | undefined {
  const stack: [XNode, Item][] = [[ours, theirs]];
  while (stack.length > 0) {
    const [a, b] = stack.pop()!;
    if (typeof b === 'string') {
      if (a.nodeType !== 3 || a.text !== b) return `text ${JSON.stringify(b.slice(0, 40))}`;
      continue;
    }
    if (a.nodeType !== b.nodeType) return `node type ${a.nodeType} vs ${b.nodeType}`;
    if (a.nodeType === 8) {
      if (a.text !== b.nodeValue) return 'comment';
      continue;
    }
    if (a.nodeType === 7) {
      const pi = b as XmPI;
      if (a.target !== pi.target || a.data !== pi.data) return `PI ${pi.target}`;
      continue;
    }
    if (a.nodeType !== 1) return `unexpected node ${a.nodeType}`;
    const e = b as XmElement;
    if (a.name !== e.tagName) return `name ${a.name} vs ${e.tagName}`;
    if (a.uri !== canonicalNs(e.namespaceURI ?? '')) return `uri of ${a.name}`;
    const attrs = a.attributes();
    if (attrs.length !== e.attributes.length) return `attribute count on ${a.name}`;
    for (let i = 0; i < attrs.length; i++) {
      const x = attrs[i]!;
      const y = e.attributes.item(i)!;
      if (x.name !== y.name) return `attribute ${x.name} vs ${y.name} on ${a.name}`;
      if (x.value !== y.value) return `value of ${x.name} on ${a.name}`;
      if (x.uri !== canonicalNs(y.namespaceURI ?? '')) return `uri of ${x.name} on ${a.name}`;
    }
    const ka = mergedKids(a);
    const kb = xmKids(e);
    if (ka.length !== kb.length) return `child count of ${a.name}: ${ka.length} vs ${kb.length}`;
    for (let i = 0; i < ka.length; i++) stack.push([ka[i]!, kb[i]!]);
  }
  return undefined;
}

/** Our children (the parser already merges adjacent character data; keep the check honest). */
function mergedKids(e: XElement): XNode[] {
  for (let i = 1; i < e.children.length; i++) {
    if (e.children[i]!.nodeType === 3 && e.children[i - 1]!.nodeType === 3) {
      throw new Error(`adjacent text nodes under ${e.name}`);
    }
  }
  return e.children as XNode[];
}

describe('differential vs @xmldom/xmldom', () => {
  it('agrees on a synthetic part exercising the tricky bits', () => {
    const x = `<?xml version="1.0"?>\r\n<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns="urn:d"><!-- c --><w:body a='1' b="x&#10;y&#9;z	q
r"><w:t xml:space="preserve"> a &amp; b &lt;&gt; &#x1F600;\r\n</w:t><?pi data?><x><![CDATA[<raw>]]>tail</x><y xmlns=""/></w:body></w:document>\r\n`;
    expect(compare(parseXml(x).root, xmldom(x))).toBeUndefined();
  });

  const parts = fixtureParts();
  it.skipIf(parts.length === 0)(`agrees on every fixture XML part (${parts.length})`, () => {
    for (const { name, bytes } of parts) {
      const doc = parseXmlBytes(bytes, { part: name });
      const text = strFromU8(bytes.subarray(doc.bom ? 3 : 0));
      expect(compare(doc.root, xmldom(text)), name).toBeUndefined();
    }
  });
});
