import { readFileSync } from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { XmlParseError } from '../src/core/errors.js';
import { DEFAULT_LIMITS } from '../src/core/limits.js';
import {
  alternateContent,
  domEqual,
  ignorablePrefixes,
  NS,
  parseXml,
  parseXmlBytes,
  selectBranch,
  serializeNode,
  serializeXml,
  serializeXmlBytes,
  setIgnorable,
  XComment,
  type XElement,
  XElement as XElementClass,
  XProcessingInstruction,
  XText,
} from '../src/xml/index.js';
import { fixtures as curatedFixtures } from './helpers/fixtures.js';

const STRICT_W = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
const STRICT_R = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const rt = (x: string): string => serializeXml(parseXml(x));
const root = (x: string): XElement => parseXml(x).root;

function parseError(x: string, opts?: { maxDepth?: number; part?: string }): XmlParseError {
  try {
    parseXml(x, opts);
  } catch (e) {
    if (e instanceof XmlParseError) return e;
    throw e;
  }
  throw new Error(`expected XmlParseError for ${JSON.stringify(x.slice(0, 60))}`);
}

describe('parseXml: structure', () => {
  it('keeps the declaration verbatim and the whitespace around the root raw', () => {
    const x = `${DECL}\r\n<a/>\r\n`;
    const doc = parseXml(x);
    expect(doc.declaration).toBe(DECL);
    expect(doc.children.map((k) => k.nodeType)).toEqual([3, 1, 3]);
    expect((doc.children[0] as XText).text).toBe('\r\n');
    expect((doc.children[2] as XText).text).toBe('\r\n');
    expect(serializeXml(doc)).toBe(x);
  });

  it('has no declaration when there is none, and a leading U+FEFF sets bom', () => {
    expect(parseXml('<a/>').declaration).toBeNull();
    const doc = parseXml('﻿<?xml version="1.0"?><a/>');
    expect(doc.bom).toBe(true);
    expect(doc.declaration).toBe('<?xml version="1.0"?>');
  });

  it('builds a clean tree with parent pointers', () => {
    const doc = parseXml('<a><b>t</b><!--c--><?p d?><c/></a>');
    expect(doc.dirty).toBe(false);
    const a = doc.root;
    expect(a.parent).toBe(doc);
    for (const k of a.children) expect(k.parent).toBe(a);
    const b = a.firstElement()!;
    expect(b.children[0]!.parent).toBe(b);
    expect(a.children.map((k) => k.nodeType)).toEqual([1, 8, 7, 1]);
  });

  it('keeps comments and processing instructions everywhere', () => {
    const x = '<?xml version="1.0"?><!-- head --><?pi one?><a><!--in--><?t?></a><!--tail-->';
    const doc = parseXml(x);
    const [c1, p1, , c2] = doc.children;
    expect((c1 as XComment).text).toBe(' head ');
    expect((p1 as XProcessingInstruction).target).toBe('pi');
    expect((p1 as XProcessingInstruction).data).toBe('one');
    expect((c2 as XComment).text).toBe('tail');
    const pi = doc.root.children[1] as XProcessingInstruction;
    expect(pi.target).toBe('t');
    expect(pi.data).toBe('');
    expect(serializeXml(doc)).toBe(x);
  });

  it('keeps all whitespace text inside the root', () => {
    const a = root('<a>\n  <b> </b>\t</a>');
    expect(a.children.map((k) => (k.nodeType === 3 ? k.text : k.nodeType))).toEqual([
      '\n  ',
      1,
      '\t',
    ]);
    expect(a.firstElement()!.text).toBe(' ');
  });

  it('turns CDATA into text merged with adjacent character data', () => {
    const a = root('<a>x<![CDATA[<y> & ]]>z</a>');
    expect(a.children).toHaveLength(1);
    expect(a.text).toBe('x<y> & z');
    expect(serializeNode(a)).toBe('<a>x&lt;y&gt; &amp; z</a>');
  });

  it('skips DOCTYPE (internal subset included) without expanding entities', () => {
    const x =
      '<?xml version="1.0"?>\n<!DOCTYPE a [ <!ENTITY xxe SYSTEM "file:///etc/passwd"> <!-- ] --> ]>\n<a>&xxe;</a>';
    const doc = parseXml(x);
    expect(doc.root.text).toBe('&xxe;');
    expect(serializeXml(doc)).toBe('<?xml version="1.0"?>\n\n<a>&amp;xxe;</a>');
  });

  it('accepts single-quoted attributes and whitespace inside tags', () => {
    const a = root(`<a  x = 'v"1'\n y="2" ></a >`);
    expect(a.attr('x')).toBe('v"1');
    expect(a.attr('y')).toBe('2');
    expect(serializeNode(a)).toBe('<a x="v&quot;1" y="2"/>');
  });
});

describe('parseXml: character data', () => {
  it('decodes the predefined entities and character references', () => {
    const a = root(
      '<a v="&lt;&gt;&amp;&quot;&apos;&#65;&#x42;&#x1F600;">&lt;&gt;&amp;&quot;&apos;&#65;&#x42;</a>',
    );
    expect(a.attr('v')).toBe(`<>&"'AB\u{1F600}`);
    expect(a.text).toBe(`<>&"'AB`);
  });

  it('keeps unknown entities, bare ampersands and invalid references literally', () => {
    expect(root('<a>&nbsp; a & b &#0; &#xD800; &#xZZ;</a>').text).toBe(
      '&nbsp; a & b &#0; &#xD800; &#xZZ;',
    );
  });

  it('normalizes line endings in text (CRLF and CR to LF)', () => {
    expect(root('<a>1\r\n2\r3\n4</a>').text).toBe('1\n2\n3\n4');
    expect(root('<a><![CDATA[1\r\n2]]></a>').text).toBe('1\n2');
    expect(root('<a><!--1\r\n2--></a>').children[0]).toMatchObject({ text: '1\n2' });
  });

  it('normalizes attribute whitespace but keeps character references', () => {
    const a = root('<a v="1\r\n2\t3\n4\r5" w="&#9;&#10;&#13;"/>');
    expect(a.attr('v')).toBe('1 2 3 4 5');
    expect(a.attr('w')).toBe('\t\n\r');
    expect(serializeNode(a)).toBe('<a v="1 2 3 4 5" w="&#9;&#10;&#13;"/>');
  });

  it('keeps a CR from a character reference in text across round-trips', () => {
    const x = '<a>1&#13;2&#xD;&#10;3</a>';
    const doc = parseXml(x);
    expect(doc.root.text).toBe('1\r2\r\n3');
    const out = serializeXml(doc);
    expect(out).toBe('<a>1&#13;2&#13;\n3</a>');
    expect(parseXml(out).root.text).toBe('1\r2\r\n3');
  });
});

describe('parseXml: namespaces', () => {
  it('resolves prefixes to canonical URIs and keeps qualified names', () => {
    const a = root(
      `<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:body><w:p r:id="1" w:val="2" plain="3" xml:space="preserve"/></w:body></w:document>`,
    );
    expect(a.name).toBe('w:document');
    expect(a.local).toBe('document');
    expect(a.uri).toBe(NS.w);
    expect(a._attrs).toEqual(['xmlns:w', NS.xmlns, NS.w, 'xmlns:r', NS.xmlns, NS.r]);
    const p = a.descendants(NS.w, 'p')[0]!;
    expect(p._attrs).toEqual([
      'r:id',
      NS.r,
      '1',
      'w:val',
      NS.w,
      '2',
      'plain',
      '',
      '3',
      'xml:space',
      NS.xml,
      'preserve',
    ]);
    expect(p.attrNS(NS.r, 'id')).toBe('1');
  });

  it('aliases Strict URIs to their transitional twins', () => {
    const a = root(
      `<w:document xmlns:w="${STRICT_W}" xmlns:r="${STRICT_R}"><w:p r:id="x"/></w:document>`,
    );
    expect(a.uri).toBe(NS.w);
    expect(a.element(NS.w, 'p')!.attrNS(NS.r, 'id')).toBe('x');
    // the declaration itself keeps the Strict URI
    expect(a.attr('xmlns:w')).toBe(STRICT_W);
  });

  it('handles the default namespace and its undeclaration', () => {
    const a = root(
      `<Types xmlns="${NS.ct}"><Default Extension="xml"/><x xmlns=""><y/></x></Types>`,
    );
    expect(a.uri).toBe(NS.ct);
    expect(a.firstElement()!.uri).toBe(NS.ct);
    expect(a.firstElement()!.attrNS('', 'Extension')).toBe('xml');
    const x = a.lastElement()!;
    expect(x.uri).toBe('');
    expect(x.firstElement()!.uri).toBe('');
  });

  it('scopes declarations to their element (nested, shadowed, self-closing)', () => {
    const a = root(
      '<p:a xmlns:p="urn:1"><p:b xmlns:p="urn:2"><p:c/></p:b><p:d xmlns:p="urn:3"/><p:e/></p:a>',
    );
    const [b, d, e] = a.elements();
    expect(a.uri).toBe('urn:1');
    expect(b!.uri).toBe('urn:2');
    expect(b!.firstElement()!.uri).toBe('urn:2');
    expect(d!.uri).toBe('urn:3');
    expect(e!.uri).toBe('urn:1');
  });

  it('applies declarations of an element to its own name and attributes', () => {
    const a = root('<q:a q:x="1" xmlns:q="urn:q"/>');
    expect(a.uri).toBe('urn:q');
    expect(a._attrs.slice(0, 3)).toEqual(['q:x', 'urn:q', '1']);
  });

  it('resolves undeclared prefixes leniently: conventional URI, else no namespace', () => {
    const a = root('<w:p w14:paraId="1" zz:x="2"><zz:y/></w:p>');
    expect(a.uri).toBe(NS.w);
    expect(a.attrNS(NS.w14, 'paraId')).toBe('1');
    expect(a._attrs[4]).toBe('');
    expect(a.firstElement()!.uri).toBe('');
  });
});

describe('parseXml: errors and limits', () => {
  it.each([
    ['<a><b></a></b>', 'Mismatched end tag'],
    ['<a><b>', 'Unclosed element <b>'],
    ['<a/><b/>', 'More than one root element'],
    ['x<a/>', 'Text outside the root element'],
    ['<a/>x', 'Text outside the root element'],
    ['<a/>&amp;', 'Text outside the root element'],
    ['</a>', 'End tag without a matching start tag'],
    ['<a><!-- x</a>', 'Unterminated comment'],
    ['<a><![CDATA[ x</a>', 'Unterminated CDATA'],
    ['<a><?pi x</a>', 'Unterminated processing instruction'],
    ['<?xml version="1.0"', 'Unterminated XML declaration'],
    ['<a b="1"', 'Unterminated start tag'],
    ['<a b="1/>', 'Unterminated attribute value'],
    ['<a b/>', 'has no value'],
    ['<a b=1/>', 'not quoted'],
    ['<a b="1" b="2"/>', 'Duplicate attribute'],
    ['<w:a xmlns:w="urn:w"><w:b xmlns:w="" w:c="1"/></w:a>', 'Empty namespace declaration xmlns:w'],
    ['<a></a', 'Unterminated end tag'],
    ['<![CDATA[x]]><a/>', 'CDATA outside the root element'],
    ['<a/><!DOCTYPE a>', 'DOCTYPE after the root element'],
    ['<a><!ELEMENT x></a>', 'Invalid markup'],
    ['<a>< b/></a>', 'Invalid markup'],
    [' <?xml version="1.0"?><a/>', 'XML declaration not at the start'],
    ['', 'No root element'],
    ['  <!-- only -->  ', 'No root element'],
  ])('%j → %s', (x, msg) => {
    const e = parseError(x);
    expect(e.message).toContain(msg);
    expect(e.code).toBe('XML_PARSE');
  });

  it('reports the offset and the part name', () => {
    const e = parseError('<a>\n<b></c></a>', { part: '/word/document.xml' });
    expect(e.offset).toBe(7);
    expect(e.part).toBe('/word/document.xml');
    expect(e.message).toContain('/word/document.xml');
  });

  it('enforces maxDepth (default DEFAULT_LIMITS.maxXmlDepth)', () => {
    const nest = (d: number): string => `${'<a>'.repeat(d - 1)}<a/>${'</a>'.repeat(d - 1)}`;
    expect(() => parseXml(nest(DEFAULT_LIMITS.maxXmlDepth))).not.toThrow();
    expect(parseError(nest(DEFAULT_LIMITS.maxXmlDepth + 1)).message).toContain('deeper than');
    expect(() => parseXml(nest(5), { maxDepth: 5 })).not.toThrow();
    expect(parseError(nest(6), { maxDepth: 5 }).message).toContain('deeper than 5');
  });

  it('parses and serializes very deep documents without recursion', () => {
    const d = 200_000;
    const x = `${'<a>'.repeat(d)}t${'</a>'.repeat(d)}`;
    const doc = parseXml(x, { maxDepth: d });
    expect(serializeXml(doc)).toBe(x);
    expect(domEqual(doc, doc.clone())).toBe(true);
  });
});

describe('parseXmlBytes', () => {
  const enc = new TextEncoder();

  it('decodes UTF-8 and records a BOM', () => {
    const body = enc.encode(`${DECL}<a>é</a>`);
    const withBom = new Uint8Array([0xef, 0xbb, 0xbf, ...body]);
    const d1 = parseXmlBytes(withBom);
    expect(d1.bom).toBe(true);
    expect(d1.root.text).toBe('é');
    expect(serializeXmlBytes(d1)).toEqual(withBom);
    const d2 = parseXmlBytes(body);
    expect(d2.bom).toBe(false);
    expect(serializeXmlBytes(d2)).toEqual(body);
  });

  it('decodes UTF-16LE/BE by BOM and re-encodes as UTF-8 without a BOM', () => {
    const x = '<?xml version="1.0" encoding="UTF-16"?><a>é€</a>';
    const le = new Uint8Array(2 + x.length * 2);
    const be = new Uint8Array(2 + x.length * 2);
    le.set([0xff, 0xfe]);
    be.set([0xfe, 0xff]);
    for (let i = 0; i < x.length; i++) {
      const c = x.charCodeAt(i);
      le[2 + 2 * i] = c & 0xff;
      le[3 + 2 * i] = c >> 8;
      be[2 + 2 * i] = c >> 8;
      be[3 + 2 * i] = c & 0xff;
    }
    for (const bytes of [le, be]) {
      const doc = parseXmlBytes(bytes);
      expect(doc.bom).toBe(false);
      expect(doc.root.text).toBe('é€');
      expect(new TextDecoder().decode(serializeXmlBytes(doc))).toBe(
        '<?xml version="1.0" encoding="UTF-8"?><a>é€</a>',
      );
    }
  });

  it('honours a declared single-byte encoding', () => {
    const bytes = new Uint8Array([
      ...enc.encode("<?xml version='1.0' encoding='windows-1252'?><a>"),
      0xe9,
      0x80,
      ...enc.encode('</a>'),
    ]);
    const doc = parseXmlBytes(bytes);
    expect(doc.root.text).toBe('é€');
    expect(serializeXml(doc)).toBe("<?xml version='1.0' encoding='UTF-8'?><a>é€</a>");
  });
});

describe('serializeXml', () => {
  it('writes Word style: <x/>, double quotes, stored attribute order', () => {
    expect(rt(`<a z='1' b="2" a="3"><b></b><c> </c></a>`)).toBe(
      '<a z="1" b="2" a="3"><b/><c> </c></a>',
    );
  });

  it('escapes text and attributes', () => {
    const doc = parseXml('<a/>');
    doc.root.setAttr('v', `<&>"'\t\n\r`);
    doc.root.text = `<&>"'\t\n\r\u0001`;
    expect(serializeXml(doc)).toBe(
      `<a v="&lt;&amp;&gt;&quot;'&#9;&#10;&#13;">&lt;&amp;&gt;"'\t\n&#13;</a>`,
    );
  });

  it('rewrites a non-UTF-8 encoding in the declaration only', () => {
    expect(rt('<?xml version="1.0" encoding="ISO-8859-1" standalone="yes"?><a/>')).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a/>',
    );
    expect(rt('<?xml version="1.0" encoding="utf-8"?><a/>')).toBe(
      '<?xml version="1.0" encoding="utf-8"?><a/>',
    );
    expect(rt('<?xml version="1.0"?><a/>')).toBe('<?xml version="1.0"?><a/>');
  });

  it('serializes fragments without namespace fix-up', () => {
    const doc = parseXml(
      `<w:document xmlns:w="${NS.w}"><w:p><w:r><w:t>x</w:t></w:r></w:p></w:document>`,
    );
    const p = doc.root.firstElement()!;
    expect(serializeNode(p)).toBe('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
    expect(serializeNode(new XText('a<b'))).toBe('a&lt;b');
    expect(serializeNode(new XComment(' c '))).toBe('<!-- c -->');
    expect(serializeNode(new XProcessingInstruction('t', 'd'))).toBe('<?t d?>');
  });

  it('round-trips byte-exactly and is idempotent', () => {
    const x = `${DECL}\r\n<w:document xmlns:w="${NS.w}" xmlns:mc="${NS.mc}" xmlns:w14="${NS.w14}" mc:Ignorable="w14"><w:body><w:p w14:paraId="1A2B3C4D"><w:r><w:t xml:space="preserve"> a &amp; b </w:t></w:r></w:p></w:body></w:document>`;
    expect(rt(x)).toBe(x);
    expect(rt(rt(x))).toBe(rt(x));
  });
});

describe('serializeXml: namespace fix-up', () => {
  const W = `<w:document xmlns:w="${NS.w}"><w:body><w:p/></w:body></w:document>`;

  it('declares a missing prefix on the root and makes it ignorable', () => {
    const doc = parseXml(W);
    const p = doc.root.descendants(NS.w, 'p')[0]!;
    p.setAttrNS(NS.w14, 'paraId', '00000001');
    const out = serializeXml(doc);
    expect(out).toBe(
      `<w:document xmlns:w="${NS.w}" xmlns:w14="${NS.w14}" xmlns:mc="${NS.mc}" mc:Ignorable="w14"><w:body><w:p w14:paraId="00000001"/></w:body></w:document>`,
    );
    const back = parseXml(out);
    expect(back.root.descendants(NS.w, 'p')[0]!.attrNS(NS.w14, 'paraId')).toBe('00000001');
    expect(rt(out)).toBe(out);
  });

  it('appends to an existing mc:Ignorable, keeping its prefix', () => {
    const doc = parseXml(
      `<w:document xmlns:m2="${NS.mc}" xmlns:w="${NS.w}" xmlns:w15="${NS.w15}" m2:Ignorable="w15"><w:body/></w:document>`,
    );
    const body = doc.root.firstElement()!;
    body.append(doc.createElement(NS.w14, 'checkbox'));
    body.append(doc.createElement(NS.r, 'x')); // not an extension namespace
    expect(serializeXml(doc)).toBe(
      `<w:document xmlns:m2="${NS.mc}" xmlns:w="${NS.w}" xmlns:w15="${NS.w15}" m2:Ignorable="w15 w14" xmlns:w14="${NS.w14}" xmlns:r="${NS.r}"><w:body><w14:checkbox/><r:x/></w:body></w:document>`,
    );
  });

  it('uses Strict twins when the root declares Strict URIs', () => {
    const doc = parseXml(`<w:document xmlns:w="${STRICT_W}"><w:body/></w:document>`);
    const body = doc.root.firstElement()!;
    const sect = doc.createElement(NS.w, 'headerReference');
    sect.setAttrNS(NS.r, 'id', 'rId1');
    body.append(sect);
    expect(serializeXml(doc)).toBe(
      `<w:document xmlns:w="${STRICT_W}" xmlns:r="${STRICT_R}"><w:body><w:headerReference r:id="rId1"/></w:body></w:document>`,
    );
  });

  it('drops undeclared prefixes from the root mc:Ignorable and mc:ProcessContent', () => {
    const x = `<w:document xmlns:w="${NS.w}" xmlns:mc="${NS.mc}" xmlns:w14="${NS.w14}" mc:Ignorable="w14 w15 wp14" mc:ProcessContent="w14:x w15:y"><w:body/></w:document>`;
    expect(rt(x)).toBe(
      `<w:document xmlns:w="${NS.w}" xmlns:mc="${NS.mc}" xmlns:w14="${NS.w14}" mc:Ignorable="w14" mc:ProcessContent="w14:x"><w:body/></w:document>`,
    );
    const y = `<a xmlns:mc="${NS.mc}" mc:Ignorable="w15"/>`;
    expect(rt(y)).toBe(`<a xmlns:mc="${NS.mc}"/>`);
  });

  it('re-declares a prefix bound to another namespace on the element itself', () => {
    const doc = parseXml('<p:a xmlns:p="urn:1"><p:b/></p:a>');
    doc.root.append(new XElementClass('p:c', 'urn:2'));
    const out = serializeXml(doc);
    expect(out).toBe('<p:a xmlns:p="urn:1"><p:b/><p:c xmlns:p="urn:2"/></p:a>');
    expect(parseXml(out).root.lastElement()!.uri).toBe('urn:2');
  });

  it('re-declares the default namespace where it differs', () => {
    const doc = parseXml(`<Types xmlns="${NS.ct}"/>`);
    const x = new XElementClass('x', '');
    x.append(new XElementClass('y', 'urn:y'));
    doc.root.append(x);
    const out = serializeXml(doc);
    expect(out).toBe(`<Types xmlns="${NS.ct}"><x xmlns=""><y xmlns="urn:y"/></x></Types>`);
    const back = parseXml(out).root;
    expect(back.firstElement()!.uri).toBe('');
    expect(back.firstElement()!.firstElement()!.uri).toBe('urn:y');
    // a root with a namespace and no declaration at all
    const d2 = parseXml('<r/>');
    d2.root.setName('Relationships', NS.rel);
    expect(serializeXml(d2)).toBe(`<Relationships xmlns="${NS.rel}"/>`);
  });

  it('does not modify the DOM', () => {
    const doc = parseXml(W);
    doc.root.firstElement()!.append(doc.createElement(NS.w14, 'x'));
    doc.dirty = false;
    const before = doc.root._attrs.slice();
    serializeXml(doc);
    expect(doc.root._attrs).toEqual(before);
    expect(doc.dirty).toBe(false);
  });

  it('fixes up each document independently (no state shared between calls)', () => {
    for (let i = 0; i < 2; i++) {
      const doc = parseXml('<a/>');
      doc.root.append(doc.createElement(NS.w14, 'x'));
      expect(serializeXml(doc)).toBe(
        `<a xmlns:w14="${NS.w14}" xmlns:mc="${NS.mc}" mc:Ignorable="w14"><w14:x/></a>`,
      );
    }
  });

  it('reuses an mc prefix it is already adding instead of declaring a second one', () => {
    const doc = parseXml('<a/>');
    const x = new XElementClass('mc:Fallback', NS.mc);
    doc.root.append(x, doc.createElement(NS.w14, 'y'));
    expect(serializeXml(doc)).toBe(
      `<a xmlns:mc="${NS.mc}" xmlns:w14="${NS.w14}" mc:Ignorable="w14"><mc:Fallback/><w14:y/></a>`,
    );
  });

  it('never re-adds a prefix the root declares; omits xmlns:p="" built through the DOM', () => {
    const doc = parseXml('<w:a xmlns:w="urn:w"/>');
    doc.root.append(
      new XElementClass('w:b', 'urn:w', ['xmlns:w', NS.xmlns, '', 'w:c', 'urn:w', '1']),
    );
    const out = serializeXml(doc);
    expect(out).toBe('<w:a xmlns:w="urn:w"><w:b w:c="1" xmlns:w="urn:w"/></w:a>');
    expect(rt(out)).toBe(out);
    const root2 = new XElementClass('w:a', 'urn:w', ['xmlns:w', NS.xmlns, '']);
    expect(serializeNode(root2)).toBe('<w:a/>');
  });

  it('leaves prefixes it cannot declare alone (undeclared unknown prefix)', () => {
    expect(rt('<zz:a/>')).toBe('<zz:a/>');
  });
});

describe('mc helpers', () => {
  const AC = `<w:r xmlns:w="${NS.w}" xmlns:mc="${NS.mc}" xmlns:wps="${NS.wps}" xmlns:v="${NS.v}"><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing/></mc:Choice><mc:Choice Requires="wps v zz"><w:x/></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>`;

  it('reads AlternateContent branches', () => {
    const ac = alternateContent(root(AC).firstElement()!);
    expect(ac.choices.map((c) => c.requires)).toEqual([['wps'], ['wps', 'v', 'zz']]);
    expect(ac.fallback?.local).toBe('Fallback');
    expect(() => alternateContent(root(AC))).toThrow(/mc:AlternateContent/);
  });

  it('selects the first understood branch, else the fallback', () => {
    const ac = alternateContent(root(AC).firstElement()!);
    expect(selectBranch(ac, () => true)).toBe(ac.choices[0]!.el);
    expect(selectBranch(ac, (u) => u === NS.v)).toBe(ac.fallback);
    expect(selectBranch({ choices: ac.choices }, (u) => u === NS.v)).toBeUndefined();
  });

  it('reads and writes mc:Ignorable', () => {
    const doc = parseXml(`<w:document xmlns:w="${NS.w}" xmlns:w14="${NS.w14}"/>`);
    expect(ignorablePrefixes(doc.root)).toEqual([]);
    setIgnorable(doc.root, ['w14', 'w14', '']);
    expect(ignorablePrefixes(doc.root)).toEqual(['w14']);
    expect(doc.dirty).toBe(true);
    expect(serializeXml(doc)).toBe(
      `<w:document xmlns:w="${NS.w}" xmlns:w14="${NS.w14}" mc:Ignorable="w14" xmlns:mc="${NS.mc}"/>`,
    );
    setIgnorable(doc.root, []);
    expect(doc.root.attrNS(NS.mc, 'Ignorable')).toBeUndefined();
  });
});

// ---- Word-saved fixtures ------------------------------------------------------------------

/** Every XML part of every .docx under the fixture directories (missing directories skipped). */
function fixtureParts(): { name: string; bytes: Uint8Array }[] {
  // The curated corpus (test/fixtures/expected.json); feature-specific fixtures are excluded.
  const files = curatedFixtures().map((f) => f.path);
  const out: { name: string; bytes: Uint8Array }[] = [];
  for (const f of files) {
    let entries: Record<string, Uint8Array>;
    try {
      entries = unzipSync(readFileSync(f));
    } catch {
      continue; // deliberately broken fixtures belong to other tests
    }
    for (const [name, bytes] of Object.entries(entries)) {
      if (/\.(xml|rels)$/i.test(name)) out.push({ name: `${f}:${name}`, bytes });
    }
  }
  return out;
}

/**
 * `text` with the spellings the DOM does not record normalized the way the serializer writes
 * them (losses by design): `<x></x>` → `<x/>`, ` />` → `/>`, hex character references → decimal,
 * CRLF inside the root → LF (whitespace around the root is kept raw).
 */
function normalizeSpelling(text: string): string {
  const m = /^(<\?xml[^>]*\?>)?(\s*)([\s\S]*?)(\s*)$/.exec(text)!;
  const body = m[3]!
    .replace(/\r\n?/g, '\n')
    .replace(/&#x([0-9a-f]+);/gi, (_, h: string) => `&#${Number.parseInt(h, 16)};`)
    .replace(/\s+\/>/g, '/>')
    .replace(/<([^\s/>!?]+)((?:\s+[^\s=]+\s*=\s*"[^"]*")*)\s*><\/\1>/g, '<$1$2/>');
  return (m[1] ?? '') + m[2] + body + m[4];
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((v, i) => v === b[i]);
}

const parts = fixtureParts();
const WORD_SAVED = '/turbomerge/test/fixtures/';

describe.skipIf(parts.length === 0)('fixtures', () => {
  it('round-trip every XML part DOM-equal, byte-exact up to the documented spelling losses', () => {
    let identical = 0;
    let word = 0;
    let wordIdentical = 0;
    for (const { name, bytes } of parts) {
      const text = strFromU8(bytes.subarray(bytes[0] === 0xef ? 3 : 0));
      const doc = parseXmlBytes(bytes, { part: name });
      expect(doc.dirty, name).toBe(false);
      const out = serializeXml(doc);
      expect(domEqual(parseXml(out), doc), name).toBe(true);
      expect(out, name).toBe(normalizeSpelling(text));
      const isWord = name.replaceAll('\\', '/').includes(WORD_SAVED);
      if (isWord) word++;
      if (out === text) {
        identical++;
        if (isWord) wordIdentical++;
      }
    }
    const pct = (a: number, b: number): string => `${a}/${b} (${((100 * a) / b).toFixed(1)}%)`;
    console.log(
      `[xml] byte-identical re-serialization: all fixtures ${pct(identical, parts.length)}` +
        (word > 0 ? `, Word-saved (turbomerge) ${pct(wordIdentical, word)}` : '') +
        '; DOM-equal 100%',
    );
  });

  it('re-serialization is idempotent', () => {
    for (const { name, bytes } of parts) {
      const once = serializeXml(parseXmlBytes(bytes));
      expect(serializeXml(parseXml(once)), name).toBe(once);
    }
  });

  it('serializeXmlBytes(parseXmlBytes(b)) is b itself when nothing needs normalizing', () => {
    for (const { name, bytes } of parts) {
      const doc = parseXmlBytes(bytes);
      const text = strFromU8(bytes.subarray(doc.bom ? 3 : 0));
      if (serializeXml(doc) !== text) continue;
      expect(sameBytes(serializeXmlBytes(doc), bytes), name).toBe(true);
    }
  });
});
