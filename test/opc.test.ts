import { describe, expect, it } from 'vitest';
import { DocxFormatError, UsageError } from '../src/core/errors.js';
import type { Warning } from '../src/core/warnings.js';
import {
  CT,
  normalizePartName,
  OpcPackage,
  partDir,
  REL,
  relativeTarget,
  relsPathFor,
  relTypeIs,
  resolveTarget,
  sourceOfRels,
} from '../src/opc/index.js';
import { NS } from '../src/xml/ns.js';
import { parseXml } from '../src/xml/parse.js';
import { makeEntry, unzipLazy, writeZip, ZipLimitError } from '../src/zip/index.js';

const enc = new TextEncoder();
const dec = new TextDecoder();
const xml = (s: string) => parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${s}`);

const W = `xmlns:w="${NS.w}" xmlns:r="${NS.r}"`;
const DOCUMENT = `<w:document ${W}><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr></w:body></w:document>`;
const HEADER = `<w:hdr ${W}><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`;
const STYLES = `<w:styles ${W}/>`;
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

/** A small package built through the API: document → styles, header → image, plus core props. */
function buildPackage(): OpcPackage {
  const pkg = OpcPackage.create();
  pkg.contentTypes.setDefault('png', 'image/png');
  const doc = pkg.addPart('word/document.xml', CT.document, xml(DOCUMENT));
  pkg.addPart('word/styles.xml', CT.styles, xml(STYLES));
  const hdr = pkg.addPart('word/header1.xml', CT.header, enc.encode(HEADER));
  pkg.addPart('word/media/image1.png', 'image/png', PNG);
  pkg.addPart(
    'docProps/core.xml',
    CT.coreProperties,
    xml(`<cp:coreProperties xmlns:cp="${NS.cp}"/>`),
  );
  pkg.rels.add(REL.officeDocument, 'word/document.xml');
  pkg.rels.add(REL.coreProperties, 'docProps/core.xml');
  doc.rels.add(REL.styles, 'styles.xml');
  doc.rels.add(REL.header, 'header1.xml');
  hdr.rels.add(REL.image, 'media/image1.png');
  doc.rels.add(REL.hyperlink, 'https://example.com/a b', { external: true });
  return pkg;
}

const buildBytes = (): Uint8Array => buildPackage().save();

/** Builds a zip directly from name → content (for malformed / unusual packages). */
function zipOf(files: Record<string, string | Uint8Array>): Uint8Array {
  return writeZip(
    Object.entries(files).map(([name, data]) =>
      makeEntry(name, typeof data === 'string' ? enc.encode(data) : data),
    ),
  );
}

const TYPES = `<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="${CT.relationships}"/><Default Extension="xml" ContentType="application/xml"/>`;

describe('uri helpers', () => {
  it('resolves relationship targets', () => {
    expect(resolveTarget('word/', '../media/x.png')).toBe('media/x.png');
    expect(resolveTarget('word/', 'media/x.png')).toBe('word/media/x.png');
    expect(resolveTarget('word/', '/word/x.xml')).toBe('word/x.xml');
    expect(resolveTarget('word/', 'media/my%20image.png')).toBe('word/media/my image.png');
    expect(resolveTarget('word/', 'x.xml#frag')).toBe('word/x.xml');
    expect(resolveTarget('word/', './a/./b/../c.xml')).toBe('word/a/c.xml');
    expect(resolveTarget('word', 'x.xml')).toBe('word/x.xml');
    expect(resolveTarget('', 'word/document.xml')).toBe('word/document.xml');
    expect(resolveTarget('word/', '../../../../etc/passwd')).toBe('etc/passwd');
    expect(resolveTarget('word/', 'bad%E0%A4%A.xml')).toBe('word/bad%E0%A4%A.xml');
    expect(resolveTarget('word/', 'a%23b.xml')).toBe('word/a#b.xml');
  });

  it('maps parts to rels parts and back', () => {
    expect(partDir('word/document.xml')).toBe('word/');
    expect(partDir('a.xml')).toBe('');
    expect(relsPathFor('word/document.xml')).toBe('word/_rels/document.xml.rels');
    expect(relsPathFor('')).toBe('_rels/.rels');
    expect(relsPathFor('/a.xml')).toBe('_rels/a.xml.rels');
    expect(sourceOfRels('word/_rels/document.xml.rels')).toBe('word/document.xml');
    expect(sourceOfRels('/_rels/.rels')).toBe('');
    expect(sourceOfRels('_rels/a.xml.rels')).toBe('a.xml');
    expect(sourceOfRels('word/document.xml')).toBeUndefined();
    expect(sourceOfRels('word/foo_rels/a.xml.rels')).toBeUndefined();
    expect(normalizePartName('/word/document.xml')).toBe('word/document.xml');
  });

  it('computes relative targets', () => {
    expect(relativeTarget('word/document.xml', 'word/media/a.png')).toBe('media/a.png');
    expect(relativeTarget('word/document.xml', 'customXml/item1.xml')).toBe(
      '../customXml/item1.xml',
    );
    expect(relativeTarget('', 'word/document.xml')).toBe('word/document.xml');
    expect(relativeTarget('word/a/b.xml', 'word/c.xml')).toBe('../c.xml');
    expect(relativeTarget('word/document.xml', 'word/my image#1.png')).toBe('my%20image%231.png');
    const t = relativeTarget('word/document.xml', 'media/my image#1.png');
    expect(resolveTarget('word/', t)).toBe('media/my image#1.png');
  });

  it('matches relationship types, transitional and Strict', () => {
    const strictHeader = 'http://purl.oclc.org/ooxml/officeDocument/relationships/header';
    const strictExt = 'http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties';
    expect(relTypeIs(REL.header, 'header')).toBe(true);
    expect(relTypeIs(strictHeader, 'header')).toBe(true);
    expect(relTypeIs(strictHeader, REL.header)).toBe(true);
    expect(relTypeIs(REL.header, strictHeader)).toBe(true);
    expect(relTypeIs(strictExt, REL.extendedProperties)).toBe(true);
    expect(relTypeIs(strictExt, 'extended-properties')).toBe(true);
    expect(relTypeIs(REL.footer, 'header')).toBe(false);
    expect(relTypeIs(REL.footer, REL.header)).toBe(false);
    expect(relTypeIs('http://example.com/x/header', REL.header)).toBe(false);
  });
});

describe('OpcPackage save', () => {
  it('returns the exact input bytes for a no-op save, even after reading parts', () => {
    const bytes = buildBytes();
    const pkg = OpcPackage.open(bytes);
    expect(pkg.save()).toBe(bytes);
    for (const part of pkg.parts()) {
      part.bytes();
      void part.contentType;
      if (part.isXml) part.xml();
      part.rels.all();
    }
    pkg.contentTypes.overrides();
    pkg.part('word/document.xml')!.relatedByType('header');
    expect(pkg.dirty).toBe(false);
    expect(pkg.parts().some((p) => p.dirty)).toBe(false);
    expect(pkg.save()).toBe(bytes);
  });

  it('copies clean entries raw after one part is edited', () => {
    const bytes = buildBytes();
    const pkg = OpcPackage.open(bytes);
    const doc = pkg.part('word/document.xml')!;
    const t = doc.xml().root.descendants(NS.w, 't')[0]!;
    t.text = 'Changed';
    expect(doc.dirty).toBe(true);
    expect(pkg.dirty).toBe(true);
    const out = pkg.save();
    expect(out).not.toBe(bytes);

    const before = unzipLazy(bytes);
    const after = unzipLazy(out);
    expect([...after.keys()]).toEqual([...before.keys()]);
    for (const [name, e] of before) {
      const a = after.get(name)!;
      if (name === 'word/document.xml') {
        expect(dec.decode(a.bytes())).toContain('Changed');
        continue;
      }
      expect(a.crc32).toBe(e.crc32);
      expect(a.compression).toBe(e.compression);
      expect(a.raw()).toEqual(e.raw());
    }
    // saving twice gives identical output (save does not reset dirty flags)
    expect(pkg.save()).toEqual(out);
    expect(doc.dirty).toBe(true);
  });

  it('writes created packages in order: content types, package rels, then the rest', () => {
    const names = [...unzipLazy(buildBytes()).keys()];
    expect(names.slice(0, 2)).toEqual(['[Content_Types].xml', '_rels/.rels']);
    expect(names).toContain('word/document.xml');
    expect(names).toContain('word/_rels/document.xml.rels');
    expect(names).toContain('word/_rels/header1.xml.rels');
  });

  it('reopens what it wrote', () => {
    const pkg = OpcPackage.open(buildBytes());
    const main = pkg.rels.byType('officeDocument')[0]!;
    const doc = pkg.part(main.partName!)!;
    expect(doc.contentType).toBe(CT.document);
    expect(doc.xml().root.text).toBe('Hello');
    const header = doc.related('rId2')!;
    expect(header.name).toBe('word/header1.xml');
    expect(header.relatedByType('image')[0]!.bytes()).toEqual(PNG);
    expect(pkg.part('word/media/image1.png')!.contentType).toBe('image/png');
    expect(pkg.part('word/media/image1.png')!.isXml).toBe(false);
    expect(doc.isXml).toBe(true);
    const link = doc.rels.byType('hyperlink')[0]!;
    expect(link.external).toBe(true);
    expect(link.partName).toBeUndefined();
    expect(doc.related(link.id)).toBeUndefined();
  });

  it('setBytes / setXml replace content and make the part dirty', () => {
    const bytes = buildBytes();
    const pkg = OpcPackage.open(bytes);
    const styles = pkg.part('word/styles.xml')!;
    styles.setBytes(enc.encode(`<w:styles ${W}><w:docDefaults/></w:styles>`));
    expect(styles.dirty).toBe(true);
    expect(styles.xml().root.elements().length).toBe(1);
    const header = pkg.part('word/header1.xml')!;
    header.setXml(xml(`<w:hdr ${W}/>`));
    expect(header.dirty).toBe(true);
    const re = OpcPackage.open(pkg.save());
    expect(re.part('word/styles.xml')!.xml().root.elements()[0]!.local).toBe('docDefaults');
    expect(re.part('word/header1.xml')!.xml().root.children.length).toBe(0);
  });
});

describe('part lookup', () => {
  it('is case-insensitive and ignores a leading slash', () => {
    const pkg = OpcPackage.open(buildBytes());
    const doc = pkg.part('word/document.xml')!;
    expect(pkg.part('WORD/Document.XML')).toBe(doc);
    expect(pkg.part('/word/document.xml')).toBe(doc);
    expect(pkg.part('[content_types].xml')!.name).toBe('[Content_Types].xml');
    expect(pkg.part('_RELS/.rels')).toBeDefined();
    expect(pkg.contentTypes.get('/WORD/DOCUMENT.XML')).toBe(CT.document);
    expect(pkg.part('word/missing.xml')).toBeUndefined();
  });

  it('uniqueName returns the first free number, case-insensitively', () => {
    const pkg = OpcPackage.open(buildBytes());
    expect(pkg.uniqueName('word/header{n}.xml')).toBe('word/header2.xml');
    pkg.addPart('word/HEADER2.xml', CT.header, xml(HEADER));
    expect(pkg.uniqueName('word/header{n}.xml')).toBe('word/header3.xml');
    expect(pkg.uniqueName('/word/footer{n}.xml')).toBe('word/footer1.xml');
    expect(() => pkg.uniqueName('word/header.xml')).toThrow(UsageError);
  });

  it('addPart throws for an existing part, whatever the case', () => {
    const pkg = OpcPackage.open(buildBytes());
    expect(() => pkg.addPart('Word/Document.xml', CT.document, xml(DOCUMENT))).toThrow(UsageError);
    expect(() => pkg.addPart('word/../x.xml', CT.xml, new Uint8Array())).toThrow(UsageError);
  });
});

describe('relationships', () => {
  it('adds, removes and numbers relationships', () => {
    const pkg = OpcPackage.create();
    const doc = pkg.addPart('word/document.xml', CT.document, xml(DOCUMENT));
    expect(doc.rels.all()).toEqual([]);
    expect(pkg.part('word/_rels/document.xml.rels')).toBeUndefined();
    expect(doc.rels.nextId()).toBe('rId1');

    const a = doc.rels.add(REL.styles, 'styles.xml');
    expect(a).toEqual({
      id: 'rId1',
      type: REL.styles,
      target: 'styles.xml',
      external: false,
      partName: 'word/styles.xml',
    });
    const relsPart = pkg.part('word/_rels/document.xml.rels')!;
    expect(relsPart.contentType).toBe(CT.relationships);
    expect(pkg.contentTypes.getOverride('word/_rels/document.xml.rels')).toBeUndefined();

    const b = doc.rels.add(REL.header, 'header1.xml');
    expect(b.id).toBe('rId2');
    doc.rels.add(REL.image, 'media/a.png', { id: 'rId10' });
    expect(doc.rels.nextId()).toBe('rId11');
    doc.rels.add(REL.image, 'media/b.png', { id: 'img' });
    expect(() => doc.rels.add(REL.image, 'media/c.png', { id: 'rId2' })).toThrow(UsageError);

    expect(doc.rels.get('rId10')!.partName).toBe('word/media/a.png');
    expect(doc.rels.byType('header').map((r) => r.id)).toEqual(['rId2']);
    expect(doc.rels.byType(REL.image).map((r) => r.id)).toEqual(['rId10', 'img']);
    expect(doc.rels.remove('rId10')).toBe(true);
    expect(doc.rels.remove('rId10')).toBe(false);
    expect(doc.rels.get('rId10')).toBeUndefined();
    expect(doc.rels.nextId()).toBe('rId3');
    expect(doc.rels.source).toBe('word/document.xml');
    expect(pkg.rels.source).toBe('');

    const re = OpcPackage.open(pkg.save());
    expect(
      re
        .part('word/document.xml')!
        .rels.all()
        .map((r) => r.id),
    ).toEqual(['rId1', 'rId2', 'img']);
  });

  it('nextId skips odd ids and stays unique', () => {
    const pkg = OpcPackage.create();
    const p = pkg.addPart('a.xml', CT.xml, xml('<a/>'));
    p.rels.add(REL.image, 'x.png', { id: 'rId0010' });
    p.rels.add(REL.image, 'y.png', { id: 'rIdX' });
    p.rels.add(REL.image, 'z.png', { id: 'rId99999999999999999999' });
    expect(p.rels.nextId()).toBe('rId11');
  });

  it('reads Strict relationship types and resolves targets', () => {
    const strict = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const bytes = zipOf({
      '[Content_Types].xml': `${TYPES}<Override PartName="/word/document.xml" ContentType="${CT.document}"/></Types>`,
      '_rels/.rels': `<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${strict}/officeDocument" Target="/word/document.xml"/><Relationship Id="rId2" Type="${strict}/extendedProperties" Target="docProps/app.xml"/></Relationships>`,
      'word/document.xml': `<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"/>`,
      'word/_rels/document.xml.rels': `<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${strict}/header" Target="header1.xml"/><Relationship Id="rId2" Type="${strict}/image" Target="../media/x.png"/><Relationship Id="rId3" Type="${strict}/image" Target="media/my%20image.png#frag"/><Relationship Id="rId4" Type="${strict}/hyperlink" Target="http://x.org/" TargetMode="External"/></Relationships>`,
      'word/header1.xml': `<w:hdr xmlns:w="${NS.w}"/>`,
      'word/media/my image.png': PNG,
    });
    const pkg = OpcPackage.open(bytes);
    expect(pkg.rels.byType('officeDocument')[0]!.partName).toBe('word/document.xml');
    expect(pkg.rels.byType(REL.officeDocument)).toHaveLength(1);
    expect(pkg.rels.byType(REL.extendedProperties)).toHaveLength(1);
    const doc = pkg.part('word/document.xml')!;
    expect(doc.rels.byType(REL.header)[0]!.partName).toBe('word/header1.xml');
    expect(doc.relatedByType('header')[0]!.name).toBe('word/header1.xml');
    const images = doc.rels.byType('image');
    expect(images.map((r) => r.partName)).toEqual(['media/x.png', 'word/media/my image.png']);
    expect(doc.related('rId3')!.bytes()).toEqual(PNG);
    expect(doc.related('rId2')).toBeUndefined(); // target part missing
    expect(doc.rels.get('rId4')).toMatchObject({ external: true, partName: undefined });
    expect(pkg.save()).toBe(bytes);
  });
});

describe('content types', () => {
  it('gets, sets and removes defaults and overrides', () => {
    const pkg = OpcPackage.open(buildBytes());
    const ct = pkg.contentTypes;
    expect(ct.getDefault('PNG')).toBe('image/png');
    expect(ct.getDefault('.xml')).toBe('application/xml');
    expect(ct.get('word/media/image1.png')).toBe('image/png');
    expect(ct.get('word/document.xml')).toBe(CT.document);
    expect(ct.get('other/thing.xml')).toBe('application/xml');
    expect(ct.get('noext')).toBeUndefined();
    expect(ct.defaults()).toEqual(
      new Map([
        ['rels', CT.relationships],
        ['xml', 'application/xml'],
        ['png', 'image/png'],
      ]),
    );
    expect(ct.overrides().get('word/document.xml')).toBe(CT.document);
    // images use the png default: no override was written
    expect(ct.getOverride('word/media/image1.png')).toBeUndefined();

    ct.setDefault('jpeg', 'image/jpeg');
    ct.setOverride('/word/new.xml', CT.footer);
    ct.setOverride('WORD/DOCUMENT.XML', CT.documentMacro);
    const kids = ct.xml.root.elements().map((e) => e.attr('Extension') ?? e.attr('PartName'));
    expect(kids.indexOf('jpeg')).toBe(3); // after the last Default
    expect(kids.at(-1)).toBe('/word/new.xml'); // new Override at the end
    expect(ct.get('word/document.xml')).toBe(CT.documentMacro);
    expect(ct.overrides().has('word/document.xml')).toBe(true); // updated in place

    expect(ct.removeOverride('/Word/New.xml')).toBe(true);
    expect(ct.removeOverride('word/new.xml')).toBe(false);
    expect(ct.get('word/new.xml')).toBe('application/xml');
    expect(pkg.part('[Content_Types].xml')!.dirty).toBe(true);

    const re = OpcPackage.open(pkg.save()).contentTypes;
    expect(re.getDefault('jpeg')).toBe('image/jpeg');
    expect(re.get('word/document.xml')).toBe(CT.documentMacro);
  });

  it('a new Default goes first when there is none', () => {
    const bytes = zipOf({
      '[Content_Types].xml': `<Types xmlns="${NS.ct}"><Override PartName="/a.xml" ContentType="${CT.xml}"/></Types>`,
      'a.xml': '<a/>',
    });
    const ct = OpcPackage.open(bytes).contentTypes;
    ct.setDefault('rels', CT.relationships);
    expect(ct.xml.root.elements().map((e) => e.local)).toEqual(['Default', 'Override']);
  });

  it('addPart adds an override only when the default does not match', () => {
    const pkg = OpcPackage.open(buildBytes());
    pkg.addPart('word/media/image2.png', 'image/png', PNG);
    pkg.addPart('word/media/image3.emf', 'image/x-emf', PNG);
    pkg.addPart('customXml/item1.xml', 'application/xml', xml('<a/>'));
    const ct = pkg.contentTypes;
    expect(ct.getOverride('word/media/image2.png')).toBeUndefined();
    expect(ct.getOverride('word/media/image3.emf')).toBe('image/x-emf');
    expect(ct.getOverride('customXml/item1.xml')).toBeUndefined();
    expect(pkg.part('word/media/image3.emf')!.contentType).toBe('image/x-emf');
  });
});

describe('removePart', () => {
  it('removes the part, its override and its rels part; save skips them', () => {
    const bytes = buildBytes();
    const pkg = OpcPackage.open(bytes);
    expect(pkg.removePart('WORD/header1.xml')).toBe(true);
    expect(pkg.removePart('word/header1.xml')).toBe(false);
    expect(pkg.part('word/header1.xml')).toBeUndefined();
    expect(pkg.part('word/_rels/header1.xml.rels')).toBeUndefined();
    expect(pkg.contentTypes.getOverride('word/header1.xml')).toBeUndefined();
    // relationships pointing at it are kept
    expect(pkg.part('word/document.xml')!.rels.byType('header')).toHaveLength(1);
    expect(pkg.dirty).toBe(true);
    const names = [...unzipLazy(pkg.save()).keys()];
    expect(names).not.toContain('word/header1.xml');
    expect(names).not.toContain('word/_rels/header1.xml.rels');
    expect(names).toContain('word/media/image1.png');
    expect(() => pkg.removePart('[Content_Types].xml')).toThrow(UsageError);
  });

  it('a removed part can be added again', () => {
    const pkg = OpcPackage.open(buildBytes());
    pkg.removePart('word/styles.xml');
    pkg.addPart('word/styles.xml', CT.styles, xml(`<w:styles ${W}><w:style/></w:styles>`));
    const re = OpcPackage.open(pkg.save());
    expect(re.part('word/styles.xml')!.xml().root.elements()).toHaveLength(1);
    expect(re.contentTypes.get('word/styles.xml')).toBe(CT.styles);
    expect(re.parts().filter((p) => p.name === 'word/styles.xml')).toHaveLength(1);
  });
});

describe('fork', () => {
  it('a fork of an unmodified package saves to the input bytes', () => {
    const bytes = buildBytes();
    const pkg = OpcPackage.open(bytes);
    pkg.part('word/document.xml')!.xml();
    const f = pkg.fork();
    expect(f.save()).toBe(bytes);
    expect(f.part('word/document.xml')).not.toBe(pkg.part('word/document.xml'));
    expect(f.part('word/document.xml')!.xml()).not.toBe(pkg.part('word/document.xml')!.xml());
  });

  it('fork and original are independent', () => {
    const bytes = buildBytes();
    const pkg = OpcPackage.open(bytes);
    const f = pkg.fork();
    f.part('word/document.xml')!.xml().root.descendants(NS.w, 't')[0]!.text = 'Fork';
    f.addPart('word/extra.xml', CT.xml, xml('<x/>'));
    f.removePart('word/styles.xml');
    f.part('word/document.xml')!.rels.add(REL.footer, 'footer1.xml');
    expect(pkg.dirty).toBe(false);
    expect(pkg.save()).toBe(bytes);
    expect(pkg.part('word/extra.xml')).toBeUndefined();
    expect(pkg.part('word/styles.xml')).toBeDefined();
    expect(pkg.part('word/document.xml')!.rels.byType('footer')).toHaveLength(0);
    expect(pkg.contentTypes.getOverride('word/styles.xml')).toBe(CT.styles);

    pkg.part('word/document.xml')!.xml().root.descendants(NS.w, 't')[0]!.text = 'Original';
    const fo = OpcPackage.open(f.save());
    expect(fo.part('word/document.xml')!.xml().root.text).toBe('Fork');
    expect(fo.part('word/extra.xml')).toBeDefined();
    expect(fo.part('word/styles.xml')).toBeUndefined();
    const po = OpcPackage.open(pkg.save());
    expect(po.part('word/document.xml')!.xml().root.text).toBe('Original');
    expect(po.part('word/extra.xml')).toBeUndefined();
  });

  it('dirty parts stay dirty in the fork, sharing replaced bytes', () => {
    const pkg = OpcPackage.open(buildBytes());
    pkg.part('word/document.xml')!.xml().root.descendants(NS.w, 't')[0]!.text = 'Edited';
    const data = enc.encode(`<w:styles ${W}><w:latentStyles/></w:styles>`);
    pkg.part('word/styles.xml')!.setBytes(data);
    const f = pkg.fork();
    expect(f.dirty).toBe(true);
    expect(f.part('word/document.xml')!.dirty).toBe(true);
    expect(f.part('word/styles.xml')!.bytes()).toBe(data);
    const re = OpcPackage.open(f.save());
    expect(re.part('word/document.xml')!.xml().root.text).toBe('Edited');
    expect(re.part('word/styles.xml')!.bytes()).toEqual(data);
  });

  it('forks of created packages keep their added parts', () => {
    const pkg = buildPackage();
    const f = pkg.fork();
    f.part('word/document.xml')!.xml().root.descendants(NS.w, 't')[0]!.text = 'F';
    expect(OpcPackage.open(pkg.save()).part('word/document.xml')!.xml().root.text).toBe('Hello');
    expect(OpcPackage.open(f.save()).part('word/document.xml')!.xml().root.text).toBe('F');
  });
});

describe('gc', () => {
  it('removes unreachable parts with their overrides and rels', () => {
    const pkg = OpcPackage.open(buildBytes());
    const orphan = pkg.addPart('word/orphan.xml', CT.footer, xml(`<w:ftr ${W}/>`));
    orphan.rels.add(REL.image, 'media/orphan.png');
    pkg.addPart('word/media/orphan.png', 'image/png', PNG);
    pkg.addPart(
      'word/_rels/ghost.xml.rels',
      CT.relationships,
      xml(`<Relationships xmlns="${NS.rel}"/>`),
    );
    // a cycle back to a reachable part must not loop
    pkg.part('word/header1.xml')!.rels.add(REL.officeDocument, 'document.xml');

    const removed = pkg.gc();
    expect(removed.sort()).toEqual(
      [
        'word/_rels/ghost.xml.rels',
        'word/_rels/orphan.xml.rels',
        'word/media/orphan.png',
        'word/orphan.xml',
      ].sort(),
    );
    expect(pkg.contentTypes.getOverride('word/orphan.xml')).toBeUndefined();
    const names = pkg.parts().map((p) => p.name);
    expect(names).toEqual(
      expect.arrayContaining([
        '[Content_Types].xml',
        '_rels/.rels',
        'word/document.xml',
        'word/_rels/document.xml.rels',
        'word/styles.xml',
        'word/header1.xml',
        'word/_rels/header1.xml.rels',
        'word/media/image1.png',
        'docProps/core.xml',
      ]),
    );
    expect(names).toHaveLength(9);
    expect(pkg.gc()).toEqual([]);
  });

  it('a clean, fully reachable package stays byte-identical', () => {
    const bytes = buildBytes();
    const pkg = OpcPackage.open(bytes);
    expect(pkg.gc()).toEqual([]);
    expect(pkg.save()).toBe(bytes);
  });
});

describe('open errors and warnings', () => {
  it('throws DocxFormatError for non-zip input', () => {
    for (const input of [new Uint8Array(0), enc.encode('not a zip at all, just some text')]) {
      let err: unknown;
      try {
        OpcPackage.open(input);
      } catch (e) {
        err = e;
      }
      expect(err).toBeInstanceOf(DocxFormatError);
      expect((err as DocxFormatError).code).toBe('DOCX_FORMAT');
      expect((err as Error).cause).toBeInstanceOf(ZipLimitError);
    }
  });

  it('throws DocxFormatError when [Content_Types].xml is missing', () => {
    expect(() => OpcPackage.open(zipOf({ 'word/document.xml': '<a/>' }))).toThrow(DocxFormatError);
    expect(() => OpcPackage.open(zipOf({ '[Content_Types].xml': '<NotTypes/>' }))).toThrow(
      DocxFormatError,
    );
  });

  it('lets limit violations through unchanged', () => {
    const bytes = buildBytes();
    let err: unknown;
    try {
      OpcPackage.open(bytes, { limits: { maxEntries: 2 } });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(ZipLimitError);
    expect(err).not.toBeInstanceOf(DocxFormatError);
  });

  it('applies maxXmlDepth when parsing parts', () => {
    const deep = `${'<a>'.repeat(50)}${'</a>'.repeat(50)}`;
    const bytes = zipOf({ '[Content_Types].xml': `${TYPES}</Types>`, 'deep.xml': deep });
    const pkg = OpcPackage.open(bytes, { limits: { maxXmlDepth: 10 } });
    expect(() => pkg.part('deep.xml')!.xml()).toThrow();
    expect(OpcPackage.open(bytes).part('deep.xml')!.xml().root.local).toBe('a');
  });

  it('warns about dangling overrides, missing content types and case duplicates', () => {
    const bytes = zipOf({
      '[Content_Types].xml': `${TYPES}<Override PartName="/word/gone.xml" ContentType="${CT.footer}"/></Types>`,
      'word/a.xml': '<first/>',
      'WORD/A.XML': '<second/>',
      'word/media/pic.bmp': PNG,
      'word/': new Uint8Array(0),
    });
    const warnings: Warning[] = [];
    const pkg = OpcPackage.open(bytes, { onWarning: (w) => warnings.push(w) });
    expect(warnings.map((w) => [w.code, w.part])).toEqual([
      ['DUPLICATE_PART', 'WORD/A.XML'],
      ['DANGLING_OVERRIDE', 'word/gone.xml'],
      ['MISSING_CONTENT_TYPE', 'word/media/pic.bmp'],
    ]);
    expect(pkg.part('WORD/A.XML')!.xml().root.local).toBe('first');
    expect(pkg.parts().map((p) => p.name)).toEqual([
      '[Content_Types].xml',
      'word/a.xml',
      'word/media/pic.bmp',
    ]);
    expect(pkg.save()).toBe(bytes);
    // a dirty save drops the ignored duplicate (and directory entries, which unzipLazy skips)
    pkg.part('word/a.xml')!.xml().root.setAttr('x', '1');
    const out = pkg.save();
    expect([...unzipLazy(out).keys()]).toEqual([
      '[Content_Types].xml',
      'word/a.xml',
      'word/media/pic.bmp',
    ]);
    expect(dec.decode(unzipLazy(out).get('word/a.xml')!.bytes())).toContain('first');
  });
});
