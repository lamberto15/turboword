// DocumentProperties: core, app and custom properties — set, read back, reopen, create from
// scratch, and reading the shapes other producers write. Packages are built with the OPC API,
// saved and reopened, so every assertion is about what lands in the file.

import { describe, expect, it } from 'vitest';
import { UsageError } from '../src/core/errors.js';
import { CT, OpcPackage, REL } from '../src/opc/index.js';
import { DocumentProperties } from '../src/wml/properties.js';
import { validatePackage } from '../src/wml/validate.js';
import { NS } from '../src/xml/ns.js';

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
const dec = (b: Uint8Array): string => new TextDecoder().decode(b);
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const DOC = `${DECL}<w:document xmlns:w="${NS.w}"><w:body><w:p/><w:sectPr/></w:body></w:document>`;

const CORE_NS =
  `xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" ` +
  `xmlns:dcmitype="${NS.dcmitype}" xmlns:xsi="${NS.xsi}"`;
const WORD_CORE =
  `${DECL}<cp:coreProperties ${CORE_NS}><dc:title>Annual Report</dc:title>` +
  '<dc:subject></dc:subject><dc:creator>Ada</dc:creator><cp:keywords>a, b</cp:keywords>' +
  '<dc:description>Notes</dc:description><cp:lastModifiedBy>Bob</cp:lastModifiedBy>' +
  '<cp:revision>7</cp:revision><cp:lastPrinted>2023-01-02T03:04:05Z</cp:lastPrinted>' +
  '<dcterms:created xsi:type="dcterms:W3CDTF">2023-01-01T09:00:00Z</dcterms:created>' +
  '<dcterms:modified xsi:type="dcterms:W3CDTF">2023-06-30T18:45:00Z</dcterms:modified>' +
  '<cp:category>Finance</cp:category><cp:contentStatus>Final</cp:contentStatus>' +
  '<dc:language>en-US</dc:language><dc:identifier>X-1</dc:identifier>' +
  '<cp:version>1.2</cp:version></cp:coreProperties>';

const APP_NS = `xmlns="${NS.ep}" xmlns:vt="${NS.vt}"`;
const WORD_APP =
  `${DECL}<Properties ${APP_NS}><Template>Normal.dotm</Template><TotalTime>12</TotalTime>` +
  '<Pages>3</Pages><Words>512</Words><Characters>2920</Characters>' +
  '<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity>' +
  '<Lines>24</Lines><Paragraphs>6</Paragraphs><ScaleCrop>false</ScaleCrop>' +
  '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Title</vt:lpstr>' +
  '</vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>' +
  '<LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>3426</CharactersWithSpaces>' +
  '<SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>' +
  '<AppVersion>16.0000</AppVersion></Properties>';

const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
const prop = (pid: number, name: string, value: string): string =>
  `<property fmtid="${FMTID}" pid="${pid}" name="${name}">${value}</property>`;

interface Extra {
  core?: string;
  app?: string;
  custom?: string;
  /** Add the parts without their package relationships. */
  noRels?: boolean;
  /** Content type for the core part (default CT.coreProperties). */
  coreType?: string;
}

/** create → main document (+ given property parts) → save → reopen. */
function build(extra: Extra = {}): OpcPackage {
  const pkg = OpcPackage.create();
  pkg.addPart('word/document.xml', CT.document, enc(DOC));
  pkg.rels.add(REL.officeDocument, 'word/document.xml');
  const add = (name: string, type: string, rel: string, xml: string | undefined) => {
    if (xml === undefined) return;
    pkg.addPart(name, type, enc(xml));
    if (!extra.noRels) pkg.rels.add(rel, name);
  };
  add('docProps/core.xml', extra.coreType ?? CT.coreProperties, REL.coreProperties, extra.core);
  add('docProps/app.xml', CT.extendedProperties, REL.extendedProperties, extra.app);
  add('docProps/custom.xml', CT.customProperties, REL.customProperties, extra.custom);
  return OpcPackage.open(pkg.save());
}

const reopen = (pkg: OpcPackage): OpcPackage => OpcPackage.open(pkg.save());
const xmlOf = (pkg: OpcPackage, name: string): string => dec(pkg.part(name)!.bytes());

describe('DocumentProperties: reading', () => {
  it('reads nothing, and creates nothing, when there are no property parts', () => {
    const pkg = build();
    const props = new DocumentProperties(pkg);
    expect(props.title).toBeUndefined();
    expect(props.created).toBeUndefined();
    expect(props.company).toBeUndefined();
    expect(props.app).toEqual({});
    expect(props.custom()).toEqual({});
    props.title = undefined;
    props.company = undefined;
    expect(props.removeCustom('x')).toBe(false);
    expect(pkg.dirty).toBe(false);
    expect(pkg.part('docProps/core.xml')).toBeUndefined();
  });

  it('reads every core property Word writes', () => {
    const p = new DocumentProperties(build({ core: WORD_CORE }));
    expect([p.title, p.subject, p.creator, p.keywords, p.description]).toEqual([
      'Annual Report',
      '',
      'Ada',
      'a, b',
      'Notes',
    ]);
    expect([p.lastModifiedBy, p.revision, p.category, p.contentStatus]).toEqual([
      'Bob',
      '7',
      'Finance',
      'Final',
    ]);
    expect([p.language, p.identifier, p.version]).toEqual(['en-US', 'X-1', '1.2']);
    expect(p.created).toEqual(new Date('2023-01-01T09:00:00Z'));
    expect(p.modified).toEqual(new Date('2023-06-30T18:45:00Z'));
    expect(p.lastPrinted).toEqual(new Date('2023-01-02T03:04:05Z'));
  });

  it('reads W3CDTF dates in every profile form, independent of the local time zone', () => {
    const at = (s: string) =>
      new DocumentProperties(
        build({
          core: `${DECL}<cp:coreProperties ${CORE_NS}><dcterms:created>${s}</dcterms:created></cp:coreProperties>`,
        }),
      ).created;
    expect(at('2024-03-05T10:20:30Z')).toEqual(new Date('2024-03-05T10:20:30Z'));
    expect(at('2024-03-05T12:20:30+02:00')).toEqual(new Date('2024-03-05T10:20:30Z'));
    expect(at('2024-03-05T05:50-04:30')).toEqual(new Date('2024-03-05T10:20:00Z'));
    expect(at('2024-03-05T10:20:30.1234567Z')).toEqual(new Date('2024-03-05T10:20:30.123Z'));
    expect(at('2024-03-05T10:20:30')).toEqual(new Date('2024-03-05T10:20:30Z'));
    expect(at(' 2024-03-05 ')).toEqual(new Date('2024-03-05T00:00:00Z'));
    expect(at('2024-03')).toEqual(new Date('2024-03-01T00:00:00Z'));
    expect(at('2024')).toEqual(new Date('2024-01-01T00:00:00Z'));
    expect(at('not a date')).toBeUndefined();
    expect(at('')).toBeUndefined();
  });

  it('ignores property parts that the package relationships do not reference', () => {
    const p = new DocumentProperties(build({ core: WORD_CORE, app: WORD_APP, noRels: true }));
    expect(p.title).toBeUndefined();
    expect(p.app).toEqual({});
  });

  it('reads app properties with numbers parsed', () => {
    const p = new DocumentProperties(build({ app: WORD_APP }));
    expect(p.app).toEqual({
      template: 'Normal.dotm',
      totalTime: 12,
      pages: 3,
      words: 512,
      characters: 2920,
      application: 'Microsoft Office Word',
      docSecurity: 0,
      lines: 24,
      paragraphs: 6,
      charactersWithSpaces: 3426,
      appVersion: '16.0000',
    });
    expect(p.company).toBeUndefined();
    expect(p.application).toBe('Microsoft Office Word');
  });

  it('reads custom properties of every common vt type', () => {
    const custom =
      `${DECL}<Properties xmlns="${NS.custom}" xmlns:vt="${NS.vt}">` +
      prop(2, 'Text', '<vt:lpwstr>hello</vt:lpwstr>') +
      prop(3, 'Ansi', '<vt:lpstr>ansi</vt:lpstr>') +
      prop(4, 'Int', '<vt:i4>-12</vt:i4>') +
      prop(5, 'Long', '<vt:i8>9007199254740991</vt:i8>') +
      prop(6, 'Float', '<vt:r8>2.5</vt:r8>') +
      prop(7, 'Yes', '<vt:bool>1</vt:bool>') +
      prop(8, 'No', '<vt:bool>false</vt:bool>') +
      prop(9, 'When', '<vt:filetime>2024-12-31T23:59:59Z</vt:filetime>') +
      prop(10, 'Money', '<vt:cy>12.5000</vt:cy>') +
      prop(11, '__proto__', '<vt:lpwstr>safe</vt:lpwstr>') +
      '</Properties>';
    const values = new DocumentProperties(build({ custom })).custom();
    expect(values).toEqual({
      Text: 'hello',
      Ansi: 'ansi',
      Int: -12,
      Long: 9007199254740991,
      Float: 2.5,
      Yes: true,
      No: false,
      When: new Date('2024-12-31T23:59:59Z'),
      Money: '12.5000',
      ['__proto__']: 'safe',
    });
    expect(Object.getPrototypeOf(values)).toBe(Object.prototype);
    expect(Object.hasOwn(values, '__proto__')).toBe(true);
  });
});

describe('DocumentProperties: core properties', () => {
  it('creates core.xml, its relationship and override on the first set', () => {
    const pkg = build();
    const p = new DocumentProperties(pkg);
    p.title = 'Hello';
    p.creator = 'Ada';
    p.created = new Date('2024-03-05T10:20:30.456Z');

    const part = pkg.part('docProps/core.xml')!;
    expect(part.dirty).toBe(true);
    expect(pkg.contentTypes.getOverride('docProps/core.xml')).toBe(CT.coreProperties);
    const rels = pkg.rels.byType('core-properties');
    expect(rels.map((r) => r.partName)).toEqual(['docProps/core.xml']);

    const again = reopen(pkg);
    const q = new DocumentProperties(again);
    expect(q.title).toBe('Hello');
    expect(q.creator).toBe('Ada');
    expect(q.created).toEqual(new Date('2024-03-05T10:20:30Z'));
    const xml = xmlOf(again, 'docProps/core.xml');
    expect(xml).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2024-03-05T10:20:30Z</dcterms:created>',
    );
    expect(xml).toContain(`xmlns:xsi="${NS.xsi}"`);
    expect(validatePackage(again)).toEqual([]);
  });

  it('round-trips every text property, written in Word order', () => {
    const names = [
      'title',
      'subject',
      'creator',
      'keywords',
      'description',
      'lastModifiedBy',
      'revision',
      'category',
      'contentStatus',
      'language',
      'identifier',
      'version',
    ] as const;
    const pkg = build();
    const p = new DocumentProperties(pkg);
    for (const n of [...names].reverse()) p[n] = `${n} <&> value`;
    const again = reopen(pkg);
    const q = new DocumentProperties(again);
    for (const n of names) expect(q[n]).toBe(`${n} <&> value`);
    const xml = xmlOf(again, 'docProps/core.xml');
    const order = ['dc:title', 'dc:subject', 'dc:creator', 'cp:keywords', 'dc:description'];
    const positions = order.map((t) => xml.indexOf(`<${t}>`));
    expect(positions.every((x) => x >= 0)).toBe(true);
    expect([...positions].sort((a, b) => a - b)).toEqual(positions);
  });

  it('writes dates without milliseconds; lastPrinted as untyped cp:lastPrinted', () => {
    const pkg = build();
    const p = new DocumentProperties(pkg);
    p.modified = new Date(Date.UTC(2025, 0, 2, 3, 4, 5, 999));
    p.lastPrinted = new Date('2025-02-03T04:05:06.007Z');
    const xml = xmlOf(reopen(pkg), 'docProps/core.xml');
    expect(xml).toContain(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-02T03:04:05Z</dcterms:modified>',
    );
    expect(xml).toContain('<cp:lastPrinted>2025-02-03T04:05:06Z</cp:lastPrinted>');
    expect(xml.indexOf('<cp:lastPrinted>')).toBeLessThan(xml.indexOf('<dcterms:modified'));
  });

  it('adds xsi:type to an existing untyped dcterms date', () => {
    const core = `${DECL}<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dcterms="${NS.dcterms}"><dcterms:created>2020-01-01</dcterms:created></cp:coreProperties>`;
    const pkg = build({ core });
    new DocumentProperties(pkg).created = new Date('2021-01-01T00:00:00Z');
    const xml = xmlOf(reopen(pkg), 'docProps/core.xml');
    expect(xml).toMatch(/<dcterms:created xsi:type="dcterms:W3CDTF">2021-01-01T00:00:00Z</);
    expect(xml).toContain(`xmlns:xsi="${NS.xsi}"`);
  });

  it('rejects an invalid Date', () => {
    const p = new DocumentProperties(build());
    expect(() => {
      p.created = new Date(Number.NaN);
    }).toThrow(UsageError);
  });

  it('setting undefined removes the element', () => {
    const pkg = build({ core: WORD_CORE });
    const p = new DocumentProperties(pkg);
    p.title = undefined;
    p.created = undefined;
    const q = new DocumentProperties(reopen(pkg));
    expect(q.title).toBeUndefined();
    expect(q.created).toBeUndefined();
    expect(q.creator).toBe('Ada');
    expect(q.modified).toEqual(new Date('2023-06-30T18:45:00Z'));
  });

  it('edits in place and keeps everything else', () => {
    const pkg = build({ core: WORD_CORE });
    const p = new DocumentProperties(pkg);
    p.title = 'New title';
    const xml = xmlOf(reopen(pkg), 'docProps/core.xml');
    // The serializer may write an empty element as `<x/>`; that is the same XML.
    const empty = (s: string) => s.replace(/<([\w:]+)><\/\1>/g, '<$1/>');
    expect(empty(xml)).toBe(empty(WORD_CORE.replace('Annual Report', 'New title')));
  });

  it('setting the current value leaves the package byte-identical', () => {
    const bytes = build({ core: WORD_CORE }).save();
    const pkg = OpcPackage.open(bytes);
    const p = new DocumentProperties(pkg);
    p.title = 'Annual Report';
    p.subject = '';
    p.created = new Date('2023-01-01T09:00:00Z');
    expect(pkg.dirty).toBe(false);
    expect(pkg.save()).toBe(bytes);
  });

  it('reuses an unreferenced core.xml, adding the relationship', () => {
    const pkg = build({ core: WORD_CORE, noRels: true });
    const p = new DocumentProperties(pkg);
    p.creator = 'Eve';
    expect(pkg.parts().filter((x) => x.name === 'docProps/core.xml')).toHaveLength(1);
    const q = new DocumentProperties(reopen(pkg));
    expect(q.creator).toBe('Eve');
    expect(q.title).toBe('Annual Report');
  });

  it('adds the content-type override when the part relies on the xml Default', () => {
    const pkg = build({ core: WORD_CORE, coreType: CT.xml });
    expect(pkg.contentTypes.getOverride('docProps/core.xml')).toBeUndefined();
    new DocumentProperties(pkg).title = 'T';
    expect(reopen(pkg).contentTypes.get('docProps/core.xml')).toBe(CT.coreProperties);
  });
});

describe('DocumentProperties: app properties', () => {
  it('creates app.xml for company and application', () => {
    const pkg = build();
    const p = new DocumentProperties(pkg);
    p.company = 'Acme';
    p.application = 'turboword';
    expect(pkg.contentTypes.getOverride('docProps/app.xml')).toBe(CT.extendedProperties);
    expect(pkg.rels.byType('extended-properties').map((r) => r.partName)).toEqual([
      'docProps/app.xml',
    ]);
    const again = reopen(pkg);
    const q = new DocumentProperties(again);
    expect(q.app).toEqual({ company: 'Acme', application: 'turboword' });
    const xml = xmlOf(again, 'docProps/app.xml');
    expect(xml).toContain(`<Properties xmlns="${NS.ep}"`);
    expect(xml.indexOf('<Application>')).toBeLessThan(xml.indexOf('<Company>'));
    expect(validatePackage(again)).toEqual([]);
  });

  it('inserts into an existing app.xml at the Word position and keeps the rest', () => {
    const pkg = build({ app: WORD_APP });
    const p = new DocumentProperties(pkg);
    p.company = 'Acme';
    const again = reopen(pkg);
    const xml = xmlOf(again, 'docProps/app.xml');
    expect(xml).toBe(WORD_APP.replace('<LinksUpToDate>', '<Company>Acme</Company><LinksUpToDate>'));
    expect(new DocumentProperties(again).app.pages).toBe(3);
    p.company = undefined;
    expect(xmlOf(reopen(pkg), 'docProps/app.xml')).toBe(WORD_APP);
  });
});

describe('DocumentProperties: custom properties', () => {
  it('creates custom.xml and writes each value type', () => {
    const pkg = build();
    const p = new DocumentProperties(pkg);
    p.setCustom('Client', 'Acme & Co');
    p.setCustom('Count', 42);
    p.setCustom('Negative', -7);
    p.setCustom('Big', 2 ** 40);
    p.setCustom('Ratio', 0.5);
    p.setCustom('Approved', true);
    p.setCustom('Due', new Date('2024-12-31T00:00:00.250Z'));
    expect(pkg.contentTypes.getOverride('docProps/custom.xml')).toBe(CT.customProperties);
    expect(pkg.rels.byType('custom-properties')).toHaveLength(1);

    const again = reopen(pkg);
    expect(new DocumentProperties(again).custom()).toEqual({
      Client: 'Acme & Co',
      Count: 42,
      Negative: -7,
      Big: 2 ** 40,
      Ratio: 0.5,
      Approved: true,
      Due: new Date('2024-12-31T00:00:00Z'),
    });
    const xml = xmlOf(again, 'docProps/custom.xml');
    expect(xml).toContain(`<Properties xmlns="${NS.custom}" xmlns:vt="${NS.vt}">`);
    expect(xml).toContain(`<property fmtid="${FMTID}" pid="2" name="Client">`);
    expect(xml).toContain('<vt:lpwstr>Acme &amp; Co</vt:lpwstr>');
    expect(xml).toContain('<vt:i4>42</vt:i4>');
    expect(xml).toContain('<vt:i4>-7</vt:i4>');
    expect(xml).toContain('<vt:r8>1099511627776</vt:r8>');
    expect(xml).toContain('<vt:r8>0.5</vt:r8>');
    expect(xml).toContain('<vt:bool>true</vt:bool>');
    expect(xml).toContain('<vt:filetime>2024-12-31T00:00:00Z</vt:filetime>');
    const pids = [...xml.matchAll(/pid="(\d+)"/g)].map((m) => Number(m[1]));
    expect(pids).toEqual([2, 3, 4, 5, 6, 7, 8]);
    expect(validatePackage(again)).toEqual([]);
  });

  it('uses i4 only inside the 32-bit range', () => {
    const pkg = build();
    const p = new DocumentProperties(pkg);
    p.setCustom('Max', 2147483647);
    p.setCustom('Over', 2147483648);
    p.setCustom('Min', -2147483648);
    p.setCustom('Under', -2147483649);
    const xml = xmlOf(reopen(pkg), 'docProps/custom.xml');
    expect(xml).toContain('<vt:i4>2147483647</vt:i4>');
    expect(xml).toContain('<vt:r8>2147483648</vt:r8>');
    expect(xml).toContain('<vt:i4>-2147483648</vt:i4>');
    expect(xml).toContain('<vt:r8>-2147483649</vt:r8>');
  });

  it('replaces a value in place (keeping pid, name matched case-insensitively)', () => {
    const pkg = build();
    const p = new DocumentProperties(pkg);
    p.setCustom('Status', 'draft');
    p.setCustom('Other', 1);
    p.setCustom('STATUS', true);
    const again = reopen(pkg);
    expect(new DocumentProperties(again).custom()).toEqual({ Status: true, Other: 1 });
    expect(xmlOf(again, 'docProps/custom.xml')).toContain(
      `<property fmtid="${FMTID}" pid="2" name="Status"><vt:bool>true</vt:bool></property>`,
    );
  });

  it('removes properties and never reuses a pid in use', () => {
    const custom =
      `${DECL}<Properties xmlns="${NS.custom}" xmlns:vt="${NS.vt}">` +
      prop(2, 'A', '<vt:lpwstr>a</vt:lpwstr>') +
      prop(5, 'B', '<vt:lpwstr>b</vt:lpwstr>') +
      '</Properties>';
    const pkg = build({ custom });
    const p = new DocumentProperties(pkg);
    expect(p.removeCustom('a')).toBe(true);
    expect(p.removeCustom('missing')).toBe(false);
    p.setCustom('C', 'c');
    const again = reopen(pkg);
    expect(new DocumentProperties(again).custom()).toEqual({ B: 'b', C: 'c' });
    expect(xmlOf(again, 'docProps/custom.xml')).toContain('pid="6" name="C"');
  });

  it('setting the current value leaves the package byte-identical', () => {
    const custom =
      `${DECL}<Properties xmlns="${NS.custom}" xmlns:vt="${NS.vt}">` +
      prop(2, 'N', '<vt:i4>3</vt:i4>') +
      '</Properties>';
    const bytes = build({ custom }).save();
    const pkg = OpcPackage.open(bytes);
    new DocumentProperties(pkg).setCustom('N', 3);
    expect(pkg.dirty).toBe(false);
    expect(pkg.save()).toBe(bytes);
  });

  it('rejects empty names and values that cannot be written', () => {
    const p = new DocumentProperties(build());
    expect(() => p.setCustom('', 'x')).toThrow(UsageError);
    expect(() => p.setCustom('N', Number.NaN)).toThrow(UsageError);
    expect(() => p.setCustom('N', Number.POSITIVE_INFINITY)).toThrow(UsageError);
    expect(() => p.setCustom('D', new Date(Number.NaN))).toThrow(UsageError);
  });
});
