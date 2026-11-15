// The style cascade (ECMA-376 §17.7), rule by rule, on small synthetic documents, plus resolved
// properties of real fixtures and cache invalidation.

import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  listLabel,
  resolveCell,
  resolveParagraph,
  resolveRun,
  resolveTable,
  themeOf,
} from '../src/resolve/index.js';
import { readDocx, type WordDocument } from '../src/wml/document.js';
import type { Paragraph } from '../src/wml/paragraph.js';
import { NS } from '../src/xml/ns.js';
import { fixture } from './helpers/fixtures.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

interface Parts {
  body: string;
  styles?: string;
  docDefaults?: string;
  numbering?: string;
  theme?: string;
  settings?: string;
}

const THEME = `<a:theme xmlns:a="${A}" name="T"><a:themeElements>
<a:clrScheme name="C"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="F"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Arab" typeface="Times New Roman"/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Mincho"/><a:font script="Hebr" typeface="David"/></a:minorFont></a:fontScheme>
</a:themeElements></a:theme>`;

/** A package with the given parts (each part only when given). */
function makeDoc(p: Parts): WordDocument {
  const ct: string[] = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
  ];
  const rels: string[] = [];
  const files: Record<string, Uint8Array> = {};
  const add = (name: string, type: string, contentType: string, xml: string): void => {
    files[`word/${name}`] = strToU8(`<?xml version="1.0" encoding="UTF-8"?>${xml}`);
    rels.push(`<Relationship Id="rId${rels.length + 1}" Type="${REL}/${type}" Target="${name}"/>`);
    ct.push(`<Override PartName="/word/${name}" ContentType="${contentType}"/>`);
  };
  const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
  if (p.styles !== undefined || p.docDefaults !== undefined) {
    add(
      'styles.xml',
      'styles',
      `${wml}.styles+xml`,
      `<w:styles xmlns:w="${W}">${p.docDefaults ?? ''}${p.styles ?? ''}</w:styles>`,
    );
  }
  if (p.numbering !== undefined) {
    add(
      'numbering.xml',
      'numbering',
      `${wml}.numbering+xml`,
      `<w:numbering xmlns:w="${W}">${p.numbering}</w:numbering>`,
    );
  }
  if (p.theme !== undefined) {
    add(
      'theme/theme1.xml',
      'theme',
      'application/vnd.openxmlformats-officedocument.theme+xml',
      p.theme,
    );
  }
  if (p.settings !== undefined) {
    add(
      'settings.xml',
      'settings',
      `${wml}.settings+xml`,
      `<w:settings xmlns:w="${W}">${p.settings}</w:settings>`,
    );
  }
  files['[Content_Types].xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${ct.join('')}</Types>`,
  );
  files['_rels/.rels'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  files['word/_rels/document.xml.rels'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
  );
  files['word/document.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W}" xmlns:r="${REL}"><w:body>${p.body}</w:body></w:document>`,
  );
  return readDocx(zipSync(files));
}

const para = (pPr: string, runs = '<w:r><w:t>x</w:t></w:r>'): string =>
  `<w:p>${pPr === '' ? '' : `<w:pPr>${pPr}</w:pPr>`}${runs}</w:p>`;
const run = (rPr: string, text = 'x'): string =>
  `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t>${text}</w:t></w:r>`;
const pStyle = (id: string, inner: string, extra = ''): string =>
  `<w:style w:type="paragraph" w:styleId="${id}"${extra}><w:name w:val="${id}"/>${inner}</w:style>`;
const cStyle = (id: string, inner: string, extra = ''): string =>
  `<w:style w:type="character" w:styleId="${id}"${extra}><w:name w:val="${id}"/>${inner}</w:style>`;
const defaults = (rPr: string, pPr = ''): string =>
  `<w:docDefaults><w:rPrDefault><w:rPr>${rPr}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>${pPr}</w:pPr></w:pPrDefault></w:docDefaults>`;

function paras(doc: WordDocument): Paragraph[] {
  return doc.body.paragraphs();
}
function firstRun(doc: WordDocument, k = 0): ReturnType<typeof resolveRun> {
  const p = paras(doc)[k]!;
  return resolveRun(p.runs()[0]!);
}

describe('style cascade: paragraphs', () => {
  it('applies docDefaults, then the default paragraph style when there is no pStyle', () => {
    const doc = makeDoc({
      docDefaults: defaults(
        '<w:sz w:val="22"/>',
        '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>',
      ),
      styles: pStyle('Normal', '<w:pPr><w:jc w:val="center"/></w:pPr>', ' w:default="1"'),
      body: para('') + para('<w:pStyle w:val="Missing"/>'),
    });
    for (const p of paras(doc)) {
      const r = resolveParagraph(p);
      expect(r.styleId).toBe('Normal');
      expect(r.alignment).toBe('center');
      expect(r.spacing).toMatchObject({ after: 160, line: 259, lineRule: 'auto', before: 0 });
      expect(r.runDefaults.size).toBe(11);
    }
  });

  it('follows basedOn (derived overrides base) and survives cycles', () => {
    const doc = makeDoc({
      styles:
        pStyle('Base', '<w:pPr><w:ind w:left="720" w:right="360"/><w:keepNext/></w:pPr>') +
        pStyle('Derived', '<w:basedOn w:val="Base"/><w:pPr><w:ind w:left="1440"/></w:pPr>') +
        pStyle('Loop1', '<w:basedOn w:val="Loop2"/><w:pPr><w:jc w:val="right"/></w:pPr>') +
        pStyle('Loop2', '<w:basedOn w:val="Loop1"/><w:pPr><w:keepLines/></w:pPr>'),
      body: para('<w:pStyle w:val="Derived"/>') + para('<w:pStyle w:val="Loop1"/>'),
    });
    const [d, l] = paras(doc).map(resolveParagraph);
    expect(d!.indent).toEqual({ left: 1440, right: 360, firstLine: 0, hanging: 0 });
    expect(d!.keepNext).toBe(true);
    expect(l!.alignment).toBe('right');
    expect(l!.keepLines).toBe(true);
  });

  it('treats firstLine/hanging as one quantity and maps start/end with bidi', () => {
    const doc = makeDoc({
      styles: pStyle('S', '<w:pPr><w:ind w:start="100" w:end="200" w:hanging="360"/></w:pPr>'),
      body:
        para('<w:pStyle w:val="S"/><w:ind w:firstLine="240"/>') +
        para('<w:pStyle w:val="S"/><w:bidi/><w:jc w:val="left"/>') +
        para('<w:ind w:firstLine="100" w:hanging="50"/>'),
    });
    const [a, b, c] = paras(doc).map(resolveParagraph);
    expect(a!.indent).toEqual({ left: 100, right: 200, firstLine: 240, hanging: 0 });
    expect(b!.indent).toEqual({ left: 200, right: 100, firstLine: 0, hanging: 360 });
    expect(b!.alignment).toBe('right');
    expect(c!.indent.hanging).toBe(50);
  });

  it('accumulates tab stops and honours clear', () => {
    const doc = makeDoc({
      styles: pStyle(
        'S',
        '<w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="9000" w:leader="dot"/></w:tabs></w:pPr>',
      ),
      body: para(
        '<w:pStyle w:val="S"/><w:tabs><w:tab w:val="clear" w:pos="720"/><w:tab w:val="center" w:pos="4500"/></w:tabs>',
      ),
    });
    expect(resolveParagraph(paras(doc)[0]!).tabs).toEqual([
      { pos: 4500, val: 'center' },
      { pos: 9000, val: 'right', leader: 'dot' },
    ]);
  });

  it('merges paragraph border sides; nil removes an inherited side', () => {
    const doc = makeDoc({
      styles: pStyle(
        'S',
        '<w:pPr><w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="FF0000"/><w:bottom w:val="double" w:sz="4"/></w:pBdr></w:pPr>',
      ),
      body: para(
        '<w:pStyle w:val="S"/><w:pBdr><w:bottom w:val="nil"/></w:pBdr><w:shd w:val="clear" w:fill="FFFF00"/>',
      ),
    });
    const r = resolveParagraph(paras(doc)[0]!);
    expect(r.borders).toEqual({ top: { style: 'single', size: 8, space: 1, color: 'FF0000' } });
    expect(r.shading).toBe('FFFF00');
  });

  it('defaults widow control on (Word only ever writes it off)', () => {
    const doc = makeDoc({ body: para('') + para('<w:widowControl w:val="0"/>') });
    expect(paras(doc).map((p) => resolveParagraph(p).widowControl)).toEqual([true, false]);
  });
});

describe('style cascade: runs', () => {
  it('parses ST_OnOff true/false/1/0/on/off and an omitted val', () => {
    const vals = [
      '',
      ' w:val="true"',
      ' w:val="1"',
      ' w:val="on"',
      ' w:val="false"',
      ' w:val="0"',
      ' w:val="off"',
    ];
    const doc = makeDoc({ body: vals.map((v) => para('', run(`<w:i${v}/>`))).join('') });
    expect(paras(doc).map((p) => resolveRun(p.runs()[0]!).italic)).toEqual([
      true,
      true,
      true,
      true,
      false,
      false,
      false,
    ]);
  });

  it('XORs toggle properties across style types; direct formatting is absolute', () => {
    const doc = makeDoc({
      styles:
        pStyle('H', '<w:rPr><w:b/><w:i/><w:caps/></w:rPr>') +
        pStyle('H2', '<w:basedOn w:val="H"/><w:rPr><w:caps w:val="0"/></w:rPr>') +
        cStyle('Strong', '<w:rPr><w:b/></w:rPr>') +
        cStyle('NotItalic', '<w:rPr><w:i w:val="0"/></w:rPr>'),
      body:
        para('<w:pStyle w:val="H"/>', run('<w:rStyle w:val="Strong"/>')) +
        para('<w:pStyle w:val="H"/>', run('<w:rStyle w:val="NotItalic"/>')) +
        para('<w:pStyle w:val="H"/>', run('<w:rStyle w:val="Strong"/><w:b/>')) +
        para('<w:pStyle w:val="H2"/>', run('')),
    });
    const [a, b, c, d] = [0, 1, 2, 3].map((k) => firstRun(doc, k));
    expect(a!.bold).toBe(false); // bold paragraph style + bold character style
    expect(a!.italic).toBe(true);
    expect(b!.italic).toBe(true); // "off" in a style leaves the state alone
    expect(c!.bold).toBe(true); // direct formatting wins absolutely
    expect(d!.caps).toBe(false); // within one chain, a derived style simply overrides
    expect(d!.bold).toBe(true);
    expect(a!.styleId).toBe('Strong');
  });

  it('does not apply the default character style implicitly', () => {
    const doc = makeDoc({
      styles: cStyle(
        'DefaultParagraphFont',
        '<w:rPr><w:color w:val="FF0000"/></w:rPr>',
        ' w:default="1"',
      ),
      body: para('', run('')),
    });
    const r = firstRun(doc);
    expect(r.color).toBe('auto');
    expect(r.styleId).toBeUndefined();
  });

  it('reads sizes in half-points, kerning, spacing, position, scale, underline, highlight', () => {
    const doc = makeDoc({
      body: para(
        '',
        run(
          '<w:sz w:val="25"/><w:szCs w:val="28"/><w:kern w:val="32"/><w:spacing w:val="-20"/><w:position w:val="6"/><w:w w:val="150"/><w:u w:val="double" w:color="00FF00"/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/><w:vanish/><w:lang w:val="fr-FR" w:eastAsia="ja-JP"/>',
        ),
      ),
    });
    expect(firstRun(doc)).toMatchObject({
      size: 12.5,
      sizeCs: 14,
      kern: 16,
      spacing: -20,
      position: 6,
      scale: 150,
      underline: 'double',
      underlineColor: '00FF00',
      highlight: 'yellow',
      vertAlign: 'superscript',
      hidden: true,
      lang: 'fr-FR',
      langEastAsia: 'ja-JP',
    });
  });

  it('falls back to Word defaults: 10 pt, Times New Roman, auto colour', () => {
    const r = firstRun(makeDoc({ body: para('', run('')) }));
    expect(r.size).toBe(10);
    expect(r.fonts.ascii).toBe('Times New Roman');
    expect(r.color).toBe('auto');
  });

  it('resolves the paragraph mark and reuses a passed paragraph result', () => {
    const doc = makeDoc({
      styles: pStyle('P', '<w:rPr><w:sz w:val="40"/></w:rPr>'),
      body: para('<w:pStyle w:val="P"/><w:rPr><w:b/></w:rPr>', run('<w:i/>')),
    });
    const p = paras(doc)[0]!;
    const props = resolveParagraph(p);
    expect(props.markRun).toMatchObject({ bold: true, size: 20 });
    expect(props.runDefaults).toMatchObject({ bold: false, size: 20 });
    expect(resolveRun(p.runs()[0]!, props)).toEqual(resolveRun(p.runs()[0]!));
    expect(resolveRun(p.runs()[0]!).italic).toBe(true);
  });
});

describe('theme', () => {
  const settings =
    '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP" w:bidi="he-IL"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark2" w:bg2="light2" w:t2="dark1" w:accent1="accent2" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>';

  it('resolves theme fonts with script fonts chosen by themeFontLang', () => {
    const doc = makeDoc({
      theme: THEME,
      settings,
      docDefaults: defaults(
        '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi" w:hint="eastAsia"/>',
      ),
      styles: pStyle(
        'Heading',
        '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"/></w:rPr>',
      ),
      body:
        para('', run('')) + para('<w:pStyle w:val="Heading"/>', run('<w:rFonts w:ascii="Arial"/>')),
    });
    expect(firstRun(doc, 0).fonts).toEqual({
      ascii: 'Calibri',
      hAnsi: 'Calibri',
      eastAsia: 'Yu Mincho',
      cs: 'David',
      hint: 'eastAsia',
    });
    const h = firstRun(doc, 1).fonts;
    expect(h.ascii).toBe('Arial'); // a later plain name replaces an earlier theme font
    expect(h.hAnsi).toBe('Calibri');
    expect(h.eastAsia).toBe('Yu Gothic Light');
    expect(themeOf(doc).fonts.minor).toEqual({
      latin: 'Calibri',
      eastAsia: 'Yu Mincho',
      cs: 'David',
    });
  });

  it('resolves theme colours through clrSchemeMapping, with tint and shade', () => {
    const doc = makeDoc({
      theme: THEME,
      settings,
      body:
        para('', run('<w:color w:val="000000" w:themeColor="text1"/>')) +
        para('', run('<w:color w:val="000000" w:themeColor="accent1"/>')) +
        para('', run('<w:color w:val="000000" w:themeColor="accent1" w:themeShade="BF"/>')) +
        para('', run('<w:color w:val="000000" w:themeColor="dark1" w:themeTint="99"/>')) +
        para('', run('<w:color w:val="123456" w:themeColor="none"/>')),
    });
    const colors = paras(doc).map((p) => resolveRun(p.runs()[0]!).color);
    expect(colors[0]).toBe('44546A'); // text1 → dark2 per the mapping
    expect(colors[1]).toBe('ED7D31'); // accent1 → accent2 per the mapping
    expect(colors[2]).toBe('C55A11'); // accent2 shaded 75% (Word writes C55A11)
    expect(colors[3]).toBe('666666'); // black tinted 60% (Word writes 666666)
    expect(colors[4]).toBe('123456');
    expect(themeOf(doc).colors).toMatchObject({
      accent1: '4472C4',
      text1: '44546A',
      hlink: '0563C1',
    });
  });

  it('is empty without a theme part', () => {
    expect(themeOf(makeDoc({ body: para('') }))).toEqual({
      fonts: { major: {}, minor: {} },
      colors: {},
    });
  });
});

describe('numbering in the cascade', () => {
  const numbering =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
    '<w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const styles =
    pStyle('ListParagraph', '<w:pPr><w:ind w:left="1000"/></w:pPr>') +
    pStyle(
      'Numbered',
      '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:ind w:left="2000"/></w:pPr>',
    );

  it('direct numPr: the level indent beats the style; numbering from the style: the style wins', () => {
    const doc = makeDoc({
      numbering,
      styles,
      body:
        para(
          '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>',
        ) +
        para('<w:pStyle w:val="Numbered"/>') +
        para(
          '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="3000"/>',
        ) +
        para('<w:pStyle w:val="Numbered"/><w:numPr><w:numId w:val="0"/></w:numPr>'),
    });
    const r = paras(doc).map(resolveParagraph);
    expect(r[0]!.indent).toMatchObject({ left: 720, hanging: 360 });
    expect(r[0]!.tabs).toEqual([{ pos: 720, val: 'num' }]);
    expect(r[1]!.indent).toMatchObject({ left: 2000, hanging: 360 });
    expect(r[2]!.indent.left).toBe(3000);
    expect(r[3]!.numbering).toBeUndefined();
    expect(r[3]!.indent).toMatchObject({ left: 2000, hanging: 0 });
    expect(paras(doc).map((p) => listLabel(p))).toEqual(['1.', '2.', '3.', undefined]);
  });

  it('formats the label with the paragraph mark plus the level rPr', () => {
    const doc = makeDoc({
      numbering,
      body: para(
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:rPr><w:i/><w:sz w:val="30"/></w:rPr>',
      ),
    });
    expect(paras(doc).map((p) => resolveParagraph(p).numbering)).toEqual([{ numId: 1, level: 0 }]);
  });
});

describe('tables', () => {
  const tableStyle = `<w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="18"/></w:rPr>
    <w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>
    <w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
    <w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:tcBorders><w:insideV w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:fill="4472C4"/></w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="firstCol"><w:rPr><w:b/><w:i/></w:rPr></w:tblStylePr>
    <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>
    <w:tblStylePr w:type="lastRow"><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="4"/></w:tcBorders></w:tcPr></w:tblStylePr>
  </w:style>`;
  const styles =
    pStyle(
      'Normal',
      '<w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr>',
      ' w:default="1"',
    ) +
    pStyle('Heading', '<w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="32"/></w:rPr>') +
    tableStyle +
    `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>`;
  const cell = (inner: string, tcPr = ''): string =>
    `<w:tc>${tcPr === '' ? '' : `<w:tcPr>${tcPr}</w:tcPr>`}${inner}</w:tc>`;
  const grid = (look: string, rows: string[][]): string =>
    `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/>${look}</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows
      .map((r) => `<w:tr>${r.join('')}</w:tr>`)
      .join('')}</w:tbl>`;
  const plain = (t: string): string => cell(para('', run('', t)));

  function tableDoc(look: string, rows?: string[][]): WordDocument {
    return makeDoc({
      styles,
      body:
        grid(
          look,
          rows ?? [
            [plain('a'), plain('b'), plain('c')],
            [plain('d'), plain('e'), plain('f')],
            [plain('g'), plain('h'), plain('i')],
            [plain('j'), plain('k'), plain('l')],
          ],
        ) + para(''),
    });
  }

  it('reads tblLook from the hex value or the explicit attributes', () => {
    const hex = resolveTable(tableDoc('<w:tblLook w:val="04A0"/>').body.tables()[0]!);
    expect(hex.look).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true,
    });
    const attrs = resolveTable(
      tableDoc(
        '<w:tblLook w:val="04A0" w:firstRow="0" w:lastRow="1" w:noVBand="0"/>',
      ).body.tables()[0]!,
    );
    expect(attrs.look).toMatchObject({
      firstRow: false,
      lastRow: true,
      firstColumn: true,
      noVBand: false,
    });
    expect(attrs.styleId).toBe('Grid');
    expect(attrs.cellMargins).toEqual({ top: 0, left: 108, bottom: 0, right: 108 });
    expect(attrs.borders.insideH).toEqual({ style: 'single', size: 4, space: 0, color: 'auto' });
  });

  it('applies the table style over the default paragraph style, below other paragraph styles', () => {
    const doc = makeDoc({
      styles,
      body: grid('<w:tblLook w:val="0000" w:noHBand="1"/>', [
        [plain('a'), cell(para('<w:pStyle w:val="Heading"/>', run('', 'h'))), plain('c')],
      ]),
    });
    const ps = doc.body.paragraphs();
    const a = resolveParagraph(ps[0]!);
    expect(a.spacing.after).toBe(0);
    expect(resolveRun(ps[0]!.runs()[0]!).size).toBe(9);
    const h = resolveParagraph(ps[1]!);
    expect(h.spacing.after).toBe(0);
    expect(resolveRun(ps[1]!.runs()[0]!).size).toBe(16);
  });

  it('applies conditional regions in precedence order (firstRow over firstCol, banding)', () => {
    const doc = tableDoc('<w:tblLook w:val="04A0"/>');
    const t = doc.body.tables()[0]!;
    const rows = t.rows();
    const run0 = (r: number, c: number) => {
      const p = rows[r]!.cells()[c]!.paragraphs()[0]!;
      return resolveRun(p.runs()[0]!);
    };
    expect(run0(0, 0)).toMatchObject({ bold: true, italic: true, color: 'FFFFFF' });
    expect(run0(0, 1)).toMatchObject({ bold: true, italic: false });
    expect(run0(1, 0)).toMatchObject({ bold: true, italic: true });
    expect(run0(1, 1)).toMatchObject({ bold: false, color: 'auto' });
    const c = (r: number, k: number) => resolveCell(rows[r]!.cells()[k]!);
    expect(c(0, 1).shading).toBe('4472C4');
    expect(c(0, 1).conditional).toEqual(['firstRow']);
    expect(c(1, 1).shading).toBe('D9E2F3');
    expect(c(1, 1).conditional).toEqual(['band1Horz']);
    expect(c(2, 1).shading).toBeUndefined();
    expect(c(3, 1).shading).toBe('D9E2F3');
    expect(c(1, 0).conditional).toEqual(['band1Horz', 'firstCol']);
  });

  it('turns insideH/insideV into cell edges and lets regions override them', () => {
    const doc = tableDoc('<w:tblLook w:val="04A0"/>');
    const rows = doc.body.tables()[0]!.rows();
    const header = resolveCell(rows[0]!.cells()[1]!);
    expect(header.borders.top).toMatchObject({ style: 'single' });
    // The header row's insideV is nil, but the body cell below still draws its own edges.
    expect(header.borders.left).toBeUndefined();
    expect(header.borders.right).toBeUndefined();
    expect(header.borders.bottom).toMatchObject({ style: 'single', size: 4 });
    const body = resolveCell(rows[1]!.cells()[1]!);
    expect(body.borders.left).toMatchObject({ style: 'single' });
    expect(body.borders.top).toMatchObject({ style: 'single' });
    expect(header.margins).toEqual({ top: 0, left: 108, bottom: 0, right: 108 });
  });

  it('settles shared edges by weight, then style, then darker colour', () => {
    const doc = tableDoc('<w:tblLook w:val="0000" w:noHBand="1"/>', [
      [
        plain('a'),
        cell(
          para('', run('', 'b')),
          '<w:tcBorders><w:left w:val="single" w:sz="12" w:color="FF0000"/></w:tcBorders>',
        ),
        cell(para('', run('', 'c')), '<w:tcBorders><w:left w:val="nil"/></w:tcBorders>'),
      ],
      [
        cell(
          para('', run('', 'd')),
          '<w:tcBorders><w:right w:val="single" w:sz="4" w:color="00FF00"/></w:tcBorders>',
        ),
        cell(
          para('', run('', 'e')),
          '<w:tcBorders><w:left w:val="single" w:sz="4" w:color="0000FF"/></w:tcBorders>',
        ),
        plain('f'),
      ],
    ]);
    const rows = doc.body.tables()[0]!.rows();
    // a|b: b's 1.5pt red beats a's 0.5pt (insideV).
    expect(resolveCell(rows[0]!.cells()[0]!).borders.right).toMatchObject({
      size: 12,
      color: 'FF0000',
    });
    expect(resolveCell(rows[0]!.cells()[1]!).borders.left).toMatchObject({
      size: 12,
      color: 'FF0000',
    });
    // b|c: c removes its left edge but b's insideV right edge remains.
    expect(resolveCell(rows[0]!.cells()[2]!).borders.left).toMatchObject({
      style: 'single',
      size: 4,
    });
    // d|e: same weight and style; blue (darker by R+B+2G) beats green.
    expect(resolveCell(rows[1]!.cells()[0]!).borders.right).toMatchObject({ color: '0000FF' });
  });

  it('uses cell margins, vAlign, widths and the default table style', () => {
    const doc = makeDoc({
      styles,
      body: `<w:tbl><w:tblPr><w:tblW w:w="50%" w:type="pct"/><w:jc w:val="center"/><w:tblInd w:w="144" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr>${cell(
        para(''),
        '<w:tcW w:w="1000" w:type="dxa"/><w:tcMar><w:top w:w="50" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/><w:noWrap/><w:textDirection w:val="btLr"/>',
      )}</w:tr></w:tbl>`,
    });
    const t = doc.body.tables()[0]!;
    const tp = resolveTable(t);
    expect(tp).toMatchObject({
      styleId: 'TableNormal',
      width: { value: 2500, type: 'pct' },
      alignment: 'center',
      indent: 144,
      layout: 'fixed',
      cellMargins: { top: 0, left: 100, bottom: 0, right: 100 },
      borders: {},
    });
    expect(resolveCell(t.rows()[0]!.cells()[0]!)).toMatchObject({
      width: { value: 1000, type: 'dxa' },
      margins: { top: 50, left: 100, bottom: 0, right: 100 },
      vAlign: 'center',
      noWrap: true,
      textDirection: 'btLr',
      borders: {},
    });
  });
});

describe('caching follows XDocument.version', () => {
  it('sees edits to the styles part and to the paragraph', () => {
    const doc = makeDoc({
      styles: pStyle('Normal', '<w:pPr><w:jc w:val="center"/></w:pPr>', ' w:default="1"'),
      body: para(''),
    });
    const p = paras(doc)[0]!;
    expect(resolveParagraph(p).alignment).toBe('center');
    expect(resolveParagraph(p)).toBe(resolveParagraph(p));
    const styles = doc.mainPart.relatedByType('styles')[0]!.xml();
    const jc = styles.root.descendants(NS.w, 'jc')[0]!;
    jc.setAttrNS(NS.w, 'val', 'right');
    expect(resolveParagraph(p).alignment).toBe('right');
    p.style = 'Missing';
    expect(resolveParagraph(p).alignment).toBe('right');
    const pPr = p.element.element(NS.w, 'pPr')!;
    const ownJc = doc.body.element.ownerDocument()!.createElement(NS.w, 'jc', { 'w:val': 'both' });
    pPr.append(ownJc);
    expect(resolveParagraph(p).alignment).toBe('both');
  });

  it('renumbers after a paragraph is inserted', () => {
    const doc = makeDoc({
      numbering:
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
      body: para('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>').repeat(2),
    });
    const [p0, p1] = paras(doc);
    expect([listLabel(p0!), listLabel(p1!)]).toEqual(['1.', '2.']);
    p0!.element.before(p0!.element.clone());
    expect(paras(doc).map((p) => listLabel(p))).toEqual(['1.', '2.', '3.']);
    expect(listLabel(p1!)).toBe('3.');
  });
});

describe('real fixtures', () => {
  it('docxcompose/numberings.docx: list indents and Symbol bullets', () => {
    const doc = readDocx(fixture('docxcompose/numberings.docx').bytes);
    const [first, , sub] = doc.body.paragraphs();
    expect(resolveParagraph(first!)).toMatchObject({
      styleId: 'Listenabsatz',
      indent: { left: 720, hanging: 360 },
      contextualSpacing: true,
      numbering: { numId: 1, level: 0 },
    });
    expect(resolveParagraph(sub!).indent.left).toBe(1440);
    expect(resolveRun(first!.runs()[0]!).fonts.ascii).toBe('Calibri');
    expect(themeOf(doc).fonts.minor.latin).toBe('Calibri');
  });

  it('python-docx/tbl-cell-access.docx: Table Grid borders on every cell', () => {
    const doc = readDocx(fixture('python-docx/tbl-cell-access.docx').bytes);
    const t = doc.body.tables()[0]!;
    expect(resolveTable(t)).toMatchObject({
      styleId: 'TableGrid',
      cellMargins: { left: 108, right: 108 },
    });
    for (const row of t.rows()) {
      for (const c of row.cells()) {
        const b = resolveCell(c).borders;
        for (const side of ['top', 'left', 'bottom', 'right'] as const) {
          expect(b[side]).toMatchObject({ style: 'single', size: 4 });
        }
      }
    }
  });

  it('open-xml-sdk/TableFntEdnt.docx: East Asian theme fonts and font hint', () => {
    const doc = readDocx(fixture('open-xml-sdk/TableFntEdnt.docx').bytes);
    const p = doc.body.paragraphs()[0]!;
    const r = resolveRun(p.runs()[0]!);
    expect(r.fonts).toMatchObject({ ascii: 'Calibri', eastAsia: '宋体', hint: 'eastAsia' });
    expect(r.size).toBe(10.5);
    expect(resolveParagraph(p).alignment).toBe('both');
  });
});
