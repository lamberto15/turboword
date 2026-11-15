// createDocx(): a new, minimal, valid document.
//
// The template is what Word itself would write for "Blank document", trimmed to the parts a
// consumer needs (no theme, web settings or glossary): the main part with Word's usual
// namespace declarations, styles (Calibri 11 pt defaults, Normal, headings, Title), settings in
// Word 2013+ compatibility mode, a font table and the two property parts. US Letter, 1"
// margins. Everything is a fixed string, so the output is byte-for-byte deterministic: core
// properties carry no timestamps (set them through `doc.properties` when wanted).

import { OpcPackage } from '../opc/package.js';
import { CT, REL } from '../opc/rel-types.js';
import { NS } from '../xml/ns.js';
import { WordDocument } from './document.js';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** `xmlns:p="uri"` declarations for the given conventional prefixes. */
function xmlns(...prefixes: (keyof typeof NS)[]): string {
  return prefixes.map((p) => `xmlns:${p}="${NS[p]}"`).join(' ');
}

const DOCUMENT_XML = `${DECL}<w:document ${xmlns(
  'wpc',
  'mc',
  'o',
  'r',
  'm',
  'v',
  'wp14',
  'wp',
  'w10',
  'w',
  'w14',
  'w15',
  'wpg',
  'wpi',
  'wne',
  'wps',
)} mc:Ignorable="w14 w15 wp14"><w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`;

const LIGHT = '<w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/>';

function heading(level: number, before: number, color: string, size: number): string {
  return (
    `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/>` +
    '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/>' +
    `${level > 1 ? '<w:unhideWhenUsed/>' : ''}<w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="0"/>` +
    `<w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
    `<w:rPr>${LIGHT}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr>` +
    '</w:style>'
  );
}

const STYLES_XML = `${DECL}<w:styles ${xmlns('mc', 'r', 'w', 'w14', 'w15')} mc:Ignorable="w14 w15">\
<w:docDefaults><w:rPrDefault><w:rPr>\
<w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>\
<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>\
</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>\
<w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>\
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>\
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">\
<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>\
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>\
<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>\
<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>\
<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>\
<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>\
<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>\
${heading(1, 240, '2F5496', 32)}${heading(2, 40, '2F5496', 26)}${heading(3, 40, '1F3763', 24)}\
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>\
<w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/>\
<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>\
<w:rPr>${LIGHT}<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>\
</w:style></w:styles>`;

const WORD_URI = 'http://schemas.microsoft.com/office/word';

function compat(name: string, val: string): string {
  return `<w:compatSetting w:name="${name}" w:uri="${WORD_URI}" w:val="${val}"/>`;
}

const SETTINGS_XML = `${DECL}<w:settings ${xmlns('mc', 'r', 'w', 'w14', 'w15')} mc:Ignorable="w14 w15">\
<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>\
<w:characterSpacingControl w:val="doNotCompress"/><w:compat>\
${compat('compatibilityMode', '15')}${compat('overrideTableStyleFontSizeAndJustification', '1')}\
${compat('enableOpenTypeFeatures', '1')}${compat('doNotFlipMirrorIndents', '1')}\
${compat('differentiateMultirowTableHeaders', '1')}</w:compat>\
<w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>`;

function font(name: string, panose: string, family: string): string {
  return (
    `<w:font w:name="${name}"><w:panose1 w:val="${panose}"/><w:charset w:val="00"/>` +
    `<w:family w:val="${family}"/><w:pitch w:val="variable"/></w:font>`
  );
}

const FONT_TABLE_XML = `${DECL}<w:fonts ${xmlns('mc', 'r', 'w', 'w14', 'w15')} mc:Ignorable="w14 w15">\
${font('Calibri', '020F0502020204030204', 'swiss')}\
${font('Times New Roman', '02020603050405020304', 'roman')}\
${font('Calibri Light', '020F0302020204030204', 'swiss')}</w:fonts>`;

const CORE_XML = `${DECL}<cp:coreProperties ${xmlns('cp', 'dc', 'dcterms', 'dcmitype', 'xsi')}/>`;

const APP_XML = `${DECL}<Properties xmlns="${NS.ep}" ${xmlns('vt')}>\
<Application>turboword</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>\
<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>\
<HyperlinksChanged>false</HyperlinksChanged></Properties>`;

/** A new blank document: one empty paragraph, US Letter, 1" margins, Calibri 11 pt. */
export function createDocx(): WordDocument {
  const enc = new TextEncoder();
  const pkg = OpcPackage.create();
  pkg.rels.add(REL.officeDocument, 'word/document.xml');
  pkg.rels.add(REL.coreProperties, 'docProps/core.xml');
  pkg.rels.add(REL.extendedProperties, 'docProps/app.xml');

  const main = pkg.addPart('word/document.xml', CT.document, enc.encode(DOCUMENT_XML));
  main.rels.add(REL.styles, 'styles.xml');
  main.rels.add(REL.settings, 'settings.xml');
  main.rels.add(REL.fontTable, 'fontTable.xml');
  pkg.addPart('word/styles.xml', CT.styles, enc.encode(STYLES_XML));
  pkg.addPart('word/settings.xml', CT.settings, enc.encode(SETTINGS_XML));
  pkg.addPart('word/fontTable.xml', CT.fontTable, enc.encode(FONT_TABLE_XML));
  pkg.addPart('docProps/core.xml', CT.coreProperties, enc.encode(CORE_XML));
  pkg.addPart('docProps/app.xml', CT.extendedProperties, enc.encode(APP_XML));
  return new WordDocument(pkg);
}
