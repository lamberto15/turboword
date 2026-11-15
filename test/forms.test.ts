// turboword/forms on synthetic documents: one control of every type and level, built in the
// test from createDocx() with a hand-written body, so each rule is checked in isolation.
// Every document that is modified is saved, validated (validatePackage: no errors), read back,
// and written to out/forms/ for `node scripts/libreoffice-check.mjs out/forms`.

import { mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { parseIsoDate } from '../src/forms/date-format.js';
import {
  bindContentControls,
  type ContentControl,
  contentControls,
  fillContentControls,
  fillLegacyFormFields,
  formatWordDate,
  legacyFormFields,
} from '../src/forms/index.js';
import { evaluateXPath, parseXPath, readHit } from '../src/forms/xpath.js';
import {
  createDocx,
  NS,
  parseXml,
  readDocx,
  serializeXml,
  UsageError,
  validatePackage,
  type WordDocument,
  writeDocx,
  type XElement,
} from '../src/index.js';

const OUT = fileURLToPath(new URL('../out/forms/', import.meta.url));
const ZIP_DATE = new Date(Date.UTC(2024, 0, 1));
const BASE = unzipSync(writeDocx(createDocx(), { zipDate: ZIP_DATE }));
const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const PLACEHOLDER_STYLE =
  '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>';
const NS_A = `xmlns:a="${NS.a}"`;
const NS_PIC = `xmlns:pic="${NS.pic}"`;

// 1×1 PNGs (red, transparent).
const PNG_RED = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);
const PNG_CLEAR = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=',
  ),
  (c) => c.charCodeAt(0),
);

/** A document whose body is `body` (plus a final sectPr); extra parts and rels optional. */
function makeDoc(
  body: string,
  extra?: { parts?: Record<string, string>; rels?: string },
): WordDocument {
  const files: Record<string, Uint8Array> = { ...BASE };
  const xml = strFromU8(files['word/document.xml']!).replace(
    /<w:body>[\s\S]*<\/w:body>/,
    `<w:body>${body}${SECT}</w:body>`,
  );
  files['word/document.xml'] = strToU8(xml);
  files['word/styles.xml'] = strToU8(
    strFromU8(files['word/styles.xml']!).replace('</w:styles>', `${PLACEHOLDER_STYLE}</w:styles>`),
  );
  for (const [name, text] of Object.entries(extra?.parts ?? {})) files[name] = strToU8(text);
  if (extra?.rels !== undefined) {
    files['word/_rels/document.xml.rels'] = strToU8(
      strFromU8(files['word/_rels/document.xml.rels']!).replace(
        '</Relationships>',
        `${extra.rels}</Relationships>`,
      ),
    );
  }
  return readDocx(zipSync(files, { mtime: ZIP_DATE }));
}

const sdt = (pr: string, content: string): string =>
  `<w:sdt><w:sdtPr>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const tag = (t: string): string => `<w:alias w:val="A ${t}"/><w:tag w:val="${t}"/>`;
const run = (text: string, rPr = ''): string =>
  `<w:r>${rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner: string, pPr = ''): string =>
  `<w:p>${pPr === '' ? '' : `<w:pPr>${pPr}</w:pPr>`}${inner}</w:p>`;

function byTag(doc: WordDocument, t: string): ContentControl {
  const c = contentControls(doc).find((x) => x.tag === t);
  if (c === undefined) throw new Error(`no control ${t}`);
  return c;
}

function xmlOf(el: XElement): string {
  const d = parseXml('<x/>');
  d.root.append(el.clone());
  return serializeXml(d);
}

/** Saves, checks validity, writes out/forms/<name>.docx, returns the re-read document. */
function save(doc: WordDocument, name: string): WordDocument {
  const bytes = writeDocx(doc, { zipDate: ZIP_DATE });
  mkdirSync(OUT, { recursive: true });
  writeFileSync(`${OUT}${name}.docx`, bytes);
  const again = readDocx(bytes);
  const errors = validatePackage(again).filter((i) => i.severity === 'error');
  expect(errors, name).toEqual([]);
  expect(writeDocx(doc, { zipDate: ZIP_DATE })).toEqual(bytes); // deterministic
  return again;
}

const CHECKBOX = (checked: string, extra = ''): string =>
  `<w14:checkbox><w14:checked w14:val="${checked}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>${extra}</w14:checkbox>`;
const GOTHIC =
  '<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/>';

// ---- discovery ---------------------------------------------------------------------------

describe('contentControls: types', () => {
  const kinds: [string, string][] = [
    ['<w:text/>', 'plainText'],
    ['<w:richText/>', 'richText'],
    ['', 'richText'],
    [CHECKBOX('0'), 'checkbox'],
    [
      '<w:dropDownList><w:listItem w:displayText="A" w:value="a"/></w:dropDownList>',
      'dropDownList',
    ],
    ['<w:comboBox/>', 'comboBox'],
    ['<w:date><w:dateFormat w:val="M/d/yyyy"/></w:date>', 'date'],
    ['<w:picture/>', 'picture'],
    ['<w:group/>', 'group'],
    [
      '<w:docPartObj><w:docPartGallery w:val="Cover Pages"/><w:docPartUnique/></w:docPartObj>',
      'buildingBlockGallery',
    ],
    [
      '<w:docPartList><w:docPartGallery w:val="Quick Parts"/></w:docPartList>',
      'buildingBlockGallery',
    ],
    ['<w:citation/>', 'citation'],
    ['<w:bibliography/>', 'bibliography'],
    ['<w:equation/>', 'equation'],
    ['<w14:entityPicker/>', 'unknown'],
  ];
  const body = kinds
    .map(([pr, _t], i) => sdt(`${tag(`t${i}`)}<w:id w:val="${100 + i}"/>${pr}`, para(run(`v${i}`))))
    .join('');
  const doc = makeDoc(body);

  it('classifies every control type from sdtPr', () => {
    const cs = contentControls(doc);
    expect(cs.map((c) => c.type)).toEqual(kinds.map(([, t]) => t));
    expect(cs.map((c) => c.tag)).toEqual(kinds.map((_k, i) => `t${i}`));
    expect(cs[0]!.alias).toBe('A t0');
    expect(cs[0]!.id).toBe(100);
    expect(cs.every((c) => c.level === 'block' && c.story.kind === 'body')).toBe(true);
    expect(cs[1]!.text).toBe('v1');
  });

  it('reads repeating sections and their items', () => {
    const d = makeDoc(
      sdt(
        `${tag('sec')}<w15:repeatingSection/>`,
        sdt('<w15:repeatingSectionItem/>', para(run('item'))),
      ),
    );
    expect(contentControls(d).map((c) => c.type)).toEqual([
      'repeatingSection',
      'repeatingSectionItem',
    ]);
  });
});

describe('contentControls: levels, nesting and order', () => {
  const tb = sdt(tag('tb'), para(run('in text box')));
  const doc = makeDoc(
    sdt(tag('block'), para(run('block'))) +
      para(
        `${run('before ')}${sdt(tag('inline'), run('inline'))}<w:hyperlink w:anchor="x">${sdt(tag('link'), run('link'))}</w:hyperlink>`,
      ) +
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      sdt(tag('row'), `<w:tr><w:tc>${para(run('row'))}</w:tc></w:tr>`) +
      `<w:tr>${sdt(tag('cell'), `<w:tc>${para(run('cell'))}</w:tc>`)}</w:tr>` +
      `<w:tr><w:tc>${sdt(tag('inCell'), para(run('in cell')))}</w:tc></w:tr>` +
      '</w:tbl>' +
      sdt(tag('outer'), sdt(tag('nested'), para(`${run('x ')}${sdt(tag('deep'), run('deep'))}`))) +
      para(
        `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="tb"/><a:graphic ${NS_A}><a:graphicData uri="${NS.wps}"><wps:wsp><wps:txbx><w:txbxContent>${tb}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>${tb}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`,
      ),
  );

  it('finds controls in document order, outer before nested, Choice only', () => {
    expect(contentControls(doc).map((c) => c.tag)).toEqual([
      'block',
      'inline',
      'link',
      'row',
      'cell',
      'inCell',
      'outer',
      'nested',
      'deep',
      'tb',
    ]);
  });

  it('reads the level from the context', () => {
    const levels = Object.fromEntries(contentControls(doc).map((c) => [c.tag, c.level]));
    expect(levels).toEqual({
      block: 'block',
      inline: 'inline',
      link: 'inline',
      row: 'row',
      cell: 'cell',
      inCell: 'block',
      outer: 'block',
      nested: 'block',
      deep: 'inline',
      tb: 'block',
    });
  });

  it('reads visible text per level', () => {
    expect(byTag(doc, 'inline').text).toBe('inline');
    expect(byTag(doc, 'row').text).toBe('row');
    expect(byTag(doc, 'outer').text).toBe('x deep');
  });

  it('writes into row and cell controls (first cell) and keeps the table valid', () => {
    const d = makeDoc(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
        sdt(
          `${tag('row')}<w:text/>`,
          `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para(run('a', '<w:b/>'))}</w:tc><w:tc>${para(run('b'))}</w:tc></w:tr>`,
        ) +
        `<w:tr>${sdt(`${tag('cell')}<w:text/>`, `<w:tc>${para(run('c'))}</w:tc>`)}<w:tc>${para(run('d'))}</w:tc></w:tr></w:tbl>${para('')}`,
    );
    byTag(d, 'row').setValue('ROW');
    byTag(d, 'cell').setValue('CELL');
    const again = save(d, 'levels-row-cell');
    expect(byTag(again, 'row').text).toBe('ROW\nb');
    expect(byTag(again, 'cell').text).toBe('CELL');
    const row = byTag(again, 'row').element;
    expect(row.descendants(NS.w, 'tcPr').length).toBe(1);
    expect(row.descendants(NS.w, 'b').length).toBe(1);
  });
});

describe('contentControls: properties', () => {
  const doc = makeDoc(
    sdt(
      `${tag('p1')}<w:id w:val="-5"/><w:lock w:val="sdtContentLocked"/><w:showingPlcHdr/><w:dataBinding w:prefixMappings="xmlns:ns0='urn:x'" w:xpath="/ns0:r[1]/ns0:a[1]" w:storeItemID="{ABC}"/><w:text/>`,
      para(run('Click here', '<w:rStyle w:val="PlaceholderText"/>')),
    ) +
      sdt(
        `${tag('p2')}<w:lock w:val="sdtLocked"/><w15:dataBinding w:xpath="/r/items[1]/item[1]"/><w15:repeatingSection/>`,
        sdt('<w15:repeatingSectionItem/>', para(run('i'))),
      ) +
      sdt(
        `${tag('p3')}<w:lock w:val="contentLocked"/><w:showingPlcHdr w:val="0"/><w:comboBox><w:listItem w:displayText="One" w:value="1"/><w:listItem w:value="2"/><w:listItem w:displayText="Three"/></w:comboBox>`,
        para(run('One')),
      ),
  );

  it('reads id, lock, placeholder, binding and list items', () => {
    const [a, b, c] = [byTag(doc, 'p1'), byTag(doc, 'p2'), byTag(doc, 'p3')];
    expect(a.id).toBe(-5);
    expect(a.locked).toEqual({ content: true, control: true });
    expect(b.locked).toEqual({ content: false, control: true });
    expect(c.locked).toEqual({ content: true, control: false });
    expect(a.placeholderShown).toBe(true);
    expect(c.placeholderShown).toBe(false);
    expect(a.value).toBeNull();
    expect(a.text).toBe('Click here');
    expect(a.binding).toEqual({
      xpath: '/ns0:r[1]/ns0:a[1]',
      storeItemId: '{ABC}',
      prefixMappings: "xmlns:ns0='urn:x'",
    });
    expect(b.binding).toEqual({ xpath: '/r/items[1]/item[1]' });
    expect(c.items).toEqual([
      { display: 'One', value: '1' },
      { display: '2', value: '2' },
      { display: 'Three', value: 'Three' },
    ]);
    expect(c.value).toBe('1');
  });

  it('refuses to set content-locked controls, not control-locked ones', () => {
    expect(() => byTag(doc, 'p1').setValue('x')).toThrow(UsageError);
    expect(() => byTag(doc, 'p3').setValue('2')).toThrow(/locked/);
    expect(() => byTag(doc, 'p2').setValue('x')).toThrow(/nested controls/);
  });
});

// ---- setValue ------------------------------------------------------------------------------

describe('setValue: text', () => {
  it('replaces inline content keeping the first run formatting, clearing the placeholder', () => {
    const doc = makeDoc(
      para(
        run('Name: ') +
          sdt(
            `${tag('name')}<w:showingPlcHdr/><w:text/>`,
            run('Click here', '<w:rStyle w:val="PlaceholderText"/><w:b/>'),
          ),
      ),
    );
    const c = byTag(doc, 'name');
    c.setValue('Ann\nLee\tX');
    expect(c.placeholderShown).toBe(false);
    expect(c.text).toBe('Ann\nLee\tX');
    expect(c.value).toBe('Ann\nLee\tX');
    const content = c.element.element(NS.w, 'sdtContent')!;
    expect(content.elements().length).toBe(1);
    const r = content.element(NS.w, 'r')!;
    expect(xmlOf(r)).toContain('<w:rPr><w:b/></w:rPr>');
    expect(r.descendants(NS.w, 'rStyle')).toEqual([]);
    expect(r.descendants(NS.w, 'br').length).toBe(1);
    expect(r.descendants(NS.w, 'tab').length).toBe(1);
    const again = save(doc, 'text-inline');
    expect(again.text()).toBe('Name: Ann\nLee\tX');
  });

  it("uses the control's own rPr while the placeholder is shown", () => {
    const doc = makeDoc(
      para(
        sdt(
          `<w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr>${tag('x')}<w:showingPlcHdr/>`,
          run('Type here', '<w:rStyle w:val="PlaceholderText"/>'),
        ),
      ),
    );
    byTag(doc, 'x').setValue('hello');
    const r = byTag(doc, 'x').element.descendants(NS.w, 'r')[0]!;
    expect(xmlOf(r)).toContain('<w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr><w:t>hello</w:t>');
  });

  it('writes rich text lines as paragraphs, keeping the first paragraph and markers', () => {
    const doc = makeDoc(
      sdt(
        `${tag('rich')}<w:richText/>`,
        `<w:p w14:paraId="1A2B3C4D" w14:textId="77777777"><w:pPr><w:jc w:val="center"/></w:pPr>${run('old', '<w:u w:val="single"/>')}</w:p>` +
          `<w:bookmarkStart w:id="7" w:name="mark"/>${para(run('second'))}<w:bookmarkEnd w:id="7"/>`,
      ),
    );
    const c = byTag(doc, 'rich');
    c.setValue('a\nb\n\nc');
    const ps = c.element.element(NS.w, 'sdtContent')!.elements(NS.w, 'p');
    expect(ps.length).toBe(4);
    expect(ps[0]!.attrNS(NS.w14, 'paraId')).toBe('1A2B3C4D');
    expect(ps.slice(1).every((p) => p.attrNS(NS.w14, 'paraId') === undefined)).toBe(true);
    expect(ps.every((p) => p.descendants(NS.w, 'jc').length === 1)).toBe(true);
    expect(c.text).toBe('a\nb\n\nc');
    expect(c.element.descendants(NS.w, 'u').length).toBe(3);
    expect(c.element.descendants(NS.w, 'bookmarkStart').length).toBe(1);
    expect(c.element.descendants(NS.w, 'bookmarkEnd').length).toBe(1);
    expect(ps[0]!.descendants(NS.w, 'bookmarkStart').length).toBe(1);
    expect(ps[3]!.descendants(NS.w, 'bookmarkEnd').length).toBe(1);
    save(doc, 'text-rich-block');
  });

  it('keeps plain text in one paragraph and moves a dropped section break to the end', () => {
    const doc = makeDoc(
      sdt(
        `${tag('plain')}<w:text w:multiLine="1"/>`,
        para(run('one')) +
          para(
            run('two'),
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
          ),
      ),
    );
    const c = byTag(doc, 'plain');
    c.setValue('x\ny');
    const ps = c.element.descendants(NS.w, 'p');
    expect(ps.length).toBe(1);
    expect(ps[0]!.descendants(NS.w, 'br').length).toBe(1);
    expect(ps[0]!.element(NS.w, 'pPr')?.element(NS.w, 'sectPr')).toBeDefined();
    expect(save(doc, 'text-plain-block').sections().length).toBe(2);
  });

  it('converts other values and rejects bytes', () => {
    const doc = makeDoc(para(sdt(tag('t'), run('x'))));
    const c = byTag(doc, 't');
    c.setValue(new Date(Date.UTC(2024, 4, 3)));
    expect(c.text).toBe('2024-05-03');
    c.setValue(true);
    expect(c.text).toBe('true');
    c.setValue(null);
    expect(c.text).toBe('');
    expect(() => c.setValue(new Uint8Array(1))).toThrow(UsageError);
  });

  it('unwraps temporary controls once set', () => {
    const doc = makeDoc(para(run('A ') + sdt(`${tag('tmp')}<w:temporary/>`, run('x'))));
    byTag(doc, 'tmp').setValue('B');
    expect(contentControls(doc)).toEqual([]);
    expect(doc.text()).toBe('A B');
  });
});

describe('setValue: checkbox', () => {
  it('sets w14:checked and swaps the glyph', () => {
    const doc = makeDoc(para(sdt(`${tag('cb')}${CHECKBOX('0')}`, run('☐', GOTHIC))));
    const c = byTag(doc, 'cb');
    expect(c.value).toBe(false);
    c.setValue(true);
    expect(c.value).toBe(true);
    expect(c.text).toBe('☒');
    c.setValue('no');
    expect(c.text).toBe('☐');
    c.setValue('X');
    expect(c.value).toBe(true);
    expect(() => c.setValue('maybe')).toThrow(UsageError);
    expect(() => c.setValue(new Date())).toThrow(UsageError);
    save(doc, 'checkbox');
  });

  it('uses the state fonts (theme fonts removed) and creates a missing w14:checked', () => {
    const doc = makeDoc(
      para(
        sdt(
          `${tag('cb')}<w14:checkbox><w14:checkedState w14:val="0050" w14:font="Wingdings 2"/><w14:uncheckedState w14:val="00A3" w14:font="Wingdings 2"/></w14:checkbox>`,
          run(
            '£',
            '<w:rFonts w:asciiTheme="minorHAnsi" w:ascii="Wingdings 2" w:hAnsi="Wingdings 2"/>',
          ),
        ),
      ),
    );
    const c = byTag(doc, 'cb');
    c.setValue(true);
    expect(c.text).toBe('P');
    const fonts = c.element.element(NS.w, 'sdtContent')!.descendants(NS.w, 'rFonts')[0]!;
    expect(fonts.attrNS(NS.w, 'asciiTheme')).toBeUndefined();
    expect(fonts.attrNS(NS.w, 'eastAsia')).toBe('Wingdings 2');
    expect(c.element.descendants(NS.w14, 'checked')[0]!.attrNS(NS.w14, 'val')).toBe('1');
  });

  it('updates a w:sym glyph', () => {
    const doc = makeDoc(
      para(
        sdt(
          `${tag('cb')}<w14:checkbox><w14:checked w14:val="1"/><w14:checkedState w14:val="F0FE" w14:font="Wingdings"/><w14:uncheckedState w14:val="F0A8" w14:font="Wingdings"/></w14:checkbox>`,
          '<w:r><w:sym w:font="Wingdings" w:char="F0FE"/></w:r>',
        ),
      ),
    );
    byTag(doc, 'cb').setValue(false);
    const sym = byTag(doc, 'cb').element.descendants(NS.w, 'sym')[0]!;
    expect(sym.attrNS(NS.w, 'char')).toBe('F0A8');
    expect(byTag(doc, 'cb').value).toBe(false);
  });
});

describe('setValue: dropdown and combo box', () => {
  const list =
    '<w:listItem w:displayText="Choose" w:value=""/><w:listItem w:displayText="Red" w:value="r"/><w:listItem w:displayText="Green" w:value="g"/>';
  it('selects by value or display text, sets lastValue, rejects unknown items', () => {
    const doc = makeDoc(
      para(
        sdt(
          `${tag('dd')}<w:showingPlcHdr/><w:dropDownList>${list}</w:dropDownList>`,
          run('Choose'),
        ),
      ),
    );
    const c = byTag(doc, 'dd');
    expect(c.value).toBeNull();
    c.setValue('g');
    expect([c.text, c.value]).toEqual(['Green', 'g']);
    expect(c.element.descendants(NS.w, 'dropDownList')[0]!.attrNS(NS.w, 'lastValue')).toBe('g');
    c.setValue('Red');
    expect([c.text, c.value]).toEqual(['Red', 'r']);
    expect(() => c.setValue('blue')).toThrow(/not an item.*"r", "g"/);
    save(doc, 'dropdown');
  });

  it('accepts free text in a combo box', () => {
    const doc = makeDoc(
      para(sdt(`${tag('cb')}<w:comboBox w:lastValue="r">${list}</w:comboBox>`, run('Red'))),
    );
    const c = byTag(doc, 'cb');
    expect(c.value).toBe('r');
    c.setValue('Purple');
    expect([c.text, c.value]).toEqual(['Purple', 'Purple']);
    expect(c.element.descendants(NS.w, 'comboBox')[0]!.attrNS(NS.w, 'lastValue')).toBeUndefined();
  });
});

describe('setValue: date', () => {
  const date = (inner: string, attrs = ''): string => `<w:date${attrs}>${inner}</w:date>`;
  it('writes w:fullDate and the formatted text', () => {
    const doc = makeDoc(
      para(
        sdt(
          `${tag('d')}<w:showingPlcHdr/>${date('<w:dateFormat w:val="dddd, MMMM d, yyyy"/><w:lid w:val="en-US"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/>')}`,
          run('Click to enter a date', '<w:rStyle w:val="PlaceholderText"/>'),
        ),
      ),
    );
    const c = byTag(doc, 'd');
    expect(c.value).toBeNull();
    c.setValue(new Date(Date.UTC(2024, 0, 15)));
    expect(c.text).toBe('Monday, January 15, 2024');
    expect(c.element.descendants(NS.w, 'date')[0]!.attrNS(NS.w, 'fullDate')).toBe(
      '2024-01-15T00:00:00Z',
    );
    expect((c.value as Date).toISOString()).toBe('2024-01-15T00:00:00.000Z');
    c.setValue('2024-02-29');
    expect(c.text).toBe('Thursday, February 29, 2024');
    expect(() => c.setValue('2024-02-30')).toThrow(UsageError);
    expect(() => c.setValue(true)).toThrow(UsageError);
    save(doc, 'date');
    c.setValue(null);
    expect(c.text).toBe('');
    expect(c.value).toBeNull();
  });

  it('formats in the w:lid language, else the run language, else M/d/yyyy en-US', () => {
    const doc = makeDoc(
      para(
        sdt(
          `${tag('de')}${date('<w:dateFormat w:val="d. MMMM yyyy"/><w:lid w:val="de-DE"/>')}`,
          run('x'),
        ),
      ) +
        para(
          sdt(
            `${tag('fr')}${date('<w:dateFormat w:val="d MMMM yyyy"/>')}`,
            run('x', '<w:lang w:val="fr-FR"/>'),
          ),
        ) +
        para(sdt(`${tag('none')}${date('')}`, run('x'))),
    );
    const d = new Date(Date.UTC(2026, 9, 22));
    for (const t of ['de', 'fr', 'none']) byTag(doc, t).setValue(d);
    expect(byTag(doc, 'de').text).toBe('22. Oktober 2026');
    expect(byTag(doc, 'fr').text).toBe('22 octobre 2026');
    expect(byTag(doc, 'none').text).toBe('10/22/2026');
  });
});

describe('formatWordDate and parseIsoDate', () => {
  const d = new Date(Date.UTC(2026, 9, 22, 15, 5, 9));
  it.each([
    ['ddd d MMM yyyy', 'en-AU', 'Thu 22 Oct 2026'],
    ['h:mm am/pm', 'en-AU', '3:05 pm'],
    ['h:mm AM/PM', 'en-AU', '3:05 PM'],
    ['HH:mm:ss', 'en-AU', '15:05:09'],
    ['hh:m:s', 'en-US', '03:5:9'],
    ['d MMMM yyyy', 'de-DE', '22 Oktober 2026'],
    ['dddd, d MMMM yyyy', 'en-AU', 'Thursday, 22 October 2026'],
    ["'Day' d 'of' MMMM", 'en-AU', 'Day 22 of October'],
    ['MMM. yyyy', 'de-DE', 'Okt. 2026'],
    ['d MMMM yyyy', 'ru-RU', '22 октября 2026'],
    ['MMMM yyyy', 'ru-RU', 'октябрь 2026'],
    ['M/d/yy', 'en-US', '10/22/26'],
    ['dd.MM.yyyy', 'not a tag!', '22.10.2026'],
    ["yyyy-MM-dd'T'HH", 'en-US', '2026-10-22T15'],
    ["d 'o''clock'", 'en-US', "22 o'clock"],
  ])('%s (%s) → %s', (pic, lid, want) => {
    expect(formatWordDate(d, pic, lid)).toBe(want);
  });

  it('parses ISO dates and date-times, rejecting impossible ones', () => {
    expect(parseIsoDate('2026-10-22')?.toISOString()).toBe('2026-10-22T00:00:00.000Z');
    expect(parseIsoDate('2026-10-22T15:05:00Z')?.toISOString()).toBe('2026-10-22T15:05:00.000Z');
    expect(parseIsoDate('2026-10-22T15:05:00+02:00')?.toISOString()).toBe(
      '2026-10-22T13:05:00.000Z',
    );
    expect(parseIsoDate('2026-10-22T15:05:00.250')?.toISOString()).toBe('2026-10-22T15:05:00.250Z');
    for (const bad of ['2026-13-45', 'not-a-date', '', '2026-02-29', '22/10/2026']) {
      expect(parseIsoDate(bad), bad).toBeUndefined();
    }
  });
});

describe('setValue: picture', () => {
  it('creates a drawing in an empty picture control, then swaps the blip keeping the extent', () => {
    const doc = makeDoc(
      para(sdt(`${tag('pic')}<w:showingPlcHdr/><w:picture/>`, '')) + para(run('after')),
    );
    const c = byTag(doc, 'pic');
    c.setValue(PNG_RED);
    expect(c.placeholderShown).toBe(false);
    expect(c.value).toEqual(PNG_RED);
    const extent = c.element.descendants(NS.wp, 'extent')[0]!;
    extent.setAttr('cx', '914400');
    extent.setAttr('cy', '457200');
    const firstRel = c.element.descendants(NS.a, 'blip')[0]!.attrNS(NS.r, 'embed');
    c.setValue(PNG_CLEAR);
    expect(c.value).toEqual(PNG_CLEAR);
    expect(c.element.descendants(NS.wp, 'extent')[0]!.attr('cx')).toBe('914400');
    expect(c.element.descendants(NS.a, 'blip').length).toBe(1);
    // The first image is no longer referenced: its relationship is gone.
    expect(doc.mainPart.rels.get(firstRel!)).toBeUndefined();
    expect(() => c.setValue('x')).toThrow(UsageError);
    const again = save(doc, 'picture');
    expect(byTag(again, 'pic').value).toEqual(PNG_CLEAR);
  });

  it(`replaces the blip of a Word picture control (${NS_PIC.length > 0 ? 'DrawingML' : ''})`, () => {
    const drawing = (rid: string): string =>
      `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1905000" cy="1905000"/><wp:docPr id="5" name="Picture 5"/><a:graphic ${NS_A}><a:graphicData uri="${NS.pic}"><pic:pic ${NS_PIC}><pic:nvPicPr><pic:cNvPr id="0" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const doc = makeDoc(para(sdt(`${tag('pic')}<w:picture/>`, drawing('rIdImg'))), {
      parts: { 'word/media/placeholder.png': '' },
      rels: '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/placeholder.png"/>',
    });
    doc.package.part('word/media/placeholder.png')!.setBytes(PNG_CLEAR);
    const c = byTag(doc, 'pic');
    expect(c.value).toEqual(PNG_CLEAR);
    c.setValue(PNG_RED);
    expect(c.value).toEqual(PNG_RED);
    expect(c.element.descendants(NS.wp, 'extent')[0]!.attr('cx')).toBe('1905000');
    expect(doc.mainPart.rels.get('rIdImg')).toBeUndefined();
    save(doc, 'picture-word');
  });
});

// ---- fill ---------------------------------------------------------------------------------

describe('fillContentControls', () => {
  const form = (): WordDocument =>
    makeDoc(
      para(
        run('To: ') +
          sdt(
            `${tag('name')}<w:showingPlcHdr/><w:text/>`,
            run('Name', '<w:rStyle w:val="PlaceholderText"/>'),
          ),
      ) +
        para(run('Again: ') + sdt(`${tag('name')}<w:text/>`, run('old'))) +
        para(sdt(`${tag('ok')}${CHECKBOX('0')}`, run('☐', GOTHIC))) +
        para(sdt(`${tag('untouched')}`, run('keep me'))),
    );

  it('fills by tag and alias, reports missing keys', () => {
    const doc = form();
    expect(fillContentControls(doc, { name: 'Ann', ok: true, nope: 'x' })).toEqual({
      filled: 3,
      missing: ['nope'],
    });
    expect(doc.text()).toBe('To: Ann\nAgain: Ann\n☒\nkeep me');
    expect(fillContentControls(doc, { 'A name': 'Bob' }, { by: 'alias' })).toEqual({
      filled: 2,
      missing: [],
    });
    expect(doc.text()).toContain('To: Bob');
    save(doc, 'fill-basic');
  });

  it('removes the filled controls, keeping their content', () => {
    const doc = form();
    fillContentControls(doc, { name: 'Ann', ok: true }, { removeControls: true });
    expect(contentControls(doc).map((c) => c.tag)).toEqual(['untouched']);
    expect(doc.text()).toBe('To: Ann\nAgain: Ann\n☒\nkeep me');
    save(doc, 'fill-remove');
  });

  const repeating = (inner: string, sectionPr = ''): string =>
    sdt(
      `${tag('items')}<w:id w:val="10"/>${sectionPr}<w15:repeatingSection/>`,
      sdt(`<w:id w:val="11"/><w15:repeatingSectionItem/>`, inner),
    );

  it('fills block-level repeating sections with fresh ids', () => {
    const item =
      `<w:p w14:paraId="00000ABC" w14:textId="00000DEF"><w:bookmarkStart w:id="3" w:name="row"/>` +
      sdt(`${tag('name')}<w:id w:val="12"/><w:text/>`, run('name', '<w:b/>')) +
      run(' × ') +
      sdt(`${tag('qty')}<w:id w:val="13"/><w:text/>`, run('0')) +
      '<w:bookmarkEnd w:id="3"/></w:p>';
    const doc = makeDoc(
      repeating(
        item,
        '<w15:dataBinding w:xpath="/r/item[1]" w:storeItemID="{11111111-2222-3333-4444-555555555555}"/>',
      ),
    );
    const res = fillContentControls(doc, {
      items: [
        { name: 'Apple', qty: '1' },
        { name: 'Pear', qty: '2' },
        { name: 'Plum', qty: '3' },
      ],
    });
    expect(res).toEqual({ filled: 7, missing: [] });
    expect(doc.text()).toBe('Apple × 1\nPear × 2\nPlum × 3');
    const all = contentControls(doc);
    const ids = all.map((c) => c.id);
    expect(new Set(ids).size).toBe(ids.length);
    const body = doc.body.element;
    const paraIds = body.descendants(NS.w, 'p').map((p) => p.attrNS(NS.w14, 'paraId'));
    expect(new Set(paraIds).size).toBe(3);
    expect(body.descendants(NS.w, 'bookmarkStart').length).toBe(0);
    expect(body.descendants(NS.w15, 'dataBinding').length).toBe(0);
    expect(body.descendants(NS.w, 'b').length).toBe(3);
    save(doc, 'fill-repeating-block');
  });

  it('fills row-level and nested repeating sections', () => {
    const cell = (inner: string): string =>
      `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para(inner)}</w:tc>`;
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell(run('SKU'))}${cell(run('Qty'))}</w:tr>` +
      sdt(
        `${tag('lines')}<w15:repeatingSection/>`,
        sdt(
          '<w15:repeatingSectionItem/>',
          `<w:tr>${cell(sdt(`${tag('sku')}<w:text/>`, run('-')))}${cell(sdt(`${tag('qty')}<w:text/>`, run('-')))}</w:tr>`,
        ),
      ) +
      '</w:tbl>';
    const doc = makeDoc(
      sdt(
        `${tag('orders')}<w15:repeatingSection/>`,
        sdt(
          '<w15:repeatingSectionItem/>',
          para(run('Order ') + sdt(`${tag('id')}<w:text/>`, run('?'))) + table + para(''),
        ),
      ),
    );
    fillContentControls(doc, {
      orders: [
        {
          id: 'A1',
          lines: [
            { sku: 'x', qty: '1' },
            { sku: 'y', qty: '2' },
          ],
        },
        { id: 'B2', lines: [{ sku: 'z', qty: '9' }] },
      ],
    });
    expect(doc.body.tables().length).toBe(2);
    expect(doc.text()).toBe('Order A1\nSKU\nQty\nx\n1\ny\n2\n\nOrder B2\nSKU\nQty\nz\n9\n');
    const again = save(doc, 'fill-repeating-rows');
    expect(again.body.element.descendants(NS.w, 'tr').length).toBe(5);
  });

  it('removes a repeating section given no records, keeping cells valid', () => {
    const doc = makeDoc(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>' +
        repeating(para(sdt(`${tag('name')}<w:text/>`, run('-')))) +
        `</w:tc></w:tr></w:tbl>${para(run('end'))}`,
    );
    expect(fillContentControls(doc, { items: [] }).filled).toBe(1);
    expect(contentControls(doc)).toEqual([]);
    expect(doc.body.element.descendants(NS.w, 'tc')[0]!.elements(NS.w, 'p').length).toBe(1);
    save(doc, 'fill-repeating-empty');
  });

  it('rejects arrays for plain controls and scalars for sections', () => {
    const doc = makeDoc(repeating(para(sdt(`${tag('name')}<w:text/>`, run('-')))));
    expect(() => fillContentControls(doc, { name: [{ a: 'b' }] })).toThrow(
      /arrays fill repeating sections/,
    );
    expect(() => fillContentControls(doc, { items: 'x' })).toThrow(/array of records/);
  });

  it('applies records before scalars, which only reach pre-existing controls', () => {
    const doc = makeDoc(
      repeating(para(sdt(`${tag('name')}<w:text/>`, run('-')))) +
        para(sdt(`${tag('name')}<w:text/>`, run('-'))),
    );
    fillContentControls(doc, { name: 'outside', items: [{ name: 'in1' }, { name: 'in2' }] });
    expect(doc.text()).toBe('in1\nin2\noutside');
  });

  it('writes bound values through to the custom XML part', () => {
    const store = '{9A8B7C6D-0000-4000-8000-000000000001}';
    const doc = makeDoc(
      para(
        sdt(
          `${tag('who')}<w:dataBinding w:prefixMappings="xmlns:ns0='urn:t'" w:xpath="/ns0:root[1]/ns0:who[1]" w:storeItemID="${store}"/><w:text/>`,
          run('old'),
        ) +
          sdt(
            `${tag('ok')}<w:dataBinding w:xpath="/root[1]/@ok" w:storeItemID="${store.toLowerCase()}"/>${CHECKBOX('0')}`,
            run('☐', GOTHIC),
          ),
      ),
    );
    bindContentControls(doc, '<root xmlns="urn:t" ok="false"><who>old</who></root>', {
      storeItemId: store,
    });
    fillContentControls(doc, { who: 'Zoë & co', ok: true });
    const xml = strFromU8(doc.package.part('customXml/item1.xml')!.bytes());
    expect(xml).toContain('<who>Zoë &amp; co</who>');
    // `/root[1]/@ok`: unprefixed step = no namespace, so it does not match the urn:t root.
    expect(xml).toContain('ok="false"');
    save(doc, 'fill-write-through');
  });
});

// ---- data binding ---------------------------------------------------------------------------

describe('bindContentControls', () => {
  const STORE = '{B1D0F00D-0000-4000-8000-000000000001}';
  const bound = (t: string, xpath: string, pr: string, content: string, store = STORE): string =>
    sdt(
      `${tag(t)}<w:dataBinding w:prefixMappings="xmlns:ns0='urn:probe' xmlns:ns1='urn:attr'" w:xpath="${xpath}" w:storeItemID="${store}"/>${pr}`,
      content,
    );
  const doc = (): WordDocument =>
    makeDoc(
      para(bound('name', '/ns0:root[1]/ns0:name[1]', '<w:text/>', run('stale', '<w:b/>'))) +
        para(bound('second', '/ns0:root[1]/ns0:list[1]/ns0:v[2]', '<w:text/>', run('stale'))) +
        para(bound('attr', '/ns0:root[1]/ns0:list[1]/@ns1:code', '<w:text/>', run('stale'))) +
        para(bound('txt', '/ns0:root[1]/ns0:mixed[1]/text()', '<w:text/>', run('stale'))) +
        para(bound('chk', '/ns0:root[1]/ns0:chk[1]', CHECKBOX('0'), run('☐', GOTHIC))) +
        para(
          bound(
            'when',
            '/ns0:root[1]/ns0:when[1]',
            '<w:date><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="en-GB"/><w:storeMappedDataAs w:val="dateTime"/></w:date>',
            run('stale'),
          ),
        ) +
        para(
          bound(
            'dd',
            '/ns0:root[1]/ns0:dd[1]',
            '<w:dropDownList><w:listItem w:displayText="One" w:value="v1"/><w:listItem w:displayText="Two" w:value="v2"/></w:dropDownList>',
            run('One'),
          ),
        ) +
        para(bound('pic', '/ns0:root[1]/ns0:pic[1]', '<w:picture/>', '')) +
        para(
          bound(
            'other',
            '/ns0:root[1]/ns0:name[1]',
            '<w:text/>',
            run('other store'),
            '{00000000-0000-4000-8000-000000000009}',
          ),
        ),
    );
  const DATA = `<?xml version="1.0" encoding="UTF-16"?><root xmlns="urn:probe" xmlns:a="urn:attr"><name>Ann</name><list a:code="C-7"><v>first</v><v>second</v></list><mixed>lead<i/>tail</mixed><chk>true</chk><when>2026-10-22T00:00:00</when><dd>v2</dd><pic>data:image/png;base64,${btoa(String.fromCharCode(...PNG_RED))}</pic></root>`;

  it('creates the store named by the controls and refreshes every bound control', () => {
    const d = doc();
    expect(bindContentControls(d, DATA)).toBe(8);
    const v = (t: string): unknown => byTag(d, t).value;
    expect(v('name')).toBe('Ann');
    expect(byTag(d, 'name').element.descendants(NS.w, 'b').length).toBe(1);
    expect(v('second')).toBe('second');
    expect(v('attr')).toBe('C-7');
    expect(v('txt')).toBe('leadtail');
    expect(v('chk')).toBe(true);
    expect(byTag(d, 'when').text).toBe('22 October 2026');
    expect(v('dd')).toBe('v2');
    expect(byTag(d, 'dd').text).toBe('Two');
    expect(v('pic')).toEqual(PNG_RED);
    expect(v('other')).toBe('other store');
    // The store parts and relationships.
    const item = d.package.part('customXml/item1.xml')!;
    expect(strFromU8(item.bytes())).toMatch(/^<\?xml version="1.0" encoding="UTF-8"\?>/);
    const props = strFromU8(d.package.part('customXml/itemProps1.xml')!.bytes());
    expect(props).toContain(`ds:itemID="${STORE}"`);
    expect(props).toContain('<ds:schemaRef ds:uri="urn:probe"/>');
    expect(d.mainPart.rels.byType('customXml').map((r) => r.target)).toEqual([
      '../customXml/item1.xml',
    ]);
    expect(item.rels.byType('customXmlProps').map((r) => r.target)).toEqual(['itemProps1.xml']);
    expect(d.package.contentTypes.get('customXml/itemProps1.xml')).toBe(
      'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
    );
    const again = save(d, 'bind-new-store');
    expect(byTag(again, 'name').text).toBe('Ann');
  });

  it('updates an existing store in place', () => {
    const d = doc();
    bindContentControls(d, DATA);
    bindContentControls(
      d,
      DATA.replace('<name>Ann</name>', '<name>Bea</name>').replace('<chk>true', '<chk>0'),
    );
    expect(byTag(d, 'name').text).toBe('Bea');
    expect(byTag(d, 'chk').value).toBe(false);
    expect(d.package.parts().filter((p) => p.name.startsWith('customXml/')).length).toBe(3);
  });

  it('derives a deterministic GUID when nothing names a store', () => {
    const make = (): WordDocument =>
      makeDoc(sdt(`${tag('n')}<w:dataBinding w:xpath="/r/n"/><w:text/>`, para(run('-'))));
    const a = make();
    const b = make();
    expect(bindContentControls(a, '<r><n>1</n></r>')).toBe(1);
    bindContentControls(b, '<r><n>1</n></r>');
    const id = (d: WordDocument): string =>
      /ds:itemID="([^"]+)"/.exec(
        strFromU8(d.package.part('customXml/itemProps1.xml')!.bytes()),
      )![1]!;
    expect(id(a)).toBe(id(b));
    expect(id(a)).toMatch(
      /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/,
    );
    const c = make();
    bindContentControls(c, '<r><n>2</n></r>');
    expect(id(c)).not.toBe(id(a));
    expect(byTag(a, 'n').text).toBe('1');
  });

  it('expands bound repeating sections, outermost first', () => {
    const d = makeDoc(
      sdt(
        `${tag('groups')}<w15:dataBinding w:xpath="/r[1]/g[1]" w:storeItemID="${STORE}"/><w15:repeatingSection/>`,
        sdt(
          '<w15:repeatingSectionItem/>',
          para(
            sdt(
              `${tag('title')}<w:dataBinding w:xpath="/r[1]/g[1]/@t" w:storeItemID="${STORE}"/><w:text/>`,
              run('-'),
            ),
          ) +
            sdt(
              `${tag('items')}<w15:dataBinding w:xpath="/r[1]/g[1]/i[1]" w:storeItemID="${STORE}"/><w15:repeatingSection/>`,
              sdt(
                '<w15:repeatingSectionItem/>',
                para(
                  run('• ') +
                    sdt(
                      `${tag('item')}<w:dataBinding w:xpath="/r[1]/g[1]/i[1]" w:storeItemID="${STORE}"/><w:text/>`,
                      run('-'),
                    ),
                ),
              ),
            ),
        ),
      ),
    );
    const n = bindContentControls(
      d,
      '<r><g t="A"><i>a1</i><i>a2</i></g><g t="B"><i>b1</i></g><g t="C"/></r>',
    );
    expect(d.text()).toBe('A\n• a1\n• a2\nB\n• b1\nC\n• -');
    // 1 outer + 3 inner sections expanded, 3 titles and 3 items set.
    expect(n).toBe(10);
    const xpaths = d.body.element
      .descendants(NS.w, 'dataBinding')
      .map((b) => b.attrNS(NS.w, 'xpath'));
    expect(xpaths).toContain('/r[1]/g[2]/i[1]');
    expect(xpaths).toContain('/r[1]/g[1]/i[2]');
    save(d, 'bind-repeating');
  });
});

describe('XPath subset', () => {
  const x = parseXml(
    '<a xmlns:p="urn:p"><b k="1">one</b><b>two<c/>three</b><p:d p:k="pk">ns</p:d></a>',
  );
  const read = (path: string, maps?: string): string | undefined => {
    const h = evaluateXPath(x, path, maps);
    return h === undefined ? undefined : readHit(h);
  };
  it('evaluates absolute paths, positions, attributes, text() and prefixes', () => {
    expect(read('/a/b')).toBe('one');
    expect(read('/a[1]/b[2]')).toBe('twothree');
    expect(read('/a/b[2]/text()')).toBe('twothree');
    expect(read('/a/b/@k')).toBe('1');
    expect(read('/a/*[3]')).toBe('ns');
    expect(read('/a/q:d/@q:k', "xmlns:q='urn:p'")).toBe('pk');
    expect(read('/a/d')).toBeUndefined();
    expect(read('/a/b[3]')).toBeUndefined();
    expect(read('/a/b[2]/@k')).toBeUndefined();
  });
  it('reports paths outside the subset as unsupported', () => {
    for (const p of ['//b', 'a/b', '/a/b[@k=1]', '/a/count(b)', '/a/u:b', '/@k', '/a/@k/b', ' ']) {
      expect(parseXPath(p), p).toBeUndefined();
    }
  });
});

// ---- legacy form fields --------------------------------------------------------------------

describe('legacy form fields', () => {
  const begin = (ff: string): string =>
    `<w:r><w:fldChar w:fldCharType="begin"><w:ffData>${ff}</w:ffData></w:fldChar></w:r>`;
  const instr = (code: string): string =>
    `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>`;
  const sep = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const EN = ' '.repeat(5);
  const doc = (): WordDocument =>
    makeDoc(
      para(
        run('Name: ') +
          begin(
            '<w:name w:val="Text1"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput><w:maxLength w:val="10"/><w:format w:val="UPPERCASE"/></w:textInput>',
          ) +
          '<w:bookmarkStart w:id="0" w:name="Text1"/>' +
          instr('FORMTEXT') +
          sep +
          run(EN, '<w:noProof/><w:i/>') +
          end +
          '<w:bookmarkEnd w:id="0"/>',
      ) +
        para(begin('<w:name w:val="NoResult"/><w:textInput/>') + instr('FORMTEXT') + end) +
        para(
          begin(
            '<w:name w:val="Check1"/><w:checkBox><w:sizeAuto/><w:default w:val="1"/></w:checkBox>',
          ) +
            instr('FORMCHECKBOX') +
            end,
        ) +
        para(
          begin(
            '<w:name w:val="Drop1"/><w:ddList><w:listEntry w:val="Alpha"/><w:listEntry w:val="Beta"/></w:ddList>',
          ) +
            instr('FORMDROPDOWN') +
            end,
        ) +
        para(`${begin('')}${instr('PAGE')}${sep}${run('1')}${end}`),
    );

  it('reads type, name and value', () => {
    expect(legacyFormFields(doc()).map((f) => [f.type, f.name, f.value])).toEqual([
      ['text', 'Text1', ''],
      ['text', 'NoResult', ''],
      ['checkbox', 'Check1', true],
      ['dropdown', 'Drop1', 'Alpha'],
    ]);
  });

  it('writes results, checked state and dropdown index', () => {
    const d = doc();
    const res = fillLegacyFormFields(d, {
      Text1: 'ann lee',
      NoResult: 'x',
      Check1: false,
      Drop1: 'Beta',
      Nope: 'z',
    });
    expect(res).toEqual({ filled: 4, missing: ['Nope'] });
    expect(legacyFormFields(d).map((f) => f.value)).toEqual(['ANN LEE', 'x', false, 'Beta']);
    expect(d.text()).toBe('Name: ANN LEE\nx\n\n\n1');
    const text1 = legacyFormFields(d)[0]!;
    // The result run keeps the formatting of the old result.
    const resultRun = (text1.element.parent!.parent as XElement)
      .descendants(NS.w, 'r')
      .find((r) => r.text === 'ANN LEE')!;
    expect(xmlOf(resultRun)).toContain('<w:noProof/><w:i/>');
    const checked = d.body.element.descendants(NS.w, 'checkBox')[0]!;
    expect(checked.elements().map((e) => e.local)).toEqual(['sizeAuto', 'default', 'checked']);
    const again = save(d, 'legacy');
    expect(legacyFormFields(again).map((f) => f.value)).toEqual(['ANN LEE', 'x', false, 'Beta']);
    legacyFormFields(d)[0]!.setValue('');
    expect(legacyFormFields(d)[0]!.value).toBe('');
    expect(d.text()).toContain(`Name: ${EN}`);
  });

  it('validates values', () => {
    const [text, , check, drop] = legacyFormFields(doc());
    expect(() => text!.setValue('far too long for it')).toThrow(/at most 10/);
    expect(() => text!.setValue(true)).toThrow(UsageError);
    expect(() => check!.setValue('perhaps')).toThrow(UsageError);
    expect(() => drop!.setValue('Gamma')).toThrow(/"Alpha", "Beta"/);
  });
});
