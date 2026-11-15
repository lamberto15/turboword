// turboword/forms on real documents: content controls and legacy form fields written by Word
// (and by docx4j's probes), from test/fixtures/docxcompose and test/fixtures/forms/<source>/
// (licences and pinned commits in test/fixtures/forms/LICENSE). Every modified document is
// saved, checked for new validation errors against its input, and written to out/forms/ for
// `node scripts/libreoffice-check.mjs out/forms`.

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  bindContentControls,
  type ContentControl,
  contentControls,
  fillContentControls,
  fillLegacyFormFields,
  legacyFormFields,
} from '../src/forms/index.js';
import {
  NS,
  readDocx,
  UsageError,
  type ValidationIssue,
  validatePackage,
  type WordDocument,
  writeDocx,
} from '../src/index.js';
import { fixture } from './helpers/fixtures.js';

const OUT = fileURLToPath(new URL('../out/forms/', import.meta.url));
const ZIP_DATE = new Date(Date.UTC(2024, 0, 1));

function load(name: string): { doc: WordDocument; bytes: Uint8Array } {
  const bytes = name.startsWith('forms/')
    ? new Uint8Array(readFileSync(fileURLToPath(new URL(`./fixtures/${name}`, import.meta.url))))
    : fixture(name).bytes;
  return { doc: readDocx(bytes), bytes };
}

function errors(doc: WordDocument): string[] {
  return validatePackage(doc)
    .filter((i: ValidationIssue) => i.severity === 'error')
    .map((i) => `${i.code} ${i.part ?? ''}`);
}

/** Saves, checks for new validation errors, writes out/forms/<name>.docx, returns the re-read doc. */
function save(doc: WordDocument, input: Uint8Array, name: string): WordDocument {
  const before = new Set(errors(readDocx(input)));
  const bytes = writeDocx(doc, { zipDate: ZIP_DATE });
  mkdirSync(OUT, { recursive: true });
  writeFileSync(`${OUT}${name}.docx`, bytes);
  const again = readDocx(bytes);
  expect(
    errors(again).filter((e) => !before.has(e)),
    name,
  ).toEqual([]);
  return again;
}

function find(doc: WordDocument, pred: (c: ContentControl) => boolean): ContentControl {
  const c = contentControls(doc).find(pred);
  if (c === undefined) throw new Error('no such control');
  return c;
}

const byAlias = (doc: WordDocument, a: string): ContentControl => find(doc, (c) => c.alias === a);
const byTag = (doc: WordDocument, t: string): ContentControl => find(doc, (c) => c.tag === t);

describe('docxcompose/content_controls.docx (Word, German placeholders)', () => {
  it('lists the controls without touching the file', () => {
    const { doc, bytes } = load('docxcompose/content_controls.docx');
    const cs = contentControls(doc);
    expect(cs.map((c) => [c.alias, c.type, c.level, c.placeholderShown])).toEqual([
      ['cc.plain_text', 'plainText', 'block', false],
      ['cc.plain_text_multiline', 'plainText', 'block', false],
      ['cc.plain_text_empty', 'plainText', 'block', true],
      ['cc.rich_text', 'richText', 'block', true],
    ]);
    expect(cs.map((c) => c.value)).toEqual([
      'Content Control Plain Text',
      'Content Control Plain Text Multiline',
      null,
      null,
    ]);
    expect(cs[2]!.text).toBe('Klicken oder tippen Sie hier, um Text einzugeben.');
    expect(cs.map((c) => c.id)).toEqual([1004477788, 1267113498, 779070967, 971171640]);
    expect(writeDocx(doc)).toBe(bytes);
  });

  it('fills by alias, dropping the localized placeholder style', () => {
    const { doc, bytes } = load('docxcompose/content_controls.docx');
    const res = fillContentControls(
      doc,
      {
        'cc.plain_text': 'Neuer Text',
        'cc.plain_text_multiline': 'Zeile 1\nZeile 2',
        'cc.plain_text_empty': 'Gefüllt',
        'cc.rich_text': 'Absatz 1\nAbsatz 2',
      },
      { by: 'alias' },
    );
    expect(res).toEqual({ filled: 4, missing: [] });
    const body = doc.body.element;
    expect(body.descendants(NS.w, 'rStyle').map((s) => s.attrNS(NS.w, 'val'))).not.toContain(
      'Platzhaltertext',
    );
    expect(body.descendants(NS.w, 'showingPlcHdr')).toEqual([]);
    // fr-CH language of the first run is kept.
    const first = byAlias(doc, 'cc.plain_text').element.descendants(NS.w, 'r')[0]!;
    expect(first.descendants(NS.w, 'lang')[0]?.attrNS(NS.w, 'val')).toBe('fr-CH');
    const again = save(doc, bytes, 'content_controls-filled');
    expect(again.text()).toBe(
      'Content Controls\nNeuer Text\nZeile 1\nZeile 2\nGefüllt\nAbsatz 1\nAbsatz 2\n',
    );
  });

  it('removes the controls when asked', () => {
    const { doc, bytes } = load('docxcompose/content_controls.docx');
    fillContentControls(
      doc,
      { 'cc.plain_text': 'x', 'cc.rich_text': 'y' },
      { removeControls: true },
    );
    expect(contentControls(doc).map((c) => c.tag)).toEqual([
      'cc.plain_text_multiline',
      'cc.plain_text_empty',
    ]);
    save(doc, bytes, 'content_controls-flattened');
  });
});

describe('Open-XML-SDK SDT files', () => {
  it.each([
    ['sdtRun.docx', 'inline', 'sdtRun'],
    ['sdtRow.docx', 'row', 'This is an sdtRow.\n\n'],
    ['simpleSdt.docx', 'block', 'Click here to enter text.'],
  ])('%s: a %s control', (file, level, text) => {
    const { doc, bytes } = load(`forms/open-xml-sdk/${file}`);
    const [c] = contentControls(doc);
    expect(c!.level).toBe(level);
    expect(c!.text).toBe(text);
    c!.setValue('Replaced');
    expect(c!.text.startsWith('Replaced')).toBe(true);
    save(doc, bytes, `oxsdk-${file.replace('.docx', '')}`);
  });

  it('sdtCell.docx: cell-level controls', () => {
    const { doc, bytes } = load('forms/open-xml-sdk/sdtCell.docx');
    const cs = contentControls(doc);
    expect(cs.map((c) => c.level)).toEqual(['cell', 'cell']);
    cs[0]!.setValue('A');
    cs[1]!.setValue('B');
    expect(cs.map((c) => c.text)).toEqual(['A', 'B']);
    save(doc, bytes, 'oxsdk-sdtCell');
  });

  it('simpleSdt.docx: the PlaceholderText style goes with the placeholder', () => {
    const { doc } = load('forms/open-xml-sdk/simpleSdt.docx');
    const c = byTag(doc, 'sdt');
    expect(c.placeholderShown).toBe(true);
    c.setValue('Hello');
    expect(c.element.descendants(NS.w, 'rStyle')).toEqual([]);
  });

  it('sdtDataBinding.docx: binding without a store id creates one', () => {
    const { doc, bytes } = load('forms/open-xml-sdk/sdtDataBinding.docx');
    expect(contentControls(doc)[0]!.binding).toEqual({ xpath: '/root/name/first' });
    expect(bindContentControls(doc, '<root><name><first>Grace</first></name></root>')).toBe(1);
    expect(contentControls(doc)[0]!.text).toBe('Grace');
    expect(doc.mainPart.relatedByType('customXml').length).toBe(1);
    save(doc, bytes, 'oxsdk-sdtDataBinding');
  });
});

describe('docx4j', () => {
  it('LegacyForms.docx: FORMTEXT, FORMCHECKBOX, FORMDROPDOWN', () => {
    const { doc, bytes } = load('forms/docx4j/LegacyForms.docx');
    const fields = legacyFormFields(doc);
    expect(fields.map((f) => [f.type, f.name, f.value])).toEqual([
      ['text', 'Text1', ''],
      ['checkbox', 'Check1', false],
      ['dropdown', 'Dropdown1', ''],
    ]);
    expect(writeDocx(doc)).toBe(bytes);
  });

  it('LegacyForms.docx: fills text and checkbox fields', () => {
    const { doc, bytes } = load('forms/docx4j/LegacyForms.docx');
    expect(() => fillLegacyFormFields(doc, { Dropdown1: 'x' })).toThrow(UsageError);
    expect(fillLegacyFormFields(doc, { Text1: 'Filled in', Check1: true })).toEqual({
      filled: 2,
      missing: [],
    });
    const again = save(doc, bytes, 'docx4j-LegacyForms');
    expect(legacyFormFields(again).map((f) => f.value)).toEqual(['Filled in', true, '']);
    expect(again.text()).toContain('Text form field Filled in');
  });

  it('invoice2013.docx: every typed control, rebound to new data', () => {
    const { doc, bytes } = load('forms/docx4j/invoice2013.docx');
    const types = contentControls(doc).map((c) => c.type);
    for (const t of [
      'picture',
      'plainText',
      'date',
      'repeatingSection',
      'repeatingSectionItem',
      'checkbox',
      'richText',
    ]) {
      expect(types, t).toContain(t);
    }
    const store = doc.mainPart.relatedByType('customXml')[0]!;
    const xml = new TextDecoder()
      .decode(store.bytes())
      .replace('John Citizen', 'Ada Lovelace')
      .replace('2015-01-29T00:00:00', '2025-12-24T00:00:00')
      .replace(
        /<lines>[\s\S]*<\/lines>/,
        '<lines>' +
          ['A', 'B', 'C']
            .map(
              (x, i) =>
                `<lineitem><productcode>P${x}</productcode><description>Item ${x}</description><quantity>${i + 1}</quantity><price>$${10 * (i + 1)}</price></lineitem>`,
            )
            .join('') +
          '</lines>',
      )
      .replace('applies="true"', 'applies="false"');
    const n = bindContentControls(doc, xml);
    expect(n).toBeGreaterThan(20);
    const text = doc.text();
    expect(text).toContain('Ada Lovelace');
    expect(text).toContain('24 December 2025');
    for (const x of ['PA', 'PB', 'PC', 'Item A', 'Item C']) expect(text).toContain(x);
    const sections = contentControls(doc).filter((c) => c.type === 'repeatingSection');
    const items = (c: ContentControl): number =>
      c.element
        .element(NS.w, 'sdtContent')!
        .elements(NS.w, 'sdt')
        .filter((s) => s.descendants(NS.w15, 'repeatingSectionItem').length > 0).length;
    expect(sections.map(items)).toEqual([3, 2]); // lines, notes
    const vat = find(doc, (c) => c.type === 'checkbox');
    expect(vat.value).toBe(false);
    const logo = find(doc, (c) => c.type === 'picture').value as Uint8Array;
    expect([...logo.subarray(1, 4)]).toEqual([0x50, 0x4e, 0x47]); // PNG
    // The rich text binding (escaped Flat OPC) is left alone.
    expect(find(doc, (c) => c.type === 'richText').text).not.toContain('pkg:package');
    const ids = contentControls(doc)
      .map((c) => c.id)
      .filter((i) => i !== undefined);
    expect(new Set(ids).size).toBe(ids.length);
    save(doc, bytes, 'docx4j-invoice2013-bound');
  });

  it('typed-controls.docx: date, dropdown, combo box and checkbox from XML', () => {
    const { doc, bytes } = load('forms/docx4j/typed-controls.docx');
    const store = doc.mainPart.relatedByType('customXml')[0]!;
    expect(bindContentControls(doc, new TextDecoder().decode(store.bytes()))).toBe(4);
    expect(['DT1', 'DD2', 'CB3', 'CK4'].map((t) => byTag(doc, t).text)).toEqual([
      '22 October 2026',
      'DD2_SECOND_ITEM',
      'CB3_XML_FREE_TEXT',
      '☒',
    ]);
    expect(byTag(doc, 'DD2').value).toBe('v2');
    expect(byTag(doc, 'CK4').value).toBe(true);
    save(doc, bytes, 'docx4j-typed-controls');
  });

  it('typed-controls-dates.docx: one date picture each', () => {
    const { doc, bytes } = load('forms/docx4j/typed-controls-dates.docx');
    const store = doc.mainPart.relatedByType('customXml')[0]!;
    bindContentControls(doc, new TextDecoder().decode(store.bytes()));
    const got = Array.from({ length: 13 }, (_v, i) => byTag(doc, `DF${i + 1}`).text);
    expect(got).toEqual([
      'Thu 22 Oct 2026',
      '3:05 pm',
      '12:05 PM',
      '15:05',
      '22 Oktober 2026',
      'Thursday, 22 October 2026',
      '22 octobre 2026', // no w:lid: the control's language (fr-FR)
      'not-a-date', // unparseable: shown as is
      '2026-13-45',
      '',
      '2026-10-22',
      'Day 22 of October',
      'Okt. 2026',
    ]);
    expect(byTag(doc, 'DF8').value).toBeNull();
    expect((byTag(doc, 'DF2').value as Date).toISOString()).toBe('2026-10-22T15:05:00.000Z');
    save(doc, bytes, 'docx4j-typed-controls-dates');
  });

  it('picture.docx: a picture bound to base64 data', () => {
    const { doc, bytes } = load('forms/docx4j/picture.docx');
    const pic = find(doc, (c) => c.type === 'picture');
    const before = pic.value as Uint8Array;
    const store = doc.mainPart.relatedByType('customXml')[0]!;
    expect(bindContentControls(doc, new TextDecoder().decode(store.bytes()))).toBe(2);
    const after = find(doc, (c) => c.type === 'picture');
    expect(after.placeholderShown).toBe(false);
    expect(after.value).not.toEqual(before);
    expect(after.element.descendants(NS.wp, 'extent')[0]!.attr('cx')).toBe('1899920');
    save(doc, bytes, 'docx4j-picture');
  });

  it('w15-repeating-section-three-items.docx: one item becomes three', () => {
    const { doc, bytes } = load('forms/docx4j/w15-repeating-section-three-items.docx');
    const store = doc.mainPart.relatedByType('customXml')[0]!;
    bindContentControls(doc, new TextDecoder().decode(store.bytes()));
    expect(
      contentControls(doc)
        .filter((c) => c.tag === 'name')
        .map((c) => c.text),
    ).toEqual(['RS_ITEM_A', 'RS_ITEM_B', 'RS_ITEM_C']);
    const again = save(doc, bytes, 'docx4j-w15-repeating');
    expect(again.text()).toContain('Item: RS_ITEM_C');
  });

  it('w15-repeating-section-three-items.docx: filled with records instead', () => {
    const { doc, bytes } = load('forms/docx4j/w15-repeating-section-three-items.docx');
    fillContentControls(doc, { Items: [{ name: 'one' }, { name: 'two' }] });
    expect(
      contentControls(doc)
        .filter((c) => c.tag === 'name')
        .map((c) => c.text),
    ).toEqual(['one', 'two']);
    // Unbound: Word must not rebuild the section from the (three-item) XML.
    expect(doc.body.element.descendants(NS.w15, 'dataBinding')).toEqual([]);
    save(doc, bytes, 'docx4j-w15-repeating-filled');
  });
});

describe('Apache POI', () => {
  it('Bug54849.docx: controls in every story and at every level', () => {
    const { doc, bytes } = load('forms/apache-poi/Bug54849.docx');
    const cs = contentControls(doc);
    expect(cs.map((c) => c.story.kind)).toEqual([
      ...Array(9).fill('body'),
      'header',
      'footer',
      'footnote',
      'endnote',
    ]);
    expect(cs.map((c) => c.type).slice(0, 7)).toEqual([
      'richText',
      'richText',
      'plainText',
      'plainText',
      'comboBox',
      'dropDownList',
      'date',
    ]);
    expect(new Set(cs.map((c) => c.level))).toEqual(new Set(['block', 'inline', 'cell']));
    expect(cs[0]!.tag).toBe('MyTag');
    expect(cs.slice(9).map((c) => c.text)).toEqual([
      'header_rich_text',
      'Footer_rich_text',
      'Footnote_sdt',
      'Endnote_sdt',
    ]);
    for (const c of cs.slice(9)) c.setValue(`new ${c.story.kind}`);
    const combo = cs[4]!;
    combo.setValue(combo.items[1]?.value ?? 'x');
    const date = cs[6]!;
    date.setValue('2013-04-17');
    const again = save(doc, bytes, 'poi-Bug54849');
    expect(
      contentControls(again)
        .slice(9)
        .map((c) => c.text),
    ).toEqual(['new header', 'new footer', 'new footnote', 'new endnote']);
    expect(contentControls(again)[6]!.text).toBe('4/17/2013');
  });

  it('Bug54771a.docx: controls bound to document properties write through to them', () => {
    const { doc, bytes } = load('forms/apache-poi/Bug54771a.docx');
    const cs = contentControls(doc);
    expect(cs[0]!.type).toBe('buildingBlockGallery');
    const title = byAlias(doc, 'Title');
    expect(title.text).toBe('Test');
    title.setValue('A New Title');
    expect(doc.properties.title).toBe('A New Title');
    expect(() =>
      bindContentControls(doc, '<x/>', { storeItemId: '{6C3C8BC8-F283-45AE-878A-BAB7291924A1}' }),
    ).toThrow(/properties/);
    save(doc, bytes, 'poi-Bug54771a');
  });

  it('checkboxes.docx: legacy checkboxes', () => {
    const { doc, bytes } = load('forms/apache-poi/checkboxes.docx');
    const fields = legacyFormFields(doc);
    expect(fields.length).toBe(9);
    expect(fields.every((f) => f.type === 'checkbox')).toBe(true);
    const flipped = fields.map((f) => !f.value);
    fields.forEach((f, i) => {
      f.setValue(flipped[i]!);
    });
    const again = save(doc, bytes, 'poi-checkboxes');
    expect(legacyFormFields(again).map((f) => f.value)).toEqual(flipped);
  });

  it('form_footnotes.docx: many legacy text and checkbox fields', () => {
    const { doc, bytes } = load('forms/apache-poi/form_footnotes.docx');
    const fields = legacyFormFields(doc);
    expect(fields.length).toBeGreaterThan(60);
    expect(new Set(fields.map((f) => f.type))).toEqual(new Set(['text', 'checkbox']));
    const res = fillLegacyFormFields(doc, { Text2: 'two', Text3: '42', Check19: true });
    expect(res.missing).toEqual([]);
    const again = save(doc, bytes, 'poi-form_footnotes');
    const named = legacyFormFields(again).filter((f) => f.name === 'Text2' || f.name === 'Check19');
    expect(named.map((f) => f.value)).toEqual(['two', true]);
  });
});
