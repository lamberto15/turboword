// The wml core: the visible-text flattener, stories, paragraphs/runs, createDocx and the
// read/load/write entry points, plus smoke tests over the fixture corpus.

import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  CT,
  createDocx,
  DOCX_MIME,
  DocumentProperties,
  DocxFormatError,
  docxBlob,
  docxStream,
  loadDocx,
  NS,
  Paragraph,
  paragraphText,
  parseXml,
  REL,
  readDocx,
  serializeXml,
  Table,
  TurbowordError,
  UsageError,
  validatePackage,
  WordDocument,
  walkVisible,
  writeDocx,
  XElement,
  XText,
} from '../src/index.js';
import { insertOrdered, PPR_ORDER, RPR_ORDER, wElement } from '../src/wml/order.js';
import { blockParagraphs } from '../src/wml/text.js';
import { fixtures as curatedFixtures } from './helpers/fixtures.js';

// ---- helpers ------------------------------------------------------------------------------

const DECLS = [
  `xmlns:w="${NS.w}"`,
  `xmlns:r="${NS.r}"`,
  `xmlns:mc="${NS.mc}"`,
  `xmlns:m="${NS.m}"`,
  `xmlns:wp="${NS.wp}"`,
  `xmlns:a="${NS.a}"`,
  `xmlns:v="${NS.v}"`,
  `xmlns:wps="${NS.wps}"`,
  `xmlns:w14="${NS.w14}"`,
].join(' ');

/** First element of an XML fragment (wrapped in a root declaring the usual prefixes). */
function frag(xml: string): XElement {
  return parseXml(`<w:wrap ${DECLS}>${xml}</w:wrap>`).root.firstElement()!;
}

const flat = (xml: string): string => paragraphText(frag(xml));
const locals = (el: XElement | undefined): string[] => el?.elements().map((e) => e.local) ?? [];
const run = (body: string): string => `<w:r>${body}</w:r>`;
const t = (s: string): string => run(`<w:t xml:space="preserve">${s}</w:t>`);

const TEXTBOX = run(
  '<w:drawing><wp:inline><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>' +
    `<w:p>${t('boxed')}</w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic>` +
    '</wp:inline></w:drawing>',
);

interface SyntheticParts {
  /** Inner XML of w:body. */
  body: string;
  footnotes?: string;
  endnotes?: string;
  comments?: string;
  header?: string;
  footer?: string;
}

/** A .docx built from scratch (independent of createDocx). */
function docxFrom(o: SyntheticParts): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const overrides: [string, string][] = [['word/document.xml', CT.document]];
  const rels: [string, string][] = [];
  const add = (name: string, ct: string, rel: string, root: string, inner: string): void => {
    files[`word/${name}`] = strToU8(`<?xml version="1.0"?><${root} ${DECLS}>${inner}</${root}>`);
    overrides.push([`word/${name}`, ct]);
    rels.push([rel, name]);
  };
  if (o.footnotes !== undefined)
    add('footnotes.xml', CT.footnotes, REL.footnotes, 'w:footnotes', o.footnotes);
  if (o.endnotes !== undefined)
    add('endnotes.xml', CT.endnotes, REL.endnotes, 'w:endnotes', o.endnotes);
  if (o.comments !== undefined)
    add('comments.xml', CT.comments, REL.comments, 'w:comments', o.comments);
  if (o.header !== undefined) add('header1.xml', CT.header, REL.header, 'w:hdr', o.header);
  if (o.footer !== undefined) add('footer1.xml', CT.footer, REL.footer, 'w:ftr', o.footer);
  files['[Content_Types].xml'] = strToU8(
    `<?xml version="1.0"?><Types xmlns="${NS.ct}">` +
      `<Default Extension="rels" ContentType="${CT.relationships}"/>` +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      overrides.map(([n, c]) => `<Override PartName="/${n}" ContentType="${c}"/>`).join('') +
      '</Types>',
  );
  files['_rels/.rels'] = strToU8(
    `<?xml version="1.0"?><Relationships xmlns="${NS.rel}">` +
      `<Relationship Id="rId1" Type="${REL.officeDocument}" Target="word/document.xml"/>` +
      '</Relationships>',
  );
  files['word/_rels/document.xml.rels'] = strToU8(
    `<?xml version="1.0"?><Relationships xmlns="${NS.rel}">` +
      rels
        .map(
          ([type, target], i) =>
            `<Relationship Id="rId${i + 1}" Type="${type}" Target="${target}"/>`,
        )
        .join('') +
      '</Relationships>',
  );
  files['word/document.xml'] = strToU8(
    `<?xml version="1.0"?><w:document ${DECLS}><w:body>${o.body}</w:body></w:document>`,
  );
  return zipSync(files);
}

/** A package whose officeDocument relationship and main part are given verbatim. */
function rawPackage(rels: string, parts: Record<string, string>): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="${CT.relationships}"/>` +
        '<Default Extension="xml" ContentType="application/xml"/></Types>',
    ),
    '_rels/.rels': strToU8(`<Relationships xmlns="${NS.rel}">${rels}</Relationships>`),
  };
  for (const [name, xml] of Object.entries(parts)) files[name] = strToU8(xml);
  return zipSync(files);
}

const officeRel = (target: string): string =>
  `<Relationship Id="rId1" Type="${REL.officeDocument}" Target="${target}"/>`;

// ---- the flattener --------------------------------------------------------------------------

describe('visible text (text.ts)', () => {
  it('maps run content to characters', () => {
    const xml =
      '<w:p><w:r><w:t>a</w:t><w:tab/><w:t xml:space="preserve"> b </w:t><w:br/><w:t>c</w:t><w:cr/>' +
      '<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/><w:noBreakHyphen/>' +
      '<w:softHyphen/><w:sym w:font="Wingdings" w:char="F04A"/><w:sym w:font="Symbol" w:char="03B1"/>' +
      '<w:sym w:font="Symbol" w:char="zz"/></w:r></w:p>';
    expect(flat(xml)).toBe('a\t b \nc\n\t‑­α');
  });

  it('skips deleted, moved-away, instruction, drawing and property content', () => {
    const xml =
      '<w:p><w:pPr><w:rPr><w:t>pPr</w:t></w:rPr></w:pPr>' +
      t('keep') +
      `<w:del w:id="1" w:author="x">${run('<w:delText>gone</w:delText>')}</w:del>` +
      `<w:moveFrom w:id="2" w:author="x">${t('moved')}</w:moveFrom>` +
      `<w:moveTo w:id="3" w:author="x">${t('+to')}</w:moveTo>` +
      `<w:ins w:id="4" w:author="x">${t('+ins')}</w:ins>` +
      run('<w:rPr><w:t>rPr</w:t></w:rPr><w:delText>dt</w:delText><w:instrText>PAGE</w:instrText>') +
      TEXTBOX +
      run(
        `<w:pict><v:shape><v:textbox><w:txbxContent><w:p>${t('vml')}</w:p></w:txbxContent></v:textbox></v:shape></w:pict>`,
      ) +
      run('<w:object><w:t>obj</w:t></w:object>') +
      '</w:p>';
    expect(flat(xml)).toBe('keep+to+ins');
  });

  it('traverses hyperlinks, smart tags, content controls, fields, custom XML, dir and bdo', () => {
    const xml =
      `<w:p><w:hyperlink r:id="rId1">${t('link')}</w:hyperlink>` +
      `<w:smartTag w:uri="u" w:element="e">${t(',tag')}</w:smartTag>` +
      `<w:sdt><w:sdtPr><w:alias w:val="A"/><w:placeholder><w:docPart w:val="x"/></w:placeholder></w:sdtPr>` +
      `<w:sdtContent>${t(',sdt')}</w:sdtContent></w:sdt>` +
      `<w:fldSimple w:instr=" DATE ">${t(',date')}</w:fldSimple>` +
      `<w:customXml w:element="x">${t(',cx')}</w:customXml>` +
      `<w:dir w:val="rtl">${t(',dir')}</w:dir><w:bdo w:val="ltr">${t(',bdo')}</w:bdo></w:p>`;
    expect(flat(xml)).toBe('link,tag,sdt,date,cx,dir,bdo');
  });

  it('reads the Fallback of mc:AlternateContent, else its first Choice', () => {
    const xml =
      '<w:p><mc:AlternateContent>' +
      `<mc:Choice Requires="w14">${t('choice')}</mc:Choice><mc:Fallback>${t('fallback')}</mc:Fallback>` +
      '</mc:AlternateContent><mc:AlternateContent>' +
      `<mc:Choice Requires="w14">${t('|first')}</mc:Choice><mc:Choice Requires="w15">${t('second')}</mc:Choice>` +
      '</mc:AlternateContent></w:p>';
    expect(flat(xml)).toBe('fallback|first');
  });

  it('shows the result of a complex field, not its instructions', () => {
    const xml =
      `<w:p>${t('Page ')}${run('<w:fldChar w:fldCharType="begin"/>')}` +
      run('<w:instrText xml:space="preserve"> PAGE </w:instrText>') +
      `${run('<w:fldChar w:fldCharType="separate"/>')}${t('3')}` +
      `${run('<w:fldChar w:fldCharType="end"/>')}</w:p>`;
    expect(flat(xml)).toBe('Page 3');
  });

  it('includes math text', () => {
    expect(flat('<w:p><m:oMath><m:r><m:rPr/><w:rPr/><m:t>x=1</m:t></m:r></m:oMath></w:p>')).toBe(
      'x=1',
    );
  });

  it('matches Strict OOXML namespaces', () => {
    const p = parseXml(
      '<w:p xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:r><w:t>strict</w:t><w:tab/></w:r></w:p>',
    ).root;
    expect(paragraphText(p)).toBe('strict\t');
  });

  it('walkVisible reports each contributing node with its text', () => {
    const p = frag(
      `<w:p>${run('<w:t>a</w:t><w:t></w:t><w:tab/>')}<w:ins>${run('<w:t>b</w:t>')}</w:ins></w:p>`,
    );
    const seen: [XElement, string][] = [];
    walkVisible(p, (node, text) => seen.push([node, text]));
    expect(seen.map(([n, s]) => [n.local, s])).toEqual([
      ['t', 'a'],
      ['tab', '\t'],
      ['t', 'b'],
    ]);
    expect(seen[0]![0]).toBe(p.descendants(NS.w, 't')[0]);
    // the root itself is subject to the rules
    expect(paragraphText(p.descendants(NS.w, 'tab')[0]!)).toBe('\t');
  });

  it('handles arbitrarily deep nesting without recursion', () => {
    let node = new XElement('w:t', NS.w, [], [new XText('deep')]);
    node = new XElement('w:r', NS.w, [], [node]);
    const p = new XElement('w:p', NS.w, [], [node]);
    let inner: XElement = p;
    for (let i = 0; i < 200_000; i++) inner = new XElement('w:customXml', NS.w, [], [inner]);
    const body = new XElement('w:body', NS.w, [], [inner]);
    expect(paragraphText(body)).toBe('deep');
    expect(blockParagraphs(body)).toEqual([p]);
  });
});

// ---- stories ------------------------------------------------------------------------------

const BODY =
  `<w:p>${t('one')}</w:p>` +
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p>${t('cell')}</w:p>` +
  `<w:tbl><w:tblGrid><w:gridCol w:w="50"/></w:tblGrid><w:tr><w:tc><w:p>${t('nested')}</w:p></w:tc></w:tr></w:tbl>` +
  '<w:p/></w:tc></w:tr></w:tbl>' +
  `<w:sdt><w:sdtPr/><w:sdtContent><w:p>${t('in sdt')}</w:p>` +
  `<w:tbl><w:tr><w:tc><w:p>${t('sdt table')}</w:p></w:tc></w:tr></w:tbl></w:sdtContent></w:sdt>` +
  `<w:customXml w:element="x"><w:p>${t('in cx')}</w:p></w:customXml>` +
  `<w:p>${t('with box')}${TEXTBOX}</w:p>` +
  '<w:bookmarkStart w:id="0" w:name="b"/><w:bookmarkEnd w:id="0"/>' +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

describe('Story', () => {
  it('paragraphs() is deep, in document order, without text boxes', () => {
    const doc = readDocx(docxFrom({ body: BODY }));
    const paras = doc.body.paragraphs();
    expect(paras.map((p) => p.text)).toEqual([
      'one',
      'cell',
      'nested',
      '',
      'in sdt',
      'sdt table',
      'in cx',
      'with box',
    ]);
    expect(paras.every((p) => p instanceof Paragraph && p.story === doc.body)).toBe(true);
    expect(doc.body.text()).toBe('one\ncell\nnested\n\nin sdt\nsdt table\nin cx\nwith box');
    expect(doc.text()).toBe(doc.body.text());
  });

  it('blocks() lists direct block-level children (no body sectPr)', () => {
    const doc = readDocx(docxFrom({ body: BODY }));
    const blocks = doc.body.blocks();
    const kinds = blocks.map((b) =>
      b instanceof Paragraph ? 'p' : b instanceof Table ? 'tbl' : (b as XElement).local,
    );
    expect(kinds).toEqual(['p', 'tbl', 'sdt', 'customXml', 'p', 'bookmarkStart', 'bookmarkEnd']);
  });

  it('tables() lists top-level tables, including ones wrapped in content controls', () => {
    const doc = readDocx(docxFrom({ body: BODY }));
    const tables = doc.body.tables();
    expect(tables).toHaveLength(2);
    expect(tables.every((tb) => tb.story === doc.body)).toBe(true);
    expect(tables[1]!.element.parent).toMatchObject({ local: 'sdtContent' });
  });

  it('exposes kind, part and element', () => {
    const doc = readDocx(docxFrom({ body: BODY }));
    expect(doc.body.kind).toBe('body');
    expect(doc.body.part).toBe(doc.mainPart);
    expect(doc.body.element.is(NS.w, 'body')).toBe(true);
    expect(doc.body.id).toBeUndefined();
    expect(doc.body.doc).toBe(doc);
  });
});

const FOOTNOTES =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  `<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>${t(' Note one')}</w:p></w:footnote>` +
  `<w:footnote w:id="2"><w:p>${t('Note two')}</w:p><w:p>${t('second para')}</w:p></w:footnote>`;
const ENDNOTES =
  '<w:endnote w:type="separator" w:id="-1"><w:p/></w:endnote>' +
  '<w:endnote w:type="continuationNotice" w:id="0"><w:p/></w:endnote>' +
  `<w:endnote w:id="1"><w:p>${t('End')}</w:p></w:endnote>`;
const COMMENTS =
  `<w:comment w:id="0" w:author="A" w:initials="A"><w:p><w:r><w:annotationRef/></w:r>${t('First comment')}</w:p></w:comment>` +
  `<w:comment w:id="5" w:author="B"><w:p>${t('Second')}</w:p></w:comment>`;

function notesDoc(): WordDocument {
  return readDocx(
    docxFrom({
      body: `<w:p>${t('Body')}</w:p><w:sectPr/>`,
      footnotes: FOOTNOTES,
      endnotes: ENDNOTES,
      comments: COMMENTS,
      header: `<w:p>${t('Head')}</w:p>`,
      footer: `<w:p>${t('Foot')}</w:p>`,
    }),
  );
}

describe('notes, comments and stories()', () => {
  it('footnotes() has one story per note, separators excluded', () => {
    const notes = notesDoc().footnotes();
    expect(notes.map((n) => n.id)).toEqual(['1', '2']);
    expect(notes.map((n) => n.text())).toEqual([' Note one', 'Note two\nsecond para']);
    expect(notes.every((n) => n.kind === 'footnote' && n.part.name === 'word/footnotes.xml')).toBe(
      true,
    );
    expect(notes[0]!.element.is(NS.w, 'footnote')).toBe(true);
  });

  it('endnotes() skips separator and continuationNotice', () => {
    const notes = notesDoc().endnotes();
    expect(notes.map((n) => [n.id, n.kind, n.text()])).toEqual([['1', 'endnote', 'End']]);
  });

  it('comments() has one story per comment with its id', () => {
    const comments = notesDoc().comments();
    expect(comments.map((c) => [c.id, c.kind, c.text()])).toEqual([
      ['0', 'comment', 'First comment'],
      ['5', 'comment', 'Second'],
    ]);
  });

  it('stories() orders body, headers/footers, footnotes, endnotes, comments', () => {
    const doc = notesDoc();
    expect(doc.stories(['body', 'footnote', 'endnote', 'comment']).map((s) => s.kind)).toEqual([
      'body',
      'footnote',
      'footnote',
      'endnote',
      'comment',
      'comment',
    ]);
    expect(doc.stories(['comment', 'body']).map((s) => s.kind)).toEqual([
      'body',
      'comment',
      'comment',
    ]);
  });

  it('stories() includes headers and footers', () => {
    const doc = notesDoc();
    const all = doc.stories();
    expect(all.map((s) => s.kind)).toEqual([
      'body',
      'header',
      'footer',
      'footnote',
      'footnote',
      'endnote',
      'comment',
      'comment',
    ]);
    expect(doc.stories(['footer']).map((s) => s.text())).toEqual(['Foot']);
    expect(doc.headersFooters().map((hf) => hf.text())).toEqual(['Head', 'Foot']);
  });

  it('a document without notes or comments has none', () => {
    const doc = createDocx();
    expect(doc.footnotes()).toEqual([]);
    expect(doc.endnotes()).toEqual([]);
    expect(doc.comments()).toEqual([]);
    expect(doc.stories()).toEqual([doc.body]);
  });

  it('appendParagraph() in a note appends at its end', () => {
    const doc = notesDoc();
    const note = doc.footnotes()[1]!;
    note.appendParagraph('third');
    expect(note.text()).toBe('Note two\nsecond para\nthird');
    const back = readDocx(writeDocx(doc));
    expect(back.footnotes()[1]!.text()).toBe('Note two\nsecond para\nthird');
  });
});

// ---- paragraphs and runs ------------------------------------------------------------------

describe('Paragraph and Run', () => {
  it('appendParagraph() keeps the body sectPr last', () => {
    const doc = createDocx();
    const body = doc.body.element;
    const p = doc.body.appendParagraph('Hello', { style: 'Heading1' });
    expect(body.lastElement()?.local).toBe('sectPr');
    expect(p.element.nextElement()?.local).toBe('sectPr');
    expect(p.style).toBe('Heading1');
    expect(p.text).toBe('Hello');
    doc.body.appendParagraph();
    expect(locals(body)).toEqual(['p', 'p', 'p', 'sectPr']);
    expect(doc.body.paragraphs().map((x) => x.text)).toEqual(['', 'Hello', '']);
    expect(doc.package.dirty).toBe(true);
  });

  it('appendParagraph() appends at the end when there is no sectPr', () => {
    const doc = readDocx(docxFrom({ body: `<w:p>${t('a')}</w:p>` }));
    doc.body.appendParagraph('b');
    expect(locals(doc.body.element)).toEqual(['p', 'p']);
    expect(doc.text()).toBe('a\nb');
  });

  it('appendRun() splits TAB/LF and preserves edge spaces', () => {
    const doc = createDocx();
    const p = doc.body.appendParagraph();
    const r = p.appendRun(' a\tb\nc\r\nd ');
    expect(locals(r.element)).toEqual(['t', 'tab', 't', 'br', 't', 'br', 't']);
    const ts = r.element.elements(NS.w, 't');
    expect(ts.map((x) => x.text)).toEqual([' a', 'b', 'c', 'd ']);
    expect(ts.map((x) => x.attr('xml:space'))).toEqual([
      'preserve',
      undefined,
      undefined,
      'preserve',
    ]);
    expect(r.text).toBe(' a\tb\nc\nd ');
    expect(p.text).toBe(' a\tb\nc\nd ');
    r.text = 'x‑y­z';
    expect(locals(r.element)).toEqual(['t', 'noBreakHyphen', 't', 'softHyphen', 't']);
    expect(r.text).toBe('x‑y­z');
  });

  it('appendRun() escapes markup characters and round-trips them', () => {
    const doc = createDocx();
    const s = '<a & "b"> \'c\'';
    doc.body.appendParagraph().appendRun(s);
    const xml = serializeXml(doc.mainPart.xml());
    expect(xml).toContain('&lt;a &amp; "b"&gt; \'c\'');
    expect(readDocx(writeDocx(doc)).body.paragraphs()[1]!.text).toBe(s);
  });

  it('run properties are written in schema order', () => {
    const doc = createDocx();
    const p = doc.body.appendParagraph();
    const a = p.appendRun('a', { bold: true, italic: true, style: 'Strong' });
    expect(locals(a.element)).toEqual(['rPr', 't']);
    expect(locals(a.element.element(NS.w, 'rPr'))).toEqual(['rStyle', 'b', 'i']);
    const b = p.appendRun('b');
    expect(locals(b.element)).toEqual(['t']);
    b.italic = true;
    b.bold = true;
    b.style = 'Emphasis';
    expect(locals(b.element.element(NS.w, 'rPr'))).toEqual(['rStyle', 'b', 'i']);
    expect([b.bold, b.italic, b.style]).toEqual([true, true, 'Emphasis']);
    expect(locals(b.element)).toEqual(['rPr', 't']);
  });

  it('bold/italic are tri-state and removable', () => {
    const doc = createDocx();
    const r = doc.body.appendParagraph().appendRun('x');
    expect(r.bold).toBeUndefined();
    r.bold = false;
    expect(r.element.element(NS.w, 'rPr')?.element(NS.w, 'b')?.attrNS(NS.w, 'val')).toBe('0');
    expect(r.bold).toBe(false);
    r.bold = true;
    expect(r.element.element(NS.w, 'rPr')?.element(NS.w, 'b')?._attrs).toEqual([]);
    r.bold = undefined;
    r.style = 'S';
    r.style = undefined;
    expect(r.element.element(NS.w, 'rPr')).toBeUndefined();
    for (const [val, want] of [
      ['false', false],
      ['off', false],
      ['0', false],
      ['1', true],
      ['true', true],
      ['on', true],
    ] as const) {
      const p = frag(`<w:p><w:r><w:rPr><w:i w:val="${val}"/></w:rPr></w:r></w:p>`);
      expect(new Paragraph(p, doc.body).runs()[0]!.italic).toBe(want);
    }
  });

  it('Run.text replaces content but keeps rPr', () => {
    const doc = createDocx();
    const r = doc.body.appendParagraph().appendRun('old', { bold: true });
    r.text = 'new\ttext';
    expect(locals(r.element)).toEqual(['rPr', 't', 'tab', 't']);
    expect(r.bold).toBe(true);
    r.text = '';
    expect(locals(r.element)).toEqual(['rPr']);
  });

  it('style setter keeps pPr first and pStyle first in pPr; undefined removes it', () => {
    const doc = createDocx();
    const p = new Paragraph(
      frag(`<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr>${t('x')}</w:p>`),
      doc.body,
    );
    p.style = 'Title';
    expect(locals(p.element.element(NS.w, 'pPr'))).toEqual(['pStyle', 'spacing', 'jc']);
    p.style = 'Heading2';
    expect(p.element.element(NS.w, 'pPr')?.elements(NS.w, 'pStyle')).toHaveLength(1);
    expect(p.style).toBe('Heading2');
    p.style = undefined;
    expect(locals(p.element.element(NS.w, 'pPr'))).toEqual(['spacing', 'jc']);
    const q = doc.body.appendParagraph('y', { style: 'S' });
    expect(locals(q.element)).toEqual(['pPr', 'r']);
    q.style = undefined;
    expect(locals(q.element)).toEqual(['r']);
  });

  it('runs() includes runs in inline containers, not deleted or text-box runs', () => {
    const doc = createDocx();
    const p = new Paragraph(
      frag(
        `<w:p><w:pPr><w:rPr/></w:pPr>${t('a')}<w:hyperlink>${t('b')}</w:hyperlink>` +
          `<w:del>${run('<w:delText>x</w:delText>')}</w:del><w:moveFrom>${t('m')}</w:moveFrom>` +
          `<w:ins>${t('c')}</w:ins><w:sdt><w:sdtContent>${t('d')}</w:sdtContent></w:sdt>` +
          `<mc:AlternateContent><mc:Choice Requires="w14">${t('no')}</mc:Choice><mc:Fallback>${t('e')}</mc:Fallback></mc:AlternateContent>` +
          `${TEXTBOX}</w:p>`,
      ),
      doc.body,
    );
    const runs = p.runs();
    expect(runs.map((r) => r.text)).toEqual(['a', 'b', 'c', 'd', 'e', '']);
    expect(runs.every((r) => r.paragraph === p)).toBe(true);
  });

  it('uses the prefix the document binds to WordprocessingML', () => {
    const custom = rawPackage(officeRel('word/document.xml'), {
      'word/document.xml': `<x:document xmlns:x="${NS.w}"><x:body><x:sectPr/></x:body></x:document>`,
    });
    const doc = readDocx(custom);
    doc.body.appendParagraph('hi', { style: 'S' }).runs()[0]!.bold = true;
    expect(serializeXml(doc.mainPart.xml())).toContain(
      '<x:p><x:pPr><x:pStyle x:val="S"/></x:pPr><x:r><x:rPr><x:b/></x:rPr><x:t>hi</x:t></x:r></x:p><x:sectPr/>',
    );

    const dflt = rawPackage(officeRel('word/document.xml'), {
      'word/document.xml': `<document xmlns="${NS.w}"><body><sectPr/></body></document>`,
    });
    const doc2 = readDocx(dflt);
    doc2.body.appendParagraph('hi', { style: 'S' });
    const back = readDocx(writeDocx(doc2));
    expect(back.body.paragraphs().map((p) => [p.text, p.style])).toEqual([['hi', 'S']]);
    expect(back.body.element.lastElement()?.local).toBe('sectPr');
  });
});

describe('order.ts', () => {
  it('insertOrdered places children by schema rank, unknown ones before *Change', () => {
    const pPr = frag(
      '<w:pPr><w:pStyle w:val="A"/><w:jc w:val="left"/><w:rPr/><w:pPrChange/></w:pPr>',
    );
    insertOrdered(pPr, wElement(pPr, 'numPr'), PPR_ORDER);
    insertOrdered(pPr, wElement(pPr, 'ind'), PPR_ORDER);
    insertOrdered(pPr, wElement(pPr, 'sectPr'), PPR_ORDER);
    insertOrdered(pPr, wElement(pPr, 'mystery'), PPR_ORDER);
    expect(locals(pPr)).toEqual([
      'pStyle',
      'numPr',
      'ind',
      'jc',
      'rPr',
      'sectPr',
      'mystery',
      'pPrChange',
    ]);

    const rPr = frag('<w:rPr><w:b/><w:sz w:val="20"/><w14:textFill/><w:rPrChange/></w:rPr>');
    insertOrdered(rPr, wElement(rPr, 'color', { val: 'FF0000' }), RPR_ORDER);
    insertOrdered(rPr, wElement(rPr, 'rFonts'), RPR_ORDER);
    insertOrdered(rPr, wElement(rPr, 'lang'), RPR_ORDER);
    expect(locals(rPr)).toEqual(['rFonts', 'b', 'color', 'sz', 'lang', 'textFill', 'rPrChange']);
    expect(rPr.element(NS.w, 'color')?.attrNS(NS.w, 'val')).toBe('FF0000');
  });
});

// ---- createDocx and the entry points --------------------------------------------------------

describe('createDocx', () => {
  it('writes a minimal package that reads back', () => {
    const bytes = writeDocx(createDocx());
    const doc = readDocx(bytes);
    expect(doc.mainPart.name).toBe('word/document.xml');
    expect(doc.body.paragraphs()).toHaveLength(1);
    expect(doc.text()).toBe('');
    const names = doc.package.parts().map((p) => p.name);
    for (const n of [
      '[Content_Types].xml',
      '_rels/.rels',
      'word/document.xml',
      'word/_rels/document.xml.rels',
      'word/styles.xml',
      'word/settings.xml',
      'word/fontTable.xml',
      'docProps/core.xml',
      'docProps/app.xml',
    ]) {
      expect(names).toContain(n);
    }
    expect(doc.package.contentTypes.get('word/document.xml')).toBe(CT.document);
    expect(doc.package.contentTypes.get('word/styles.xml')).toBe(CT.styles);
    expect(doc.package.contentTypes.get('docProps/core.xml')).toBe(CT.coreProperties);
    expect(doc.package.contentTypes.get('docProps/app.xml')).toBe(CT.extendedProperties);
    expect(doc.package.rels.all().map((r) => [r.type, r.partName])).toEqual([
      [REL.officeDocument, 'word/document.xml'],
      [REL.coreProperties, 'docProps/core.xml'],
      [REL.extendedProperties, 'docProps/app.xml'],
    ]);
    expect(doc.mainPart.rels.all().map((r) => r.partName)).toEqual([
      'word/styles.xml',
      'word/settings.xml',
      'word/fontTable.xml',
    ]);
  });

  it('has the expected template content', () => {
    const doc = createDocx();
    const root = doc.mainPart.xml().root;
    expect(root.attrNS(NS.mc, 'Ignorable')).toBe('w14 w15 wp14');
    const sectPr = doc.body.element.lastElement()!;
    expect(sectPr.local).toBe('sectPr');
    const pgSz = sectPr.element(NS.w, 'pgSz')!;
    expect([pgSz.attrNS(NS.w, 'w'), pgSz.attrNS(NS.w, 'h')]).toEqual(['12240', '15840']);
    const pgMar = sectPr.element(NS.w, 'pgMar')!;
    expect(
      ['top', 'right', 'bottom', 'left', 'header', 'footer'].map((k) => pgMar.attrNS(NS.w, k)),
    ).toEqual(['1440', '1440', '1440', '1440', '720', '720']);
    const styles = doc.mainPart.relatedByType('styles')[0]!.xml().root;
    expect(styles.elements(NS.w, 'style').map((s) => s.attrNS(NS.w, 'styleId'))).toEqual([
      'Normal',
      'DefaultParagraphFont',
      'TableNormal',
      'NoList',
      'Heading1',
      'Heading2',
      'Heading3',
      'Title',
    ]);
    const settings = doc.mainPart.relatedByType('settings')[0]!.xml().root;
    const mode = settings
      .descendants(NS.w, 'compatSetting')
      .find((c) => c.attrNS(NS.w, 'name') === 'compatibilityMode');
    expect(mode?.attrNS(NS.w, 'val')).toBe('15');
    const app = doc.package.part('docProps/app.xml')!.xml().root;
    expect(app.element(NS.ep, 'Application')?.text).toBe('turboword');
    const core = doc.package.part('docProps/core.xml')!.xml().root;
    expect(core.descendants(NS.dcterms)).toEqual([]);
  });

  it('is deterministic and passes validation', () => {
    const a = writeDocx(createDocx());
    const b = writeDocx(createDocx(), { validate: true });
    expect(b).toEqual(a);
    expect(validatePackage(createDocx()).filter((i) => i.severity === 'error')).toEqual([]);
    expect(createDocx().sections()).toHaveLength(1);
  });

  it('a no-op save of a read document returns the input bytes', () => {
    const bytes = writeDocx(createDocx());
    const doc = readDocx(bytes);
    doc.text();
    doc.stories();
    expect(writeDocx(doc)).toEqual(bytes);
  });
});

describe('readDocx / writeDocx', () => {
  it('finds the main part through the officeDocument relationship', () => {
    const bytes = rawPackage(officeRel('/content/main.xml'), {
      'content/main.xml': `<w:document xmlns:w="${NS.w}"><w:body><w:p><w:r><w:t>moved</w:t></w:r></w:p></w:body></w:document>`,
    });
    const doc = readDocx(new Uint8Array(bytes).buffer);
    expect(doc.mainPart.name).toBe('content/main.xml');
    expect(doc.text()).toBe('moved');
  });

  it('reads Strict OOXML documents', () => {
    const strict = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
    const bytes = rawPackage(
      '<Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="word/document.xml"/>',
      {
        'word/document.xml': `<w:document xmlns:w="${strict}"><w:body><w:p><w:r><w:t>strict</w:t></w:r></w:p></w:body></w:document>`,
      },
    );
    expect(readDocx(bytes).text()).toBe('strict');
  });

  it('throws DocxFormatError for packages that are not WordprocessingML documents', () => {
    const cases = [
      rawPackage('', {}),
      rawPackage(officeRel('word/document.xml'), {}),
      rawPackage(officeRel('xl/workbook.xml'), {
        'xl/workbook.xml':
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
      }),
      rawPackage(officeRel('word/document.xml'), {
        'word/document.xml': `<w:document xmlns:w="${NS.w}"/>`,
      }),
    ];
    for (const bytes of cases) {
      expect(() => readDocx(bytes)).toThrow(DocxFormatError);
    }
  });

  it('passes limits through to the package', () => {
    const bytes = writeDocx(createDocx());
    expect(() => readDocx(bytes, { limits: { maxEntries: 2 } })).toThrow();
  });

  it('validate: true throws a VALIDATION error listing the problems', () => {
    const doc = createDocx();
    doc.mainPart.rels.add(REL.image, 'media/missing.png');
    let err: unknown;
    try {
      writeDocx(doc, { validate: true });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(TurbowordError);
    expect((err as TurbowordError).code).toBe('VALIDATION');
    expect((err as Error).message).toContain('REL_TARGET_MISSING');
    expect(() => writeDocx(doc)).not.toThrow();
  });

  it('fork() is independent of the original', () => {
    const bytes = writeDocx(createDocx());
    const doc = readDocx(bytes);
    const copy = doc.fork();
    expect(copy).toBeInstanceOf(WordDocument);
    copy.body.appendParagraph('only in the copy');
    expect(copy.text()).toBe('\nonly in the copy');
    expect(doc.text()).toBe('');
    expect(writeDocx(doc)).toEqual(bytes);
    expect(readDocx(writeDocx(copy)).text()).toBe('\nonly in the copy');
  });

  it('gc() removes unreachable parts', () => {
    const doc = createDocx();
    doc.package.addPart('word/orphan.xml', CT.xml, strToU8('<x/>'));
    expect(doc.gc()).toEqual(['word/orphan.xml']);
    expect(doc.package.part('word/orphan.xml')).toBeUndefined();
  });

  it('properties are created once', () => {
    const doc = createDocx();
    expect(doc.properties).toBeInstanceOf(DocumentProperties);
    expect(doc.properties).toBe(doc.properties);
  });
});

describe('loadDocx / docxBlob / docxStream', () => {
  const doc = createDocx();
  doc.body.appendParagraph('loaded');
  const bytes = writeDocx(doc);

  it('loads from a Blob', async () => {
    expect((await loadDocx(new Blob([new Uint8Array(bytes)]))).text()).toBe('\nloaded');
  });

  it('loads from a chunked ReadableStream', async () => {
    const stream = new ReadableStream<Uint8Array>({
      start(c) {
        for (let i = 0; i < bytes.length; i += 1000) c.enqueue(bytes.slice(i, i + 1000));
        c.close();
      },
    });
    expect((await loadDocx(stream)).text()).toBe('\nloaded');
  });

  it('loads from an async iterable', async () => {
    async function* chunks(): AsyncGenerator<Uint8Array> {
      for (let i = 0; i < bytes.length; i += 777) yield bytes.subarray(i, i + 777);
    }
    expect((await loadDocx(chunks())).text()).toBe('\nloaded');
  });

  it('loads from bytes and ArrayBuffer', async () => {
    expect((await loadDocx(bytes)).text()).toBe('\nloaded');
    expect((await loadDocx(bytes.slice().buffer)).text()).toBe('\nloaded');
  });

  it('rejects other sources and text chunks', async () => {
    await expect(loadDocx(42 as never)).rejects.toThrow(UsageError);
    async function* text(): AsyncGenerator<string> {
      yield 'not bytes';
    }
    await expect(loadDocx(text() as never)).rejects.toThrow(UsageError);
  });

  it('docxBlob() has the .docx MIME type and round-trips', async () => {
    const blob = docxBlob(doc);
    expect(blob.type).toBe(DOCX_MIME);
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(bytes);
    expect((await loadDocx(blob)).text()).toBe('\nloaded');
  });

  it('docxStream() round-trips', async () => {
    expect((await loadDocx(docxStream(doc))).text()).toBe('\nloaded');
  });
});

// ---- fixture corpus -------------------------------------------------------------------------

const FIXTURES = fileURLToPath(new URL('./fixtures/', import.meta.url));

function _findDocx(dir: string): string[] {
  const out: string[] = [];
  let names: string[];
  try {
    names = readdirSync(dir);
  } catch {
    return out;
  }
  for (const name of names) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) out.push(..._findDocx(p));
    else if (name.toLowerCase().endsWith('.docx')) out.push(p);
  }
  return out.sort();
}

// The curated corpus only (test/fixtures/expected.json).
const fixtures = curatedFixtures().map((f) => relative(FIXTURES, f.path).split(sep).join('/'));

const fixture = (name: string): Uint8Array | undefined => {
  const p = join(FIXTURES, name);
  return existsSync(p) ? new Uint8Array(readFileSync(p)) : undefined;
};

describe('fixture content', () => {
  const delins = fixture('apache-poi/delins.docx');
  it.skipIf(delins === undefined)('tracked changes: insertions shown, deletions hidden', () => {
    const text = readDocx(delins!).text();
    expect(text).toContain('March 2009: Apache Tika Release');
    expect(text).not.toContain('Also known as Maang Tika.');
    expect(text).not.toContain('Tika Waylan');
  });

  const notes = fixture('open-xml-sdk/TableFntEdnt.docx');
  it.skipIf(notes === undefined)('footnotes and endnotes', () => {
    const doc = readDocx(notes!);
    expect(doc.footnotes().map((n) => [n.id, n.text()])).toEqual([
      ['2', ' Footnote 1'],
      ['3', ' Footnote 2'],
      ['4', ' Footnote 3'],
    ]);
    expect(doc.endnotes().map((n) => n.text())).toEqual([' Endnote 1', ' Endnote 2', ' Endnote 3']);
  });

  const strict = fixture('open-xml-sdk/strict-CommentRangeStart.docx');
  it.skipIf(strict === undefined)('Strict OOXML comments', () => {
    const doc = readDocx(strict!);
    expect(doc.text()).toBe('This is some text.');
    expect(doc.comments().map((c) => [c.id, c.text()])).toEqual([['1', 'some test comment']]);
  });
});

describe.skipIf(fixtures.length === 0)('fixture smoke tests', () => {
  it.each(fixtures)('%s', (name) => {
    const bytes = new Uint8Array(readFileSync(join(FIXTURES, name)));
    const doc = readDocx(bytes);
    expect(typeof doc.text()).toBe('string');
    for (const story of doc.stories()) {
      story.text();
      story.blocks();
      story.tables();
      for (const p of story.paragraphs()) {
        expect(typeof p.text).toBe('string');
        void p.style;
        for (const r of p.runs()) {
          expect(typeof r.text).toBe('string');
          void r.bold;
          void r.italic;
          void r.style;
        }
      }
    }
    // reading never dirties: a no-op save returns the input bytes
    expect(doc.package.dirty).toBe(false);
    expect(writeDocx(doc)).toEqual(bytes);

    const before = doc.body.paragraphs().length;
    const endsWithSectPr = doc.body.element.lastElement()?.is(NS.w, 'sectPr') === true;
    doc.body.appendParagraph('turboword smoke', { style: 'Normal' }).appendRun(' & more');
    const back = readDocx(writeDocx(doc));
    const paras = back.body.paragraphs();
    expect(paras).toHaveLength(before + 1);
    expect(paras[paras.length - 1]!.text).toBe('turboword smoke & more');
    if (endsWithSectPr) expect(back.body.element.lastElement()?.local).toBe('sectPr');
  });
});
