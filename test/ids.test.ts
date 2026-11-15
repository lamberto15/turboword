// DocumentIds: the document-wide id registry (src/wml/ids.ts).
import { describe, expect, it } from 'vitest';
import { createDocx } from '../src/wml/create.js';
import { readDocx, type WordDocument } from '../src/wml/document.js';
import { DocumentIds } from '../src/wml/ids.js';
import { NS } from '../src/xml/ns.js';
import { parseXml } from '../src/xml/parse.js';
import { fixture } from './helpers/fixtures.js';

const W = NS.w;

function open(name: string): WordDocument {
  return readDocx(fixture(name).bytes);
}

/** Ids of `kind`-elements in a part's DOM. */
function attrValues(
  doc: WordDocument,
  partName: string,
  uri: string,
  local: string,
  attr: string,
): string[] {
  const root = doc.package.part(partName)!.xml().root;
  return root.descendants(uri, local).map((e) => e.attr(attr) ?? '');
}

/** Appends a paragraph holding raw XML to the body (before the final sectPr). */
function appendXml(doc: WordDocument, xml: string): void {
  const frag = parseXml(
    `<w:root xmlns:w="${W}" xmlns:wp="${NS.wp}" xmlns:w14="${NS.w14}" xmlns:o="${NS.o}" xmlns:v="${NS.v}">${xml}</w:root>`,
  ).root;
  const body = doc.body.element;
  const last = body.lastElement();
  for (const k of frag.elements()) {
    if (last?.is(W, 'sectPr')) body.insertBefore(k, last);
    else body.append(k);
  }
}

describe('DocumentIds', () => {
  it('is cached per document package; a fork gets its own registry', () => {
    const doc = createDocx();
    expect(DocumentIds.of(doc)).toBe(DocumentIds.of(doc));
    expect(DocumentIds.of(doc.fork())).not.toBe(DocumentIds.of(doc));
  });

  it('scans the main document, headers, notes and comments', () => {
    const hdr = open('docxcompose/header_with_image.docx');
    const headerDocPr = attrValues(hdr, 'word/header1.xml', NS.wp, 'docPr', 'id');
    expect(headerDocPr.length).toBeGreaterThan(0);
    for (const id of headerDocPr) expect(DocumentIds.of(hdr).has('docPr', id)).toBe(true);

    const notes = open('open-xml-sdk/TableFntEdnt.docx');
    const ids = DocumentIds.of(notes);
    for (const id of attrValues(notes, 'word/footnotes.xml', W, 'footnote', 'w:id')) {
      expect(ids.has('footnote', id)).toBe(true);
    }
    for (const id of attrValues(notes, 'word/endnotes.xml', W, 'endnote', 'w:id')) {
      expect(ids.has('endnote', id)).toBe(true);
    }

    const comments = open('python-docx/comments-rich-para.docx');
    for (const id of attrValues(comments, 'word/comments.xml', W, 'comment', 'w:id')) {
      expect(DocumentIds.of(comments).has('comment', id)).toBe(true);
    }
    const paraIds = attrValues(comments, 'word/comments.xml', W, 'p', 'w14:paraId').filter(Boolean);
    expect(paraIds.length).toBeGreaterThan(0);
    for (const id of paraIds)
      expect(DocumentIds.of(comments).has('paraId', id.toLowerCase())).toBe(true);

    const revs = open('apache-poi/delins.docx');
    const insIds = attrValues(revs, 'word/document.xml', W, 'ins', 'w:id');
    expect(insIds.length).toBeGreaterThan(0);
    for (const id of insIds) expect(DocumentIds.of(revs).has('revision', id)).toBe(true);

    const sdts = open('docxcompose/content_controls.docx');
    const sdtIds = sdts.mainPart
      .xml()
      .root.descendants(W, 'sdtPr')
      .map((p) => p.element(W, 'id')?.attrNS(W, 'val'))
      .filter((v): v is string => v !== undefined);
    expect(sdtIds.length).toBeGreaterThan(0);
    for (const id of sdtIds) expect(DocumentIds.of(sdts).has('sdt', id)).toBe(true);
  });

  it('hands out numeric ids above every id in use, and reserves them', () => {
    const doc = open('docxcompose/header_with_image.docx');
    const ids = DocumentIds.of(doc);
    const used = attrValues(doc, 'word/header1.xml', NS.wp, 'docPr', 'id').map(Number);
    const a = ids.next('docPr');
    expect(a).toBe(Math.max(...used, 0) + 1);
    expect(ids.has('docPr', a)).toBe(true);
    expect(ids.next('docPr')).toBe(a + 1);
    expect(ids.next('footnote')).toBeGreaterThanOrEqual(1);
    expect(ids.next('bookmark')).toBeGreaterThanOrEqual(0);
  });

  it('hands out deterministic 8-hex ids below 0x80000000 that are not in use', () => {
    const seq = (): string[] => {
      const ids = DocumentIds.of(open('python-docx/comments-rich-para.docx'));
      return Array.from({ length: 50 }, () => ids.nextHex('paraId'));
    };
    const a = seq();
    expect(seq()).toEqual(a);
    expect(new Set(a).size).toBe(a.length);
    for (const h of a) {
      expect(h).toMatch(/^[0-9A-F]{8}$/);
      expect(Number.parseInt(h, 16)).toBeLessThan(0x80000000);
      expect(Number.parseInt(h, 16)).toBeGreaterThan(0);
    }
    const doc = open('python-docx/comments-rich-para.docx');
    const used = new Set(
      attrValues(doc, 'word/comments.xml', W, 'p', 'w14:paraId').map((v) => v.toUpperCase()),
    );
    for (const h of a) expect(used.has(h)).toBe(false);
  });

  it('re-scans when a part changes, and after invalidate()', () => {
    const doc = createDocx();
    const ids = DocumentIds.of(doc);
    expect(ids.next('docPr')).toBe(1);
    appendXml(
      doc,
      '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="41" name="x"/></wp:inline></w:drawing></w:r></w:p>',
    );
    expect(ids.has('docPr', 41)).toBe(true);
    expect(ids.next('docPr')).toBe(42);
    ids.invalidate();
    expect(ids.has('docPr', 41)).toBe(true);
    // Reserved ids survive invalidation: 42 was handed out.
    expect(ids.has('docPr', 42)).toBe(true);
    expect(ids.next('docPr')).toBe(43);
  });

  it('sees story parts added later (main part relationships changed)', () => {
    const doc = createDocx();
    const ids = DocumentIds.of(doc);
    expect(ids.has('bookmark', 7)).toBe(false);
    const enc = new TextEncoder();
    doc.package.addPart(
      'word/header1.xml',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      enc.encode(
        `<w:hdr xmlns:w="${W}"><w:p><w:bookmarkStart w:id="7" w:name="b"/><w:bookmarkEnd w:id="7"/></w:p></w:hdr>`,
      ),
    );
    doc.mainPart.rels.add(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
      'header1.xml',
    );
    expect(ids.has('bookmark', 7)).toBe(true);
    expect(ids.next('bookmark')).toBe(8);
  });

  it('hold() suspends re-scans until released', () => {
    const doc = createDocx();
    const ids = DocumentIds.of(doc);
    const release = ids.hold();
    appendXml(doc, '<w:p w14:paraId="0000ABCD"/>');
    expect(ids.has('paraId', '0000abcd')).toBe(false);
    release();
    release(); // idempotent
    expect(ids.has('paraId', '0000ABCD')).toBe(true);
    expect(ids.has('paraId', 'abcd')).toBe(true);
  });

  it('reserve() and has() cover every kind', () => {
    const ids = DocumentIds.of(createDocx());
    ids.reserve('sdt', -5);
    expect(ids.has('sdt', '-5')).toBe(true);
    ids.reserve('durableId', '1a2b3c4d');
    expect(ids.has('durableId', '1A2B3C4D')).toBe(true);
    expect(ids.has('textId', 'zz')).toBe(false);
    expect(ids.next('revision')).toBe(0);
    expect(ids.next('revision')).toBe(1);
  });

  it('hands out VML shape ids above the numbers in use', () => {
    const doc = open('apache-poi/shapes-with-text.docx');
    const ids = DocumentIds.of(doc);
    const xml = new TextDecoder().decode(doc.mainPart.bytes());
    const nums = [...xml.matchAll(/_x0000_[a-z](\d+)/g)].map((m) => Number(m[1]));
    expect(nums.length).toBeGreaterThan(0);
    const id = ids.nextVmlShapeId();
    expect(id).toMatch(/^_x0000_s\d+$/);
    expect(Number(id.slice(8))).toBeGreaterThan(Math.max(...nums));
    expect(ids.hasVmlShapeId(id)).toBe(true);
    expect(DocumentIds.of(createDocx()).nextVmlShapeId()).toBe('_x0000_s1025');
  });
});
