// Shared helpers of the text-engine tests (test/text*.test.ts, test/revisions.test.ts):
// synthetic documents built from WordprocessingML snippets, structural invariants, a reference
// "original view" text (deleted content shown, inserted content hidden) and output writing.

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strToU8, zipSync } from 'fflate';
import { expect } from 'vitest';
import {
  CT,
  createDocx,
  NS,
  parseXml,
  REL,
  readDocx,
  validatePackage,
  type WordDocument,
  writeDocx,
  type XElement,
} from '../src/index.js';
import { REVISION_ELEMENTS } from '../src/wml/ids.js';
import { chooseAlternate } from '../src/wml/text.js';
import { serializeNode } from '../src/xml/serialize.js';

export const DECLS: string = [
  `xmlns:w="${NS.w}"`,
  `xmlns:r="${NS.r}"`,
  `xmlns:mc="${NS.mc}"`,
  `xmlns:m="${NS.m}"`,
  `xmlns:wp="${NS.wp}"`,
  `xmlns:a="${NS.a}"`,
  `xmlns:v="${NS.v}"`,
  `xmlns:o="${NS.o}"`,
  `xmlns:wps="${NS.wps}"`,
  `xmlns:w14="${NS.w14}"`,
].join(' ');

/** Absolute path of out/text/ (created on demand). */
export const OUT_DIR: string = fileURLToPath(new URL('../out/text/', import.meta.url));

/** `<w:r>` with optional rPr inner XML and `<w:t>` text. */
export function r(text: string, rPr = ''): string {
  const pr = rPr === '' ? '' : `<w:rPr>${rPr}</w:rPr>`;
  return `<w:r>${pr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
}

/** A body paragraph from inner XML. */
export function p(inner: string, pPr = ''): string {
  return `<w:p>${pPr === '' ? '' : `<w:pPr>${pPr}</w:pPr>`}${inner}</w:p>`;
}

/** createDocx() with `bodyXml` as the body content (before the final sectPr). */
export function docWith(bodyXml: string): WordDocument {
  const doc = createDocx();
  const body = doc.body.element;
  for (const k of body.elements(NS.w, 'p')) k.remove();
  const sectPr = body.element(NS.w, 'sectPr')!;
  const frag = parseXml(`<w:body ${DECLS}>${bodyXml}</w:body>`).root;
  for (const k of frag.children.slice()) body.insertBefore(k, sectPr);
  return doc;
}

export interface Parts {
  body: string;
  header?: string;
  footer?: string;
  footnotes?: string;
  endnotes?: string;
  comments?: string;
  styles?: string;
  numbering?: string;
}

/** A .docx built from scratch with the given stories (header/footer referenced from sectPr). */
export function docxFrom(o: Parts): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const overrides: [string, string][] = [['word/document.xml', CT.document]];
  const rels: [string, string, string][] = [];
  const add = (name: string, ct: string, rel: string, root: string, inner: string): string => {
    files[`word/${name}`] = strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${root} ${DECLS}>${inner}</${root}>`,
    );
    overrides.push([`word/${name}`, ct]);
    const id = `rId${rels.length + 1}`;
    rels.push([id, rel, name]);
    return id;
  };
  let refs = '';
  if (o.styles !== undefined) add('styles.xml', CT.styles, REL.styles, 'w:styles', o.styles);
  if (o.numbering !== undefined) {
    add('numbering.xml', CT.numbering, REL.numbering, 'w:numbering', o.numbering);
  }
  if (o.footnotes !== undefined) {
    add('footnotes.xml', CT.footnotes, REL.footnotes, 'w:footnotes', o.footnotes);
  }
  if (o.endnotes !== undefined) {
    add('endnotes.xml', CT.endnotes, REL.endnotes, 'w:endnotes', o.endnotes);
  }
  if (o.comments !== undefined) {
    add('comments.xml', CT.comments, REL.comments, 'w:comments', o.comments);
  }
  if (o.header !== undefined) {
    const id = add('header1.xml', CT.header, REL.header, 'w:hdr', o.header);
    refs += `<w:headerReference w:type="default" r:id="${id}"/>`;
  }
  if (o.footer !== undefined) {
    const id = add('footer1.xml', CT.footer, REL.footer, 'w:ftr', o.footer);
    refs += `<w:footerReference w:type="default" r:id="${id}"/>`;
  }
  files['[Content_Types].xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${NS.ct}">` +
      `<Default Extension="rels" ContentType="${CT.relationships}"/>` +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      overrides.map(([n, c]) => `<Override PartName="/${n}" ContentType="${c}"/>`).join('') +
      '</Types>',
  );
  files['_rels/.rels'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}">` +
      `<Relationship Id="rId1" Type="${REL.officeDocument}" Target="word/document.xml"/>` +
      '</Relationships>',
  );
  files['word/_rels/document.xml.rels'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}">` +
      rels
        .map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`)
        .join('') +
      '</Relationships>',
  );
  const sectPr =
    `<w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/>` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  files['word/document.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${DECLS}><w:body>${o.body}${sectPr}</w:body></w:document>`,
  );
  return zipSync(files);
}

/** First element of an XML fragment (prefixes of DECLS in scope). */
export function frag(xml: string): XElement {
  return parseXml(`<w:wrap ${DECLS}>${xml}</w:wrap>`).root.firstElement()!;
}

/** Serialized element without namespace fix-up (for structural comparisons). */
export function xml(el: XElement): string {
  return serializeNode(el);
}

/** Root elements of every XML part of a document. */
export function partRoots(doc: WordDocument): { name: string; root: XElement }[] {
  const out: { name: string; root: XElement }[] = [];
  for (const part of doc.package.parts()) {
    if (!part.isXml || !/^word\/.*\.xml$/.test(part.name)) continue;
    const root = part.xml().children.find((k): k is XElement => k.nodeType === 1);
    if (root) out.push({ name: part.name, root });
  }
  return out;
}

/** Every revision element (ids.ts REVISION_ELEMENTS + delText/delInstrText) left in `doc`. */
export function revisionsLeft(doc: WordDocument): string[] {
  const out: string[] = [];
  for (const { name, root } of partRoots(doc)) {
    for (const el of root.descendants(NS.w)) {
      if (
        REVISION_ELEMENTS.has(el.local) ||
        el.local === 'delText' ||
        el.local === 'delInstrText'
      ) {
        out.push(`${name}: ${el.name}`);
      }
    }
  }
  return out;
}

/**
 * Structural invariants of a story tree after editing: bookmarks, comment ranges and
 * permissions balanced (one start and one end per id, start first), no nested runs, no markers
 * inside runs, no empty `w:t`.
 */
export function expectSane(root: XElement): void {
  const order = new Map<XElement, number>();
  const all = root.descendants();
  for (let i = 0; i < all.length; i++) order.set(all[i]!, i);
  for (const [s, e] of [
    ['bookmarkStart', 'bookmarkEnd'],
    ['commentRangeStart', 'commentRangeEnd'],
    ['permStart', 'permEnd'],
  ] as const) {
    const starts = new Map<string, XElement[]>();
    const ends = new Map<string, XElement[]>();
    for (const el of root.descendants(NS.w, s)) {
      const id = el.attrNS(NS.w, 'id')!;
      starts.set(id, [...(starts.get(id) ?? []), el]);
    }
    for (const el of root.descendants(NS.w, e)) {
      const id = el.attrNS(NS.w, 'id')!;
      ends.set(id, [...(ends.get(id) ?? []), el]);
    }
    for (const [id, list] of starts) {
      expect(list.length, `${s} ${id} count`).toBe(1);
      const end = ends.get(id);
      expect(end?.length, `${e} ${id} count`).toBe(1);
      expect(order.get(list[0]!)!, `${s} ${id} before its end`).toBeLessThan(order.get(end![0]!)!);
    }
    for (const id of ends.keys()) expect(starts.has(id), `${e} ${id} has a start`).toBe(true);
  }
  for (const run of root.descendants(NS.w, 'r')) {
    let a = run.parent;
    while (a !== null && a.nodeType === 1) {
      // Runs nest only inside ruby and text boxes (their own paragraphs).
      if (a.is(NS.w, 'p')) break;
      expect(a.is(NS.w, 'r'), 'run nested in a run').toBe(false);
      a = a.parent;
    }
    for (const k of run.elements(NS.w)) {
      expect([
        'bookmarkStart',
        'bookmarkEnd',
        'commentRangeStart',
        'commentRangeEnd',
        'proofErr',
      ]).not.toContain(k.local);
    }
  }
  for (const t of root.descendants(NS.w, 't')) expect(t.text, 'empty w:t').not.toBe('');
}

/**
 * Text of a paragraph as it read before any tracked change: deleted and moved-away content
 * shown (delText as text), inserted and moved-to content hidden. Otherwise walkVisible's rules.
 */
export function originalText(p: XElement): string {
  let out = '';
  const stack: XElement[] = [p];
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el.uri === NS.w) {
      switch (el.local) {
        case 't':
        case 'delText':
          out += el.text;
          continue;
        case 'tab':
        case 'ptab':
          out += '\t';
          continue;
        case 'br':
        case 'cr':
          out += '\n';
          continue;
        case 'noBreakHyphen':
          out += '‑';
          continue;
        case 'softHyphen':
          out += '­';
          continue;
        case 'sym': {
          const hex = el.attrNS(NS.w, 'char');
          if (hex && /^[0-9A-Fa-f]{1,6}$/.test(hex)) out += String.fromCodePoint(parseInt(hex, 16));
          continue;
        }
        case 'ins':
        case 'moveTo':
        case 'instrText':
        case 'delInstrText':
        case 'drawing':
        case 'pict':
        case 'object':
          continue;
      }
    } else if (el.uri === NS.mc && el.local === 'AlternateContent') {
      const b = chooseAlternate(el);
      if (b) stack.push(b);
      continue;
    } else if (el.uri === NS.m && el.local === 't') {
      out += el.text;
      continue;
    }
    if (el.local.endsWith('Pr')) continue;
    const kids = el.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
  }
  return out;
}

/** Saves, validates (no errors), reopens and returns the reopened document. */
export function roundTrip(doc: WordDocument, outName?: string): WordDocument {
  const errors = validatePackage(doc).filter((i) => i.severity === 'error');
  expect(errors).toEqual([]);
  const bytes = writeDocx(doc);
  if (outName !== undefined) {
    mkdirSync(OUT_DIR, { recursive: true });
    writeFileSync(join(OUT_DIR, outName), bytes);
  }
  const back = readDocx(bytes);
  expect(validatePackage(back).filter((i) => i.severity === 'error')).toEqual([]);
  return back;
}
