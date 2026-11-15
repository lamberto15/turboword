// Builders for synthetic documents with fields (shared by the test/fields*.test.ts files).
//
//   docWith(para('Total: ', cf('= 2 + 3', '0')))
//
// `cf` builds a complex field (begin / instruction / separate / result / end runs); its
// instruction parts may be other fields (nested), whose results are then written as
// `w:instrText`, as Word does. `sf` builds a `w:fldSimple`. Strings passed to `para` become
// plain text runs.

import { listFields } from '../src/fields/index.js';
import { createDocx, type WordDocument } from '../src/index.js';
import { NS } from '../src/xml/ns.js';
import { parseXml } from '../src/xml/parse.js';

/** A markup fragment (as opposed to plain text). */
export interface Xml {
  xml: string;
}

export const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** A text run (`rPr` is inner rPr markup). */
export function run(text: string, rPr = ''): Xml {
  const t = text === '' ? '' : `<w:t xml:space="preserve">${esc(text)}</w:t>`;
  return { xml: `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t}</w:r>` };
}

export interface FieldOptions {
  /** Nested in another field's instruction (result written as instrText). */
  nested?: boolean;
  /** rPr markup of the result run. */
  rPr?: string;
  /** rPr markup of the instruction runs. */
  codeRPr?: string;
  lock?: boolean;
  dirty?: boolean;
  /** Omit the separate marker (field without a result section). */
  noSeparate?: boolean;
}

/** A complex field. */
export function cf(instr: string | Xml | (string | Xml)[], result = '', o: FieldOptions = {}): Xml {
  const parts = Array.isArray(instr) ? instr : [instr];
  const code = o.codeRPr ? `<w:rPr>${o.codeRPr}</w:rPr>` : '';
  const attrs = `${o.lock ? ' w:fldLock="true"' : ''}${o.dirty ? ' w:dirty="true"' : ''}`;
  let xml = `<w:r><w:fldChar w:fldCharType="begin"${attrs}/></w:r>`;
  for (const p of parts) {
    xml +=
      typeof p === 'string'
        ? `<w:r>${code}<w:instrText xml:space="preserve">${esc(p)}</w:instrText></w:r>`
        : p.xml;
  }
  if (!o.noSeparate) {
    xml += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    if (result !== '') {
      const tag = o.nested ? 'instrText' : 't';
      const rPr = o.rPr ? `<w:rPr>${o.rPr}</w:rPr>` : '';
      xml += `<w:r>${rPr}<w:${tag} xml:space="preserve">${esc(result)}</w:${tag}></w:r>`;
    }
  }
  xml += '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  return { xml };
}

/** A simple field. */
export function sf(instr: string, result = '', o: FieldOptions = {}): Xml {
  const tag = o.nested ? 'instrText' : 't';
  const rPr = o.rPr ? `<w:rPr>${o.rPr}</w:rPr>` : '';
  const attrs = `${o.lock ? ' w:fldLock="1"' : ''}${o.dirty ? ' w:dirty="1"' : ''}`;
  const r =
    result === ''
      ? ''
      : `<w:r>${rPr}<w:${tag} xml:space="preserve">${esc(result)}</w:${tag}></w:r>`;
  return { xml: `<w:fldSimple w:instr="${esc(instr)}"${attrs}>${r}</w:fldSimple>` };
}

/** A paragraph; strings become text runs. `pPr` is inner pPr markup. */
export function para(...parts: (string | Xml)[]): Xml {
  return {
    xml: `<w:p>${parts.map((p) => (typeof p === 'string' ? run(p).xml : p.xml)).join('')}</w:p>`,
  };
}

/** A paragraph with properties. */
export function ppara(pPr: string, ...parts: (string | Xml)[]): Xml {
  const inner = para(...parts).xml.slice(5, -6);
  return { xml: `<w:p><w:pPr>${pPr}</w:pPr>${inner}</w:p>` };
}

/** A heading paragraph (built-in Heading1..3 styles of createDocx). */
export function heading(level: number, text: string): Xml {
  return ppara(`<w:pStyle w:val="Heading${level}"/>`, text);
}

/** A table: rows of cells; a cell is text or a list of paragraphs. */
export function table(rows: (string | Xml[])[][]): Xml {
  const cols = Math.max(...rows.map((r) => r.length));
  const grid = `<w:tblGrid>${'<w:gridCol w:w="1000"/>'.repeat(cols)}</w:tblGrid>`;
  const body = rows
    .map(
      (r) =>
        `<w:tr>${r
          .map(
            (c) =>
              `<w:tc>${typeof c === 'string' ? para(c).xml : c.map((x) => x.xml).join('')}</w:tc>`,
          )
          .join('')}</w:tr>`,
    )
    .join('');
  return { xml: `<w:tbl><w:tblPr/>${grid}${body}</w:tbl>` };
}

/** A bookmark around `parts`. */
export function bookmark(name: string, id: number, ...parts: (string | Xml)[]): Xml {
  return {
    xml: `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${parts
      .map((p) => (typeof p === 'string' ? run(p).xml : p.xml))
      .join('')}<w:bookmarkEnd w:id="${id}"/>`,
  };
}

/** A blank document (createDocx) whose body is `blocks`. */
export function docWith(...blocks: Xml[]): WordDocument {
  const doc = createDocx();
  const body = doc.body.element;
  const frag = parseXml(
    `<w:body xmlns:w="${NS.w}">${blocks.map((b) => b.xml).join('')}</w:body>`,
  ).root;
  const sectPr = body.element(NS.w, 'sectPr') ?? null;
  for (const k of body.children.slice()) if (k !== sectPr) k.remove();
  for (const k of frag.children.slice()) body.insertBefore(k, sectPr);
  return doc;
}

/** Results of every field, in document order. */
export function results(doc: WordDocument): string[] {
  return listFields(doc).map((f) => f.result);
}
