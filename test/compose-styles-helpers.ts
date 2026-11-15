// Test helpers of the compose-styles agent (test/compose-styles.test.ts and
// test/compose-numbering.test.ts): .docx files built from scratch, a minimal ImportContext (real
// documents, a simple part copier, fake ids) so importStyles/importNumbering run without the
// orchestrator, and structural checks a merged document must pass.

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strToU8, zipSync } from 'fflate';
import { importNumbering } from '../src/compose/numbering.js';
import { importStyles } from '../src/compose/styles.js';
import type { ImportContext, ImportOptions } from '../src/compose/types.js';
import { readDocx, validatePackage, type WordDocument, writeDocx } from '../src/index.js';
import type { Part } from '../src/opc/part.js';
import { relativeTarget } from '../src/opc/uri.js';
import type { DocumentIds } from '../src/wml/ids.js';
import type { XElement } from '../src/xml/dom.js';
import { NS } from '../src/xml/ns.js';
import { fixture } from './helpers/fixtures.js';

// ---- builders -----------------------------------------------------------------------------

export const W: string = NS.w;
export const DECL: string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
export const XMLNS: string = `xmlns:w="${NS.w}" xmlns:r="${NS.r}"`;
export const OUT: string = fileURLToPath(new URL('../out/compose-styles/', import.meta.url));

export interface ExtraPart {
  name: string;
  contentType: string;
  xml: string;
  /** Relationship type from the main part (omit for none). */
  rel?: string;
  relId?: string;
}

export interface DocSpec {
  /** Inner XML of w:styles (undefined: no styles part). */
  styles?: string;
  /** Inner XML of w:body (a sectPr is appended). */
  body: string;
  /** Inner XML of w:numbering. */
  numbering?: string;
  /** Extra sectPr children (header references). */
  sectPr?: string;
  parts?: ExtraPart[];
  /** Raw extra zip entries (rels of other parts, media). */
  files?: Record<string, string | Uint8Array>;
}

export const CT_BASE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
export const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A .docx built from scratch (fflate), opened with readDocx. */
export function makeDocx(spec: DocSpec): WordDocument {
  const parts: ExtraPart[] = [...(spec.parts ?? [])];
  if (spec.styles !== undefined) {
    parts.push({
      name: 'word/styles.xml',
      contentType: `${CT_BASE}.styles+xml`,
      xml: `${DECL}<w:styles ${XMLNS}>${spec.styles}</w:styles>`,
      rel: `${REL_BASE}/styles`,
    });
  }
  if (spec.numbering !== undefined) {
    parts.push({
      name: 'word/numbering.xml',
      contentType: `${CT_BASE}.numbering+xml`,
      xml: `${DECL}<w:numbering ${XMLNS}>${spec.numbering}</w:numbering>`,
      rel: `${REL_BASE}/numbering`,
    });
  }
  const files: Record<string, Uint8Array> = {};
  const overrides = parts
    .map((p) => `<Override PartName="/${p.name}" ContentType="${p.contentType}"/>`)
    .join('');
  files['[Content_Types].xml'] = strToU8(
    `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      `<Override PartName="/word/document.xml" ContentType="${CT_BASE}.document.main+xml"/>${overrides}</Types>`,
  );
  files['_rels/.rels'] = strToU8(
    `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL_BASE}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  let n = 0;
  const rels = parts
    .filter((p) => p.rel !== undefined)
    .map(
      (p) =>
        `<Relationship Id="${p.relId ?? `rIdX${++n}`}" Type="${p.rel}" Target="${relativeTarget('word/document.xml', p.name)}"/>`,
    )
    .join('');
  files['word/_rels/document.xml.rels'] = strToU8(
    `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
  );
  files['word/document.xml'] = strToU8(
    `${DECL}<w:document ${XMLNS}><w:body>${spec.body}<w:sectPr>${spec.sectPr ?? ''}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  for (const p of parts) files[p.name] = strToU8(p.xml);
  for (const [name, data] of Object.entries(spec.files ?? {})) {
    files[name] = typeof data === 'string' ? strToU8(data) : data;
  }
  return readDocx(zipSync(files));
}

export function loadFixture(name: string): WordDocument {
  return readDocx(fixture(`compose-styles/${name}`).bytes);
}

export const DEFAULT_OPTIONS: ImportContext['options'] = {
  styles: 'useDestinationStyles',
  smartStyleBehavior: true,
  lists: 'keepSource',
  restartListsAtImport: false,
  sectionStart: 'nextPage',
  headersFooters: 'keepSource',
  pageNumbering: 'keepSource',
  keepSourceSections: true,
  theme: 'destination',
  bookmarks: 'rename',
  comments: 'keep',
  revisions: 'keep',
  validate: false,
};

/** Copies a related part (bytes and content type) into the target, deduped per source part. */
export function simpleCopier(target: WordDocument): ImportContext['copier'] {
  const copied = new Map<Part, Part>();
  const copyPart = (p: Part): Part => {
    const hit = copied.get(p);
    if (hit) return hit;
    const dot = p.name.lastIndexOf('.');
    const name = target.package.uniqueName(`${p.name.slice(0, dot)}_{n}${p.name.slice(dot)}`);
    const out = target.package.addPart(
      name,
      p.contentType ?? 'application/octet-stream',
      p.bytes(),
    );
    copied.set(p, out);
    return out;
  };
  return {
    copyPart,
    copyRelationship(sourcePart, relId, targetPart) {
      const rel = sourcePart.rels.get(relId);
      if (rel === undefined) throw new Error(`no relationship ${relId}`);
      if (rel.external) return targetPart.rels.add(rel.type, rel.target, { external: true }).id;
      const p = sourcePart.related(relId);
      if (p === undefined) throw new Error(`dangling relationship ${relId}`);
      const out = copyPart(p);
      return targetPart.rels.add(rel.type, relativeTarget(targetPart.name, out.name)).id;
    },
  };
}

export function fakeIds(): DocumentIds {
  let n = 1000;
  return {
    next: () => ++n,
    nextHex: () => (++n).toString(16).toUpperCase().padStart(8, '0'),
    nextVmlShapeId: () => `_x0000_s${++n}`,
    has: () => false,
    reserve: () => undefined,
    invalidate: () => undefined,
  } as unknown as DocumentIds;
}

export interface Run {
  ctx: ImportContext;
  blocks: XElement[];
  warnings: string[];
}

/** Clones the source body into the target and runs importStyles (+ importNumbering). */
export function runImport(
  target: WordDocument,
  source: WordDocument,
  o: ImportOptions = {},
  numbering = true,
): Run {
  const warnings: string[] = [];
  const ctx: ImportContext = {
    target,
    source,
    options: { ...DEFAULT_OPTIONS, ...o } as ImportContext['options'],
    ids: fakeIds(),
    copier: simpleCopier(target),
    warn: (code, message) => warnings.push(`${code}: ${message}`),
    styleMap: new Map(),
    numIdMap: new Map(),
    bookmarkRenames: new Map(),
    commentIdMap: new Map(),
    footnoteIdMap: new Map(),
    endnoteIdMap: new Map(),
    addedParts: [],
    importedStyles: [],
  };
  const body = target.body.element;
  const sectPr = body.element(W, 'sectPr') ?? null;
  const blocks: XElement[] = [];
  for (const k of source.body.element.elements()) {
    if (k.is(W, 'sectPr')) continue;
    const c = k.clone();
    body.insertBefore(c, sectPr);
    blocks.push(c);
  }
  importStyles(ctx, blocks);
  if (numbering) importNumbering(ctx, blocks);
  return { ctx, blocks, warnings };
}

// ---- inspection ---------------------------------------------------------------------------

export function stylesRoot(doc: WordDocument): XElement {
  return doc.mainPart.relatedByType('styles')[0]!.xml().root;
}

export function styleById(doc: WordDocument, id: string): XElement | undefined {
  return stylesRoot(doc)
    .elements(W, 'style')
    .find((s) => s.attrNS(W, 'styleId') === id);
}

export function styleIds(doc: WordDocument): string[] {
  return stylesRoot(doc)
    .elements(W, 'style')
    .map((s) => s.attrNS(W, 'styleId') ?? '');
}

export const val = (el: XElement | undefined, local: string): string | undefined =>
  el?.element(W, local)?.attrNS(W, 'val');

export function pStyleOf(p: XElement): string | undefined {
  return val(p.element(W, 'pPr'), 'pStyle');
}

/** The paragraphs among the imported blocks, by their text. */
export function para(run: Run, text: string): XElement {
  for (const b of run.blocks) {
    for (const p of [b, ...b.descendants(W, 'p')]) {
      if (p.is(W, 'p') && p.text === text) return p;
    }
  }
  throw new Error(`no imported paragraph "${text}"`);
}

export function firstRun(p: XElement): XElement {
  return p.descendants(W, 'r')[0]!;
}

/** Structural rules a merged document must keep: resolvable, unique styles. */
export function styleIntegrity(doc: WordDocument): string[] {
  const problems: string[] = [];
  const root = stylesRoot(doc);
  const ids = new Map<string, string>();
  const names = new Set<string>();
  for (const s of root.elements(W, 'style')) {
    const id = s.attrNS(W, 'styleId') ?? '';
    if (ids.has(id)) problems.push(`duplicate styleId ${id}`);
    ids.set(id, s.attrNS(W, 'type') ?? 'paragraph');
    const name = val(s, 'name')?.toLowerCase();
    if (name !== undefined) {
      if (names.has(name)) problems.push(`duplicate style name ${name}`);
      names.add(name);
    }
  }
  for (const s of root.elements(W, 'style')) {
    for (const local of ['basedOn', 'next', 'link']) {
      const v = val(s, local);
      if (v !== undefined && !ids.has(v)) problems.push(`${s.attrNS(W, 'styleId')} ${local} ${v}`);
    }
  }
  const want: Record<string, string> = {
    pStyle: 'paragraph',
    rStyle: 'character',
    tblStyle: 'table',
  };
  for (const el of doc.body.element.descendants(W)) {
    const type = want[el.local];
    if (type === undefined) continue;
    const v = el.attrNS(W, 'val') ?? '';
    if (ids.get(v) !== type) problems.push(`${el.local} ${v} does not resolve to a ${type} style`);
  }
  return problems;
}

/** writeDocx → readDocx round trip; returns the validatePackage errors. */
export function savedErrors(doc: WordDocument, file?: string): string[] {
  const bytes = writeDocx(doc);
  if (file !== undefined) {
    mkdirSync(OUT, { recursive: true });
    writeFileSync(join(OUT, file), bytes);
  }
  const back = readDocx(bytes);
  return validatePackage(back)
    .filter((i) => i.severity === 'error')
    .map((i) => `${i.code}: ${i.message}`);
}

// ---- synthetic styles ----------------------------------------------------------------------

export const DEFAULTS = (sz = '22', font = 'Calibri'): string =>
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${sz}"/></w:rPr></w:rPrDefault>` +
  '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>';

export const style = (type: string, id: string, name: string, inner = '', attrs = ''): string =>
  `<w:style w:type="${type}" w:styleId="${id}"${attrs}><w:name w:val="${name}"/>${inner}</w:style>`;

export const NORMAL = (id = 'Normal', inner = ''): string =>
  style('paragraph', id, 'Normal', `<w:qFormat/>${inner}`, ' w:default="1"');
export const DPF = (id = 'DefaultParagraphFont'): string =>
  style('character', id, 'Default Paragraph Font', '<w:uiPriority w:val="1"/>', ' w:default="1"');
export const TABLE_NORMAL = (id = 'TableNormal', margins = '108'): string =>
  style(
    'table',
    id,
    'Normal Table',
    `<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:left w:w="${margins}" w:type="dxa"/><w:right w:w="${margins}" w:type="dxa"/></w:tblCellMar></w:tblPr>`,
    ' w:default="1"',
  );
export const HEADING = (
  id: string,
  color: string,
  sz: string,
  extra = '',
  basedOn = 'Normal',
): string =>
  style(
    'paragraph',
    id,
    'heading 1',
    `<w:basedOn w:val="${basedOn}"/><w:next w:val="${basedOn}"/><w:uiPriority w:val="9"/><w:qFormat/>` +
      `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr>` +
      `<w:rPr><w:color w:val="${color}"/><w:sz w:val="${sz}"/>${extra}</w:rPr>`,
  );

export const P = (text: string, pStyle?: string, rPr = ''): string =>
  `<w:p>${pStyle ? `<w:pPr><w:pStyle w:val="${pStyle}"/></w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t>${text}</w:t></w:r></w:p>`;

export function target(extraStyles = '', sz = '22'): WordDocument {
  return makeDocx({
    styles: `${DEFAULTS(sz)}${NORMAL()}${DPF()}${TABLE_NORMAL()}${HEADING('Heading1', '2F5496', '32', '<w:i/>')}${extraStyles}`,
    body: P('Target heading', 'Heading1') + P('Target text'),
  });
}
