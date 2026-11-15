// buildIR: the layout IR built from real fixtures and from small synthetic documents.

import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readDocx, writeDocx } from '../src/index.js';
import type { FontManager } from '../src/layout/engine.js';
import {
  buildIR,
  type IRAnchorExtras,
  type IRInlineExtras,
  type IRParagraphExtras,
  wantFonts,
} from '../src/layout/ir.js';
import type { IRBlock, IRDocument, IRInline, IRParagraph, IRTable } from '../src/layout/types.js';
import type { ResolvedRunProps } from '../src/resolve/types.js';
import { NS } from '../src/xml/ns.js';
import { fixture, fixtures, strictFixtures } from './helpers/fixtures.js';

// ---- helpers -------------------------------------------------------------------------------

function paragraphs(blocks: readonly IRBlock[]): IRParagraph[] {
  return blocks.filter((b): b is IRParagraph => b.kind === 'paragraph');
}

/** Every paragraph of the IR body (tables and text boxes included), in document order. */
function allParagraphs(ir: IRDocument): IRParagraph[] {
  const out: IRParagraph[] = [];
  const stack: IRBlock[] = ir.sections.flatMap((s) => s.blocks).reverse();
  while (stack.length > 0) {
    const b = stack.pop()!;
    if (b.kind === 'table') {
      const cells = b.rows.flatMap((r) => r.cells.flatMap((c) => c.blocks));
      for (let i = cells.length - 1; i >= 0; i--) stack.push(cells[i]!);
      continue;
    }
    out.push(b);
    const inner = b.anchors.flatMap((a) => a.blocks ?? []);
    for (let i = inner.length - 1; i >= 0; i--) stack.push(inner[i]!);
  }
  return out;
}

function textOf(p: IRParagraph): string {
  return p.inlines
    .map((i) => {
      switch (i.kind) {
        case 'text':
          return i.text;
        case 'tab':
          return '\t';
        case 'break':
          return '\n';
        case 'field':
          return `{${i.type}:${i.fallback}}`;
        case 'noteRef':
          return `[${i.text}]`;
        default:
          return `<${i.kind}>`;
      }
    })
    .join('');
}

const DECLS = [
  `xmlns:w="${NS.w}"`,
  `xmlns:r="${NS.r}"`,
  `xmlns:wp="${NS.wp}"`,
  `xmlns:a="${NS.a}"`,
  `xmlns:pic="${NS.pic}"`,
  `xmlns:mc="${NS.mc}"`,
  `xmlns:wps="${NS.wps}"`,
  `xmlns:m="${NS.m}"`,
  `xmlns:v="${NS.v}"`,
  `xmlns:o="${NS.o}"`,
  `xmlns:w10="${NS.w10}"`,
].join(' ');

const CT_MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG_1x1 = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 3,
  0, 0, 0, 2, 8, 6, 0, 0, 0,
]);

interface SynthOptions {
  settings?: string;
  styles?: string;
  footnotes?: string;
  header?: string;
  core?: string;
  sectPr?: string;
}

/** A minimal .docx around `body` (document.xml content of w:body, final sectPr added). */
function synth(body: string, o: SynthOptions = {}): Uint8Array {
  const overrides: string[] = [
    `<Override PartName="/word/document.xml" ContentType="${CT_MAIN}.document.main+xml"/>`,
  ];
  const rels: string[] = [
    `<Relationship Id="rIdImg" Type="${REL}/image" Target="media/image1.png"/>`,
    `<Relationship Id="rIdLink" Type="${REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>`,
  ];
  const files: Record<string, Uint8Array> = {};
  const add = (name: string, type: string, rel: string, content: string): void => {
    overrides.push(`<Override PartName="/word/${name}" ContentType="${type}"/>`);
    rels.push(
      `<Relationship Id="rId${name.replace(/\W/g, '')}" Type="${REL}/${rel}" Target="${name}"/>`,
    );
    files[`word/${name}`] = strToU8(content);
  };
  if (o.settings !== undefined)
    add(
      'settings.xml',
      `${CT_MAIN}.settings+xml`,
      'settings',
      `<w:settings ${DECLS}>${o.settings}</w:settings>`,
    );
  if (o.styles !== undefined)
    add(
      'styles.xml',
      `${CT_MAIN}.styles+xml`,
      'styles',
      `<w:styles ${DECLS}>${o.styles}</w:styles>`,
    );
  if (o.footnotes !== undefined)
    add(
      'footnotes.xml',
      `${CT_MAIN}.footnotes+xml`,
      'footnotes',
      `<w:footnotes ${DECLS}>${o.footnotes}</w:footnotes>`,
    );
  if (o.header !== undefined)
    add('header1.xml', `${CT_MAIN}.header+xml`, 'header', `<w:hdr ${DECLS}>${o.header}</w:hdr>`);
  const pkgRels = [
    `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`,
  ];
  if (o.core !== undefined) {
    overrides.push(
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    );
    pkgRels.push(
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    );
    files['docProps/core.xml'] = strToU8(
      `<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}">${o.core}</cp:coreProperties>`,
    );
  }
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Default Extension="png" ContentType="image/png"/>' +
        `${overrides.join('')}</Types>`,
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${pkgRels.join('')}</Relationships>`,
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
    ),
    'word/document.xml': strToU8(
      `<w:document ${DECLS}><w:body>${body}${o.sectPr ?? '<w:sectPr/>'}</w:body></w:document>`,
    ),
    'word/media/image1.png': PNG_1x1,
    ...files,
  });
}

function irOf(body: string, o?: SynthOptions): IRDocument {
  return buildIR(readDocx(synth(body, o)));
}

const r = (inner: string, rPr = ''): string =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${inner}</w:r>`;
const t = (s: string): string => `<w:t xml:space="preserve">${s}</w:t>`;
const p = (inner: string, pPr = ''): string =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
const fld = (instr: string, result: string, rPr = ''): string =>
  r('<w:fldChar w:fldCharType="begin"/>') +
  r(`<w:instrText xml:space="preserve">${instr}</w:instrText>`) +
  r('<w:fldChar w:fldCharType="separate"/>') +
  r(t(result), rPr) +
  r('<w:fldChar w:fldCharType="end"/>');

// ---- real fixtures ---------------------------------------------------------------------------

describe('buildIR: real fixtures', () => {
  for (const f of [
    ...fixtures(),
    ...strictFixtures([
      'docx-mailmerge/test_footnote_header_footer.docx',
      'docxcompose/header_with_image.docx',
    ]),
  ]) {
    it(`${f.name}: builds and leaves the document untouched`, () => {
      const doc = readDocx(f.bytes);
      const ir = buildIR(doc);
      expect(ir.sections.length).toBe(doc.sections().length);
      expect(ir.sections.reduce((n, s) => n + s.blocks.length, 0)).toBeGreaterThan(0);
      // Read-only: nothing became dirty, so a save returns the input bytes.
      expect(writeDocx(doc)).toEqual(f.bytes);
      // Every body paragraph's visible text is in the IR (fields aside).
      expect(allParagraphs(ir).length).toBeGreaterThanOrEqual(doc.body.paragraphs().length);
    });
  }

  it('test_footnote_header_footer: headers, footers, notes, page breaks, settings', () => {
    const ir = buildIR(readDocx(fixture('docx-mailmerge/test_footnote_header_footer.docx').bytes));
    expect(ir.evenAndOddHeaders).toBe(true);
    expect(ir.defaultTabStop).toBe(36);
    const s = ir.sections[0]!;
    expect(s).toMatchObject({
      type: 'continuous',
      pageWidth: 612,
      pageHeight: 792,
      titlePage: true,
    });
    expect(s.margins).toMatchObject({ top: 72, left: 72, header: 35.4, gutter: 0 });
    expect(paragraphs(s.headers.default!).map(textOf)).toEqual(['Header on every page: \t\t']);
    expect(paragraphs(s.headers.first!).map(textOf)).toEqual(['Header on first page: ']);
    expect(paragraphs(s.headers.even!).map(textOf)).toEqual(['Header even: ']);
    expect(textOf(paragraphs(s.footers.first!)[0]!)).toBe('Footer on first page ');
    const body = paragraphs(s.blocks);
    // Footnote reference numbered 1; MERGEFIELD (fldSimple) shows its result.
    expect(textOf(body[0]!)).toBe('[1]Merge : «fieldname»');
    expect(body[0]!.bookmarks).toEqual(['_Ref150455263', 'OLE_LINK1', 'OLE_LINK2']);
    expect(body[3]!.inlines).toEqual([{ kind: 'break', type: 'page' }]);
    // The note itself starts with its mark (w:footnoteRef).
    expect(textOf(paragraphs(ir.footnotes.get('1')!)[0]!)).toBe('1 Merge : «fieldname»');
    expect(ir.metadata).toMatchObject({
      author: 'redbrick',
      creator: 'Microsoft Office Word',
      lang: 'en-US',
    });
  });

  it('docproperties_header_footer_3_sections: effective (inherited) headers are shared', () => {
    const ir = buildIR(
      readDocx(fixture('docxcompose/docproperties_header_footer_3_sections.docx').bytes),
    );
    expect(ir.sections).toHaveLength(3);
    expect(textOf(paragraphs(ir.sections[0]!.blocks)[0]!)).toBe('Section 1');
    expect(textOf(paragraphs(ir.sections[1]!.blocks)[0]!)).toBe('Section 2');
    expect(paragraphs(ir.sections[0]!.headers.default!).map(textOf)).toEqual(['Text: Foo Bar']);
    // Shared parts are built once: identical arrays.
    for (const s of ir.sections.slice(1)) {
      for (const k of ['default', 'first', 'even'] as const) {
        const own = s.headers[k];
        if (own !== undefined) expect(ir.sections.some((o) => o.headers[k] === own)).toBe(true);
      }
    }
    expect(ir.sections[0]!.pageWidth).toBeCloseTo(595.3);
  });

  it('header_with_image: inline picture in a header with its image resource', () => {
    const ir = buildIR(readDocx(fixture('docxcompose/header_with_image.docx').bytes));
    const img = paragraphs(ir.sections[0]!.headers.default!)[0]!.inlines[0]!;
    expect(img).toMatchObject({ kind: 'image', image: 'word/media/image1.jpg' });
    if (img.kind !== 'image') throw new Error('not an image');
    expect(img.width).toBeCloseTo(607101 / 12700, 6);
    expect(img.height).toBeCloseTo(403722 / 12700, 6);
    const res = ir.images.get('word/media/image1.jpg')!;
    expect(res).toMatchObject({ contentType: 'image/jpeg', pixelWidth: 200, pixelHeight: 133 });
    expect(res.bytes.length).toBeGreaterThan(1000);
  });

  it('TOC2: PAGEREF fields with fallbacks, internal hyperlinks, settings compat', () => {
    const ir = buildIR(readDocx(fixture('open-xml-sdk/TOC2.docx').bytes));
    const first = paragraphs(ir.sections[0]!.blocks)[0]!;
    expect(textOf(first)).toBe('阿道发送发萨发萨发送颠fs\t{PAGEREF:1}');
    const field = first.inlines.find((i) => i.kind === 'field');
    expect(field).toMatchObject({
      kind: 'field',
      type: 'PAGEREF',
      arg: '_Toc167001275',
      fallback: '1',
    });
    const linked = first.inlines.filter(
      (i) => i.kind === 'text' && i.link?.anchor === '_Toc167001275',
    );
    expect(linked).toHaveLength(2);
    expect((first as IRParagraph & IRParagraphExtras).compat).toMatchObject({
      doNotExpandShiftReturn: true,
    });
    expect(first.props.tabs).toEqual([{ pos: 8296, val: 'right', leader: 'dot' }]);
  });

  it('TableFntEdnt: table grid, cells and note marks in reference order', () => {
    const ir = buildIR(readDocx(fixture('open-xml-sdk/TableFntEdnt.docx').bytes));
    const table = ir.sections[0]!.blocks.find((b): b is IRTable => b.kind === 'table')!;
    expect(table.grid).toEqual([154, 154.05, 154.05]);
    expect(table.rows[0]!.cells).toHaveLength(3);
    expect(table.rows[0]!.cells[0]!.margins).toEqual({ top: 0, right: 5.4, bottom: 0, left: 5.4 });
    expect(table.rows[0]!.cells[0]!.borders.top).toMatchObject({ style: 'single', size: 4 });
    const cellText = table.rows.map((row) =>
      row.cells.map((c) => paragraphs(c.blocks).map(textOf).join('|')),
    );
    expect(cellText[0]![0]).toBe('test cell[i][1]');
    expect(cellText[1]![0]).toBe('cell[2]');
    expect(cellText[1]![1]).toBe('test[ii]');
    expect([...ir.footnotes.keys()]).toEqual(['2', '3', '4']);
    expect([...ir.endnotes.keys()]).toEqual(['2', '3', '4']);
    // The note's own mark follows the reference numbering (footnote 2 is referenced first).
    expect(textOf(paragraphs(ir.footnotes.get('2')!)[0]!).startsWith('1')).toBe(true);
  });

  it('tbl-cell-access: gridSpan and vMerge', () => {
    const ir = buildIR(readDocx(fixture('python-docx/tbl-cell-access.docx').bytes));
    const tables = ir.sections[0]!.blocks.filter((b): b is IRTable => b.kind === 'table');
    const cells = tables.flatMap((tb) => tb.rows.flatMap((row) => row.cells));
    expect(cells.some((c) => c.gridSpan > 1)).toBe(true);
    expect(cells.some((c) => c.vMerge === 'restart')).toBe(true);
    expect(cells.some((c) => c.vMerge === 'continue')).toBe(true);
  });

  it('shapes-with-text: anchored text boxes (DrawingML choice) and shapes', () => {
    const ir = buildIR(readDocx(fixture('apache-poi/shapes-with-text.docx').bytes));
    const anchors = paragraphs(ir.sections[0]!.blocks)[0]!.anchors;
    const box = anchors[0]! as (typeof anchors)[number] & IRAnchorExtras;
    expect(box).toMatchObject({
      kind: 'textbox',
      width: 183,
      height: 45,
      horizontal: { relativeTo: 'column', offset: 18 },
      vertical: { relativeTo: 'paragraph', offset: 0 },
      wrap: 'none',
      behindText: false,
      distance: { top: 0, right: 9, bottom: 0, left: 9 },
      zIndex: 251659264,
    });
    expect(box.source.is(NS.w, 'drawing')).toBe(true);
    expect(box.insets).toEqual({ left: 7.2, top: 3.6, right: 7.2, bottom: 3.6 });
    expect(paragraphs(box.blocks!).map(textOf)).toEqual(['Floating text box']);
    expect(anchors.filter((a) => a.kind === 'drawing').length).toBeGreaterThan(0);
  });

  it('smart_art: a non-picture inline graphic is a drawing inline', () => {
    const ir = buildIR(readDocx(fixture('docxcompose/smart_art.docx').bytes));
    const d = allParagraphs(ir)
      .flatMap((pp) => pp.inlines)
      .find((i) => i.kind === 'drawing');
    expect(d).toMatchObject({ kind: 'drawing', width: 432, height: 252 });
  });

  it('numberings: list labels with suffix and alignment', () => {
    const ir = buildIR(readDocx(fixture('docxcompose/numberings.docx').bytes));
    const labelled = paragraphs(ir.sections[0]!.blocks).filter((pp) => pp.label !== undefined);
    expect(labelled.length).toBeGreaterThan(5);
    expect(labelled[2]!.label).toMatchObject({ text: 'o', suffix: 'tab', alignment: 'left' });
  });
});

// ---- synthetic documents ----------------------------------------------------------------------

describe('buildIR: sections and settings', () => {
  it('page geometry in points, columns, page numbering, vAlign, type', () => {
    const sect1 =
      '<w:sectPr><w:type w:val="oddPage"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
      '<w:pgMar w:top="1000" w:right="800" w:bottom="1200" w:left="900" w:header="500" w:footer="600" w:gutter="100"/>' +
      '<w:cols w:num="2" w:space="400" w:sep="1"/><w:pgNumType w:start="5" w:fmt="upperRoman"/><w:vAlign w:val="center"/><w:titlePg/></w:sectPr>';
    const ir = irOf(p(r(t('one')), sect1) + p(r(t('two'))), {
      sectPr:
        '<w:sectPr><w:cols w:num="2" w:equalWidth="0"><w:col w:w="3000" w:space="200"/><w:col w:w="5000"/></w:cols></w:sectPr>',
      settings: '<w:defaultTabStop w:val="708"/><w:evenAndOddHeaders/>',
    });
    expect(ir.sections).toHaveLength(2);
    const [a, b] = ir.sections;
    expect(a).toMatchObject({
      index: 0,
      type: 'oddPage',
      pageWidth: 841.9,
      pageHeight: 595.3,
      margins: { top: 50, right: 40, bottom: 60, left: 45, header: 25, footer: 30, gutter: 5 },
      columns: { count: 2, space: 20, separator: true },
      titlePage: true,
      pageNumberStart: 5,
      pageNumberFormat: 'upperRoman',
      vAlign: 'center',
    });
    expect(paragraphs(a!.blocks).map(textOf)).toEqual(['one']);
    expect(b).toMatchObject({
      type: 'nextPage',
      pageWidth: 612,
      pageHeight: 792,
      vAlign: 'top',
      pageNumberFormat: 'decimal',
    });
    expect(b!.columns.widths).toEqual([
      { width: 150, space: 10 },
      { width: 250, space: 0 },
    ]);
    expect(paragraphs(b!.blocks).map(textOf)).toEqual(['two']);
    expect(ir.defaultTabStop).toBeCloseTo(35.4);
    expect(ir.evenAndOddHeaders).toBe(true);
  });

  it('metadata from the core properties', () => {
    const ir = irOf(p(r(t('x'))), {
      core: '<dc:title>T</dc:title><dc:creator>A</dc:creator><dc:subject>S</dc:subject><dc:language>fr-FR</dc:language>',
    });
    expect(ir.metadata).toEqual({ title: 'T', author: 'A', subject: 'S', lang: 'fr-FR' });
  });

  it('a header referenced by the final section', () => {
    const ir = irOf(p(r(t('body'))), {
      header: p(r(t('head'))),
      sectPr: '<w:sectPr><w:headerReference w:type="default" r:id="rIdheader1xml"/></w:sectPr>',
    });
    expect(paragraphs(ir.sections[0]!.headers.default!).map(textOf)).toEqual(['head']);
  });
});

describe('buildIR: inline content', () => {
  it('splits text by run, keeps tabs/breaks, skips hidden and deleted text, shows insertions', () => {
    const body = p(
      r(t('Hello ')) +
        r(t('bold'), '<w:b/>') +
        r(
          `${t('a')}<w:tab/>${t('b')}<w:br/>${t('c')}<w:br w:type="page"/><w:br w:type="column"/><w:br w:clear="all"/><w:cr/>`,
        ) +
        r(t('hidden'), '<w:vanish/>') +
        `<w:del w:id="1" w:author="x"><w:r><w:delText>gone</w:delText></w:r></w:del>` +
        `<w:ins w:id="2" w:author="x">${r(t('new'))}</w:ins>` +
        r('<w:noBreakHyphen/><w:softHyphen/>'),
    );
    const para = paragraphs(irOf(body).sections[0]!.blocks)[0]!;
    expect(textOf(para)).toBe('Hello bolda\tb\nc\n\n\n\nnew‑­');
    const kinds = para.inlines.map((i) => i.kind);
    expect(kinds.slice(0, 2)).toEqual(['text', 'text']);
    expect(para.inlines[1]).toMatchObject({ kind: 'text', text: 'bold' });
    expect((para.inlines[1] as { props: ResolvedRunProps }).props.bold).toBe(true);
    const breaks = para.inlines.filter((i) => i.kind === 'break');
    expect(breaks).toEqual([
      { kind: 'break', type: 'line' },
      { kind: 'break', type: 'page' },
      { kind: 'break', type: 'column' },
      { kind: 'break', type: 'line', clear: 'all' },
      { kind: 'break', type: 'line' },
    ]);
    // Text of one run stays one inline even across w:t elements interrupted by a tab.
    expect(para.inlines[2]).toMatchObject({ kind: 'text', text: 'a' });
  });

  it('symbols carry their font; positional tabs carry their alignment', () => {
    const body = p(
      r('<w:sym w:font="Wingdings" w:char="F0E0"/>') +
        r('<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/>'),
    );
    const para = paragraphs(irOf(body).sections[0]!.blocks)[0]!;
    const sym = para.inlines[0]!;
    expect(sym).toMatchObject({ kind: 'text', text: '' });
    expect((sym as { props: ResolvedRunProps }).props.fonts.ascii).toBe('Wingdings');
    expect((para.inlines[1] as IRInline & IRInlineExtras).ptab).toEqual({
      alignment: 'right',
      relativeTo: 'margin',
      leader: 'dot',
    });
  });

  it('page fields: format switch, current result as fallback, result formatting', () => {
    const body = p(
      r(t('Page ')) +
        fld(' PAGE \\* roman \\* MERGEFORMAT ', 'iv', '<w:b/>') +
        r(t(' of ')) +
        r('<w:fldChar w:fldCharType="begin"/>') +
        r('<w:instrText> NUMPAGES </w:instrText>') +
        r('<w:fldChar w:fldCharType="end"/>') +
        '<w:fldSimple w:instr=" SECTIONPAGES "><w:r><w:t>3</w:t></w:r></w:fldSimple>' +
        r('<w:pgNum/>'),
    );
    const para = paragraphs(irOf(body).sections[0]!.blocks)[0]!;
    const fields = para.inlines.filter((i) => i.kind === 'field');
    expect(fields.map((f) => [f.type, f.fallback, f.format])).toEqual([
      ['PAGE', 'iv', 'roman'],
      ['NUMPAGES', '', undefined],
      ['SECTIONPAGES', '3', undefined],
      ['PAGE', '', undefined],
    ]);
    expect(fields[0]!.props.bold).toBe(true);
    expect(textOf(para)).toBe('Page {PAGE:iv} of {NUMPAGES:}{SECTIONPAGES:3}{PAGE:}');
  });

  it('other fields show their result; nested fields in instructions show nothing', () => {
    const body = p(
      fld(' DATE \\@ "d.M.yyyy" ', '1.2.2024') +
        r('<w:fldChar w:fldCharType="begin"/>') +
        r('<w:instrText xml:space="preserve"> IF </w:instrText>') +
        fld(' PAGE ', '1') +
        r('<w:instrText xml:space="preserve"> = 1 "first" "other" </w:instrText>') +
        r('<w:fldChar w:fldCharType="separate"/>') +
        r(t('first')) +
        r('<w:fldChar w:fldCharType="end"/>'),
    );
    const para = paragraphs(irOf(body).sections[0]!.blocks)[0]!;
    expect(textOf(para)).toBe('1.2.2024first');
  });

  it('hyperlinks: w:hyperlink (external and anchor) and HYPERLINK fields', () => {
    const body = p(
      `<w:hyperlink r:id="rIdLink">${r(t('ext'))}</w:hyperlink>` +
        `<w:hyperlink w:anchor="bm">${r(t('int'))}</w:hyperlink>` +
        fld(' HYPERLINK "https://f.test/" \\l "sec" ', 'fld'),
    );
    const para = paragraphs(irOf(body).sections[0]!.blocks)[0]!;
    const links = para.inlines.map((i) => (i.kind === 'text' ? i.link : undefined));
    expect(links).toEqual([
      { url: 'https://example.com/' },
      { anchor: 'bm' },
      { url: 'https://f.test/', anchor: 'sec' },
    ]);
  });

  it('bookmarks: positions within the paragraph, block-level starts join the next paragraph', () => {
    const body =
      '<w:bookmarkStart w:id="0" w:name="before"/>' +
      p(`${r(t('a'))}<w:bookmarkStart w:id="1" w:name="mid"/>${r(t('b'))}`);
    const para = paragraphs(irOf(body).sections[0]!.blocks)[0]! as IRParagraph & IRParagraphExtras;
    expect(para.bookmarks).toEqual(['before', 'mid']);
    expect(para.bookmarkAt).toEqual([0, 1]);
  });

  it('footnote references: numbering, custom marks and formats from settings', () => {
    const notes = ['1', '2', '3']
      .map(
        (id) =>
          `<w:footnote w:id="${id}">${p(r('<w:footnoteRef/>') + r(t(` note ${id}`)))}</w:footnote>`,
      )
      .join('');
    const body = p(
      r('<w:footnoteReference w:id="2"/>') +
        r(`<w:footnoteReference w:customMarkFollows="1" w:id="3"/>${t('*')}`) +
        r('<w:footnoteReference w:id="1"/>'),
    );
    const ir = irOf(body, {
      footnotes: `<w:footnote w:type="separator" w:id="-1">${p(r('<w:separator/>'))}</w:footnote>${notes}`,
      settings:
        '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="2"/></w:footnotePr>',
    });
    const para = paragraphs(ir.sections[0]!.blocks)[0]!;
    expect(textOf(para)).toBe('[b][*][c]');
    expect(para.inlines.map((i) => (i.kind === 'noteRef' ? i.id : ''))).toEqual(['2', '3', '1']);
    expect([...ir.footnotes.keys()]).toEqual(['1', '2', '3']);
    expect(textOf(paragraphs(ir.footnotes.get('2')!)[0]!)).toBe('b note 2');
    expect(textOf(paragraphs(ir.footnotes.get('1')!)[0]!)).toBe('c note 1');
  });

  it('inline pictures, VML pictures and math', () => {
    const blip = (cx: number, cy: number): string =>
      `<w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:blipFill><a:blip r:embed="rIdImg"/><a:srcRect l="10000" r="20000"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const vml =
      '<w:pict><v:shape style="width:36pt;height:0.5in"><v:imagedata r:id="rIdImg"/></v:shape></w:pict>';
    const math = '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>';
    const para = paragraphs(
      irOf(p(r(blip(127000, 254000)) + r(vml) + math)).sections[0]!.blocks,
    )[0]!;
    expect(para.inlines[0]).toMatchObject({
      kind: 'image',
      image: 'word/media/image1.png',
      width: 10,
      height: 20,
    });
    expect((para.inlines[0] as IRInline & IRInlineExtras).crop).toEqual({
      left: 0.1,
      top: 0,
      right: 0.2,
      bottom: 0,
    });
    expect(para.inlines[1]).toMatchObject({
      kind: 'image',
      image: 'word/media/image1.png',
      width: 36,
      height: 36,
    });
    expect(para.inlines[2]).toMatchObject({ kind: 'math' });
    const ir = irOf(p(r(blip(1, 1))));
    expect(ir.images.get('word/media/image1.png')).toMatchObject({
      contentType: 'image/png',
      pixelWidth: 3,
      pixelHeight: 2,
    });
  });

  it('anchored pictures: alignment, wrap, behind text, distances', () => {
    const anchor =
      '<w:drawing><wp:anchor distT="12700" distB="25400" distL="0" distR="0" simplePos="0" relativeHeight="7" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>' +
      '<wp:positionV relativeFrom="page"><wp:posOffset>127000</wp:posOffset></wp:positionV><wp:extent cx="254000" cy="127000"/>' +
      '<wp:wrapTopAndBottom/><wp:docPr id="3" name="a"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>';
    const para = paragraphs(irOf(p(r(anchor) + r(t('text')))).sections[0]!.blocks)[0]!;
    expect(para.inlines.map((i) => i.kind)).toEqual(['text']);
    expect(para.anchors[0]).toMatchObject({
      kind: 'image',
      image: 'word/media/image1.png',
      width: 20,
      height: 10,
      horizontal: { relativeTo: 'margin', align: 'center' },
      vertical: { relativeTo: 'page', offset: 10 },
      wrap: 'topAndBottom',
      behindText: true,
      distance: { top: 1, bottom: 2, left: 0, right: 0 },
      zIndex: 7,
    });
  });

  it('nested tables and text boxes are built without recursion and capped in depth', () => {
    const cell = (inner: string): string =>
      `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>${inner}${p('')}</w:tc></w:tr></w:tbl>`;
    let inner = p(r(t('deep')));
    for (let i = 0; i < 40; i++) inner = cell(inner);
    const warnings: string[] = [];
    const ir = buildIR(readDocx(synth(inner)), { warn: (c) => warnings.push(c) });
    let depth = 0;
    let blocks = ir.sections[0]!.blocks;
    while (blocks[0]?.kind === 'table') {
      depth++;
      blocks = blocks[0].rows[0]!.cells[0]!.blocks;
    }
    expect(depth).toBe(33);
    expect(warnings).toContain('LAYOUT_NESTING');
  });

  it('mc:AlternateContent: the DrawingML choice, else the fallback', () => {
    const ac = (requires: string): string =>
      `<mc:AlternateContent><mc:Choice Requires="${requires}">${r(t('choice'))}</mc:Choice><mc:Fallback>${r(t('fallback'))}</mc:Fallback></mc:AlternateContent>`;
    const ir = irOf(p(ac('wps')) + p(ac('w99')).replace('<w:p>', `<w:p xmlns:w99="urn:unknown">`));
    expect(paragraphs(ir.sections[0]!.blocks).map(textOf)).toEqual(['choice', 'fallback']);
  });
});

describe('wantFonts', () => {
  it('registers every run format of body, headers, notes, cells and labels', () => {
    const ir = buildIR(readDocx(fixture('open-xml-sdk/TableFntEdnt.docx').bytes));
    const wanted: ResolvedRunProps[] = [];
    const fm = { want: (pp: ResolvedRunProps) => wanted.push(pp) } as unknown as FontManager;
    wantFonts(ir, fm);
    const families = new Set(wanted.map((w) => w.fonts.ascii));
    expect(wanted.length).toBeGreaterThan(10);
    expect(families.size).toBeGreaterThan(0);
  });
});
