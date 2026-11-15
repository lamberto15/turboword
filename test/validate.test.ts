// validatePackage: every code is checked both ways (reported when broken, silent when fine) on
// synthetic packages built with the OPC API, saved and reopened, so the checks see what a reader
// of the written file sees.

import { describe, expect, it } from 'vitest';
import { CT, OpcPackage, REL } from '../src/opc/index.js';
import { type ValidationIssue, validatePackage } from '../src/wml/validate.js';
import { NS } from '../src/xml/ns.js';

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NSDECL = [
  `xmlns:w="${NS.w}"`,
  `xmlns:r="${NS.r}"`,
  `xmlns:wp="${NS.wp}"`,
  `xmlns:a="${NS.a}"`,
  `xmlns:dgm="${NS.dgm}"`,
  `xmlns:mc="${NS.mc}"`,
].join(' ');
const P = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>';
const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

const documentXml = (body: string, rootAttrs = NSDECL): string =>
  `${DECL}<w:document ${rootAttrs}><w:body>${body}</w:body></w:document>`;
const headerXml = (content: string): string => `${DECL}<w:hdr ${NSDECL}>${content}</w:hdr>`;

const drawing = (id: string): string =>
  `<w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  '<a:graphic><a:graphicData uri="urn:test"/></a:graphic></wp:inline></w:drawing>';
const drawingP = (id: string): string => `<w:p><w:r>${drawing(id)}</w:r></w:p>`;

const table = (cell: string): string =>
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${cell}</w:tc></w:tr></w:tbl>`;

interface BuildOptions {
  /** Body content (default: one paragraph and the body sectPr). */
  body?: string;
  /** Whole document.xml (overrides `body`). */
  xml?: string;
  /** Content type of the main part (default CT.document). */
  mainType?: string;
  /** No main document part or relationship at all. */
  noMain?: boolean;
  /** Further edits before saving. */
  edit?: (pkg: OpcPackage) => void;
}

/** create → addPart → save → reopen. */
function build(o: BuildOptions = {}): OpcPackage {
  const pkg = OpcPackage.create();
  if (!o.noMain) {
    const xml = o.xml ?? documentXml(o.body ?? P + SECT);
    pkg.addPart('word/document.xml', o.mainType ?? CT.document, enc(xml));
    pkg.rels.add(REL.officeDocument, 'word/document.xml');
  }
  o.edit?.(pkg);
  return OpcPackage.open(pkg.save());
}

const main = (pkg: OpcPackage) => pkg.part('word/document.xml')!;
const codes = (issues: ValidationIssue[]): string[] => issues.map((i) => i.code);
const only = (issues: ValidationIssue[], code: string): ValidationIssue[] =>
  issues.filter((i) => i.code === code);

/** Adds word/header1.xml related from the main part with id `rIdH`. */
function addHeader(pkg: OpcPackage, content: string): void {
  pkg.addPart('word/header1.xml', CT.header, enc(headerXml(content)));
  main(pkg).rels.add(REL.header, 'header1.xml', { id: 'rIdH' });
}
const SECT_WITH_HEADER =
  '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>' +
  '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

describe('validatePackage', () => {
  it('reports nothing for a minimal valid package', () => {
    expect(validatePackage(build())).toEqual([]);
  });

  it('accepts anything with a `package` (a WordDocument)', () => {
    const pkg = build({ body: P });
    expect(validatePackage({ package: pkg })).toEqual(validatePackage(pkg));
  });

  it('does not modify the package', () => {
    const bytes = build().save();
    const pkg = OpcPackage.open(bytes);
    validatePackage(pkg);
    expect(pkg.dirty).toBe(false);
    expect(pkg.save()).toBe(bytes);
  });

  describe('XML_MALFORMED', () => {
    it('reports an XML part that does not parse', () => {
      const issues = validatePackage(
        build({ edit: (pkg) => pkg.addPart('customXml/item1.xml', CT.xml, enc('<a><b></a>')) }),
      );
      expect(issues).toHaveLength(1);
      expect(issues[0]).toMatchObject({
        severity: 'error',
        code: 'XML_MALFORMED',
        part: 'customXml/item1.xml',
      });
    });

    it('reports a malformed rels part, and skips checks that need it', () => {
      const pkg = build({
        body: `<w:p><w:hyperlink r:id="rId1"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>${SECT}`,
        edit: (p) => p.addPart('word/_rels/document.xml.rels', CT.relationships, enc('<Rel')),
      });
      expect(codes(validatePackage(pkg))).toEqual(['XML_MALFORMED']);
    });

    it('does not parse media or other binary parts', () => {
      const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x3c, 0x3c]);
      const issues = validatePackage(
        build({
          edit: (pkg) => {
            pkg.addPart('word/media/image1.png', 'image/png', png);
            pkg.addPart('customXml/item1.xml', CT.xml, enc('<root><x/></root>'));
          },
        }),
      );
      expect(issues).toEqual([]);
    });
  });

  describe('CT_MISSING', () => {
    // OpcPackage.open() refuses a zip without [Content_Types].xml, so the check is exercised on a
    // package view that hides it (as a package assembled by other means could lack it).
    const isCt = (n: string) => n.replace(/^\/+/, '').toLowerCase() === '[content_types].xml';
    function hideContentTypes(pkg: OpcPackage): OpcPackage {
      return new Proxy(pkg, {
        get(target, prop) {
          if (prop === 'part') return (n: string) => (isCt(n) ? undefined : target.part(n));
          if (prop === 'parts') return () => target.parts().filter((p) => !isCt(p.name));
          if (prop === 'contentTypes') throw new Error('no [Content_Types].xml');
          const v: unknown = Reflect.get(target, prop, target);
          return typeof v === 'function' ? v.bind(target) : v;
        },
      });
    }

    it('reports a package without [Content_Types].xml (and no content-type noise)', () => {
      const issues = validatePackage(hideContentTypes(build()));
      expect(issues).toEqual([
        { severity: 'error', code: 'CT_MISSING', message: expect.any(String) },
      ]);
    });

    it('is silent when [Content_Types].xml exists', () => {
      expect(codes(validatePackage(build()))).not.toContain('CT_MISSING');
    });
  });

  describe('NO_CONTENT_TYPE', () => {
    it('reports a part with neither Default nor Override', () => {
      const issues = validatePackage(
        build({
          edit: (pkg) => {
            pkg.addPart('word/embeddings/blob.bin', 'application/octet-stream', enc('x'));
            pkg.contentTypes.removeOverride('word/embeddings/blob.bin');
          },
        }),
      );
      expect(issues).toEqual([
        {
          severity: 'error',
          code: 'NO_CONTENT_TYPE',
          message: expect.stringContaining('word/embeddings/blob.bin'),
          part: 'word/embeddings/blob.bin',
        },
      ]);
    });

    it('accepts a content type from an Override or from a Default', () => {
      const issues = validatePackage(
        build({
          edit: (pkg) => {
            pkg.addPart('word/embeddings/blob.bin', 'application/octet-stream', enc('x'));
            pkg.contentTypes.setDefault('png', 'image/png');
            pkg.addPart('word/media/image1.png', 'image/png', new Uint8Array([0x89, 0x50]));
          },
        }),
      );
      expect(issues).toEqual([]);
    });
  });

  describe('DANGLING_OVERRIDE', () => {
    it('warns about an Override for a part that does not exist', () => {
      const issues = validatePackage(
        build({ edit: (pkg) => pkg.contentTypes.setOverride('word/ghost.xml', CT.header) }),
      );
      expect(issues).toHaveLength(1);
      expect(issues[0]).toMatchObject({
        severity: 'warning',
        code: 'DANGLING_OVERRIDE',
        part: 'word/ghost.xml',
      });
    });

    it('is silent when every Override names an existing part', () => {
      const issues = validatePackage(
        build({
          edit: (pkg) => pkg.addPart('word/styles.xml', CT.styles, enc(`<w:styles ${NSDECL}/>`)),
        }),
      );
      expect(issues).toEqual([]);
    });
  });

  describe('REL_TARGET_MISSING', () => {
    it('reports an internal relationship whose target part is missing', () => {
      const issues = validatePackage(
        build({ edit: (pkg) => main(pkg).rels.add(REL.styles, 'styles.xml') }),
      );
      expect(issues).toHaveLength(1);
      expect(issues[0]).toMatchObject({
        severity: 'error',
        code: 'REL_TARGET_MISSING',
        part: 'word/_rels/document.xml.rels',
      });
      expect(issues[0]!.message).toContain('word/styles.xml');
    });

    it('ignores external targets and accepts existing ones (relative or absolute)', () => {
      const issues = validatePackage(
        build({
          edit: (pkg) => {
            pkg.addPart('word/styles.xml', CT.styles, enc(`<w:styles ${NSDECL}/>`));
            main(pkg).rels.add(REL.styles, '/word/styles.xml');
            main(pkg).rels.add(REL.hyperlink, 'https://example.com/a b', { external: true });
          },
        }),
      );
      expect(issues).toEqual([]);
    });
  });

  describe('DUPLICATE_REL_ID', () => {
    function withTwoRels(sameId: boolean): OpcPackage {
      return build({
        edit: (pkg) => {
          pkg.addPart('word/styles.xml', CT.styles, enc(`<w:styles ${NSDECL}/>`));
          pkg.addPart('word/settings.xml', CT.settings, enc(`<w:settings ${NSDECL}/>`));
          main(pkg).rels.add(REL.styles, 'styles.xml');
          main(pkg).rels.add(REL.settings, 'settings.xml');
          if (sameId) {
            const rels = pkg.part('word/_rels/document.xml.rels')!.xml().root;
            rels.elements(undefined, 'Relationship')[1]!.setAttr('Id', 'rId1');
          }
        },
      });
    }

    it('reports two relationships with the same Id in one rels part', () => {
      const issues = validatePackage(withTwoRels(true));
      expect(issues).toEqual([
        {
          severity: 'error',
          code: 'DUPLICATE_REL_ID',
          message: expect.stringContaining('"rId1"'),
          part: 'word/_rels/document.xml.rels',
        },
      ]);
    });

    it('is silent for distinct ids', () => {
      expect(validatePackage(withTwoRels(false))).toEqual([]);
    });
  });

  describe('REL_ID_UNRESOLVED', () => {
    it('reports r:* attributes of any name that are not ids of the part', () => {
      const body =
        '<w:p><w:hyperlink r:id="rId42"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>' +
        '<w:p><w:hyperlink r:id="rId42"><w:r><w:t>same id again</w:t></w:r></w:hyperlink></w:p>' +
        '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="d"/><a:graphic><a:graphicData uri="u">' +
        '<a:blip r:embed="rId7"/><dgm:relIds r:dm="rId8" r:lo="rId9" r:qs="rId10" r:cs="rId11"/>' +
        `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>${SECT}`;
      const issues = only(validatePackage(build({ body })), 'REL_ID_UNRESOLVED');
      expect(issues.map((i) => i.message.split(' ')[0])).toEqual([
        'r:id="rId42"',
        'r:embed="rId7"',
        'r:dm="rId8"',
        'r:lo="rId9"',
        'r:qs="rId10"',
        'r:cs="rId11"',
      ]);
      expect(issues.every((i) => i.part === 'word/document.xml' && i.severity === 'error')).toBe(
        true,
      );
    });

    it('ignores empty references, which Word writes routinely (r:blip="")', () => {
      const body =
        '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="d"/><a:graphic><a:graphicData uri="u">' +
        `<a:blip r:embed="" r:link=""/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>${SECT}`;
      expect(only(validatePackage(build({ body })), 'REL_ID_UNRESOLVED')).toEqual([]);
    });

    it('resolves ids against the relationships of the part that uses them', () => {
      const issues = validatePackage(
        build({
          body: `<w:p><w:hyperlink r:id="rIdL"><w:r><w:t>ok</w:t></w:r></w:hyperlink></w:p>${SECT_WITH_HEADER}`,
          edit: (pkg) => {
            main(pkg).rels.add(REL.hyperlink, 'https://example.com', {
              external: true,
              id: 'rIdL',
            });
            // rIdL exists in document.xml.rels but not in header1.xml.rels.
            addHeader(
              pkg,
              '<w:p><w:hyperlink r:id="rIdL"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>',
            );
          },
        }),
      );
      expect(issues).toEqual([
        {
          severity: 'error',
          code: 'REL_ID_UNRESOLVED',
          message: expect.stringContaining('r:id="rIdL"'),
          part: 'word/header1.xml',
        },
      ]);
    });

    it('is silent when every id resolves', () => {
      const issues = validatePackage(
        build({
          body: `<w:p><w:hyperlink r:id="rIdL"><w:r><w:t>ok</w:t></w:r></w:hyperlink></w:p>${SECT}`,
          edit: (pkg) =>
            main(pkg).rels.add(REL.hyperlink, 'https://example.com', {
              external: true,
              id: 'rIdL',
            }),
        }),
      );
      expect(issues).toEqual([]);
    });
  });

  describe('NO_MAIN_DOCUMENT', () => {
    it('reports a package without an officeDocument relationship', () => {
      expect(codes(validatePackage(build({ noMain: true })))).toEqual(['NO_MAIN_DOCUMENT']);
    });

    it('reports an officeDocument relationship whose target is missing', () => {
      const pkg = build({
        noMain: true,
        edit: (p) => p.rels.add(REL.officeDocument, 'word/document.xml'),
      });
      const issues = validatePackage(pkg);
      expect(codes(issues)).toEqual(['REL_TARGET_MISSING', 'NO_MAIN_DOCUMENT']);
      expect(issues[1]!.part).toBe('word/document.xml');
    });

    it('is silent when the main part exists', () => {
      expect(codes(validatePackage(build()))).not.toContain('NO_MAIN_DOCUMENT');
    });
  });

  describe('MAIN_CONTENT_TYPE', () => {
    it('reports a main part that is not a WordprocessingML main part', () => {
      const issues = validatePackage(build({ mainType: CT.xml }));
      expect(issues).toEqual([
        {
          severity: 'error',
          code: 'MAIN_CONTENT_TYPE',
          message: expect.stringContaining('application/xml'),
          part: 'word/document.xml',
        },
      ]);
    });

    it('accepts documents, templates and their macro-enabled variants', () => {
      for (const mainType of [CT.document, CT.documentMacro, CT.template, CT.templateMacro]) {
        expect(validatePackage(build({ mainType }))).toEqual([]);
      }
    });
  });

  describe('DUPLICATE_DOCPR_ID', () => {
    it('reports a docPr id used twice in the body', () => {
      const issues = validatePackage(
        build({ body: drawingP('1') + drawingP('2') + drawingP('1') + SECT }),
      );
      expect(issues).toEqual([
        {
          severity: 'error',
          code: 'DUPLICATE_DOCPR_ID',
          message: expect.stringContaining('id="1" is used 2 times'),
          part: 'word/document.xml',
        },
      ]);
    });

    it('reports ids repeated across the main document and a header (numerically)', () => {
      const issues = validatePackage(
        build({
          body: drawingP('7') + SECT_WITH_HEADER,
          edit: (pkg) => addHeader(pkg, drawingP('007')),
        }),
      );
      expect(issues).toHaveLength(1);
      expect(issues[0]).toMatchObject({ code: 'DUPLICATE_DOCPR_ID', part: 'word/header1.xml' });
      expect(issues[0]!.message).toContain('word/document.xml, word/header1.xml');
    });

    it('is silent for unique ids, and for the Fallback copy of an AlternateContent', () => {
      const alt =
        `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps">${drawing('5')}</mc:Choice>` +
        `<mc:Fallback>${drawing('5')}</mc:Fallback></mc:AlternateContent></w:r></w:p>`;
      const issues = validatePackage(
        build({
          body: drawingP('1') + alt + SECT_WITH_HEADER,
          edit: (pkg) => addHeader(pkg, drawingP('2')),
        }),
      );
      expect(issues).toEqual([]);
    });
  });

  describe('SECTPR_PLACEMENT', () => {
    it('reports a body sectPr that is not the last child of w:body', () => {
      const issues = validatePackage(build({ body: P + SECT + P }));
      expect(issues).toEqual([
        {
          severity: 'error',
          code: 'SECTPR_PLACEMENT',
          message: expect.stringContaining('last child'),
          part: 'word/document.xml',
        },
      ]);
    });

    it('reports a sectPr outside w:body and w:pPr', () => {
      const issues = validatePackage(build({ body: `<w:p><w:r>${SECT}</w:r></w:p>${SECT}` }));
      expect(codes(issues)).toEqual(['SECTPR_PLACEMENT']);
      expect(issues[0]!.message).toContain('<w:r>');
    });

    it('accepts paragraph sectPrs, the last body sectPr and tracked sectPr changes', () => {
      const body =
        `<w:p><w:pPr>${SECT}</w:pPr></w:p>${P}` +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
        '<w:sectPrChange w:id="1" w:author="a"><w:sectPr/></w:sectPrChange></w:sectPr>';
      expect(validatePackage(build({ body }))).toEqual([]);
    });
  });

  describe('CELL_NO_PARAGRAPH', () => {
    it('reports a cell that ends with a table', () => {
      const issues = validatePackage(build({ body: table(P + table(P)) + P + SECT }));
      expect(issues).toEqual([
        {
          severity: 'error',
          code: 'CELL_NO_PARAGRAPH',
          message: expect.stringContaining('<w:tbl>'),
          part: 'word/document.xml',
        },
      ]);
    });

    it('reports a cell without block content, in headers too', () => {
      const issues = validatePackage(
        build({ body: P + SECT_WITH_HEADER, edit: (pkg) => addHeader(pkg, table('') + P) }),
      );
      expect(issues).toHaveLength(1);
      expect(issues[0]).toMatchObject({ code: 'CELL_NO_PARAGRAPH', part: 'word/header1.xml' });
      expect(issues[0]!.message).toContain('no block content');
    });

    it('accepts cells ending with a paragraph, after markup or inside a content control', () => {
      const cells = [
        P,
        `${table(P)}${P}<w:bookmarkEnd w:id="0"/>`,
        `<w:sdt><w:sdtPr/><w:sdtContent>${P}</w:sdtContent></w:sdt>`,
      ];
      const body = cells.map(table).join(P) + P + SECT;
      expect(validatePackage(build({ body }))).toEqual([]);
    });
  });

  describe('UNDECLARED_IGNORABLE', () => {
    const w14 = `xmlns:w14="${NS.w14}"`;
    const w15 = `xmlns:w15="${NS.w15}"`;

    it('reports a prefix listed in mc:Ignorable but not declared', () => {
      const xml = documentXml(P + SECT, `${NSDECL} ${w14} mc:Ignorable="w14 w15"`);
      const issues = validatePackage(build({ xml }));
      expect(issues).toEqual([
        {
          severity: 'error',
          code: 'UNDECLARED_IGNORABLE',
          message: expect.stringContaining('"w15"'),
          part: 'word/document.xml',
        },
      ]);
    });

    it('is silent when every listed prefix is declared', () => {
      const xml = documentXml(P + SECT, `${NSDECL} ${w14} ${w15} mc:Ignorable="w14  w15"`);
      expect(validatePackage(build({ xml }))).toEqual([]);
    });
  });

  it('returns issues in a deterministic, documented order', () => {
    const make = () =>
      build({
        body: `<w:p><w:hyperlink r:id="rId99"/></w:p>${SECT}${drawingP('1')}${drawingP('1')}`,
        edit: (pkg) => {
          pkg.addPart('word/zz.xml', CT.xml, enc('<broken'));
          pkg.addPart('word/aa.bin', 'application/x-test', enc('x'));
          pkg.contentTypes.removeOverride('word/aa.bin');
          pkg.contentTypes.setOverride('word/ghost.xml', CT.footer);
          main(pkg).rels.add(REL.numbering, 'numbering.xml');
        },
      });
    const first = validatePackage(make());
    expect(validatePackage(make())).toEqual(first);
    expect(codes(first)).toEqual([
      'XML_MALFORMED',
      'NO_CONTENT_TYPE',
      'DANGLING_OVERRIDE',
      'REL_TARGET_MISSING',
      'REL_ID_UNRESOLVED',
      'SECTPR_PLACEMENT',
      'DUPLICATE_DOCPR_ID',
    ]);
  });
});
