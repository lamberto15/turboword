// importStyles (src/compose/styles.ts): matching by id and name, the three style modes,
// docDefaults, dependency chains, table styles and the scope rule, on synthetic documents and on
// the docxcompose style fixtures (test/fixtures/compose-styles, MIT).
//
// The step runs against a minimal ImportContext (test/compose-styles-helpers.ts: real source and
// target documents, a simple part copier, fake ids), so it is tested independently of the
// orchestrator; the last block re-runs the fixtures through the real appendDocument.

import { describe, expect, it } from 'vitest';
import { appendDocument } from '../src/compose/index.js';
import { remapStyleRefs } from '../src/compose/styles.js';
import type { ImportOptions } from '../src/compose/types.js';
import { createDocx, type WordDocument, writeDocx } from '../src/index.js';
import {
  CT_BASE,
  DECL,
  DEFAULTS,
  DPF,
  firstRun,
  HEADING,
  loadFixture,
  makeDocx,
  NORMAL,
  P,
  para,
  pStyleOf,
  REL_BASE,
  runImport,
  savedErrors,
  style,
  styleById,
  styleIds,
  styleIntegrity,
  stylesRoot,
  TABLE_NORMAL,
  target,
  val,
  W,
  XMLNS,
} from './compose-styles-helpers.js';

// ---- tests --------------------------------------------------------------------------------

describe('importStyles: matching', () => {
  it('maps a localized built-in by name ("Überschrift1" named "heading 1" → Heading1)', () => {
    const src = makeDocx({
      styles: `${DEFAULTS()}${NORMAL('Standard')}${DPF('Absatz-Standardschriftart')}${HEADING('Überschrift1', '2F5496', '32', '<w:i/>', 'Standard')}`,
      body: P('Kapitel', 'Überschrift1') + P('Text'),
    });
    const tgt = target();
    const run = runImport(tgt, src);
    expect(run.ctx.styleMap.get('Überschrift1')).toBe('Heading1');
    expect(run.ctx.styleMap.get('Standard')).toBe('Normal');
    expect(pStyleOf(para(run, 'Kapitel'))).toBe('Heading1');
    expect(pStyleOf(para(run, 'Text'))).toBeUndefined(); // default style maps to the default
    expect(run.ctx.importedStyles).toEqual([]);
    expect(styleIds(tgt)).not.toContain('Überschrift1');
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt)).toEqual([]);
  });

  it('merges docxcompose styles_de into styles_en: built-ins by name, custom styles copied', () => {
    const tgt = loadFixture('styles_en.docx');
    const run = runImport(tgt, loadFixture('styles_de.docx'));
    const ids = styleIds(tgt);
    for (const id of ['Heading1', 'Heading2', 'Strong', 'Quote', 'QuoteChar', 'MyStyle1']) {
      expect(ids).toContain(id);
    }
    for (const id of ['berschrift1', 'berschrift2', 'Fett', 'Zitat', 'Standard']) {
      expect(ids).not.toContain(id);
    }
    expect(ids).toContain('MeineFormatvorlage');
    expect(ids).toContain('MeineFormatvorlageZchn');
    const used = run.blocks.flatMap((b) =>
      b.descendants(W).filter((e) => e.local === 'pStyle' || e.local === 'rStyle'),
    );
    expect(used.map((e) => e.attrNS(W, 'val')).sort()).toEqual(
      ['Heading1', 'Heading2', 'MeineFormatvorlage', 'Strong', 'ZitatZchn'].sort(),
    );
    // The copied linked character style keeps a mutual link only.
    expect(val(styleById(tgt, 'MeineFormatvorlageZchn'), 'link')).toBe('MeineFormatvorlage');
    expect(val(styleById(tgt, 'MeineFormatvorlage'), 'link')).toBe('MeineFormatvorlageZchn');
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'styles_en+de.docx')).toEqual([]);
  });

  it('prefers a name match over an id match whose name differs (docxcompose common_stylename)', () => {
    const tgt = loadFixture('common_stylename_different_id1.docx');
    const run = runImport(tgt, loadFixture('common_stylename_different_id2.docx'));
    // Source Aufzhlung10 is "Aufzählung 1", which the target calls Aufzhlung1.
    expect(run.ctx.styleMap.get('Aufzhlung10')).toBe('Aufzhlung1');
    expect(
      run.blocks.some((b) =>
        b.descendants(W, 'pStyle').some((e) => e.attrNS(W, 'val') === 'Aufzhlung10'),
      ),
    ).toBe(false);
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt)).toEqual([]);
  });

  it('treats styles with equal resolved properties as identical, whatever their chains', () => {
    // Target: color on Heading1 itself; source: color inherited through an intermediate style.
    const src = makeDocx({
      styles:
        `${DEFAULTS()}${NORMAL()}${DPF()}` +
        style(
          'paragraph',
          'Base',
          'Base',
          '<w:basedOn w:val="Normal"/><w:rsid w:val="00AB12CD"/><w:pPr><w:keepNext w:val="1"/><w:spacing w:before="240"/></w:pPr><w:rPr><w:color w:val="2f5496"/><w:i w:val="true"/></w:rPr>',
          ' w:customStyle="1"',
        ) +
        style(
          'paragraph',
          'Heading1',
          'heading 1',
          '<w:basedOn w:val="Base"/><w:pPr><w:spacing w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr>',
        ),
      body: P('Same', 'Heading1'),
    });
    const tgt = target();
    const run = runImport(tgt, src, { styles: 'keepDifferentStyles' });
    expect(run.ctx.styleMap.get('Heading1')).toBe('Heading1');
    expect(styleIds(tgt).filter((id) => id.startsWith('Heading1'))).toEqual(['Heading1']);
  });

  it('copies a style whose id belongs to another type under a new id and name', () => {
    const src = makeDocx({
      styles: `${DEFAULTS()}${NORMAL()}${DPF()}${style('paragraph', 'Strong', 'Strong', '<w:rPr><w:b/></w:rPr>')}`,
      body: P('Bold para', 'Strong'),
    });
    const tgt = target(style('character', 'Strong', 'Strong', '<w:rPr><w:b/></w:rPr>'));
    const run = runImport(tgt, src);
    expect(run.ctx.styleMap.get('Strong')).toBe('Strong_1');
    const copy = styleById(tgt, 'Strong_1')!;
    expect(copy.attrNS(W, 'type')).toBe('paragraph');
    expect(val(copy, 'name')).toBe('Strong_1');
    expect(pStyleOf(para(run, 'Bold para'))).toBe('Strong_1');
    expect(styleIntegrity(tgt)).toEqual([]);
  });

  it('leaves latentStyles alone and does not copy the source docDefaults', () => {
    const tgt = target();
    const before = stylesRoot(tgt).element(W, 'docDefaults')!.clone();
    runImport(
      tgt,
      makeDocx({
        styles: `${DEFAULTS('28', 'Arial')}<w:latentStyles w:defQFormat="0"/>${NORMAL()}${style('paragraph', 'X', 'X', '', ' w:customStyle="1"')}`,
        body: P('x', 'X'),
      }),
    );
    expect(stylesRoot(tgt).elements(W, 'latentStyles')).toHaveLength(0);
    expect(stylesRoot(tgt).element(W, 'docDefaults')!.text).toBe(before.text);
    expect(stylesRoot(tgt).elements(W, 'docDefaults')).toHaveLength(1);
  });
});

describe('importStyles: useDestinationStyles', () => {
  it('keeps the target definition of conflicting styles and copies the others', () => {
    const src = makeDocx({
      styles:
        `${DEFAULTS()}${NORMAL()}${DPF()}${HEADING('Heading1', 'FF0000', '40')}` +
        style(
          'paragraph',
          'MyStyle',
          'My Style',
          '<w:basedOn w:val="Normal"/><w:rPr><w:color w:val="00FF00"/></w:rPr>',
          ' w:customStyle="1"',
        ),
      body: P('Red heading', 'Heading1') + P('Green', 'MyStyle'),
    });
    const tgt = target();
    const before = styleById(tgt, 'Heading1')!.clone();
    const run = runImport(tgt, src);
    expect(run.ctx.styleMap.get('Heading1')).toBe('Heading1');
    expect(run.ctx.styleMap.get('MyStyle')).toBe('MyStyle');
    expect(styleById(tgt, 'Heading1')!.text).toBe(before.text);
    expect(val(styleById(tgt, 'Heading1')!.element(W, 'rPr'), 'color')).toBe('2F5496');
    expect(run.ctx.importedStyles.map((s) => s.attrNS(W, 'styleId'))).toEqual(['MyStyle']);
    expect(pStyleOf(para(run, 'Red heading'))).toBe('Heading1');
    // No direct formatting is added in this mode.
    expect(firstRun(para(run, 'Red heading')).element(W, 'rPr')).toBeUndefined();
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt)).toEqual([]);
  });
});

describe('importStyles: keepDifferentStyles', () => {
  const src = (): WordDocument =>
    makeDocx({
      styles:
        `${DEFAULTS()}${NORMAL('Normal', '<w:rPr><w:color w:val="333333"/></w:rPr>')}${DPF()}` +
        // a built-in heading without explicit outline level: the name gives level 0
        style(
          'paragraph',
          'Heading1',
          'heading 1',
          '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:color w:val="FF0000"/><w:sz w:val="40"/></w:rPr>',
        ) +
        style(
          'paragraph',
          'Heading2',
          'heading 2',
          '<w:basedOn w:val="Heading1"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr>',
        ),
      body: P('Red heading', 'Heading1') + P('Sub', 'Heading2') + P('Plain'),
    });

  it('copies conflicting styles renamed (Heading1_1), keeping outline levels and chains', () => {
    const tgt = target(
      style(
        'paragraph',
        'Heading2',
        'heading 2',
        '<w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:sz w:val="28"/></w:rPr>',
      ),
    );
    const run = runImport(tgt, src(), { styles: 'keepDifferentStyles' });
    expect(run.ctx.styleMap.get('Heading1')).toBe('Heading1_1');
    expect(run.ctx.styleMap.get('Normal')).toBe('Normal_1');
    const h1 = styleById(tgt, 'Heading1_1')!;
    expect(val(h1, 'name')).toBe('heading 1_1');
    expect(h1.attrNS(W, 'customStyle')).toBe('1');
    expect(h1.attrNS(W, 'default')).toBeUndefined();
    expect(val(h1.element(W, 'pPr'), 'outlineLvl')).toBe('0');
    expect(val(h1, 'basedOn')).toBe('Normal_1');
    expect(val(h1, 'next')).toBe('Normal_1');
    const h2 = styleById(tgt, 'Heading2_1')!;
    expect(val(h2, 'basedOn')).toBe('Heading1_1');
    expect(val(h2.element(W, 'pPr'), 'outlineLvl')).toBe('1');
    expect(styleById(tgt, 'Normal_1')!.attrNS(W, 'default')).toBeUndefined();
    expect(pStyleOf(para(run, 'Red heading'))).toBe('Heading1_1');
    // Unstyled source paragraphs used the source's (different) Normal: now explicit.
    expect(pStyleOf(para(run, 'Plain'))).toBe('Normal_1');
    // The target's own content is untouched.
    expect(pStyleOf(tgt.body.element.elements(W, 'p')[0]!)).toBe('Heading1');
    expect(pStyleOf(tgt.body.element.elements(W, 'p')[1]!)).toBeUndefined();
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'keep-different.docx')).toEqual([]);
  });

  it('reuses an identical renamed copy from an earlier import instead of adding _2', () => {
    const tgt = target();
    runImport(tgt, src(), { styles: 'keepDifferentStyles' });
    const once = styleIds(tgt);
    const run2 = runImport(tgt, src(), { styles: 'keepDifferentStyles' });
    expect(styleIds(tgt)).toEqual(once);
    expect(run2.ctx.styleMap.get('Heading1')).toBe('Heading1_1');
    expect(run2.ctx.importedStyles).toEqual([]);
  });

  it('docxcompose styles_preserve: MyCustomStyle_1, never duplicated on repeated appends', () => {
    const tgt = loadFixture('styles_preserve1.docx');
    const run = runImport(tgt, loadFixture('styles_preserve2.docx'), {
      styles: 'keepDifferentStyles',
    });
    expect(run.ctx.styleMap.get('MyCustomStyle')).toBe('MyCustomStyle_1');
    runImport(tgt, loadFixture('styles_preserve2.docx'), { styles: 'keepDifferentStyles' });
    const back = runImport(tgt, loadFixture('styles_preserve1.docx'), {
      styles: 'keepDifferentStyles',
    });
    expect(back.ctx.styleMap.get('MyCustomStyle')).toBe('MyCustomStyle');
    expect(styleIds(tgt).filter((id) => id.startsWith('MyCustomStyle'))).toEqual([
      'MyCustomStyle',
      'MyCustomStyleZchn',
      'MyCustomStyle_1',
      'MyCustomStyleZchn_1',
    ]);
    // The renamed pair links to each other.
    expect(val(styleById(tgt, 'MyCustomStyle_1'), 'link')).toBe('MyCustomStyleZchn_1');
    expect(val(styleById(tgt, 'MyCustomStyleZchn_1'), 'link')).toBe('MyCustomStyle_1');
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'styles_preserve.docx')).toEqual([]);
  });

  it('ignores conflicting same-id styles in useDestinationStyles (docxcompose)', () => {
    const tgt = loadFixture('styles_preserve1.docx');
    runImport(tgt, loadFixture('styles_preserve2.docx'));
    expect(styleIds(tgt)).not.toContain('MyCustomStyle_1');
  });
});

describe('importStyles: keepSourceFormatting', () => {
  const src = (sz = '22'): WordDocument =>
    makeDocx({
      styles:
        `${DEFAULTS(sz)}${NORMAL()}${DPF()}${HEADING('Heading1', 'FF0000', '40', '<w:b/>')}` +
        style(
          'paragraph',
          'Quote2',
          'Quote 2',
          '<w:basedOn w:val="Heading1"/><w:rPr><w:u w:val="single"/></w:rPr>',
          ' w:customStyle="1"',
        ) +
        style('character', 'Emph', 'Emph', '<w:rPr><w:caps/></w:rPr>', ' w:customStyle="1"'),
      body:
        P('Red heading', 'Heading1') +
        P('Own color', 'Heading1', '<w:color w:val="00FF00"/>') +
        P('Quoted', 'Quote2') +
        `<w:p><w:r><w:rPr><w:rStyle w:val="Emph"/></w:rPr><w:t>Caps</w:t></w:r></w:p>`,
    });

  it('expands conflicting styles into direct formatting on paragraphs and runs', () => {
    const tgt = target();
    const before = styleById(tgt, 'Heading1')!.clone();
    const run = runImport(tgt, src(), { styles: 'keepSourceFormatting' });
    expect(run.ctx.styleMap.get('Heading1')).toBe('Heading1');
    expect(styleById(tgt, 'Heading1')!.text).toBe(before.text);
    const p = para(run, 'Red heading');
    expect(pStyleOf(p)).toBe('Heading1');
    const rPr = firstRun(p).element(W, 'rPr')!;
    expect(val(rPr, 'color')).toBe('FF0000');
    expect(val(rPr, 'sz')).toBe('40');
    expect(rPr.element(W, 'b')).toBeDefined();
    // The target heading is italic, the source's is not: reset explicitly.
    expect(rPr.element(W, 'i')?.attrNS(W, 'val')).toBe('0');
    // The paragraph mark gets the same run formatting.
    expect(val(p.element(W, 'pPr')!.element(W, 'rPr'), 'color')).toBe('FF0000');
    // Direct formatting already there wins.
    expect(val(firstRun(para(run, 'Own color')).element(W, 'rPr'), 'color')).toBe('00FF00');
    // Schema order inside rPr: b, i, color, sz.
    expect(rPr.elements().map((e) => e.local)).toEqual(['b', 'i', 'color', 'sz']);
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'keep-source-formatting.docx')).toEqual([]);
  });

  it('copies non-conflicting styles (smartStyleBehavior) and materializes inherited differences', () => {
    const tgt = target();
    const run = runImport(tgt, src(), { styles: 'keepSourceFormatting' });
    const q = styleById(tgt, 'Quote2')!;
    expect(val(q, 'basedOn')).toBe('Heading1');
    // Quote2 inherits from the conflicting Heading1: the source's values are now its own.
    expect(val(q.element(W, 'rPr'), 'color')).toBe('FF0000');
    expect(val(q.element(W, 'rPr'), 'sz')).toBe('40');
    expect(q.element(W, 'rPr')!.element(W, 'i')?.attrNS(W, 'val')).toBe('0');
    expect(pStyleOf(para(run, 'Quoted'))).toBe('Quote2');
    expect(firstRun(para(run, 'Quoted')).element(W, 'rPr')).toBeUndefined();
    expect(styleById(tgt, 'Emph')).toBeDefined();
  });

  it('without smartStyleBehavior expands every paragraph/character style', () => {
    const tgt = target();
    const run = runImport(tgt, src(), {
      styles: 'keepSourceFormatting',
      smartStyleBehavior: false,
    });
    expect(styleById(tgt, 'Quote2')).toBeUndefined();
    expect(styleById(tgt, 'Emph')).toBeUndefined();
    const quoted = para(run, 'Quoted');
    expect(pStyleOf(quoted)).toBe('Normal');
    const rPr = firstRun(quoted).element(W, 'rPr')!;
    expect(val(rPr, 'u')).toBe('single');
    expect(val(rPr, 'color')).toBe('FF0000');
    expect(val(quoted.element(W, 'pPr'), 'outlineLvl')).toBe('0');
    const caps = firstRun(para(run, 'Caps')).element(W, 'rPr')!;
    expect(caps.element(W, 'rStyle')).toBeUndefined();
    expect(caps.element(W, 'caps')).toBeDefined();
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt)).toEqual([]);
  });

  it('expands differing docDefaults as direct formatting', () => {
    const tgt = target('', '22');
    const run = runImport(tgt, src('28'), { styles: 'keepSourceFormatting' });
    expect(val(firstRun(para(run, 'Caps')).element(W, 'rPr'), 'sz')).toBe('28');
    // Same docDefaults: nothing is added to plain runs.
    const tgt2 = target();
    const run2 = runImport(tgt2, src(), { styles: 'keepSourceFormatting' });
    const caps = firstRun(para(run2, 'Caps')).element(W, 'rPr')!;
    expect(caps.elements().map((e) => e.local)).toEqual(['rStyle']);
  });

  it('docxcompose styles_default: the differing Normal formatting is kept as direct formatting', () => {
    const tgt = loadFixture('styles_default1.docx');
    const run = runImport(tgt, loadFixture('styles_default2.docx'), {
      styles: 'keepSourceFormatting',
    });
    // Source Normal: Arial, accent6 green; source List Paragraph adds purple. The target's are
    // blue-ish and orange, with no font: every run gets the source's values directly.
    const runs = run.blocks.flatMap((b) => b.descendants(W, 'r'));
    expect(runs.length).toBeGreaterThan(5);
    for (const r of runs) {
      const rPr = r.element(W, 'rPr')!;
      const listItem = pStyleOf(r.closest(W, 'p')!) === 'Listenabsatz';
      expect(val(rPr, 'color')).toBe(listItem ? '7030A0' : '70AD47');
      expect(rPr.element(W, 'rFonts')?.attrNS(W, 'ascii')).toBe('Arial');
    }
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'styles_default.docx')).toEqual([]);
  });

  it('copies conflicting table styles renamed (they cannot be expanded)', () => {
    const tbl = (sz: string): string =>
      style(
        'table',
        'TableGrid',
        'Table Grid',
        `<w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="${sz}" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>`,
      );
    const tgt = target(tbl('4'));
    const source = makeDocx({
      styles: `${DEFAULTS()}${NORMAL()}${DPF()}${TABLE_NORMAL()}${tbl('12')}`,
      body: `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    });
    const run = runImport(tgt, source, { styles: 'keepSourceFormatting' });
    expect(run.ctx.styleMap.get('TableGrid')).toBe('TableGrid_1');
    expect(val(run.blocks[0]!.element(W, 'tblPr'), 'tblStyle')).toBe('TableGrid_1');
    expect(styleIntegrity(tgt)).toEqual([]);
  });
});

describe('importStyles: table styles', () => {
  it('copies a localized table style with its conditional formatting (docxcompose table.docx)', () => {
    const tgt = createDocx();
    const run = runImport(tgt, loadFixture('table.docx'));
    const id = 'Rastertabelle1hell-Akzent5';
    expect(run.ctx.styleMap.get(id)).toBe(id);
    const copy = styleById(tgt, id)!;
    // Its base, the German "NormaleTabelle", is the target's "Normal Table".
    expect(val(copy, 'basedOn')).toBe('TableNormal');
    expect(copy.elements(W, 'tblStylePr').length).toBeGreaterThan(0);
    expect(styleIds(tgt)).not.toContain('NormaleTabelle');
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'table.docx')).toEqual([]);
  });

  it('gives unstyled tables an explicit style when the default table styles differ', () => {
    const tgt = target();
    const source = makeDocx({
      styles: `${DEFAULTS()}${NORMAL()}${DPF()}${TABLE_NORMAL('TableNormal', '200')}`,
      body: `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
    });
    const run = runImport(tgt, source, { styles: 'keepDifferentStyles' });
    const tblPr = run.blocks[0]!.element(W, 'tblPr')!;
    expect(val(tblPr, 'tblStyle')).toBe('TableNormal_1');
    expect(tblPr.elements().map((e) => e.local)).toEqual(['tblStyle', 'tblW']);
    expect(styleIntegrity(tgt)).toEqual([]);
  });

  it('matches identical table styles and conditional formatting', () => {
    const t = style(
      'table',
      'Grid',
      'Grid Table 1',
      '<w:basedOn w:val="TableNormal"/><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>',
    );
    const tgt = target(t);
    const source = makeDocx({
      styles: `${DEFAULTS()}${NORMAL()}${TABLE_NORMAL()}${t.replace('<w:b/>', '<w:b w:val="1"/>')}`,
      body: `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
    });
    const run = runImport(tgt, source, { styles: 'keepDifferentStyles' });
    expect(run.ctx.styleMap.get('Grid')).toBe('Grid');
    const conflicting = makeDocx({
      styles: `${DEFAULTS()}${NORMAL()}${TABLE_NORMAL()}${t.replace('<w:b/>', '<w:i/>')}`,
      body: `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
    });
    const run2 = runImport(target(t), conflicting, { styles: 'keepDifferentStyles' });
    expect(run2.ctx.styleMap.get('Grid')).toBe('Grid_1');
  });
});

describe('importStyles: dependencies and scope', () => {
  it('copies basedOn / next / link chains and drops references to undefined styles', () => {
    const source = makeDocx({
      styles:
        `${DEFAULTS()}${NORMAL()}${DPF()}` +
        style(
          'paragraph',
          'A',
          'A',
          '<w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr>',
          ' w:customStyle="1"',
        ) +
        style(
          'paragraph',
          'B',
          'B',
          '<w:basedOn w:val="A"/><w:next w:val="Nowhere"/>',
          ' w:customStyle="1"',
        ) +
        style(
          'paragraph',
          'C',
          'C',
          '<w:basedOn w:val="B"/><w:next w:val="D"/><w:link w:val="CChar"/>',
          ' w:customStyle="1"',
        ) +
        style('paragraph', 'D', 'D', '<w:basedOn w:val="Normal"/>', ' w:customStyle="1"') +
        style(
          'character',
          'CChar',
          'C Char',
          '<w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="C"/>',
          ' w:customStyle="1"',
        ) +
        style('paragraph', 'Unused', 'Unused', '', ' w:customStyle="1"'),
      body: P('deep', 'C'),
    });
    const tgt = target();
    const run = runImport(tgt, source);
    expect(run.ctx.importedStyles.map((s) => s.attrNS(W, 'styleId'))).toEqual([
      'A',
      'B',
      'C',
      'D',
      'CChar',
    ]);
    expect(val(styleById(tgt, 'B'), 'next')).toBeUndefined();
    expect(val(styleById(tgt, 'C'), 'link')).toBe('CChar');
    expect(val(styleById(tgt, 'CChar'), 'link')).toBe('C');
    expect(styleById(tgt, 'Unused')).toBeUndefined();
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt)).toEqual([]);
  });

  it('imports styles used only in headers and footnotes (scope rule) and remaps later content', () => {
    const hdr = `${DECL}<w:hdr ${XMLNS}>${P('Header', 'HeaderStyle')}</w:hdr>`;
    const notes = `${DECL}<w:footnotes ${XMLNS}><w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="NoteText"/></w:pPr><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>`;
    const source = makeDocx({
      styles:
        `${DEFAULTS()}${NORMAL('Standard')}${DPF()}` +
        style(
          'paragraph',
          'HeaderStyle',
          'My Header',
          '<w:rPr><w:b/></w:rPr>',
          ' w:customStyle="1"',
        ) +
        style(
          'paragraph',
          'NoteText',
          'footnote text',
          '<w:basedOn w:val="Standard"/><w:rPr><w:sz w:val="16"/></w:rPr>',
        ),
      body: P('Body'),
      parts: [
        {
          name: 'word/header1.xml',
          contentType: `${CT_BASE}.header+xml`,
          xml: hdr,
          rel: `${REL_BASE}/header`,
        },
        {
          name: 'word/footnotes.xml',
          contentType: `${CT_BASE}.footnotes+xml`,
          xml: notes,
          rel: `${REL_BASE}/footnotes`,
        },
      ],
    });
    const tgt = target();
    const run = runImport(tgt, source);
    expect(run.ctx.styleMap.get('HeaderStyle')).toBe('HeaderStyle');
    expect(run.ctx.styleMap.get('NoteText')).toBe('NoteText');
    expect(val(styleById(tgt, 'NoteText'), 'basedOn')).toBe('Normal');
    // Content copied later (by the parts steps) goes through remapStyleRefs.
    const later = source.mainPart.relatedByType('footnotes')[0]!.xml().root.clone();
    remapStyleRefs(run.ctx, [later]);
    expect(later.descendants(W, 'pStyle')[0]!.attrNS(W, 'val')).toBe('NoteText');
  });

  it('imports list styles reached through numStyleLink (docxcompose custom_list_style)', () => {
    const tgt = loadFixture('master.docx');
    const run = runImport(tgt, loadFixture('custom_list_style.docx'));
    const s = styleById(tgt, 'Style1')!;
    expect(s.attrNS(W, 'type')).toBe('numbering');
    expect(run.ctx.importedStyles).toContain(s);
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'custom_list_style.docx')).toEqual([]);
  });

  it('creates the target styles part when it is missing', () => {
    const tgt = makeDocx({ body: P('no styles') });
    const run = runImport(tgt, loadFixture('styles_en.docx'));
    expect(tgt.mainPart.relatedByType('styles')).toHaveLength(1);
    expect(run.ctx.addedParts).toContain('word/styles.xml');
    expect(styleIds(tgt)).toContain('Heading1');
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt)).toEqual([]);
  });

  it('is deterministic', () => {
    const make = (): Uint8Array => {
      const tgt = loadFixture('styles_en.docx');
      runImport(tgt, loadFixture('styles_de.docx'), { styles: 'keepDifferentStyles' });
      return writeDocx(tgt, { zipDate: new Date(Date.UTC(2020, 0, 1)) });
    };
    expect(make()).toEqual(make());
  });
});

describe('importStyles through appendDocument', () => {
  const pairs: [string, string, NonNullable<ImportOptions['styles']>][] = [
    ['styles_en.docx', 'styles_de.docx', 'useDestinationStyles'],
    ['styles_de.docx', 'styles_en.docx', 'keepDifferentStyles'],
    ['styles_preserve1.docx', 'styles_preserve2.docx', 'keepDifferentStyles'],
    ['styles_default1.docx', 'styles_default2.docx', 'keepSourceFormatting'],
    [
      'common_stylename_different_id1.docx',
      'common_stylename_different_id2.docx',
      'useDestinationStyles',
    ],
    ['master.docx', 'table.docx', 'keepSourceFormatting'],
  ];
  for (const [t, s, mode] of pairs) {
    it(`${t} + ${s} (${mode})`, () => {
      const tgt = loadFixture(t);
      const result = appendDocument(tgt, loadFixture(s), { styles: mode });
      expect(result.styleMap.size).toBeGreaterThan(0);
      expect(styleIntegrity(tgt)).toEqual([]);
      expect(savedErrors(tgt, `append-${t.replace('.docx', '')}+${s}`)).toEqual([]);
    });
  }
});
