// importNumbering (src/compose/numbering.ts): keepSource (fresh nsid), mergeWithDestination,
// restartListsAtImport, numId 0, list styles (numStyleLink/styleLink), styles' numPr, picture
// bullets (VML and DrawingML) and the scope rule, on synthetic documents and on the docxcompose
// numbering fixtures. Runs against the minimal ImportContext of test/compose-styles-helpers.ts;
// the last block goes through the real appendDocument.

import { describe, expect, it } from 'vitest';
import { appendDocument } from '../src/compose/index.js';
import { remapNumIds } from '../src/compose/numbering.js';
import { abstractFingerprint, NumberingIndex } from '../src/compose/styles-props.js';
import type { ImportOptions } from '../src/compose/types.js';
import { readDocx, type WordDocument, writeDocx } from '../src/index.js';
import type { XElement } from '../src/xml/dom.js';
import { NS } from '../src/xml/ns.js';
import {
  CT_BASE,
  DECL,
  DEFAULTS,
  DPF,
  loadFixture,
  makeDocx,
  NORMAL,
  P,
  para,
  REL_BASE,
  runImport,
  savedErrors,
  style,
  styleById,
  styleIntegrity,
  val,
  W,
  XMLNS,
} from './compose-styles-helpers.js';
import { fixture } from './helpers/fixtures.js';

// ---- builders -----------------------------------------------------------------------------

const LVL = (ilvl: number, fmt: string, start = '1', extra = ''): string =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="${start}"/><w:numFmt w:val="${fmt}"/>` +
  `<w:lvlText w:val="${fmt === 'bullet' ? '•' : `%${ilvl + 1}.`}"/><w:lvlJc w:val="left"/>${extra}` +
  `<w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
const DECIMAL = LVL(0, 'decimal') + LVL(1, 'lowerLetter') + LVL(2, 'lowerRoman', '3');
const BULLETS = LVL(0, 'bullet') + LVL(1, 'bullet');
const ABS = (id: number, nsid: string, lvls: string, extra = ''): string =>
  `<w:abstractNum w:abstractNumId="${id}"><w:nsid w:val="${nsid}"/>` +
  `<w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="0409000F"/>${extra}${lvls}</w:abstractNum>`;
const NUM = (id: number, abs: number, inner = ''): string =>
  `<w:num w:numId="${id}"><w:abstractNumId w:val="${abs}"/>${inner}</w:num>`;
const LI = (text: string, numId: number, ilvl = 0, pStyle?: string): string =>
  `<w:p><w:pPr>${pStyle ? `<w:pStyle w:val="${pStyle}"/>` : ''}<w:numPr><w:ilvl w:val="${ilvl}"/>` +
  `<w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const BASE_STYLES = `${DEFAULTS()}${NORMAL()}${DPF()}`;

/** A 1x1 PNG. */
const PNG = (seed: number): Uint8Array => {
  const b = Uint8Array.from(
    atob(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=',
    ),
    (c) => c.charCodeAt(0),
  );
  b[b.length - 13] = seed; // inside the IDAT CRC: distinct bytes, same size (dedupe is by content)
  return b;
};

function numberingRoot(doc: WordDocument): XElement {
  return doc.mainPart.relatedByType('numbering')[0]!.xml().root;
}

function index(doc: WordDocument): NumberingIndex {
  return new NumberingIndex(doc.mainPart.relatedByType('numbering')[0]);
}

function nsids(doc: WordDocument): string[] {
  return numberingRoot(doc)
    .elements(W, 'abstractNum')
    .map((a) => val(a, 'nsid') ?? '');
}

/** numIds of the numPr of `elements`, in document order. */
function numIdsIn(elements: readonly XElement[]): number[] {
  return elements.flatMap((e) =>
    [e, ...e.descendants(W, 'numPr')]
      .filter((n) => n.is(W, 'numPr'))
      .map((n) => Number(val(n, 'numId'))),
  );
}

/** Structural rules of a merged numbering part. */
function numberingIntegrity(doc: WordDocument): string[] {
  const problems: string[] = [];
  const root = numberingRoot(doc);
  const idx = index(doc);
  const order = root.elements(W).map((e) => e.local);
  const rank: Record<string, number> = {
    numPicBullet: 0,
    abstractNum: 1,
    num: 2,
    numIdMacAtCleanup: 3,
  };
  for (let i = 1; i < order.length; i++) {
    if ((rank[order[i]!] ?? 9) < (rank[order[i - 1]!] ?? 9))
      problems.push(`order: ${order[i]} after ${order[i - 1]}`);
  }
  const count = (local: string, attr: string): void => {
    const seen = new Set<string>();
    for (const e of root.elements(W, local)) {
      const v = e.attrNS(W, attr) ?? '';
      if (seen.has(v)) problems.push(`duplicate ${local} ${v}`);
      seen.add(v);
    }
  };
  count('num', 'numId');
  count('abstractNum', 'abstractNumId');
  count('numPicBullet', 'numPicBulletId');
  const ns = nsids(doc);
  if (new Set(ns).size !== ns.length) problems.push(`duplicate nsid in ${ns.join(',')}`);
  for (const [id, num] of idx.nums) {
    if (idx.abstractOfNum(id) === undefined)
      problems.push(`num ${id} → missing abstractNum ${idx.abstractIdOf(num)}`);
  }
  for (const pic of root.descendants(W, 'lvlPicBulletId')) {
    if (!idx.picBullets.has(Number(pic.attrNS(W, 'val'))))
      problems.push(`lvlPicBulletId ${pic.attrNS(W, 'val')}`);
  }
  const styles = doc.mainPart.relatedByType('styles')[0]?.xml().root;
  const styleType = (id: string): string | undefined =>
    styles
      ?.elements(W, 'style')
      .find((s) => s.attrNS(W, 'styleId') === id)
      ?.attrNS(W, 'type');
  for (const a of root.elements(W, 'abstractNum')) {
    for (const local of ['numStyleLink', 'styleLink']) {
      const v = val(a, local);
      if (v !== undefined && styleType(v) !== 'numbering')
        problems.push(`${local} ${v} is not a numbering style`);
    }
    for (const lvl of a.elements(W, 'lvl')) {
      const v = val(lvl, 'pStyle');
      if (v !== undefined && styleType(v) !== 'paragraph') problems.push(`lvl pStyle ${v}`);
    }
  }
  const refs = [doc.body.element, ...(styles ? [styles] : [])];
  for (const n of numIdsIn(refs)) {
    if (n !== 0 && !idx.nums.has(n)) problems.push(`numId ${n} does not resolve`);
  }
  return problems;
}

const LIST_DOC = (numbering: string, body: string, styles = BASE_STYLES): WordDocument =>
  makeDocx({ styles, numbering, body });

// ---- tests --------------------------------------------------------------------------------

describe('importNumbering: keepSource', () => {
  it('copies every list with new ids and a fresh nsid (docxcompose numberings.docx)', () => {
    const tgt = readDocx(fixture('docxcompose/numberings.docx').bytes);
    const src = readDocx(fixture('docxcompose/numberings.docx').bytes);
    const before = nsids(tgt);
    const run = runImport(tgt, src);
    expect([...run.ctx.numIdMap]).toEqual([
      [1, 4],
      [2, 5],
      [3, 6],
    ]);
    const after = nsids(tgt);
    expect(after).toHaveLength(6);
    expect(after.slice(0, 3)).toEqual(before);
    expect(new Set(after).size).toBe(6);
    // Copies define the same levels as their originals.
    const t = index(tgt);
    const s = index(src);
    for (const [from, to] of run.ctx.numIdMap) {
      expect(abstractFingerprint(t.abstractOfNum(to))).toBe(
        abstractFingerprint(s.abstractOfNum(from)),
      );
    }
    expect(numIdsIn(run.blocks).sort()).toEqual(
      [...Array(8).fill(4), ...Array(8).fill(5), 6, 6].sort(),
    );
    // The target's own paragraphs keep their lists.
    const own = tgt.body.element.elements(W, 'p').filter((p) => !run.blocks.includes(p));
    expect(new Set(numIdsIn(own))).toEqual(new Set([1, 2, 3]));
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'numberings+numberings.docx')).toEqual([]);
  });

  it('generates unique, deterministic nsids across repeated imports', () => {
    const build = (): WordDocument => {
      const tgt = readDocx(fixture('docxcompose/numberings.docx').bytes);
      for (let i = 0; i < 3; i++)
        runImport(tgt, readDocx(fixture('docxcompose/numberings.docx').bytes));
      return tgt;
    };
    const a = build();
    const ids = nsids(a);
    expect(ids).toHaveLength(12);
    expect(new Set(ids).size).toBe(12);
    for (const id of ids) expect(id).toMatch(/^[0-9A-F]{8}$/);
    expect(nsids(build())).toEqual(ids);
    const date = new Date(Date.UTC(2020, 0, 1));
    expect(writeDocx(build(), { zipDate: date })).toEqual(writeDocx(a, { zipDate: date }));
    expect(numberingIntegrity(a)).toEqual([]);
  });

  it('shares one target num between override-less nums of one list; overrides get their own', () => {
    const src = LIST_DOC(
      ABS(0, '11111111', DECIMAL) +
        NUM(1, 0) +
        NUM(2, 0) +
        NUM(3, 0, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>'),
      LI('one', 1) + LI('two', 2) + LI('five', 3),
    );
    const tgt = LIST_DOC(ABS(0, '22222222', BULLETS) + NUM(1, 0), LI('t', 1));
    const run = runImport(tgt, src);
    const m = run.ctx.numIdMap;
    expect(m.get(1)).toBe(m.get(2));
    expect(m.get(3)).not.toBe(m.get(1));
    const own = index(tgt).nums.get(m.get(3)!)!;
    expect(val(own.element(W, 'lvlOverride'), 'startOverride')).toBe('5');
    expect(index(tgt).abstractIdOf(own)).toBe(
      index(tgt).abstractIdOf(index(tgt).nums.get(m.get(1)!)!),
    );
    expect(numberingIntegrity(tgt)).toEqual([]);
  });

  it('keeps numId 0 ("no list") (docxcompose numbering_reference_to_numbering_zero)', () => {
    const tgt = loadFixture('numbering_reference_to_numbering_zero.docx');
    const run = runImport(tgt, loadFixture('numbering_reference_to_numbering_zero.docx'));
    // This fixture writes w:ilvl/w:numId straight into w:pPr (no w:numPr): left alone too.
    const zeros = tgt.body.element
      .descendants(W, 'numId')
      .filter((n) => n.attrNS(W, 'val') === '0');
    expect(zeros).toHaveLength(2);
    expect(run.ctx.numIdMap.has(0)).toBe(false);
    expect(numberingIntegrity(tgt)).toEqual([]);
  });

  it('maps numIds the source does not define to 0, with a warning', () => {
    const src = LIST_DOC(ABS(0, '11111111', DECIMAL) + NUM(1, 0), LI('ok', 1) + LI('dangling', 42));
    const tgt = LIST_DOC(ABS(0, '22222222', DECIMAL) + NUM(1, 0), LI('t', 1));
    const run = runImport(tgt, src);
    expect(run.ctx.numIdMap.get(42)).toBe(0);
    expect(val(para(run, 'dangling').element(W, 'pPr')!.element(W, 'numPr'), 'numId')).toBe('0');
    expect(run.warnings.some((w) => w.startsWith('NUMID_UNDEFINED'))).toBe(true);
  });

  it('creates the numbering part, its content type and relationship when missing', () => {
    const tgt = makeDocx({ styles: BASE_STYLES, body: P('plain') });
    const run = runImport(tgt, LIST_DOC(ABS(3, '11111111', DECIMAL) + NUM(7, 3), LI('a', 7)));
    const part = tgt.mainPart.relatedByType('numbering')[0]!;
    expect(part.name).toBe('word/numbering.xml');
    expect(part.contentType).toBe(`${CT_BASE}.numbering+xml`);
    expect(run.ctx.addedParts).toContain('word/numbering.xml');
    expect(run.ctx.numIdMap.get(7)).toBe(1);
    expect(index(tgt).abstractIdOf(index(tgt).nums.get(1)!)).toBe(0);
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt)).toEqual([]);
  });
});

describe('importNumbering: mergeWithDestination', () => {
  it('reuses a target list with the same levels and continues it', () => {
    const src = LIST_DOC(
      ABS(0, 'AAAAAAAA', DECIMAL) + ABS(1, 'BBBBBBBB', BULLETS) + NUM(1, 0) + NUM(2, 1),
      LI('a', 1) + LI('b', 2),
    );
    const tgt = LIST_DOC(
      ABS(5, 'CCCCCCCC', DECIMAL.replace('<w:start w:val="1"/>', '<w:start w:val="1"/>')) +
        NUM(9, 5),
      LI('t', 9),
    );
    const run = runImport(tgt, src, { lists: 'mergeWithDestination' });
    expect(run.ctx.numIdMap.get(1)).toBe(9); // continues the target's list
    const bullets = run.ctx.numIdMap.get(2)!;
    expect(bullets).not.toBe(9); // no matching target definition: copied
    expect(numberingRoot(tgt).elements(W, 'abstractNum')).toHaveLength(2);
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'merge-lists.docx')).toEqual([]);
  });

  it('reuses a target list linked to the same list style (numStyleLink)', () => {
    const listStyle = style(
      'numbering',
      'MyList',
      'My List',
      '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>',
      ' w:customStyle="1"',
    );
    const def = (nsid: string): string => ABS(0, nsid, DECIMAL, '<w:styleLink w:val="MyList"/>');
    const link = (id: number, nsid: string): string =>
      ABS(id, nsid, '', '<w:numStyleLink w:val="MyList"/>');
    const src = LIST_DOC(
      def('A1A1A1A1') + link(1, 'B1B1B1B1') + NUM(1, 0) + NUM(2, 1),
      LI('x', 2),
      BASE_STYLES + listStyle,
    );
    const tgt = LIST_DOC(
      def('C1C1C1C1') + link(1, 'D1D1D1D1') + NUM(1, 0) + NUM(2, 1),
      LI('t', 2),
      BASE_STYLES + listStyle,
    );
    const run = runImport(tgt, src, { lists: 'mergeWithDestination' });
    expect(run.ctx.numIdMap.get(2)).toBe(2);
    expect(numberingRoot(tgt).elements(W, 'abstractNum')).toHaveLength(2);
    expect(numberingIntegrity(tgt)).toEqual([]);
  });

  it('never merges into a list whose levels differ', () => {
    const src = LIST_DOC(ABS(0, 'AAAAAAAA', DECIMAL) + NUM(1, 0), LI('a', 1));
    const tgt = LIST_DOC(ABS(0, 'CCCCCCCC', DECIMAL.replace('%1.', '%1)')) + NUM(1, 0), LI('t', 1));
    const run = runImport(tgt, src, { lists: 'mergeWithDestination' });
    expect(run.ctx.numIdMap.get(1)).toBe(2);
    expect(numberingRoot(tgt).elements(W, 'abstractNum')).toHaveLength(2);
  });
});

describe('importNumbering: restartListsAtImport', () => {
  it('starts each imported list over with startOverride on counted levels only', () => {
    const src = LIST_DOC(
      ABS(0, 'AAAAAAAA', DECIMAL) + ABS(1, 'BBBBBBBB', BULLETS) + NUM(1, 0) + NUM(2, 0) + NUM(3, 1),
      LI('a', 1) + LI('b', 2) + LI('c', 3),
    );
    const tgt = LIST_DOC(ABS(0, 'CCCCCCCC', DECIMAL) + NUM(1, 0), LI('t', 1));
    const run = runImport(tgt, src, { restartListsAtImport: true });
    const num = index(tgt).nums.get(run.ctx.numIdMap.get(1)!)!;
    const overrides = num
      .elements(W, 'lvlOverride')
      .map((o) => [o.attrNS(W, 'ilvl'), val(o, 'startOverride')]);
    expect(overrides).toEqual([
      ['0', '1'],
      ['1', '1'],
      ['2', '3'],
    ]);
    // The second num of the same list continues the restarted one.
    expect(run.ctx.numIdMap.get(2)).toBe(run.ctx.numIdMap.get(1));
    // Bullets count nothing: no override.
    const bullets = index(tgt).nums.get(run.ctx.numIdMap.get(3)!)!;
    expect(bullets.elements(W, 'lvlOverride')).toHaveLength(0);
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt)).toEqual([]);
  });

  it('with mergeWithDestination: reuses the target definition through a new, restarting num', () => {
    const src = LIST_DOC(ABS(0, 'AAAAAAAA', DECIMAL) + NUM(1, 0), LI('a', 1));
    const tgt = LIST_DOC(ABS(0, 'CCCCCCCC', DECIMAL) + NUM(1, 0), LI('t', 1));
    const run = runImport(tgt, src, { lists: 'mergeWithDestination', restartListsAtImport: true });
    const n = run.ctx.numIdMap.get(1)!;
    expect(n).not.toBe(1);
    const num = index(tgt).nums.get(n)!;
    expect(index(tgt).abstractIdOf(num)).toBe(0);
    expect(val(num.element(W, 'lvlOverride'), 'startOverride')).toBe('1');
    expect(numberingRoot(tgt).elements(W, 'abstractNum')).toHaveLength(1);
  });

  it('keeps the source overrides and adds none to bullets (docxcompose numberings_restart)', () => {
    const tgt = loadFixture('numberings_restart.docx');
    const run = runImport(tgt, loadFixture('numberings_restart.docx'), {
      restartListsAtImport: true,
    });
    const idx = index(tgt);
    for (const to of new Set(run.ctx.numIdMap.values())) {
      const num = idx.nums.get(to)!;
      const def = idx.definitionOfNum(to, { get: () => undefined } as never);
      const fmts = def!.elements(W, 'lvl').map((l) => val(l, 'numFmt'));
      if (fmts.every((f) => f === 'bullet' || f === 'none')) {
        expect(num.descendants(W, 'startOverride')).toHaveLength(0);
      }
    }
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'numberings_restart.docx')).toEqual([]);
  });
});

describe('importNumbering: styles and list styles', () => {
  const listNumber = style(
    'paragraph',
    'ListNumber',
    'List Number',
    '<w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>',
  );
  const styledDoc = (nsid: string, body: string, extraNums = ''): WordDocument =>
    LIST_DOC(ABS(0, nsid, DECIMAL) + NUM(1, 0) + extraNums, body, BASE_STYLES + listNumber);

  it('joins a list numbering a style that keeps the target definition (one sequence)', () => {
    // Source: item 1 numbered directly (restarting num 2), item 2 through the style (num 1).
    const src = styledDoc(
      'AAAAAAAA',
      LI('first', 2, 0, 'ListNumber') + P('second', 'ListNumber'),
      NUM(2, 0, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'),
    );
    const tgt = styledDoc('CCCCCCCC', P('t1', 'ListNumber') + P('t2', 'ListNumber'));
    const run = runImport(tgt, src);
    const idx = index(tgt);
    const first = idx.nums.get(run.ctx.numIdMap.get(2)!)!;
    // Same abstractNum as the style's list (target num 1), restarting there.
    expect(idx.abstractIdOf(first)).toBe(idx.abstractIdOf(idx.nums.get(1)!));
    expect(val(first.element(W, 'lvlOverride'), 'startOverride')).toBe('1');
    expect(numberingRoot(tgt).elements(W, 'abstractNum')).toHaveLength(1);
    expect(numberingIntegrity(tgt)).toEqual([]);
  });

  it('restarts lists followed only through a target style when restartListsAtImport is set', () => {
    const src = styledDoc('AAAAAAAA', P('s1', 'ListNumber') + P('s2', 'ListNumber'));
    const plain = runImport(styledDoc('CCCCCCCC', P('t1', 'ListNumber')), src);
    expect(para(plain, 's1').element(W, 'pPr')!.element(W, 'numPr')).toBeUndefined();
    const tgt = styledDoc('CCCCCCCC', P('t1', 'ListNumber'));
    const run = runImport(tgt, src, { restartListsAtImport: true });
    const numPr = para(run, 's1').element(W, 'pPr')!.element(W, 'numPr')!;
    const num = index(tgt).nums.get(Number(val(numPr, 'numId')))!;
    expect(val(num.element(W, 'lvlOverride'), 'startOverride')).toBe('1');
    expect(index(tgt).abstractIdOf(num)).toBe(0);
    // Only the first paragraph: the next one continues the restarted list through the style.
    expect(para(run, 's2').element(W, 'pPr')!.element(W, 'numPr')).toBeUndefined();
    expect(
      para(run, 's1')
        .element(W, 'pPr')!
        .elements()
        .map((e) => e.local),
    ).toEqual(['pStyle', 'numPr']);
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt)).toEqual([]);
  });

  it('keeps a restarted docxcompose list in one sequence (numberings_restart)', () => {
    const tgt = loadFixture('numberings_restart.docx');
    const run = runImport(tgt, loadFixture('numberings_restart.docx'));
    const idx = index(tgt);
    // Source num 11 restarts the list of the "Listennummer" style (num 5): still that list.
    expect(idx.abstractIdOf(idx.nums.get(run.ctx.numIdMap.get(11)!)!)).toBe(
      idx.abstractIdOf(idx.nums.get(5)!),
    );
    expect(numberingIntegrity(tgt)).toEqual([]);
  });

  it('renumbers the numPr of copied styles (a list reached only through a style)', () => {
    const src = LIST_DOC(
      ABS(0, 'AAAAAAAA', DECIMAL, '') + NUM(1, 0),
      P('item', 'ListNumber'),
      BASE_STYLES +
        style(
          'paragraph',
          'ListNumber',
          'List Number',
          '<w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>',
        ),
    );
    const tgt = LIST_DOC(ABS(0, 'CCCCCCCC', BULLETS) + NUM(1, 0), LI('t', 1));
    const run = runImport(tgt, src);
    const copy = styleById(tgt, 'ListNumber')!;
    const n = Number(val(copy.element(W, 'pPr')!.element(W, 'numPr'), 'numId'));
    expect(n).toBe(run.ctx.numIdMap.get(1));
    expect(n).not.toBe(1);
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(styleIntegrity(tgt)).toEqual([]);
  });

  it('copies a numStyleLink chain consistently (docxcompose custom_list_style into master)', () => {
    const tgt = loadFixture('master.docx');
    const run = runImport(tgt, loadFixture('custom_list_style.docx'));
    const idx = index(tgt);
    const used = new Set(numIdsIn(run.blocks));
    expect(used.size).toBe(1);
    const abs = idx.abstractOfNum([...used][0]);
    expect(val(abs, 'numStyleLink')).toBe('Style1');
    const s = styleById(tgt, 'Style1')!;
    const defNum = Number(val(s.element(W, 'pPr')!.element(W, 'numPr'), 'numId'));
    expect(val(idx.abstractOfNum(defNum), 'styleLink')).toBe('Style1');
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'master+custom_list_style.docx')).toEqual([]);
  });

  it("reuses the target's definition of a list style that already exists (styleLink)", () => {
    const listStyle = style(
      'numbering',
      'MyList',
      'My List',
      '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>',
      ' w:customStyle="1"',
    );
    const def = (nsid: string): string => ABS(0, nsid, DECIMAL, '<w:styleLink w:val="MyList"/>');
    const src = LIST_DOC(def('A1A1A1A1') + NUM(1, 0), LI('direct', 1), BASE_STYLES + listStyle);
    const tgt = LIST_DOC(def('C1C1C1C1') + NUM(1, 0), LI('t', 1), BASE_STYLES + listStyle);
    const run = runImport(tgt, src);
    expect(run.ctx.numIdMap.get(1)).toBe(1);
    const links = numberingRoot(tgt)
      .elements(W, 'abstractNum')
      .filter((a) => val(a, 'styleLink') === 'MyList');
    expect(links).toHaveLength(1);
    expect(numberingIntegrity(tgt)).toEqual([]);
  });

  it('maps level pStyle links through the style map (docxcompose numberings_styles)', () => {
    const tgt = loadFixture('master.docx');
    const run = runImport(tgt, loadFixture('numberings_styles.docx'));
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(styleIntegrity(tgt)).toEqual([]);
    expect(run.ctx.numIdMap.size).toBeGreaterThan(0);
    expect(savedErrors(tgt, 'master+numberings_styles.docx')).toEqual([]);
  });

  it('imports the lists keepSourceFormatting writes as direct formatting', () => {
    const lp = (inner: string): string =>
      style('paragraph', 'ListPara', 'List Paragraph', `<w:basedOn w:val="Normal"/>${inner}`);
    const src = LIST_DOC(
      ABS(0, 'AAAAAAAA', DECIMAL) + NUM(4, 0),
      P('styled', 'ListPara'),
      BASE_STYLES +
        lp('<w:pPr><w:numPr><w:numId w:val="4"/></w:numPr><w:ind w:left="720"/></w:pPr>'),
    );
    const tgt = LIST_DOC(
      ABS(0, 'CCCCCCCC', BULLETS) + NUM(1, 0),
      LI('t', 1),
      BASE_STYLES + lp('<w:pPr><w:ind w:left="1440"/></w:pPr>'),
    );
    const run = runImport(tgt, src, { styles: 'keepSourceFormatting' });
    const pPr = para(run, 'styled').element(W, 'pPr')!;
    expect(val(pPr, 'pStyle')).toBe('ListPara');
    const n = Number(val(pPr.element(W, 'numPr'), 'numId'));
    expect(n).toBe(run.ctx.numIdMap.get(4));
    expect(index(tgt).nums.has(n)).toBe(true);
    expect(pPr.element(W, 'ind')!.attrNS(W, 'left')).toBe('720');
    expect(numberingIntegrity(tgt)).toEqual([]);
  });
});

describe('importNumbering: picture bullets', () => {
  const VML = (rid: string): string =>
    `<w:numPicBullet w:numPicBulletId="0" xmlns:v="${NS.v}" xmlns:o="${NS.o}"><w:pict><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:9pt;height:9pt" o:bullet="t"><v:imagedata r:id="${rid}" o:title=""/></v:shape></w:pict></w:numPicBullet>`;
  const DRAWING = (rid: string): string =>
    `<w:numPicBullet w:numPicBulletId="0" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="114300" cy="114300"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="b.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="114300" cy="114300"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:numPicBullet>`;
  const PIC_LVL = LVL(0, 'bullet', '1', '<w:lvlPicBulletId w:val="0"/>');
  const rels = (target: string): string =>
    `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_BASE}/image" Target="${target}"/></Relationships>`;
  const picDoc = (bullet: string, seed: number, text: string): WordDocument =>
    makeDocx({
      styles: BASE_STYLES,
      numbering: bullet + ABS(0, seed === 1 ? '11111111' : '22222222', PIC_LVL) + NUM(1, 0),
      body: LI(text, 1),
      files: {
        'word/_rels/numbering.xml.rels': rels('media/image1.png'),
        'word/media/image1.png': PNG(seed),
      },
    });

  it("copies VML picture bullets with the numbering part's own relationships", () => {
    const tgt = picDoc(VML('rId1'), 2, 'target');
    const src = picDoc(VML('rId1'), 1, 'source');
    const run = runImport(tgt, src);
    const root = numberingRoot(tgt);
    const pics = root.elements(W, 'numPicBullet');
    expect(pics.map((p) => p.attrNS(W, 'numPicBulletId'))).toEqual(['0', '1']);
    const abs = index(tgt).abstractOfNum(run.ctx.numIdMap.get(1));
    expect(val(abs!.element(W, 'lvl'), 'lvlPicBulletId')).toBe('1');
    const rid = pics[1]!.descendants(NS.v, 'imagedata')[0]!.attrNS(NS.r, 'id')!;
    const part = tgt.mainPart.relatedByType('numbering')[0]!;
    const image = part.related(rid)!;
    expect(image.bytes()).toEqual(PNG(1));
    expect(
      part.related(pics[0]!.descendants(NS.v, 'imagedata')[0]!.attrNS(NS.r, 'id')!)!.bytes(),
    ).toEqual(PNG(2));
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'picture-bullets-vml.docx')).toEqual([]);
  });

  it('copies DrawingML picture bullets and renumbers their drawing ids', () => {
    const tgt = makeDocx({ styles: BASE_STYLES, body: P('no lists') });
    const run = runImport(tgt, picDoc(DRAWING('rId1'), 1, 'source'));
    const pic = numberingRoot(tgt).element(W, 'numPicBullet')!;
    expect(pic.attrNS(W, 'numPicBulletId')).toBe('0');
    expect(Number(pic.descendants(NS.wp, 'docPr')[0]!.attr('id'))).toBeGreaterThan(1000);
    const rid = pic.descendants(NS.a, 'blip')[0]!.attrNS(NS.r, 'embed')!;
    expect(tgt.mainPart.relatedByType('numbering')[0]!.related(rid)!.bytes()).toEqual(PNG(1));
    expect(run.ctx.numIdMap.get(1)).toBe(1);
    expect(numberingIntegrity(tgt)).toEqual([]);
    expect(savedErrors(tgt, 'picture-bullets-drawing.docx')).toEqual([]);
  });

  it('never merges picture-bullet lists by fingerprint', () => {
    const tgt = picDoc(VML('rId1'), 2, 'target');
    const run = runImport(tgt, picDoc(VML('rId1'), 1, 'source'), { lists: 'mergeWithDestination' });
    expect(run.ctx.numIdMap.get(1)).not.toBe(1);
  });
});

describe('importNumbering: scope', () => {
  it('imports lists used only in a header and remaps later content', () => {
    const hdr = `${DECL}<w:hdr ${XMLNS}>${LI('header item', 2)}</w:hdr>`;
    const src = makeDocx({
      styles: BASE_STYLES,
      numbering: ABS(0, 'AAAAAAAA', DECIMAL) + ABS(1, 'BBBBBBBB', BULLETS) + NUM(1, 0) + NUM(2, 1),
      body: LI('body item', 1),
      parts: [
        {
          name: 'word/header1.xml',
          contentType: `${CT_BASE}.header+xml`,
          xml: hdr,
          rel: `${REL_BASE}/header`,
        },
      ],
    });
    const tgt = LIST_DOC(ABS(0, 'CCCCCCCC', DECIMAL) + NUM(1, 0), LI('t', 1));
    const run = runImport(tgt, src);
    expect(run.ctx.numIdMap.has(2)).toBe(true);
    const later = src.mainPart.relatedByType('header')[0]!.xml().root.clone();
    remapNumIds(run.ctx, [later]);
    expect(numIdsIn([later])).toEqual([run.ctx.numIdMap.get(2)]);
    expect(numberingIntegrity(tgt)).toEqual([]);
  });
});

describe('importNumbering through appendDocument', () => {
  const cases: [string, string, ImportOptions][] = [
    ['numberings_restart.docx', 'numberings_restart.docx', {}],
    ['numberings_styles.docx', 'numberings_styles.docx', { lists: 'mergeWithDestination' }],
    ['master.docx', 'numbering_with_paragraphs_in_between.docx', { restartListsAtImport: true }],
    ['master.docx', 'custom_list_style.docx', {}],
    ['broken_listing.docx', 'switched_listing_style.docx', { styles: 'keepDifferentStyles' }],
    ['fail_duplicate_list.docx', 'fail_duplicate_list.docx', { styles: 'keepSourceFormatting' }],
  ];
  for (const [t, s, o] of cases) {
    it(`${t} + ${s} ${JSON.stringify(o)}`, () => {
      const tgt = loadFixture(t);
      const known = styleIntegrity(tgt); // problems the input already has
      appendDocument(tgt, loadFixture(s), o);
      expect(numberingIntegrity(tgt)).toEqual([]);
      expect(styleIntegrity(tgt).filter((p) => !known.includes(p))).toEqual([]);
      expect(savedErrors(tgt, `append-${t.replace('.docx', '')}+${s}`)).toEqual([]);
    });
  }
});
