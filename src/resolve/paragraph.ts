// Paragraph and run resolution: the full cascade for one paragraph and its runs.
//
// Paragraph properties, lowest to highest priority:
//   docDefaults pPr
//   → [in a table] the default paragraph style's chain, then the table style (whole table +
//     the cell's conditional regions), then the rest of the paragraph style chain
//     [outside tables] the paragraph style chain
//   → the numbering level's pPr (just above the docDefaults/table style when the numbering
//     comes from the paragraph style, so the style's own indents win; just below direct
//     formatting when the paragraph sets `w:numPr` itself)
//   → direct `w:pPr`.
// Run properties follow the same order for the paragraph's base (`runDefaults`), then the run's
// character style chain, then direct `w:rPr`; toggles combine as described in rpr.ts. The
// paragraph mark (`markRun`) is a run whose rPr is `w:pPr/w:rPr`. The numbering level's rPr
// formats only the list label (see labels.ts).

import type { Story } from '../wml/story.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { DocContext } from './context.js';
import { cached, store } from './context.js';
import type { EffectiveLevel } from './numbering.js';
import { effectiveLevel } from './numbering.js';
import type { ParaLayer } from './ppr.js';
import { alignmentOf, finishBorders, finishTabs, mergePara, readPPr } from './ppr.js';
import type { RunLayer, RunState } from './rpr.js';
import { applyRun, finishRun, initialRunState, readRPr, rStyleOf } from './rpr.js';
import { cellStyleLayers, enclosingCell, tableInfo } from './table.js';
import type { ResolvedParagraphProps, ResolvedRunProps } from './types.js';

/** Everything resolved for one paragraph. */
export interface ParaResult {
  readonly props: ResolvedParagraphProps;
  /** Internal run state of the paragraph base (toggle layers already combined). */
  readonly runDefaults: RunState;
  /** Run state of the paragraph mark. */
  readonly markState: RunState;
  /** Effective list membership. */
  readonly list: ListMembership | undefined;
}

/** A paragraph's list membership. */
export interface ListMembership {
  readonly numId: number;
  readonly ilvl: number;
  readonly eff: EffectiveLevel;
  /** False when the numbering comes from a style that does not own the level (no label). */
  readonly counted: boolean;
}

/** Public props object → its internal result (lets resolveRun reuse a caller's props). */
const resultOf = new WeakMap<ResolvedParagraphProps, ParaResult>();

/** The internal result behind a props object returned by {@link paragraphResult}. */
export function resultForProps(props: ResolvedParagraphProps): ParaResult | undefined {
  return resultOf.get(props);
}

/** Resolves a paragraph (cached per element and story-part version). */
export function paragraphResult(ctx: DocContext, p: XElement, story: Story): ParaResult {
  const hit = cached(ctx.paragraphs, p);
  if (hit !== undefined) return hit;
  const pPrEl = p.element(NS.w, 'pPr');
  const direct = readPPr(ctx, pPrEl);
  const ps = ctx.paraStyle(pPrEl?.element(NS.w, 'pStyle')?.attrNS(NS.w, 'val'));

  let tblP: ParaLayer | undefined;
  let tblR: RunLayer | undefined;
  const cell = enclosingCell(p);
  if (cell !== undefined) {
    const layers = cellStyleLayers(ctx, tableInfo(ctx, cell.tbl, story), cell.tc);
    tblP = layers.pPr;
    tblR = layers.rPr;
  }

  // List membership: numId and ilvl merge field by field (style chain, then direct); an omitted
  // ilvl is level 0. Numbering that comes from a style chain without the level's owner applies
  // the level's indents but shows no label and is not counted (see numbering.ts).
  const numId = direct.numId ?? ps.pFull.numId;
  const fromDirect = direct.numId !== undefined || direct.ilvl !== undefined;
  let list: ParaResult['list'];
  if (numId !== undefined && numId > 0) {
    const ilvl = Math.min(8, Math.max(0, direct.ilvl ?? ps.pFull.ilvl ?? 0));
    const eff = effectiveLevel(ctx.numbering, ctx.styles, numId, ilvl, ctx.levels);
    if (eff !== undefined) {
      const owner = fromDirect ? undefined : ctx.levelOwner(eff.abstractId, ilvl);
      list = {
        numId,
        ilvl,
        eff,
        counted: owner === undefined || ps.chain.some((st) => st.id === owner),
      };
    }
  }
  let numP = list === undefined ? undefined : readPPr(ctx, list.eff.level.pPr);
  const linked = list?.eff.level.pStyle;
  if (
    numP !== undefined &&
    linked !== undefined &&
    numP.indStart === undefined &&
    numP.indEnd === undefined &&
    numP.firstLine === undefined &&
    ctx.styles.byId.get(linked)?.type === 'paragraph'
  ) {
    // A level stating no indent takes the one of the style it is linked to.
    const ls = ctx.paraStyle(linked).pFull;
    numP = { ...numP };
    if (ls.indStart !== undefined) numP.indStart = ls.indStart;
    if (ls.indEnd !== undefined) numP.indEnd = ls.indEnd;
    if (ls.firstLine !== undefined) numP.firstLine = ls.firstLine;
  }

  let pl: ParaLayer = ctx.docPPr;
  if (tblP !== undefined) {
    pl = mergePara(pl, ps.pLower);
    pl = mergePara(pl, tblP);
    if (numP !== undefined && !fromDirect) pl = mergePara(pl, numP);
    pl = mergePara(pl, ps.pUpper);
  } else {
    if (numP !== undefined && !fromDirect) pl = mergePara(pl, numP);
    pl = mergePara(pl, ps.pFull);
  }
  if (numP !== undefined && fromDirect) pl = mergePara(pl, numP);
  pl = mergePara(pl, direct);

  let rs = applyRun(initialRunState(), ctx.docRPr, 'abs');
  if (tblR !== undefined) {
    rs = applyRun(rs, ps.rLower, 'plain');
    rs = applyRun(rs, tblR, 'style');
    rs = applyRun(rs, ps.rUpper, 'plain');
    rs = applyRun(rs, ps.rFull, 'toggles');
  } else {
    rs = applyRun(rs, ps.rFull, 'style');
  }
  const markRPr = pPrEl?.element(NS.w, 'rPr');
  const markStyle = ctx.charStyle(rStyleOf(markRPr));
  let ms = applyRun(rs, markStyle.r, 'style');
  ms = applyRun(ms, readRPr(ctx, markRPr), 'abs');

  const bidi = pl.bidi ?? false;
  const start = pl.indStart ?? 0;
  const end = pl.indEnd ?? 0;
  const first = pl.firstLine ?? 0;
  const props: ResolvedParagraphProps = {
    alignment: alignmentOf(pl.jc, bidi),
    indent: {
      left: bidi ? end : start,
      right: bidi ? start : end,
      firstLine: Math.max(first, 0),
      hanging: first < 0 ? -first : 0,
    },
    spacing: {
      before: pl.before ?? 0,
      after: pl.after ?? 0,
      beforeAutospacing: pl.beforeAuto ?? false,
      afterAutospacing: pl.afterAuto ?? false,
      line: pl.line ?? 240,
      lineRule: pl.lineRule ?? 'auto',
    },
    contextualSpacing: pl.contextualSpacing ?? false,
    keepNext: pl.keepNext ?? false,
    keepLines: pl.keepLines ?? false,
    pageBreakBefore: pl.pageBreakBefore ?? false,
    // Word never writes widow control on, only `w:val="0"` to turn it off: it defaults on.
    widowControl: pl.widowControl ?? true,
    suppressAutoHyphens: pl.suppressAutoHyphens ?? false,
    bidi,
    tabs: finishTabs(pl.tabs),
    borders: finishBorders(pl.borders),
    runDefaults: finishRun(ctx, rs, undefined),
    markRun: finishRun(ctx, ms, markStyle.id),
  };
  if (ps.id !== undefined) props.styleId = ps.id;
  if (pl.outlineLvl !== undefined && pl.outlineLvl >= 0 && pl.outlineLvl <= 8) {
    props.outlineLevel = pl.outlineLvl;
  }
  if (list?.counted === true) props.numbering = { numId: list.numId, level: list.ilvl };
  if (pl.shd !== undefined && pl.shd !== 'none') props.shading = pl.shd;
  if (pl.framePr !== undefined) props.framePr = { ...pl.framePr };

  const result: ParaResult = { props, runDefaults: rs, markState: ms, list };
  resultOf.set(props, result);
  return store(ctx.paragraphs, p, result);
}

/** Resolves a run given its paragraph's result. */
export function runProps(ctx: DocContext, r: XElement, para: ParaResult): ResolvedRunProps {
  const rPr = r.element(NS.w, 'rPr');
  const cs = ctx.charStyle(rStyleOf(rPr));
  let s = applyRun(para.runDefaults, cs.r, 'style');
  s = applyRun(s, readRPr(ctx, rPr), 'abs');
  return finishRun(ctx, s, cs.id);
}

/** Label run props: the paragraph mark with the level's rPr applied as direct formatting. */
export function labelRunProps(
  ctx: DocContext,
  para: ParaResult,
  level: EffectiveLevel,
): ResolvedRunProps {
  const s = applyRun(para.markState, readRPr(ctx, level.level.rPr), 'abs');
  return finishRun(ctx, s, para.props.markRun.styleId);
}
