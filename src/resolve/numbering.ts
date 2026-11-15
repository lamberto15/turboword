// The numbering part: abstract definitions, instances, overrides and level ownership.
//
// A paragraph's `w:numId` names a `w:num` (a list instance), which points at a `w:abstractNum`
// (the level definitions). Indirections and overrides:
// - an abstractNum with `w:numStyleLink` takes its levels from the numbering style it names
//   (that style's `w:numPr/w:numId` → num → abstractNum, usually the one with `w:styleLink`);
// - a num's `w:lvlOverride` replaces a level (`w:lvl`) and/or restarts it (`w:startOverride`).
//
// Word's behaviour where the specification is silent (measured against Word 365 by the docx4j
// project's probes, and matching LibreOffice where LibreOffice implements it):
// - counters belong to the abstract definition: every num over it continues the same count, and
//   a num's `w:startOverride` restarts that count the first time the num is used;
// - levels are keyed by `w:ilvl`, a later definition replacing an earlier one; a `w:lvl` without
//   `w:ilvl` counts as an empty level 0 (no label);
// - a level belongs to one paragraph style: the one its `w:pStyle` names, else the first
//   paragraph style in styles.xml whose own `w:numPr` reaches it. A paragraph whose numbering
//   comes from its style gets a label only when the owner is in its style chain (the style
//   itself or one it is based on); direct `w:numPr` always gets one.

import type { XDocument, XElement } from '../xml/dom.js';
import { alternateContent, selectBranch } from '../xml/mc.js';
import { NS } from '../xml/ns.js';
import type { StyleIndex } from './styles.js';
import { styleChain } from './styles.js';
import { int, intVal, onOff, onOffValue, wAttr } from './values.js';

/** One numbering level (`w:lvl`). */
export interface LevelDef {
  readonly ilvl: number;
  readonly start: number;
  readonly numFmt: string;
  /** `w:numFmt/@w:format` (custom formats). */
  readonly numFmtCustom: string | undefined;
  /** `w:lvlText/@w:val`; undefined when absent or `w:null`. */
  readonly lvlText: string | undefined;
  /** `w:lvlRestart` (1-based level; 0 = never); undefined = restart after any higher level. */
  readonly lvlRestart: number | undefined;
  readonly pStyle: string | undefined;
  readonly isLgl: boolean;
  readonly suff: 'tab' | 'space' | 'nothing';
  readonly lvlJc: 'left' | 'center' | 'right';
  readonly pPr: XElement | undefined;
  readonly rPr: XElement | undefined;
  readonly picBullet: boolean;
}

/** A level slot: the `w:lvl`, or null for an empty level (a `w:lvl` without `w:ilvl`). */
type LevelSlot = XElement | null;

interface AbstractDef {
  readonly id: number;
  readonly levels: ReadonlyMap<number, LevelSlot>;
  readonly numStyleLink: string | undefined;
}

interface Override {
  start: number | undefined;
  lvl: XElement | undefined;
}

interface NumDef {
  readonly id: number;
  readonly abstractId: number;
  readonly overrides: ReadonlyMap<number, Override>;
}

/** Parsed numbering part. */
export interface NumberingIndex {
  readonly abstracts: ReadonlyMap<number, AbstractDef>;
  readonly nums: ReadonlyMap<number, NumDef>;
}

/** Effective definition of one list level of one list instance. */
export interface EffectiveLevel {
  readonly numId: number;
  /** Abstract definition actually used (after numStyleLink indirection); also the counter key. */
  readonly abstractId: number;
  readonly level: LevelDef;
  /** Every level of the instance (0-8) for %N placeholders and restarts. */
  readonly levels: readonly (LevelDef | undefined)[];
  /** Levels the num restarts (`w:startOverride`) the first time it is used. */
  readonly startOverrides: readonly number[];
}

/** Indexes a numbering part (empty when there is none). */
export function indexNumbering(doc: XDocument | undefined): NumberingIndex {
  const abstracts = new Map<number, AbstractDef>();
  const nums = new Map<number, NumDef>();
  let root: XElement | undefined;
  try {
    root = doc?.root;
  } catch {
    root = undefined;
  }
  if (root === undefined || !root.is(NS.w, 'numbering')) return { abstracts, nums };
  for (const el of root.elements(NS.w, 'abstractNum')) {
    const id = int(wAttr(el, 'abstractNumId'));
    if (id === undefined || abstracts.has(id)) continue;
    const levels = new Map<number, LevelSlot>();
    for (const lvl of el.elements(NS.w, 'lvl')) {
      const raw = wAttr(lvl, 'ilvl');
      if (raw === undefined) {
        levels.set(0, null);
        continue;
      }
      const ilvl = int(raw);
      if (ilvl !== undefined && ilvl >= 0 && ilvl <= 8) levels.set(ilvl, lvl);
    }
    const numStyleLink = el.element(NS.w, 'numStyleLink')?.attrNS(NS.w, 'val');
    abstracts.set(id, { id, levels, numStyleLink });
  }
  for (const el of root.elements(NS.w, 'num')) {
    const id = int(wAttr(el, 'numId'));
    const abstractId = intVal(el.element(NS.w, 'abstractNumId'));
    if (id === undefined || abstractId === undefined || nums.has(id)) continue;
    const overrides = new Map<number, Override>();
    for (const o of el.elements(NS.w, 'lvlOverride')) {
      const ilvl = int(wAttr(o, 'ilvl'));
      if (ilvl === undefined || ilvl < 0 || ilvl > 8) continue;
      const start = intVal(o.element(NS.w, 'startOverride'));
      const lvl = o.element(NS.w, 'lvl');
      if (start === undefined && lvl === undefined) continue;
      overrides.set(ilvl, { start, lvl });
    }
    nums.set(id, { id, abstractId, overrides });
  }
  return { abstracts, nums };
}

/** Maximum numStyleLink hops followed. */
const MAX_LINKS = 8;

/** The abstract definition of a num, following numStyleLink → numbering style → num → abstract. */
function abstractOf(
  idx: NumberingIndex,
  styles: StyleIndex,
  abstractId: number,
): AbstractDef | undefined {
  let abs = idx.abstracts.get(abstractId);
  const seen = new Set<number>();
  for (let hop = 0; abs?.numStyleLink !== undefined && hop < MAX_LINKS; hop++) {
    if (seen.has(abs.id)) break;
    seen.add(abs.id);
    const style = styles.byId.get(abs.numStyleLink);
    if (style === undefined || style.type !== 'numbering') break;
    const linkedNumId = intVal(
      style.el.element(NS.w, 'pPr')?.element(NS.w, 'numPr')?.element(NS.w, 'numId'),
    );
    const linkedNum = linkedNumId === undefined ? undefined : idx.nums.get(linkedNumId);
    const next = linkedNum === undefined ? undefined : idx.abstracts.get(linkedNum.abstractId);
    if (next === undefined || next === abs) break;
    abs = next;
  }
  return abs;
}

/** The resolved abstract definition id of list instance `numId`. */
export function abstractIdOf(
  idx: NumberingIndex,
  styles: StyleIndex,
  numId: number,
): number | undefined {
  const num = idx.nums.get(numId);
  return num === undefined ? undefined : abstractOf(idx, styles, num.abstractId)?.id;
}

/** Namespaces whose `mc:Choice` branches the level reader understands. */
const UNDERSTOOD: ReadonlySet<string> = new Set([NS.w, NS.w14, NS.w15]);

/** Child elements with `mc:AlternateContent` replaced by the branch Word 2010+ would read. */
function levelChildren(el: XElement): XElement[] {
  const out: XElement[] = [];
  for (const k of el.elements()) {
    if (k.is(NS.mc, 'AlternateContent')) {
      const branch = selectBranch(alternateContent(k), (uri) => UNDERSTOOD.has(uri));
      if (branch !== undefined) out.push(...branch.elements());
    } else out.push(k);
  }
  return out;
}

function readLevel(ilvl: number, el: XElement, base: LevelDef | undefined): LevelDef {
  const kids = levelChildren(el);
  const child = (local: string): XElement | undefined =>
    kids.find((k) => k.uri === NS.w && k.local === local);
  const startEl = child('start');
  const fmtEl = child('numFmt');
  const textEl = child('lvlText');
  const restartEl = child('lvlRestart');
  const suffEl = child('suff');
  const jcEl = child('lvlJc');
  const suffVal = wAttr(suffEl, 'val');
  const jcVal = wAttr(jcEl, 'val');
  const lvlText =
    textEl === undefined
      ? base?.lvlText
      : onOffValue(wAttr(textEl, 'null')) === true
        ? undefined
        : (wAttr(textEl, 'val') ?? '');
  return {
    ilvl,
    // ST_DecimalNumber default for w:start is 0 (§17.9.25).
    start: startEl === undefined ? (base?.start ?? 0) : (intVal(startEl) ?? 0),
    numFmt: fmtEl === undefined ? (base?.numFmt ?? 'decimal') : (wAttr(fmtEl, 'val') ?? 'decimal'),
    numFmtCustom: fmtEl === undefined ? base?.numFmtCustom : wAttr(fmtEl, 'format'),
    lvlText,
    lvlRestart: restartEl === undefined ? base?.lvlRestart : intVal(restartEl),
    pStyle: child('pStyle')?.attrNS(NS.w, 'val') ?? base?.pStyle,
    isLgl: child('isLgl') === undefined ? (base?.isLgl ?? false) : onOff(child('isLgl')) === true,
    suff:
      suffEl === undefined
        ? (base?.suff ?? 'tab')
        : suffVal === 'space' || suffVal === 'nothing'
          ? suffVal
          : 'tab',
    lvlJc:
      jcEl === undefined
        ? (base?.lvlJc ?? 'left')
        : jcVal === 'center'
          ? 'center'
          : jcVal === 'right' || jcVal === 'end'
            ? 'right'
            : 'left',
    pPr: child('pPr') ?? base?.pPr,
    rPr: child('rPr') ?? base?.rPr,
    picBullet: child('lvlPicBulletId') !== undefined || (base?.picBullet ?? false),
  };
}

/**
 * The effective level `ilvl` of list instance `numId` (undefined when the instance, its abstract
 * definition or the level does not exist, or the level is empty). Memoized in `memo`.
 */
export function effectiveLevel(
  idx: NumberingIndex,
  styles: StyleIndex,
  numId: number,
  ilvl: number,
  memo: Map<string, EffectiveLevel | undefined>,
): EffectiveLevel | undefined {
  const key = `${numId}:${ilvl}`;
  if (memo.has(key)) return memo.get(key);
  let out: EffectiveLevel | undefined;
  const num = idx.nums.get(numId);
  const abs = num === undefined ? undefined : abstractOf(idx, styles, num.abstractId);
  if (num !== undefined && abs !== undefined) {
    const levels: (LevelDef | undefined)[] = [];
    const startOverrides: number[] = [];
    for (let k = 0; k <= 8; k++) {
      const el = abs.levels.get(k);
      let def = el === undefined || el === null ? undefined : readLevel(k, el, undefined);
      const o = num.overrides.get(k);
      if (o?.lvl !== undefined) def = readLevel(k, o.lvl, def);
      if (o?.start !== undefined && def !== undefined) {
        def = { ...def, start: o.start };
        startOverrides.push(k);
      }
      levels.push(def);
    }
    const level = ilvl >= 0 && ilvl <= 8 ? levels[ilvl] : undefined;
    if (level !== undefined) out = { numId, abstractId: abs.id, level, levels, startOverrides };
  }
  memo.set(key, out);
  return out;
}

/**
 * The owner style of every claimed level, keyed `${abstractId}:${ilvl}`: the paragraph style a
 * level's `w:pStyle` names, else the first paragraph style (styles.xml order) whose own
 * `w:numPr` reaches the level (completed through basedOn; an omitted `w:ilvl` meaning level 0).
 */
export function levelOwners(idx: NumberingIndex, styles: StyleIndex): Map<string, string> {
  const owners = new Map<string, string>();
  for (const abs of idx.abstracts.values()) {
    for (const [ilvl, el] of abs.levels) {
      const ps = el?.element(NS.w, 'pStyle')?.attrNS(NS.w, 'val');
      if (ps !== undefined && styles.byId.get(ps)?.type === 'paragraph') {
        owners.set(`${abs.id}:${ilvl}`, ps);
      }
    }
  }
  for (const style of styles.byId.values()) {
    if (style.type !== 'paragraph') continue;
    if (style.el.element(NS.w, 'pPr')?.element(NS.w, 'numPr') === undefined) continue;
    // The style's numPr may name only the level and inherit the list (or the reverse).
    let numId: number | undefined;
    let ilvl: number | undefined;
    for (const s of styleChain(styles, style.id, 'paragraph')) {
      const numPr = s.el.element(NS.w, 'pPr')?.element(NS.w, 'numPr');
      numId ??= intVal(numPr?.element(NS.w, 'numId'));
      ilvl ??= intVal(numPr?.element(NS.w, 'ilvl'));
    }
    if (numId === undefined || numId <= 0) continue;
    const absId = abstractIdOf(idx, styles, numId);
    if (absId === undefined) continue;
    const key = `${absId}:${ilvl ?? 0}`;
    if (!owners.has(key)) owners.set(key, style.id);
  }
  return owners;
}
