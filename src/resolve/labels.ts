// List counters and rendered labels.
//
// Counters run in document order over one *flow*, as Word numbers them: the body; all headers
// and footers together (headers first); all footnotes of the footnotes part; all endnotes; all
// comments; and each text box on its own (text-box paragraphs are not part of their story's
// flow). Within a flow, counters belong to the abstract definition (see numbering.ts):
// - the first paragraph of a num with `w:startOverride` restarts those levels;
// - a paragraph at level L advances counter L (or starts it at the level's start value);
// - deeper levels restart when level L is used, unless their `w:lvlRestart` says otherwise
//   (0 = never; n = only after a level up to n−1 is used);
// - `w:lvlText` placeholders %1-%9 show the counters of levels 0-8 in each level's own format
//   (decimal for the higher levels under `w:isLgl`); a higher level not used yet shows its start
//   value and counts as used, so its next paragraph shows start + 1.
// Counters are cached per flow and the versions of the parts holding it (the numbering and
// styles parts are covered by the context's own invalidation).

import type { Story } from '../wml/story.js';
import { blockParagraphs } from '../wml/text.js';
import type { XDocument, XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { DocContext } from './context.js';
import { formatNumberValue } from './format.js';
import type { EffectiveLevel } from './numbering.js';
import { paragraphResult } from './paragraph.js';

/** Counter state of one paragraph's label. */
export interface LabelEntry {
  readonly eff: EffectiveLevel;
  readonly ilvl: number;
  /** Counter values of levels 0..ilvl. */
  readonly counters: readonly number[];
  readonly label: string;
}

/** Labels of every list paragraph of one story. */
export interface StoryCounters {
  readonly entries: ReadonlyMap<XElement, LabelEntry>;
}

interface Group {
  readonly counters: number[];
  readonly started: boolean[];
}

/** True when using level `used` restarts level `level` (whose definition is `lvlRestart`). */
function restarts(lvlRestart: number | undefined, level: number, used: number): boolean {
  const r = lvlRestart ?? level;
  if (r <= 0) return false;
  return used < r;
}

/** Renders `w:lvlText` with the current counters. */
export function renderLabel(
  eff: EffectiveLevel,
  ilvl: number,
  counters: readonly number[],
): string {
  const text = eff.level.lvlText;
  if (text === undefined) return '';
  return text.replace(/%([1-9])/g, (_m, d: string) => {
    const k = Number(d) - 1;
    const def = eff.levels[k];
    if (def === undefined) return '';
    const value = counters[k] ?? def.start;
    const legal = eff.level.isLgl && k < ilvl && def.numFmt !== 'none' && def.numFmt !== 'bullet';
    return formatNumberValue(value, legal ? 'decimal' : def.numFmt, def.numFmtCustom);
  });
}

/** One counter flow: the containers whose paragraphs are counted together, in order. */
interface Flow {
  readonly key: object;
  readonly containers: readonly XElement[];
}

/** The flow a paragraph is counted in (see the file comment). */
function flowOf(ctx: DocContext, p: XElement, story: Story): Flow {
  const txbx = p.closest(NS.w, 'txbxContent');
  if (txbx !== undefined) return { key: txbx, containers: [txbx] };
  const doc = story.doc;
  switch (story.kind) {
    case 'header':
    case 'footer': {
      const hf = doc.headersFooters();
      const containers = [
        ...hf.filter((h) => h.kind === 'header'),
        ...hf.filter((h) => h.kind === 'footer'),
      ].map((h) => h.story.element);
      return {
        key: ctx.hfKey,
        containers: containers.includes(story.element) ? containers : [story.element],
      };
    }
    case 'footnote':
    case 'endnote':
    case 'comment': {
      const all =
        story.kind === 'footnote'
          ? doc.footnotes()
          : story.kind === 'endnote'
            ? doc.endnotes()
            : doc.comments();
      const containers = all.filter((s) => s.part === story.part).map((s) => s.element);
      const root = story.element.parent;
      const key = root !== null && root.nodeType === 1 ? root : story.element;
      return { key, containers: containers.includes(story.element) ? containers : [story.element] };
    }
    default:
      return { key: story.element, containers: [story.element] };
  }
}

function partsOf(containers: readonly XElement[]): XDocument[] {
  const out: XDocument[] = [];
  for (const c of containers) {
    const d = c.ownerDocument();
    if (d !== undefined && !out.includes(d)) out.push(d);
  }
  return out;
}

/** Labels of the flow `p` is counted in (computed once per flow and part versions). */
export function flowCounters(ctx: DocContext, p: XElement, story: Story): StoryCounters {
  const flow = flowOf(ctx, p, story);
  const docs = partsOf(flow.containers);
  const hit = ctx.flows.get(flow.key);
  if (
    hit !== undefined &&
    hit.docs.length === docs.length &&
    hit.docs.every((d, k) => d === docs[k] && d.version === hit.versions[k])
  ) {
    return hit.value;
  }
  const value = countFlow(ctx, flow.containers, story);
  ctx.flows.set(flow.key, { docs, versions: docs.map((d) => d.version), value });
  return value;
}

/** Runs the counters over the paragraphs of `containers`. */
function countFlow(ctx: DocContext, containers: readonly XElement[], story: Story): StoryCounters {
  const entries = new Map<XElement, LabelEntry>();
  const groups = new Map<number, Group>();
  const seenNums = new Set<number>();
  for (const container of containers) {
    // blockParagraphs skips drawings, so text-box paragraphs stay out of the story's flow.
    for (const p of blockParagraphs(container)) {
      const list = paragraphResult(ctx, p, story).list;
      if (list === undefined || !list.counted) continue;
      const { eff, ilvl } = list;
      let g = groups.get(eff.abstractId);
      if (g === undefined) {
        g = { counters: new Array<number>(9).fill(0), started: new Array<boolean>(9).fill(false) };
        groups.set(eff.abstractId, g);
      }
      if (!seenNums.has(eff.numId)) {
        seenNums.add(eff.numId);
        for (const k of eff.startOverrides) g.started[k] = false;
      }
      for (let m = ilvl + 1; m <= 8; m++) {
        if (g.started[m] && restarts(eff.levels[m]?.lvlRestart, m, ilvl)) g.started[m] = false;
      }
      if (g.started[ilvl]) g.counters[ilvl]!++;
      else {
        g.counters[ilvl] = eff.level.start;
        g.started[ilvl] = true;
      }
      // Higher levels not used yet show their start value and count as used from here on, so
      // the next paragraph at such a level shows start + 1.
      for (let k = 0; k < ilvl; k++) {
        if (!g.started[k]) {
          g.counters[k] = eff.levels[k]?.start ?? 0;
          g.started[k] = true;
        }
      }
      const counters = g.counters.slice(0, ilvl + 1);
      entries.set(p, { eff, ilvl, counters, label: renderLabel(eff, ilvl, counters) });
    }
  }
  return { entries };
}
