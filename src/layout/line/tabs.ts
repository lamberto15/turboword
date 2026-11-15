// Tab stops, as Word resolves them.
//
// Positions are measured from the column's start edge (left for LTR paragraphs, right for RTL),
// the same origin as the indents. For a tab at pen position x:
// 1. the first custom stop after x (bar stops never stop a tab, they only draw a line);
// 2. on the first line of a paragraph with a hanging indent, the left indent itself acts as a stop
//    when it comes before that custom stop (this is how a list label's tab reaches the text);
// 3. otherwise the next default stop (multiples of `defaultTabStop` from the origin): Word clears
//    default stops left of the last custom stop, so defaults only apply past it.
// `num` stops (list tabs) behave as left stops.

import type { TabStop } from '../../resolve/types.js';

export type TabAlign = 'left' | 'center' | 'right' | 'decimal';
export type TabLeader = NonNullable<TabStop['leader']>;

/** A resolved stop for one tab. */
export interface ResolvedStop {
  pos: number;
  align: TabAlign;
  leader?: TabLeader;
}

interface Stop {
  pos: number;
  align: TabAlign;
  leader?: TabLeader;
}

/** A tiny tolerance so a pen exactly on a stop moves on to the next one. */
const EPS = 1e-6;

export class TabStops {
  readonly #stops: Stop[];
  readonly #bars: number[];
  readonly #dflt: number;
  readonly #hanging: number | undefined;

  /**
   * `stops` in points from the start edge; `hangingAt` = the start indent when the first line
   * hangs out of it (first-line offset < 0), else undefined.
   */
  constructor(
    stops: readonly { pos: number; val: TabStop['val']; leader?: TabLeader }[],
    defaultTab: number,
    hangingAt: number | undefined,
  ) {
    this.#stops = [];
    this.#bars = [];
    for (const s of stops) {
      if (s.val === 'clear') continue;
      if (s.val === 'bar') {
        this.#bars.push(s.pos);
        continue;
      }
      const align: TabAlign = s.val === 'num' ? 'left' : s.val;
      const st: Stop = { pos: s.pos, align };
      if (s.leader !== undefined) st.leader = s.leader;
      this.#stops.push(st);
    }
    this.#stops.sort((a, b) => a.pos - b.pos);
    this.#dflt = defaultTab > 0 ? defaultTab : 36;
    this.#hanging = hangingAt;
  }

  /** Positions of bar tabs (vertical lines). */
  get bars(): readonly number[] {
    return this.#bars;
  }

  /** The stop a tab at pen position `x` goes to. */
  next(x: number, firstLine: boolean): ResolvedStop {
    let custom: Stop | undefined;
    for (const s of this.#stops) {
      if (s.pos > x + EPS) {
        custom = s;
        break;
      }
    }
    const h = this.#hanging;
    if (
      firstLine &&
      h !== undefined &&
      h > x + EPS &&
      (custom === undefined || h < custom.pos - EPS)
    ) {
      return { pos: h, align: 'left' };
    }
    if (custom !== undefined) {
      const out: ResolvedStop = { pos: custom.pos, align: custom.align };
      if (custom.leader !== undefined) out.leader = custom.leader;
      return out;
    }
    const d = this.#dflt;
    return { pos: (Math.floor((x + EPS) / d) + 1) * d, align: 'left' };
  }
}
