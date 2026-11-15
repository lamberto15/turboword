// Greedy line breaking and horizontal positioning (in "start coordinates": x measured from the
// column's start edge, left for LTR and right for RTL paragraphs; mirrored at emission).
//
// Word fills lines greedily: content is added until it no longer fits, and the line ends at the
// last break opportunity (UAX #14) before that point; trailing spaces hang past the end and never
// cause a break. A word wider than an empty line is broken at the character that overflows.
// Line breaks end a line; so do page and column breaks.
//
// Justified paragraphs in Word 2013+ (compatibility mode 15) may shrink their inter-word spaces
// by up to JUSTIFY_SHRINK of their width to fit one more word on a line (LibreOffice emulates
// this for such documents too, tdf#119908): the line overflows by at most that much and emit.ts
// distributes the negative free space over the spaces.
//
// Tabs are sized while filling: a left stop moves the pen to the stop; right, centre and decimal
// stops keep a "pending" tab whose width shrinks as the text after it grows (right: the text ends
// at the stop; centre: it is centred on it; decimal: the decimal symbol sits on it) until the next
// tab or the end of the line. A right/centre/decimal stop past the right indent is honoured (Word
// lets such text run into the margin); a left stop past it takes what is left of the line.

import type { Atom } from './shape.js';
import type { TabAlign, TabLeader, TabStops } from './tabs.js';

/** Positioning tolerance, points. */
const EPS = 0.005;

/** Fraction of the inter-word space width a justified line may give up (Word 2013+). */
export const JUSTIFY_SHRINK = 0.2;

/** Horizontal extent of one line: pen start and right limit, start coordinates. */
export interface LineGeometry {
  start: number;
  end: number;
}

/** How to find the stop for a tab atom. */
export interface TabResolver {
  stops: TabStops;
  /** Explicit stop for a tab atom (positional tabs), or undefined to use the stops. */
  explicit(atom: number): { pos: number; align: TabAlign; leader?: TabLeader } | undefined;
}

/** The positions of the atoms of one line. */
export interface PlacedLine {
  from: number;
  to: number;
  geometry: LineGeometry;
  /** Start x of each atom (index - from). */
  xs: number[];
  /** Width of each atom (tabs sized). */
  ws: number[];
  /** Leader of each tab atom, by atom index. */
  leaders: Map<number, TabLeader>;
  /** Index of the last atom that is not hanging whitespace (from - 1 when none). */
  lastContent: number;
  /** Pen position after the last content atom. */
  contentEnd: number;
}

interface Pending {
  atom: number;
  x0: number;
  stop: number;
  align: TabAlign;
  seg: number;
  beforeDecimal: number | undefined;
}

/** Walks atoms left to right, sizing tabs; optionally records positions. */
class Filler {
  x: number;
  contentEnd: number;
  pending: Pending | undefined;
  /** Highest allowed content end (raised by right/centre/decimal stops past the line end). */
  limit: number;
  readonly firstLine: boolean;
  readonly tabs: TabResolver;
  readonly decimal: string;
  readonly text: string;
  readonly rec:
    | { from: number; xs: number[]; ws: number[]; leaders: Map<number, TabLeader> }
    | undefined;
  /** Shrinkable fraction of inner spaces (0 = no shrinking). */
  shrink = 0;
  /** Width of the spaces between content atoms since the last tab (shrinkable). */
  innerSpace = 0;
  /** Width of spaces after the last content atom (inner once more content follows). */
  pendingSpace = 0;

  constructor(
    g: LineGeometry,
    x0: number,
    firstLine: boolean,
    tabs: TabResolver,
    decimal: string,
    text: string,
    rec?: Filler['rec'],
  ) {
    this.x = x0;
    this.contentEnd = x0;
    this.pending = undefined;
    this.limit = g.end;
    this.firstLine = firstLine;
    this.tabs = tabs;
    this.decimal = decimal;
    this.text = text;
    this.rec = rec;
  }

  #tabWidth(p: Pending): number {
    switch (p.align) {
      case 'right':
        return Math.max(0, p.stop - p.x0 - p.seg);
      case 'center':
        return Math.max(0, p.stop - p.x0 - p.seg / 2);
      default:
        return Math.max(0, p.stop - p.x0 - (p.beforeDecimal ?? p.seg));
    }
  }

  /** Fixes the pending tab's width (shifting the atoms recorded after it). */
  finishPending(): void {
    const p = this.pending;
    if (p === undefined) return;
    const w = this.#tabWidth(p);
    const rec = this.rec;
    if (rec !== undefined) {
      rec.ws[p.atom - rec.from] = w;
      for (let k = p.atom - rec.from + 1; k < rec.xs.length; k++) rec.xs[k]! += w;
    }
    this.pending = undefined;
  }

  add(i: number, a: Atom): void {
    const rec = this.rec;
    if (a.kind === 'tab') {
      this.finishPending();
      this.innerSpace = 0;
      this.pendingSpace = 0;
      const stop = this.tabs.explicit(i) ?? this.tabs.stops.next(this.x, this.firstLine);
      if (rec !== undefined) {
        rec.xs.push(this.x);
        if (stop.leader !== undefined) rec.leaders.set(i, stop.leader);
      }
      if (stop.align === 'left' || stop.pos <= this.x) {
        // A left stop past the line end takes only what is left of the line.
        const target =
          stop.align === 'left' && stop.pos > this.limit
            ? Math.max(this.x, this.limit)
            : Math.max(this.x, stop.pos);
        rec?.ws.push(target - this.x);
        this.x = target;
      } else {
        rec?.ws.push(0);
        this.pending = {
          atom: i,
          x0: this.x,
          stop: stop.pos,
          align: stop.align,
          seg: 0,
          beforeDecimal: undefined,
        };
        if (stop.pos > this.limit) this.limit = stop.pos;
        this.x = Math.max(this.x, stop.pos);
      }
      this.contentEnd = this.x;
      return;
    }
    const p = this.pending;
    if (p !== undefined) {
      if (rec !== undefined) {
        rec.xs.push(p.x0 + p.seg);
        rec.ws.push(a.width);
      }
      if (
        p.align === 'decimal' &&
        p.beforeDecimal === undefined &&
        this.text.slice(a.start, a.end).includes(this.decimal)
      ) {
        p.beforeDecimal = p.seg;
      }
      p.seg += a.width;
      this.x = p.x0 + this.#tabWidth(p) + p.seg;
    } else {
      if (rec !== undefined) {
        rec.xs.push(this.x);
        rec.ws.push(a.width);
      }
      this.x += a.width;
    }
    if (a.space) this.pendingSpace += a.width;
    if (!a.white) {
      this.contentEnd = this.x;
      this.innerSpace += this.pendingSpace;
      this.pendingSpace = 0;
    }
  }

  overflows(): boolean {
    return this.contentEnd > this.limit + EPS + this.innerSpace * this.shrink;
  }
}

export interface BreakInput {
  atoms: readonly Atom[];
  /** UAX #14 opportunities over the paragraph text (1 = may break before, 2 = must). */
  opportunities: Uint8Array;
  text: string;
  geometry(line: number): LineGeometry;
  /** Pen offset of the first line relative to its start (a right/centre-aligned label starts before it). */
  firstLineShift: number;
  tabs: TabResolver;
  decimal: string;
  /** Shrinkable fraction of inter-word spaces (JUSTIFY_SHRINK for justified Word 2013+ text, else 0). */
  shrink?: number;
}

/** Breaks the atoms into lines; returns [from, to) atom ranges (at least one line). */
export function breakLines(input: BreakInput): { from: number; to: number }[] {
  const { atoms, opportunities: opp } = input;
  const n = atoms.length;
  const lines: { from: number; to: number }[] = [];
  let from = 0;
  while (from < n) {
    const line = lines.length;
    const g = input.geometry(line);
    const f = new Filler(
      g,
      g.start + (line === 0 ? input.firstLineShift : 0),
      line === 0,
      input.tabs,
      input.decimal,
      input.text,
    );
    f.shrink = input.shrink ?? 0;
    let to = n;
    let lastOpp = -1;
    for (let j = from; j < n; j++) {
      const a = atoms[j]!;
      if (j > from) {
        const o = opp[a.start] ?? 0;
        if (o === 2 && atoms[j - 1]!.kind !== 'break') {
          to = j;
          break;
        }
        if (o >= 1) lastOpp = j;
      }
      if (a.kind === 'break') {
        to = j + 1;
        break;
      }
      f.add(j, a);
      if (!a.white && f.overflows()) {
        if (lastOpp > from) to = lastOpp;
        else to = j > from ? j : j + 1;
        break;
      }
    }
    lines.push({ from, to });
    from = to;
  }
  // A final line break leaves the paragraph mark alone on a new (empty) line.
  const last = atoms[n - 1];
  if (n === 0 || last?.breakType === 'line') lines.push({ from: n, to: n });
  return lines;
}

/** Positions the atoms of line `[from, to)`. */
export function placeLine(input: BreakInput, line: number, from: number, to: number): PlacedLine {
  const g = input.geometry(line);
  const rec = {
    from,
    xs: [] as number[],
    ws: [] as number[],
    leaders: new Map<number, TabLeader>(),
  };
  const f = new Filler(
    g,
    g.start + (line === 0 ? input.firstLineShift : 0),
    line === 0,
    input.tabs,
    input.decimal,
    input.text,
    rec,
  );
  let lastContent = from - 1;
  for (let j = from; j < to; j++) {
    const a = input.atoms[j]!;
    if (a.kind === 'break') {
      rec.xs.push(f.x);
      rec.ws.push(0);
      continue;
    }
    f.add(j, a);
    if (!a.white) lastContent = j;
  }
  f.finishPending();
  let contentEnd = g.start + (line === 0 ? input.firstLineShift : 0);
  if (lastContent >= from) contentEnd = rec.xs[lastContent - from]! + rec.ws[lastContent - from]!;
  return {
    from,
    to,
    geometry: g,
    xs: rec.xs,
    ws: rec.ws,
    leaders: rec.leaders,
    lastContent,
    contentEnd,
  };
}
