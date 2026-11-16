// updateFields: recompute every field's result in place, as Word's "Update field" (F9) does,
// without touching field codes.
//
// Order: stories in `doc.stories()` order, fields in document order. A field's instruction is
// evaluated after the fields nested in it (their new results are what it reads); fields nested
// in a result are updated only when their parent is not (an updated parent replaces its result,
// nested fields included). Locked fields (`w:fldLock`) and everything inside them stay as they
// are. A parent whose nested MERGEFIELD/FILLIN/ASK lacked data is skipped too, so a template
// without data never turns into "IF  = ..." nonsense.
// Fields inside the `mc:AlternateContent` branch readers do not show are updated alongside
// (Word reads the other branch), but only the shown ones are counted and reported.

import type { WordDocument } from '../wml/document.js';
import type { Story } from '../wml/story.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { DocInfo } from './docinfo.js';
import {
  BUILTINS,
  KNOWN_TYPES,
  type MergeState,
  type Outcome,
  type UpdateState,
} from './evaluate.js';
import { toField, toFields } from './model.js';
import { attached, joinPieces, onOff, type Scan, scanStory } from './scan.js';
import { setUpdateFieldsOnOpen } from './settings.js';
import { rebuildToc } from './toc.js';
import { findSwitch, hasSwitch, type ParsedInstruction, parseInstruction } from './tokenize.js';
import type {
  EvalContext,
  Field,
  FieldEvaluator,
  PageInfoProvider,
  TocOptions,
  UpdateFieldsOptions,
  UpdateFieldsResult,
} from './types.js';
import { clearDirty, templateProps, writeResult } from './write.js';

interface Run {
  st: UpdateState;
  types: Set<string> | undefined;
  custom: Readonly<Record<string, FieldEvaluator>>;
  scans: Map<Story, Scan[]>;
  /** New results of updated fields; current results of the others (set when done). */
  current: Map<Scan, string>;
  updated: number;
  skipped: Field[];
  /** Fields whose result was evaluated (written, or evaluated without a result: SET, NEXT...). */
  done: Set<Scan>;
  toc: { pages?: PageInfoProvider | undefined; levels?: [number, number] | undefined };
}

function makeState(doc: WordDocument, o: UpdateFieldsOptions): UpdateState {
  return {
    info: new DocInfo(doc),
    values: o.values ?? {},
    fmt: { locale: o.locale ?? 'en-US', timeZone: o.timeZone },
    now: o.now,
    pages: o.pages,
    overlay: new Map(),
    seq: new Map(),
    noData: false,
  };
}

/** Instruction text with nested fields read as their current (possibly new) results. */
function instructionOf(run: Run, s: Scan): string {
  if (s.kind === 'simple') return s.instrText;
  return joinPieces(s.instr, (c) => run.current.get(c) ?? c.resultText);
}

/** The bookmark-name type of an implicit REF (`{ MyBookmark }`), else the parsed type. */
function effectiveType(run: Run, p: ParsedInstruction): string {
  if (p.type === '' || KNOWN_TYPES.has(p.type) || run.custom[p.type] !== undefined) return p.type;
  const t = p.type;
  if (/^[\p{L}_][\p{L}\p{N}_]*$/u.test(t)) {
    if (run.st.overlay.has(t.toLowerCase()) || run.st.info.bookmark(t) !== undefined) return 'REF';
  }
  return p.type;
}

/** Evaluation context handed to custom evaluators. */
function evalContext(run: Run, s: Scan): EvalContext {
  const st = run.st;
  const ctx: EvalContext = {
    values: st.values,
    now: st.now ?? new Date(Number.NaN),
    locale: st.fmt.locale,
    evaluate: (instruction: string) =>
      evaluateText(run, s, parseInstruction(instruction), instruction),
  };
  if (st.fmt.timeZone !== undefined) ctx.timeZone = st.fmt.timeZone;
  if (st.pages !== undefined) ctx.pages = st.pages;
  return ctx;
}

/** Evaluates a parsed instruction for field `s` (custom evaluators first). */
function evaluate(
  run: Run,
  s: Scan,
  p: ParsedInstruction,
  instruction: string,
): Outcome | undefined {
  const type = effectiveType(run, p);
  const custom = run.custom[type];
  if (custom !== undefined) {
    const text = custom(toField(s, instruction), evalContext(run, s));
    return text === undefined ? undefined : { text };
  }
  const builtin = BUILTINS[type];
  if (builtin === undefined) return undefined;
  const q: ParsedInstruction =
    type === 'REF' && p.type !== 'REF'
      ? { type: 'REF', args: [{ text: p.type, quoted: false }, ...p.args], switches: p.switches }
      : p;
  const out = builtin(s, q, run.st);
  return typeof out === 'string' ? { text: out } : out;
}

function evaluateText(
  run: Run,
  s: Scan,
  p: ParsedInstruction,
  instruction: string,
): string | undefined {
  return evaluate(run, s, p, instruction)?.text;
}

/** SEQ values of a story, in document order, with `\s` heading resets. */
function computeSeq(st: UpdateState, story: Story, scans: readonly Scan[]): void {
  const info = st.info;
  const order = info.paragraphOrder(story);
  const headings: { at: number; level: number }[] = [];
  for (const [p, at] of order) {
    const level = info.headingLevel(p, true);
    if (level !== undefined) headings.push({ at, level });
  }
  const counters = new Map<string, { value: number; last: number }>();
  const twins = new Map<XElement, Scan[]>();
  for (const s of scans) {
    if (s.parsed.type !== 'SEQ') continue;
    const id = (s.parsed.args[0]?.text ?? '').toLowerCase();
    if (id === '') continue;
    if (s.secondary) {
      const ac = s.begin.closest(NS.mc, 'AlternateContent');
      if (ac) twins.set(ac, [...(twins.get(ac) ?? []), s]);
      continue;
    }
    const para = s.begin.closest(NS.w, 'p');
    const c = counters.get(id) ?? { value: 0, last: -1 };
    const at = para === undefined ? c.last : (order.get(para) ?? c.last);
    const reset = Number(findSwitch(s.parsed.switches, '\\s')?.arg);
    if (Number.isInteger(reset) && reset >= 1 && reset <= 9) {
      if (headings.some((h) => h.level <= reset && h.at > c.last && h.at <= at)) c.value = 0;
    }
    const restart = findSwitch(s.parsed.switches, '\\r')?.arg;
    if (restart !== undefined && Number.isFinite(Number(restart)))
      c.value = Math.trunc(Number(restart));
    else if (!hasSwitch(s.parsed.switches, '\\c')) c.value++;
    c.last = at;
    counters.set(id, c);
    st.seq.set(s, c.value);
  }
  // A SEQ in the hidden branch of mc:AlternateContent takes the value of its twin (same
  // position among the SEQ fields of the shown branch).
  for (const [ac, hidden] of twins) {
    const shown = scans.filter(
      (s) =>
        s.parsed.type === 'SEQ' &&
        !s.secondary &&
        s.begin.closest(NS.mc, 'AlternateContent') === ac,
    );
    hidden.forEach((s, i) => {
      const twin = shown[i];
      const v = twin === undefined ? undefined : st.seq.get(twin);
      if (v !== undefined) st.seq.set(s, v);
    });
  }
}

/** Evaluates and writes one field; true when its result was updated. */
function updateOne(run: Run, s: Scan): boolean {
  const st = run.st;
  st.noData = false;
  const instruction = instructionOf(run, s);
  const p = parseInstruction(instruction);
  const type = effectiveType(run, p);
  if (run.types !== undefined && !run.types.has(type)) return false;
  if (type === 'TOC' && run.custom.TOC === undefined) {
    const scans = run.scans.get(s.story) ?? [];
    const ok = rebuildToc(st, s, p, scans, run.toc);
    if (ok) st.info.invalidateBookmarks();
    return ok;
  }
  const out = evaluate(run, s, p, instruction);
  if (out === undefined) return false;
  if (out.noWrite) {
    run.done.add(s);
    return true;
  }
  const charFormat = p.switches.some((x) => x.name === '\\*' && /^charformat$/i.test(x.arg ?? ''));
  const rPr = templateProps(s, charFormat);
  if (!writeResult(s, out.text, rPr, out.format)) return false;
  clearDirty(s);
  run.current.set(s, out.text);
  run.done.add(s);
  return true;
}

/** Updates field `root` and everything nested in it (explicit stack). */
function processTree(run: Run, root: Scan): void {
  interface Frame {
    s: Scan;
    stage: 0 | 1 | 2;
    updated: boolean;
  }
  const frames: Frame[] = [{ s: root, stage: 0, updated: false }];
  const noData = new Set<Scan>();
  while (frames.length > 0) {
    const f = frames[frames.length - 1]!;
    const s = f.s;
    if (f.stage === 0) {
      if (!attached(s.begin, s.story.element)) {
        frames.pop();
        continue;
      }
      if (onOff(s.begin.attrNS(NS.w, 'fldLock'))) {
        if (!s.secondary) run.skipped.push(toField(s));
        frames.pop();
        continue;
      }
      f.stage = 1;
      const inner = s.children.filter((c) => c.inInstr);
      for (let i = inner.length - 1; i >= 0; i--) {
        frames.push({ s: inner[i]!, stage: 0, updated: false });
      }
      continue;
    }
    if (f.stage === 1) {
      f.stage = 2;
      const blocked = s.children.some((c) => c.inInstr && noData.has(c));
      if (!blocked) {
        f.updated = updateOne(run, s);
        if (!f.updated && run.st.noData) noData.add(s);
      } else {
        noData.add(s);
      }
      if (f.updated) {
        if (!s.secondary) run.updated++;
      } else {
        if (!s.secondary) run.skipped.push(toField(s, instructionOf(run, s)));
        const rest = s.children.filter((c) => !c.inInstr);
        for (let i = rest.length - 1; i >= 0; i--) {
          frames.push({ s: rest[i]!, stage: 0, updated: false });
        }
      }
      continue;
    }
    frames.pop();
    if (!run.current.has(s)) {
      run.current.set(
        s,
        joinPieces(s.result, (c) => run.current.get(c) ?? c.resultText),
      );
    }
  }
}

/** Recomputes field results in place (see the file header). */
export function updateFieldsImpl(
  doc: WordDocument,
  o: UpdateFieldsOptions = {},
): UpdateFieldsResult {
  return runUpdate(doc, o).result;
}

/** What a mail merge (src/mailmerge) needs on top of the public options. */
export interface MergeHooks {
  merge: MergeState;
  /** Called before the fields of each story are updated (the body comes first). */
  enterStory(story: Story, st: UpdateState): void;
}

/** One update run; also returns the scans and which of them were evaluated (for src/mailmerge). */
export function runUpdate(
  doc: WordDocument,
  o: UpdateFieldsOptions = {},
  hooks?: MergeHooks,
): { result: UpdateFieldsResult; scans: Map<Story, Scan[]>; done: Set<Scan> } {
  const st = makeState(doc, o);
  if (hooks !== undefined) st.merge = hooks.merge;
  const run: Run = {
    st,
    types: o.types === undefined ? undefined : new Set(o.types.map((t) => t.toUpperCase())),
    custom: Object.fromEntries(
      Object.entries(o.evaluators ?? {}).map(([k, v]) => [k.toUpperCase(), v]),
    ),
    scans: new Map(),
    current: new Map(),
    updated: 0,
    skipped: [],
    done: new Set(),
    toc: { pages: o.pages },
  };
  const stories = doc.stories();
  for (const story of stories) {
    const scans = scanStory(story, true);
    run.scans.set(story, scans);
    computeSeq(st, story, scans);
  }
  for (const story of stories) {
    hooks?.enterStory(story, st);
    for (const s of run.scans.get(story)!) if (s.parent === undefined) processTree(run, s);
  }
  if (o.markForWord === true) setUpdateFieldsOnOpen(doc, true);
  return {
    result: { updated: run.updated, skipped: run.skipped },
    scans: run.scans,
    done: run.done,
  };
}

/** Rebuilds every TOC field of the body; returns how many were rebuilt. */
export function updateTocImpl(doc: WordDocument, o: TocOptions = {}): number {
  const st = makeState(doc, o.pages === undefined ? {} : { pages: o.pages });
  const scans = scanStory(doc.body, false);
  let n = 0;
  for (const s of scans) {
    if (s.parsed.type !== 'TOC' || !attached(s.begin, doc.body.element)) continue;
    if (onOff(s.begin.attrNS(NS.w, 'fldLock'))) continue;
    if (rebuildToc(st, s, s.parsed, scans, { pages: o.pages, levels: o.levels })) {
      st.info.invalidateBookmarks();
      n++;
    }
  }
  return n;
}

/** Every field of the stories, as public views. */
export function listScans(stories: readonly Story[]): Field[] {
  const out: Field[] = [];
  for (const story of stories) out.push(...toFields(scanStory(story, false)));
  return out;
}
