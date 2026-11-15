// replaceText(): replace visible text across run boundaries without breaking the document.
//
// Planning comes first (no mutation): every target paragraph is flattened, matched, and each
// match is checked. A match that cuts through a field (a `w:fldChar` or a `w:fldSimple` edge
// strictly inside it) is skipped with a warning or throws (`onFieldCrossing`), so a throw leaves
// the document untouched. Matches touching math or ruby text are skipped (TEXT_UNSUPPORTED).
// Function replacements are called here, in document order, and `limit` is applied here.
//
// Then each paragraph's matches are applied from last to first (see runs.ts for why that keeps
// the plan valid):
// 1. the run holding the last matched character is split after it and the run holding the
//    first one before it (text leaves split too, rPr cloned); the matched leaves are then
//    isolated into runs of their own;
// 2. range markers strictly inside the match are lifted out: starts go to the match start, ends
//    to the match end, so bookmarks/comment ranges stay balanced; `w:proofErr` is dropped;
// 3. plain: the formatting run (first matched run, or last with `formatFrom: 'last'`) receives
//    the replacement (TAB/LF become `w:tab`/`w:br`, `xml:space` as needed) and the other matched
//    runs are removed, with wrappers they leave empty (hyperlinks, smart tags, insertions);
//    `RunInput[]` builds new runs from the formatting run's rPr plus the given formatting;
//    tracked: the matched runs are wrapped in `w:del` (`w:t` → `w:delText`) followed by a `w:ins`
//    with the new runs; ids come from DocumentIds, `w:date` only when given;
// 4. the pieces of split runs that ended up adjacent and alike are merged back.
//
// Non-text content inside a match (drawings, footnote references, deleted runs) is never
// removed: it simply ends up next to the replacement. Text in the other branches of an
// `mc:AlternateContent` (the DrawingML twin of a VML text box) receives the same edits when it
// reads the same (see stories.ts).

import { TurbowordError } from '../core/errors.js';
import { warning } from '../core/warnings.js';
import type { WordDocument } from '../wml/document.js';
import { DocumentIds } from '../wml/ids.js';
import { ensureChild, ensureFirstChild, RPR_ORDER, setWAttr, wElement } from '../wml/order.js';
import { Paragraph, Run } from '../wml/paragraph.js';
import { Story } from '../wml/story.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { compileQuery, matchText } from './find.js';
import { type Flat, flattenElement } from './flatten.js';
import { isolate, mergePieces, Pieces, splitRunAfter, splitRunBefore, splitText } from './runs.js';
import { paragraphOf, type Target, targetParagraphs } from './stories.js';
import type { Replacement, ReplaceOptions, RunInput, TextMatch } from './types.js';

/** One planned replacement in a paragraph. */
interface Edit {
  readonly start: number;
  readonly end: number;
  readonly value: string | readonly RunInput[];
}

interface Plan {
  readonly target: Target;
  readonly flat: Flat;
  readonly edits: Edit[];
}

/** What applying an edit needs besides the edit itself. */
interface Context {
  readonly para: Paragraph;
  readonly formatFrom: 'first' | 'last';
  readonly track: RevisionAttrs | undefined;
}

/** Attributes of a new `w:ins`/`w:del` (id assigned per element). */
interface RevisionAttrs {
  readonly next: () => number;
  readonly author: string;
  readonly date: string | undefined;
}

/** `w:` wrappers removed when a replacement leaves them without content. */
const PRUNABLE: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'hyperlink',
  'smartTag',
  'ins',
  'moveTo',
  'customXml',
  'dir',
  'bdo',
]);

function docOf(target: WordDocument | Story | Paragraph): WordDocument {
  if (target instanceof Paragraph) return target.story.doc;
  if (target instanceof Story) return target.doc;
  return target;
}

/** `2009-07-24T15:05:00Z` (Word's form: no milliseconds). */
function isoDate(d: Date): string {
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** True when a field boundary lies strictly inside [s, e). */
function crossesField(flat: Flat, s: number, e: number): boolean {
  for (const m of flat.markers) {
    if (m.kind === 'field' && m.pos > s && m.pos < e) return true;
  }
  return false;
}

/** True when every character of [s, e) sits in a plain `w:r` (no math, no ruby). */
function editable(flat: Flat, s: number, e: number, p: XElement): boolean {
  const i0 = flat.segmentIndex(s);
  const i1 = flat.segmentIndex(e - 1);
  if (i0 < 0 || i1 < 0) return false;
  for (let i = i0; i <= i1; i++) {
    const n = flat.segments[i]!.node;
    if (n.uri !== NS.w) return false;
    const r = n.parent;
    if (r === null || r.nodeType !== 1 || !r.is(NS.w, 'r')) return false;
    let a = r.parent;
    while (a !== null && a.nodeType === 1 && a !== p) {
      if (a.is(NS.w, 'r') || a.uri === NS.m) return false;
      a = a.parent;
    }
  }
  return true;
}

/** Expands `$&`, `$1`, `$<name>`, ... of `template` for the match of `re` at `start`. */
function expand(template: string, re: RegExp, text: string, start: number, end: number): string {
  if (!template.includes('$')) return template;
  const sticky = new RegExp(re.source, `${re.flags.replace('g', '')}y`);
  sticky.lastIndex = start;
  const out = text.replace(sticky, template);
  return out.slice(start, out.length - (text.length - end));
}

/** Applies RunInput formatting to run `r` (whose rPr was cloned from the formatting run). */
function format(r: XElement, input: RunInput, para: Paragraph): void {
  const run = new Run(r, para);
  if (input.style !== undefined) run.style = input.style;
  if (input.bold !== undefined) run.bold = input.bold;
  if (input.italic !== undefined) run.italic = input.italic;
  if (input.underline !== undefined || input.color !== undefined || input.size !== undefined) {
    const rPr = ensureFirstChild(r, 'rPr');
    if (input.color !== undefined) {
      setWAttr(ensureChild(rPr, 'color', RPR_ORDER), 'val', input.color.replace(/^#/, ''));
    }
    if (input.size !== undefined) {
      const hp = String(Math.round(input.size * 2));
      setWAttr(ensureChild(rPr, 'sz', RPR_ORDER), 'val', hp);
      setWAttr(ensureChild(rPr, 'szCs', RPR_ORDER), 'val', hp);
    }
    if (input.underline !== undefined) {
      setWAttr(ensureChild(rPr, 'u', RPR_ORDER), 'val', input.underline ? 'single' : 'none');
    }
  }
}

/**
 * A new run named in `scope`'s prefixes, inserted by `place`, with the formatting run's rPr
 * (minus revision records) and `text`.
 */
function newRun(
  scope: XElement,
  fr: XElement,
  input: RunInput,
  formatted: boolean,
  para: Paragraph,
  place: (r: XElement) => void,
): XElement {
  const r = wElement(scope, 'r');
  place(r);
  const rPr = fr.element(NS.w, 'rPr');
  if (rPr !== undefined) {
    const c = rPr.clone();
    for (const k of c.elements(NS.w, 'rPrChange')) k.remove();
    if (c.firstElement() !== undefined || c._attrs.length > 0) r.append(c);
  }
  if (formatted) format(r, input, para);
  new Run(r, para).text = input.text;
  return r;
}

/** Removes `el`, then every wrapper it leaves without content. */
function removePruning(el: XElement): void {
  let parent = el.parent;
  el.remove();
  while (parent !== null && parent.nodeType === 1 && parent.uri === NS.w) {
    if (!PRUNABLE.has(parent.local)) break;
    const content = parent.elements().some((k) => !k.local.endsWith('Pr'));
    if (content) break;
    const up: typeof parent.parent = parent.parent;
    parent.remove();
    parent = up;
  }
}

function revision(scope: XElement, local: 'ins' | 'del', t: RevisionAttrs): XElement {
  const attrs: Record<string, string> = { id: String(t.next()), author: t.author };
  if (t.date !== undefined) attrs.date = t.date;
  return wElement(scope, local, attrs);
}

/** Applies one edit to a flattened paragraph (see the file comment, steps 1-3). */
function applyEdit(flat: Flat, edit: Edit, ctx: Context, pieces: Pieces): void {
  const { start: s, end: e } = edit;
  const segs = flat.segments;
  const i0 = flat.segmentIndex(s);
  const i1 = flat.segmentIndex(e - 1);
  const first = segs[i0]!;
  const last = segs[i1]!;

  // 1. Boundaries: split after the last matched character, then before the first.
  const lastNode = last.node;
  const offEnd = e - last.start;
  if (lastNode.local === 't' && offEnd < lastNode.text.length) splitText(lastNode, offEnd);
  splitRunAfter(lastNode.parent as XElement, lastNode, pieces);
  let firstLeaf = first.node;
  const offStart = s - first.start;
  if (firstLeaf.local === 't' && offStart > 0 && offStart < firstLeaf.text.length) {
    firstLeaf = splitText(firstLeaf, offStart);
  }
  splitRunBefore(firstLeaf.parent as XElement, firstLeaf, pieces);
  const leaves: XElement[] = [firstLeaf];
  for (let i = i0 + 1; i < i1; i++) leaves.push(segs[i]!.node);
  if (i1 > i0) leaves.push(lastNode);
  const covered = new Set(leaves);
  const seen = new Set<XElement>();
  for (const l of leaves) {
    const r = l.parent as XElement;
    if (seen.has(r)) continue;
    seen.add(r);
    isolate(r, covered, pieces);
  }
  const runs: XElement[] = [];
  for (const l of leaves) {
    const r = l.parent as XElement;
    if (runs[runs.length - 1] !== r) runs.push(r);
  }

  // 2. Markers strictly inside the match.
  const starts: XElement[] = [];
  const ends: XElement[] = [];
  for (const m of flat.markers) {
    if (m.pos <= s || m.pos >= e) continue;
    if (m.kind === 'proof') m.node.remove();
    else if (m.kind === 'start') starts.push(m.node);
    else if (m.kind === 'end') ends.push(m.node);
  }
  for (const n of starts) n.remove();
  for (const n of ends) n.remove();

  // 3. The replacement.
  const fr = ctx.formatFrom === 'last' ? runs[runs.length - 1]! : runs[0]!;
  const value = edit.value;
  if (ctx.track !== undefined) {
    applyTracked(runs, fr, value, ctx, starts, ends);
    return;
  }
  let repl: XElement[];
  if (typeof value === 'string') {
    if (value === '') repl = [];
    else {
      new Run(fr, ctx.para).text = value;
      repl = [fr];
    }
  } else {
    repl = [];
    for (const input of value) {
      if (input.text === '') continue;
      repl.push(newRun(fr, fr, input, true, ctx.para, (r) => fr.before(r)));
    }
  }
  if (repl.length > 0) {
    const head = repl[0]!;
    for (const n of starts) head.before(n);
    let tail = repl[repl.length - 1]!;
    for (const n of ends) {
      tail.after(n);
      tail = n;
    }
  } else {
    for (const n of starts) fr.before(n);
    for (const n of ends) fr.before(n);
  }
  for (const r of runs) if (r !== repl[0]) removePruning(r);
}

/** Step 3, tracked: `w:del` around the matched runs, then a `w:ins` with the new text. */
function applyTracked(
  runs: XElement[],
  fr: XElement,
  value: string | readonly RunInput[],
  ctx: Context,
  starts: XElement[],
  ends: XElement[],
): void {
  const track = ctx.track!;
  const dels: XElement[] = [];
  let frDel: XElement | undefined;
  for (let i = 0; i < runs.length; ) {
    const r = runs[i]!;
    const del = revision(r, 'del', track);
    r.before(del);
    del.append(r);
    let j = i + 1;
    while (j < runs.length && del.nextElement() === runs[j]) del.append(runs[j++]!);
    for (let k = i; k < j; k++) {
      for (const t of runs[k]!.elements(NS.w, 't'))
        t.setName(t.name.replace(/t$/, 'delText'), NS.w);
      if (runs[k] === fr) frDel = del;
    }
    dels.push(del);
    i = j;
  }
  const inputs: RunInput[] =
    typeof value === 'string' ? (value === '' ? [] : [{ text: value }]) : [...value];
  let ins: XElement | undefined;
  const anchor = frDel ?? dels[dels.length - 1]!;
  for (const input of inputs) {
    if (input.text === '') continue;
    if (ins === undefined) {
      ins = revision(anchor, 'ins', track);
      anchor.after(ins);
    }
    const into = ins;
    newRun(into, fr, input, typeof value !== 'string', ctx.para, (r) => into.append(r));
  }
  const head = dels[0]!;
  for (const n of starts) head.before(n);
  let tail = ins ?? dels[dels.length - 1]!;
  for (const n of ends) {
    tail.after(n);
    tail = n;
  }
}

/** Applies a paragraph's edits (last first), then merges split pieces. */
function applyAll(flat: Flat, edits: readonly Edit[], ctx: Context): void {
  const pieces = new Pieces();
  for (let i = edits.length - 1; i >= 0; i--) applyEdit(flat, edits[i]!, ctx, pieces);
  mergePieces(pieces);
}

/**
 * Replaces every match of `q` in `target` (see {@link findText} for matching and `o.in`) and
 * returns how many were replaced. The replacement takes the formatting of the first matched run
 * (`formatFrom: 'last'`: the last); bookmarks, comment ranges and other range markers inside a
 * match move to its edges, proofing marks inside it are dropped, and non-text content (images,
 * note references) is kept. For a RegExp query, a string replacement expands `$&`, `$1`,
 * `$<name>`, `$$` like `String.prototype.replace`; a string query's replacement is literal. A
 * function receives each {@link TextMatch} in document order. `RunInput[]` creates formatted runs
 * (TAB/LF become tabs/breaks). Matches crossing a field boundary are skipped with a
 * `TEXT_FIELD_CROSSING` warning, or throw before anything changes with
 * `onFieldCrossing: 'throw'`; replacing inside a field result is fine. `track` records the
 * change as a tracked deletion + insertion.
 */
export function replaceText(
  target: WordDocument | Story | Paragraph,
  q: string | RegExp,
  r: Replacement,
  o: ReplaceOptions = {},
): number {
  const re = compileQuery(q, o);
  const limit = o.limit ?? Number.POSITIVE_INFINITY;
  const wholeWord = o.wholeWord === true;
  const plans: Plan[] = [];
  let count = 0;
  for (const t of targetParagraphs(target, o.in)) {
    if (count >= limit) break;
    const flat = flattenElement(t.el);
    if (flat.text === '') continue;
    const found = matchText(flat.text, re, wholeWord);
    if (found.length === 0) continue;
    let para: Paragraph | undefined;
    const edits: Edit[] = [];
    for (const m of found) {
      if (count >= limit) break;
      if (crossesField(flat, m.start, m.end)) {
        const msg = `Match "${m.text}" crosses a field boundary`;
        if (o.onFieldCrossing === 'throw') {
          throw new TurbowordError('TEXT_FIELD_CROSSING', `${msg} (in ${t.story.part.name})`);
        }
        o.onWarning?.(warning('TEXT_FIELD_CROSSING', `${msg}; skipped`, t.story.part.name));
        continue;
      }
      if (!editable(flat, m.start, m.end, t.el)) {
        o.onWarning?.(
          warning(
            'TEXT_UNSUPPORTED',
            `Match "${m.text}" touches math or ruby text; skipped`,
            t.story.part.name,
          ),
        );
        continue;
      }
      let value: string | readonly RunInput[];
      if (typeof r === 'function') {
        para ??= paragraphOf(t);
        const match: TextMatch = {
          story: t.story,
          paragraph: para,
          start: m.start,
          end: m.end,
          text: m.text,
          groups: m.groups,
        };
        value = r(match);
      } else if (typeof r === 'string') {
        value = typeof q === 'string' ? r : expand(r, re, flat.text, m.start, m.end);
      } else {
        value = r;
      }
      edits.push({ start: m.start, end: m.end, value });
      count++;
    }
    if (edits.length > 0) plans.push({ target: t, flat, edits });
  }
  if (plans.length === 0) return 0;

  const doc = docOf(target);
  let release: (() => void) | undefined;
  let track: RevisionAttrs | undefined;
  if (o.track !== undefined) {
    const ids = DocumentIds.of(doc);
    release = ids.hold();
    track = {
      next: () => ids.next('revision'),
      author: o.track.author,
      date: o.track.date === undefined ? undefined : isoDate(o.track.date),
    };
  }
  try {
    for (const plan of plans) {
      const { target: t, flat, edits } = plan;
      const formatFrom = o.formatFrom ?? 'first';
      applyAll(flat, edits, { para: paragraphOf(t), formatFrom, track });
      for (const m of t.mirrors) {
        const mf = flattenElement(m);
        if (mf.text !== flat.text) continue;
        const ok = edits.filter(
          (ed) => !crossesField(mf, ed.start, ed.end) && editable(mf, ed.start, ed.end, m),
        );
        const para = new Paragraph(m, t.story);
        applyAll(mf, ok, { para, formatFrom, track });
      }
    }
  } finally {
    release?.();
  }
  return count;
}
