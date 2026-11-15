// Which paragraphs a search covers, in document order.
//
// A WordDocument target expands to stories (body, headers/footers, footnotes, endnotes,
// comments: the order of `doc.stories()`), a Story to its own paragraphs, a Paragraph to itself.
// Text boxes (`w:txbxContent` inside drawings and VML shapes) are included on request; their
// paragraphs follow the paragraph that anchors them, which is plain pre-order.
//
// `mc:AlternateContent` is read through the branch `chooseAlternate` picks (the Fallback), like
// every other reader in turboword. Word, however, displays the Choice (the DrawingML text box),
// so a replacement made only in the Fallback would be invisible there. Each paragraph of the
// chosen branch therefore carries its *mirrors*: the paragraphs at the same index in the other
// branches. Replace applies the same edits to them; find never reports them (no duplicates).

import type { WordDocument } from '../wml/document.js';
import { Paragraph } from '../wml/paragraph.js';
import { Story, type StoryKind } from '../wml/story.js';
import { chooseAlternate } from '../wml/text.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { FindOptions } from './types.js';

/** One paragraph to search, with its story and its copies in other mc:AlternateContent branches. */
export interface Target {
  readonly story: Story;
  readonly el: XElement;
  readonly mirrors: readonly XElement[];
}

const NO_MIRRORS: readonly XElement[] = [];

/** Paragraphs under `root` in pre-order (see the file comment); `mirrors` collects copies. */
function collect(root: XElement, boxes: boolean, mirrors?: Map<XElement, XElement[]>): XElement[] {
  const out: XElement[] = [];
  const stack: XElement[] = [root];
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el.uri === NS.w) {
      if (el.local === 'p') {
        out.push(el);
        if (!boxes) continue;
      } else if (el.local === 'del' || el.local === 'moveFrom') {
        continue;
      } else if (
        !boxes &&
        (el.local === 'drawing' || el.local === 'pict' || el.local === 'object')
      ) {
        continue;
      }
    } else if (el.uri === NS.mc && el.local === 'AlternateContent') {
      const chosen = chooseAlternate(el);
      if (chosen === undefined) continue;
      if (mirrors !== undefined) recordMirrors(el, chosen, boxes, mirrors);
      stack.push(chosen);
      continue;
    }
    if (el !== root && el.local.endsWith('Pr')) continue;
    const kids = el.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
  }
  return out;
}

/** Pairs the chosen branch's paragraphs with those of the other branches, index by index. */
function recordMirrors(
  ac: XElement,
  chosen: XElement,
  boxes: boolean,
  mirrors: Map<XElement, XElement[]>,
): void {
  let primary: XElement[] | undefined;
  for (const branch of ac.elements(NS.mc)) {
    if (branch === chosen || (branch.local !== 'Choice' && branch.local !== 'Fallback')) continue;
    primary ??= collect(chosen, boxes);
    if (primary.length === 0) return;
    const other = collect(branch, boxes);
    if (other.length !== primary.length) continue;
    for (let k = 0; k < primary.length; k++) {
      const p = primary[k]!;
      const list = mirrors.get(p);
      if (list) list.push(other[k]!);
      else mirrors.set(p, [other[k]!]);
    }
  }
}

/** The paragraphs of one story (text boxes on request), with their mirrors. */
export function storyTargets(story: Story, boxes: boolean, boxesOnly = false): Target[] {
  const mirrors = new Map<XElement, XElement[]>();
  const out: Target[] = [];
  for (const el of collect(story.element, boxes, mirrors)) {
    if (boxesOnly && el.closest(NS.w, 'txbxContent') === undefined) continue;
    out.push({ story, el, mirrors: mirrors.get(el) ?? NO_MIRRORS });
  }
  return out;
}

/** Every paragraph `target` covers under `scope`, in document order. */
export function targetParagraphs(
  target: WordDocument | Story | Paragraph,
  scope: FindOptions['in'],
): Target[] {
  if (target instanceof Paragraph) {
    return [{ story: target.story, el: target.element, mirrors: NO_MIRRORS }];
  }
  let kinds: StoryKind[] | undefined = ['body'];
  let boxes = false;
  let boxesOnly = false;
  if (scope === 'all') {
    kinds = undefined;
    boxes = true;
  } else if (scope !== undefined) {
    boxes = scope.includes('textbox');
    kinds = scope.filter((k): k is StoryKind => k !== 'textbox');
    if (kinds.length === 0 && boxes) {
      kinds = undefined;
      boxesOnly = true;
    }
  }
  const stories = target instanceof Story ? [target] : target.stories(kinds);
  const out: Target[] = [];
  for (const s of stories) out.push(...storyTargets(s, boxes, boxesOnly));
  return out;
}

/** A Paragraph view of a target. */
export function paragraphOf(t: Target): Paragraph {
  return new Paragraph(t.el, t.story);
}
