// Public `Field` views built from the internal index (scan.ts).
//
// A Field is a snapshot: its texts are what the document said when it was listed. Text uses
// '\n' for both paragraph marks and line breaks, like `Story.text()`.

import type { Story } from '../wml/story.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { onOff, type Scan } from './scan.js';
import type { Field, FieldSwitch } from './types.js';

const lf = (s: string): string => s.replace(/\r\n?/g, '\n');

class FieldView implements Field {
  readonly kind: 'complex' | 'simple';
  readonly type: string;
  readonly instruction: string;
  readonly code: string;
  readonly args: readonly string[];
  readonly switches: readonly FieldSwitch[];
  readonly story: Story;
  readonly start: XElement;
  readonly separate: XElement | undefined;
  readonly end: XElement;
  readonly result: string;
  readonly nested: readonly Field[];
  readonly locked: boolean;
  readonly dirty: boolean;
  parent?: Field;

  constructor(s: Scan, nested: readonly Field[], instruction?: string) {
    const run = (el: XElement | undefined): XElement | undefined =>
      el === undefined || s.kind === 'simple' ? el : ((el.parent as XElement | null) ?? el);
    this.kind = s.kind;
    this.type = s.parsed.type;
    this.instruction = lf(instruction ?? s.instrText);
    this.code = lf(s.code);
    this.args = s.parsed.args.map((t) => lf(t.text));
    this.switches = s.parsed.switches.map((w) =>
      w.arg === undefined ? w : { name: w.name, arg: lf(w.arg) },
    );
    this.story = s.story;
    this.start = run(s.begin)!;
    this.separate = run(s.separate);
    this.end = run(s.end)!;
    this.result = lf(s.resultText);
    this.nested = nested;
    this.locked = onOff(s.begin.attrNS(NS.w, 'fldLock'));
    this.dirty = onOff(s.begin.attrNS(NS.w, 'dirty'));
  }
}

function viewMap(scans: readonly Scan[]): Map<Scan, FieldView> {
  const views = new Map<Scan, FieldView>();
  for (let i = scans.length - 1; i >= 0; i--) {
    const s = scans[i]!;
    const nested: Field[] = [];
    for (const c of s.children) {
      const v = c.inInstr ? views.get(c) : undefined;
      if (v) nested.push(v);
    }
    views.set(s, new FieldView(s, nested));
  }
  for (const s of scans) {
    const parent = s.parent === undefined ? undefined : views.get(s.parent);
    if (parent) views.get(s)!.parent = parent;
  }
  return views;
}

/** Public views of `scans` (in document order; `nested` and `parent` linked). */
export function toFields(scans: readonly Scan[]): Field[] {
  const views = viewMap(scans);
  return scans.map((s) => views.get(s)!);
}

/** One view (with `instruction` overridden, e.g. after nested fields were updated). */
export function toField(s: Scan, instruction?: string): Field {
  const inner: Scan[] = [];
  const stack = [...s.children];
  while (stack.length > 0) {
    const c = stack.pop()!;
    inner.push(c);
    stack.push(...c.children);
  }
  inner.sort((a, b) => a.index - b.index);
  const views = viewMap(inner);
  const nested = s.children.filter((c) => c.inInstr).map((c) => views.get(c)!);
  return new FieldView(s, nested, instruction);
}
