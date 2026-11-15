// Footnote and endnote reference marks.
//
// Word numbers notes by the order of their references in the body (text boxes included), not by
// `w:id`. A reference with `w:customMarkFollows` shows its own mark (the run's following text) and
// does not consume a number. The format and start come from `w:footnotePr`/`w:endnotePr` in the
// settings, overridden per section by the section's own `w:footnotePr`; `numRestart="eachSect"`
// restarts at each section (`eachPage` needs pagination and is treated as continuous).
// Footnotes default to decimal numbers, endnotes to lower-case roman (§17.11.11).

import { formatNumber } from '../../resolve/index.js';
import type { WordDocument } from '../../wml/document.js';
import { sectionsOf } from '../../wml/section.js';
import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import { intAttr, layoutAlternate, onOffValue, settingsRoot } from './ir-util.js';

type NoteKind = 'footnote' | 'endnote';

interface NoteFormat {
  fmt: string;
  start: number;
  restart: string;
}

/** Computed marks for the note references of a document. */
export interface NoteMarks {
  /** Mark text of a `w:footnoteReference`/`w:endnoteReference` element. */
  ofReference(ref: XElement): string | undefined;
  /** Mark text of note `id` (for the `w:footnoteRef`/`w:endnoteRef` inside the note itself). */
  ofNote(kind: NoteKind, id: string): string | undefined;
}

function readFormat(pr: XElement | undefined, base: NoteFormat): NoteFormat {
  if (pr === undefined) return base;
  const fmt = pr.element(NS.w, 'numFmt')?.attrNS(NS.w, 'val') ?? base.fmt;
  const start = intAttr(pr.element(NS.w, 'numStart'), NS.w, 'val', base.start);
  const restart = pr.element(NS.w, 'numRestart')?.attrNS(NS.w, 'val') ?? base.restart;
  return { fmt, start, restart };
}

/** Marker pushed after a section's last paragraph so the walk knows the section ended. */
const SECTION_END: unique symbol = Symbol('section-end');

/** Walks the body in document order and numbers its note references. */
export function computeNoteMarks(doc: WordDocument): NoteMarks {
  const settings = settingsRoot(doc);
  const docFmt: Record<NoteKind, NoteFormat> = {
    footnote: readFormat(settings?.element(NS.w, 'footnotePr'), {
      fmt: 'decimal',
      start: 1,
      restart: 'continuous',
    }),
    endnote: readFormat(settings?.element(NS.w, 'endnotePr'), {
      fmt: 'lowerRoman',
      start: 1,
      restart: 'continuous',
    }),
  };
  const sections = sectionsOf(doc);
  const formatAt = (kind: NoteKind, si: number): NoteFormat => {
    const sectPr = sections[Math.min(si, sections.length - 1)]?.element;
    return readFormat(sectPr?.element(NS.w, `${kind}Pr`), docFmt[kind]);
  };

  const refMarks = new Map<XElement, string>();
  const noteMarks = new Map<string, string>();
  const counters: Record<NoteKind, number> = { footnote: 0, endnote: 0 };
  let section = 0;

  const stack: (XElement | typeof SECTION_END)[] = [doc.body.element];
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el === SECTION_END) {
      section++;
      for (const kind of ['footnote', 'endnote'] as const) {
        if (formatAt(kind, section).restart === 'eachSect') counters[kind] = 0;
      }
      continue;
    }
    if (el.uri === NS.w) {
      const local = el.local;
      if (local === 'footnoteReference' || local === 'endnoteReference') {
        const kind: NoteKind = local === 'footnoteReference' ? 'footnote' : 'endnote';
        const id = el.attrNS(NS.w, 'id');
        if (onOffValue(el.attrNS(NS.w, 'customMarkFollows'), false)) continue;
        const f = formatAt(kind, section);
        const n = f.start + counters[kind]++;
        const mark = formatNumber(n, f.fmt);
        refMarks.set(el, mark);
        if (id !== undefined && !noteMarks.has(`${kind}:${id}`))
          noteMarks.set(`${kind}:${id}`, mark);
        continue;
      }
      if (local === 'del' || local === 'moveFrom' || local === 'instrText' || local === 'delText')
        continue;
      if (local === 'p' && el.element(NS.w, 'pPr')?.element(NS.w, 'sectPr') !== undefined) {
        if (el.parent === doc.body.element || el.closest(NS.w, 'tbl') === undefined) {
          stack.push(SECTION_END);
        }
      }
    } else if (el.uri === NS.mc && el.local === 'AlternateContent') {
      const branch = layoutAlternate(el);
      if (branch !== undefined) stack.push(branch);
      continue;
    }
    if (el.local.endsWith('Pr') && el.uri === NS.w) continue;
    const kids = el.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
  }

  return {
    ofReference: (ref) => refMarks.get(ref),
    ofNote: (kind, id) => noteMarks.get(`${kind}:${id}`),
  };
}
