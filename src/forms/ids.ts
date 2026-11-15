// Fresh ids for cloned content (repeating-section items).
//
// A clone carries copies of ids Word requires to be unique document-wide: content-control ids
// (`w:sdtPr/w:id`), paragraph ids (`w14:paraId`/`w14:textId`) and drawing ids (`wp:docPr/@id`).
// The document-wide registry is `DocumentIds` (src/wml/ids.ts). While it is not available (a
// stub during development), a small scan of the story parts stands in, so filling works either
// way; both hand out ids deterministically.

import type { WordDocument } from '../wml/document.js';
import { DocumentIds, REVISION_ELEMENTS } from '../wml/ids.js';
import { NS } from '../xml/ns.js';

/** The id kinds a clone needs. */
export interface IdSource {
  sdt(): number;
  paraId(): string;
  textId(): string;
  docPr(): number;
  revision(): number;
  /**
   * Suspends re-scans for a batch of edits that only write ids from this source; returns the
   * release function.
   */
  hold(): () => void;
}

/** True for the error a contract stub throws. */
function isStub(e: unknown): boolean {
  return e instanceof Error && e.message === 'not implemented';
}

/** Ids for `doc`: the document registry when available, else a local scan. */
export function idSource(doc: WordDocument): IdSource {
  try {
    const ids = DocumentIds.of(doc);
    return {
      sdt: () => ids.next('sdt'),
      paraId: () => ids.nextHex('paraId'),
      textId: () => ids.nextHex('textId'),
      docPr: () => ids.next('docPr'),
      revision: () => ids.next('revision'),
      hold: () => ids.hold(),
    };
  } catch (e) {
    if (!isStub(e)) throw e;
  }
  return scanIds(doc);
}

/** Fallback registry: counters that skip every id used in the story parts. */
function scanIds(doc: WordDocument): IdSource {
  const sdt = new Set<number>();
  const hex = new Set<number>();
  let docPr = 0;
  let revision = 0;
  const seen = new Set<string>();
  for (const story of doc.stories()) {
    if (seen.has(story.part.name)) continue;
    seen.add(story.part.name);
    const root = story.part.xml().root;
    for (const el of root.descendants()) {
      if (el.uri === NS.w && el.local === 'id' && el.parent?.nodeType === 1) {
        const p = el.parent;
        if (p.uri === NS.w && p.local === 'sdtPr') sdt.add(Number(el.attrNS(NS.w, 'val')));
      } else if (el.uri === NS.w && REVISION_ELEMENTS.has(el.local)) {
        revision = Math.max(revision, Number(el.attrNS(NS.w, 'id')) || 0);
      } else if (el.uri === NS.wp && el.local === 'docPr') {
        docPr = Math.max(docPr, Number(el.attr('id')) || 0);
      }
      for (const k of ['paraId', 'textId']) {
        const v = el.attrNS(NS.w14, k);
        if (v !== undefined) hex.add(Number.parseInt(v, 16));
      }
    }
  }
  let nSdt = 0;
  let nHex = 0;
  const nextHex = (): string => {
    do nHex++;
    while (hex.has(nHex));
    hex.add(nHex);
    return nHex.toString(16).toUpperCase().padStart(8, '0');
  };
  return {
    sdt: () => {
      do nSdt++;
      while (sdt.has(nSdt));
      sdt.add(nSdt);
      return nSdt;
    },
    paraId: nextHex,
    textId: nextHex,
    docPr: () => ++docPr,
    revision: () => ++revision,
    hold: () => () => {},
  };
}
