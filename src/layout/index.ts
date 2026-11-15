// turboword/layout (experimental until docs/fidelity.md meets its targets): pagination into a JSON-able PageModel, the page
// count, and a PageInfoProvider over a layout for updateFields / updateToc.
//
// Page indices in DocumentLayout (paragraphPages, destinations, outline) are 0-based indices
// into `pages`; PageInfoProvider reports 1-based page numbers as fields expect.

import type { PageInfoProvider } from '../fields/types.js';
import type { WordDocument } from '../wml/document.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { paginateDocument } from './paginate.js';
import type { DocumentLayout, LayoutOptions } from './types.js';

export { FONT_SUBSTITUTIONS, loadFont } from './fonts/index.js';
export { basicShaper, bidiLevels, lineBreakOpportunities } from './text/index.js';
export type * from './types.js';

/** Lays the document out into pages (experimental: see docs/fidelity.md). */
export function paginate(doc: WordDocument, o?: LayoutOptions): Promise<DocumentLayout> {
  return paginateDocument(doc, o);
}

/** Number of pages. */
export async function pageCount(doc: WordDocument, o?: LayoutOptions): Promise<number> {
  return (await paginate(doc, o)).pages.length;
}

/** The paragraph a node belongs to: itself, its w:p ancestor, or the next paragraph after it. */
function paragraphOf(node: XElement, pages: ReadonlyMap<XElement, number>): number | undefined {
  const own = node.is(NS.w, 'p') ? node : node.closest(NS.w, 'p');
  if (own !== undefined) {
    const i = pages.get(own);
    if (i !== undefined) return i;
  }
  // body-level markers (a bookmarkStart between paragraphs): the next paragraph's page
  let cur: XElement | undefined = node;
  while (cur !== undefined) {
    for (let n = cur.nextElement(); n !== undefined; n = n.nextElement()) {
      const p = n.is(NS.w, 'p') ? n : n.descendants(NS.w, 'p')[0];
      if (p !== undefined) return pages.get(p);
    }
    const parent: unknown = cur.parent;
    cur =
      parent !== null && (parent as { nodeType: number }).nodeType === 1
        ? (parent as XElement)
        : undefined;
  }
  return undefined;
}

/** A PageInfoProvider over a layout (for updateFields / updateToc). */
export function pageInfo(layout: DocumentLayout): PageInfoProvider {
  const index = (node: XElement): number | undefined => paragraphOf(node, layout.paragraphPages);
  return {
    pageOf(node) {
      const i = index(node);
      return i === undefined ? undefined : i + 1;
    },
    pageNumberText(node) {
      const i = index(node);
      return i === undefined ? undefined : layout.pages[i]?.number;
    },
    pageCount() {
      return layout.pages.length;
    },
    sectionPageCount(node) {
      const i = index(node);
      const page = i === undefined ? undefined : layout.pages[i];
      if (page === undefined) return undefined;
      let n = 0;
      for (const p of layout.pages) if (p.section === page.section) n++;
      return n;
    },
  };
}
