// importContent: the shared "make copied markup valid here" routine. The parts steps use it for
// content that does not travel with the body blocks (header/footer, note and comment content);
// the orchestrator runs the equivalent steps on the body blocks itself.
//
// Order matters: namespaces first (later steps compare canonical URIs only, but the target root
// must declare what the content uses), then the style and list ids (through the styles and
// numbering steps' own remappers, which also apply keepSourceFormatting's expansion), the
// relationships (copying the parts they reference) and finally the document-wide ids.

import type { Part } from '../opc/part.js';
import type { XElement } from '../xml/dom.js';
import { reconcileNamespaces } from './namespaces.js';
import { remapNumIds } from './numbering.js';
import { remapRelationships } from './part-copier.js';
import { renumberIds } from './renumber.js';
import { remapStyleRefs } from './styles.js';
import type { ImportContext } from './types.js';

/**
 * Makes arbitrary copied content valid in the target: namespaces reconciled, pStyle/rStyle/tblStyle
 * mapped through ctx.styleMap, numIds through ctx.numIdMap, r:* relationships copied from
 * `sourcePart` to `targetPart` (skip with remapRelationships: false when the part's rels were
 * copied already), and docPr/paraId/sdt/revision/VML ids renumbered. Used by the parts steps for
 * header/footer, note and comment content. `elements` must already be inside `targetPart`.
 */
export function importContent(
  ctx: ImportContext,
  elements: XElement[],
  sourcePart: Part,
  targetPart: Part,
  o?: { remapRelationships?: boolean },
): void {
  if (elements.length === 0) return;
  reconcileNamespaces(ctx, elements, targetPart, sourcePart);
  // The styles/numbering steps own the mappings (and keepSourceFormatting's expansion).
  remapStyleRefs(ctx, elements);
  remapNumIds(ctx, elements);
  if (o?.remapRelationships !== false) {
    remapRelationships(ctx, sourcePart, targetPart, elements);
  }
  renumberIds(ctx, elements);
}
