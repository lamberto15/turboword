// The text engine: find/replace across runs and accepting/rejecting tracked changes.
//
// Built on the shared visible-text walker (wml/text.ts) so search sees exactly what `text()`
// shows. See flatten.ts (character map), stories.ts (what is searched), find.ts (matching),
// replace.ts (the edit rules) and revisions.ts (accept/reject).

export { findText } from './find.js';
export {
  type Flat,
  flatten,
  flattenElement,
  type Marker,
  type MarkerKind,
  type Segment,
} from './flatten.js';
export { replaceText } from './replace.js';
export { acceptAllRevisions, rejectAllRevisions } from './revisions.js';
export type {
  FindOptions,
  FlatText,
  Replacement,
  ReplaceOptions,
  RunInput,
  SearchScope,
  TextMatch,
} from './types.js';
