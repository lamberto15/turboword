// CONTRACT (compose agents): options, results and the shared ImportContext.
//
// Feature bar: Aspose appendDocument + Syncfusion list modes, in pure TypeScript.
// Pipeline (orchestrated by compose/index.ts, compose-core agent), each step its own module:
//   1. clone the source blocks into the target (compose/index.ts)
//   2. reconcileNamespaces(ctx, blocks)        compose/namespaces.ts  [core]
//   3. importStyles(ctx, blocks)               compose/styles.ts      [styles]
//   4. importNumbering(ctx, blocks)            compose/numbering.ts   [styles]
//   5. importNotes(ctx, blocks)                compose/notes.ts       [parts]
//   6. importComments(ctx, blocks)             compose/comments.ts    [parts]
//   7. importBookmarks(ctx, blocks)            compose/bookmarks.ts   [parts]
//   8. remapRelationships(ctx, from, to, els)  compose/part-copier.ts [core] (r:* attributes)
//   9. renumberIds(ctx, blocks)                compose/renumber.ts    [core] (docPr, paraId, sdt, revisions, VML)
//  10. importSections(ctx, blocks, placement)  compose/sections.ts    [parts] (sectPr + headers/footers)
//  11. mergeSettings(ctx)                      compose/settings.ts    [parts] (settings, fonts, theme, webSettings)
// A step reads/writes only through ctx and the blocks; steps never call each other except through
// ctx.copier (part copying), ctx.ids (fresh ids) and importContent() (compose/content.ts).
//
// Scope rule: importStyles and importNumbering bring over every style / list used ANYWHERE in the
// source (body, headers, footers, notes, comments, styles' own pPr/numPr), not only in `blocks`,
// so content copied later by the parts steps finds its mappings in ctx.styleMap / ctx.numIdMap.
//
// Step signatures (each exported from its module):
//   reconcileNamespaces(ctx, blocks: XElement[], targetPart: Part): void
//   importStyles(ctx, blocks: XElement[]): void
//   importNumbering(ctx, blocks: XElement[]): void
//   importNotes(ctx, blocks: XElement[]): void
//   importComments(ctx, blocks: XElement[]): void
//   importBookmarks(ctx, blocks: XElement[]): void
//   remapRelationships(ctx, sourcePart: Part, targetPart: Part, elements: XElement[]): void
//   renumberIds(ctx, elements: XElement[]): void
//   importSections(ctx, blocks: XElement[], placement: SectionPlacement): number
//   mergeSettings(ctx): void

import type { WarningHandler } from '../core/warnings.js';
import type { Part } from '../opc/part.js';
import type { WordDocument } from '../wml/document.js';
import type { DocumentIds } from '../wml/ids.js';
import type { Paragraph } from '../wml/paragraph.js';
import type { SectionType } from '../wml/section.js';
import type { Story } from '../wml/story.js';
import type { XElement } from '../xml/dom.js';

export interface ImportOptions {
  /**
   * - 'useDestinationStyles' (default): same-id styles take the target definition.
   * - 'keepSourceFormatting': conflicting source styles are expanded into direct formatting.
   * - 'keepDifferentStyles': conflicting source styles are copied renamed (`Heading1_1`), outline level kept.
   * Styles are matched by styleId, then by `w:name` for built-ins (localized ids), and compared
   * by a fingerprint of their resolved properties.
   */
  styles?: 'useDestinationStyles' | 'keepSourceFormatting' | 'keepDifferentStyles';
  /** With keepSourceFormatting: keep the source style when the target has none of that id. Default true. */
  smartStyleBehavior?: boolean;
  /** 'keepSource' (default): copied list definitions with fresh nsid; 'mergeWithDestination': reuse matching target lists. */
  lists?: 'keepSource' | 'mergeWithDestination';
  /** Start imported lists at their start value instead of continuing. Default false. */
  restartListsAtImport?: boolean;
  /** Section break type of the first imported section. Default 'nextPage'. */
  sectionStart?: SectionType;
  /** 'keepSource' (default): imported sections keep their own headers/footers; 'linkToPrevious': inherit the target's; 'none': no headers/footers. */
  headersFooters?: 'keepSource' | 'linkToPrevious' | 'none';
  /** 'keepSource' (default) keeps pgNumType/@w:start; 'continue' removes it; 'restart' sets start=1. */
  pageNumbering?: 'keepSource' | 'continue' | 'restart';
  /** Import source section properties (default true). insertDocument with false adopts the host section. */
  keepSourceSections?: boolean;
  /** Theme for the result when they differ: keep the target's (default) or take the source's. */
  theme?: 'destination' | 'source';
  /** Conflicting bookmark names: rename (default, references rewritten) or drop. */
  bookmarks?: 'rename' | 'drop';
  /** Source comments: keep (default) or drop. */
  comments?: 'keep' | 'drop';
  /** Source tracked changes: keep (default), accept or reject before importing. */
  revisions?: 'keep' | 'accept' | 'reject';
  /** Run validatePackage afterwards and throw on new errors. Default false. */
  validate?: boolean;
  onWarning?: WarningHandler;
}

export interface ImportResult {
  /** The imported block elements, now in the target. */
  blocks: XElement[];
  /** Number of sections imported. */
  sections: number;
  /** Source styleId → target styleId. */
  styleMap: Map<string, string>;
  /** Source numId → target numId. */
  numIdMap: Map<number, number>;
  /** Renamed bookmarks: source name → target name. */
  bookmarkRenames: Map<string, string>;
  /** Names of parts added to the target. */
  addedParts: string[];
}

/** Where insertDocument puts the content. */
export type InsertTarget =
  | { placeholder: string }
  | { bookmark: string }
  | { contentControl: { tag?: string; alias?: string } }
  | { paragraph: Paragraph; mode: 'before' | 'after' | 'replace' }
  | { story: Story; index: number };

/** Copies source parts into the target package (recursively, with their own rels). */
export interface PartCopier {
  /**
   * Ensures the target of source relationship `relId` of `sourcePart` exists in the target
   * package as a relationship of `targetPart`, copying the referenced part (and everything it
   * references) when needed; media are deduped by CRC+size+bytes. Returns the new relationship id
   * (external relationships are re-created as external).
   */
  copyRelationship(sourcePart: Part, relId: string, targetPart: Part): string;
  /** Copies a whole part (and its graph) into the target; returns the target part (memoized per source part). */
  copyPart(sourcePart: Part): Part;
}

/** Mutable state shared by the import steps. */
export interface ImportContext {
  readonly target: WordDocument;
  readonly source: WordDocument;
  readonly options: Readonly<Required<Omit<ImportOptions, 'onWarning'>>>;
  readonly ids: DocumentIds;
  readonly copier: PartCopier;
  warn(code: string, message: string, part?: string): void;
  readonly styleMap: Map<string, string>;
  readonly numIdMap: Map<number, number>;
  readonly bookmarkRenames: Map<string, string>;
  /** Source comment id → target comment id. */
  readonly commentIdMap: Map<number, number>;
  readonly footnoteIdMap: Map<number, number>;
  readonly endnoteIdMap: Map<number, number>;
  readonly addedParts: string[];
  /**
   * Target style elements created by importStyles (importNumbering rewrites the numIds inside
   * their pPr/numPr too).
   */
  readonly importedStyles: XElement[];
}

/** Where imported sections go (computed by the orchestrator for importSections). */
export interface SectionPlacement {
  /** 'append': at the end of the target body; 'insert': inside it before/after a block. */
  mode: 'append' | 'insert';
  /** The target story part receiving the blocks (main document for the body). */
  part: Part;
}
