// Public types of the text engine (find/replace across runs, revisions).
//
// Kept apart from the implementation so every module of src/text can import them without
// cycles; src/text/index.ts re-exports them unchanged.

import type { WarningHandler } from '../core/warnings.js';
import type { Paragraph } from '../wml/paragraph.js';
import type { Story, StoryKind } from '../wml/story.js';
import type { XElement } from '../xml/dom.js';

/** Visible text of a paragraph with a map back to the DOM (built on wml/text.ts walkVisible). */
export interface FlatText {
  readonly text: string;
  /** DOM position of character `index` (an element contributing text and the offset within it). */
  at(index: number): { node: XElement; offset: number } | undefined;
}

export interface TextMatch {
  readonly story: Story;
  readonly paragraph: Paragraph;
  /** UTF-16 offsets in the paragraph's visible text. */
  readonly start: number;
  readonly end: number;
  readonly text: string;
  /** RegExp capture groups. */
  readonly groups: readonly (string | undefined)[];
}

/**
 * What to search: story kinds, plus `'textbox'` for text-box content (`w:txbxContent` inside
 * drawings and VML shapes) of the selected stories. `['textbox']` alone searches the text boxes
 * of every story.
 */
export type SearchScope = StoryKind | 'textbox';

export interface FindOptions {
  /** Stories to search (default ['body']); 'all' = every story incl. headers, footers, notes, comments, text boxes. */
  in?: SearchScope[] | 'all';
  /** Default true for strings. */
  caseSensitive?: boolean;
  wholeWord?: boolean;
}

/** One run of replacement text with optional formatting. */
export interface RunInput {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  style?: string;
  color?: string;
  /** Font size in points. */
  size?: number;
}

export type Replacement = string | RunInput[] | ((m: TextMatch) => string | RunInput[]);

export interface ReplaceOptions extends FindOptions {
  /** Which run of the match gives the replacement its formatting. Default 'first'. */
  formatFrom?: 'first' | 'last';
  /** Matches crossing a field boundary: skip with a warning (default) or throw. */
  onFieldCrossing?: 'skip' | 'throw';
  /** Emit the change as tracked w:del + w:ins by this author. */
  track?: { author: string; date?: Date; initials?: string };
  /** Max replacements (default unlimited). */
  limit?: number;
  /** Receives skipped matches (`TEXT_FIELD_CROSSING`, `TEXT_UNSUPPORTED`). */
  onWarning?: WarningHandler;
}
