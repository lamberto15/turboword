// Public types of turboword/forms (re-exported by index.ts). Kept apart so the implementation
// modules can share them without importing the entry point.

import type { Story } from '../wml/story.js';
import type { XElement } from '../xml/dom.js';

export type ContentControlType =
  | 'richText'
  | 'plainText'
  | 'checkbox'
  | 'dropDownList'
  | 'comboBox'
  | 'date'
  | 'picture'
  | 'group'
  | 'repeatingSection'
  | 'repeatingSectionItem'
  | 'buildingBlockGallery'
  | 'citation'
  | 'bibliography'
  | 'equation'
  | 'unknown';

export type ContentControlValue = string | boolean | Date | Uint8Array | null;

export interface ContentControl {
  readonly element: XElement;
  readonly story: Story;
  readonly type: ContentControlType;
  readonly tag: string | undefined;
  readonly alias: string | undefined;
  readonly id: number | undefined;
  /** Block, inline (run-level), row or cell content control. */
  readonly level: 'block' | 'inline' | 'row' | 'cell';
  /** Visible text of the content. */
  readonly text: string;
  /** Typed value: checkbox → boolean, date → Date, dropdown → selected value, picture → bytes. */
  readonly value: ContentControlValue;
  /** Data binding (w:dataBinding), if any. */
  readonly binding: { xpath: string; storeItemId?: string; prefixMappings?: string } | undefined;
  readonly placeholderShown: boolean;
  readonly locked: { content: boolean; control: boolean };
  /** Dropdown/combo items. */
  readonly items: readonly { display: string; value: string }[];
  /** Sets the value (formatting of the first run kept; placeholder state cleared). */
  setValue(v: ContentControlValue): void;
}

/**
 * A value for {@link fillContentControls}: a scalar for a control, or records for a repeating
 * section (records may hold arrays again, for nested sections).
 */
export type FillValue = ContentControlValue | readonly FillRecord[];

/** One repeating-section item: values for the controls inside it, by key. */
export interface FillRecord {
  readonly [key: string]: FillValue;
}

export interface FillResult {
  filled: number;
  /** Keys in `values` with no matching control. */
  missing: string[];
}

export interface LegacyFormField {
  readonly type: 'text' | 'checkbox' | 'dropdown';
  readonly name: string | undefined;
  readonly value: string | boolean;
  readonly element: XElement;
  setValue(v: string | boolean): void;
}
