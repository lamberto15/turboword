// Optional data the IR builder attaches to IR objects beyond the shared contract in types.ts.
//
// These are plain extra properties on the same objects (IRParagraph, IRInline, IRAnchor), so
// code that only knows the contract passes them through untouched and layoutParagraph picks them
// up. Everything here is optional: an IR built by hand without them still lays out (with the
// defaults documented on each field). Candidates for folding into types.ts later.

import type { IRAnchor, IRInline, IRParagraph } from '../types.js';

/** Document-level settings that change how lines are filled (from `w:settings`). */
export interface LineCompat {
  /** `w:compat/w:doNotExpandShiftReturn`: lines ending in a line break are not justified. */
  doNotExpandShiftReturn?: boolean;
  /** `w:decimalSymbol` (decimal tab alignment), default '.'. */
  decimalSymbol?: string;
  /**
   * Word compatibility mode (`w:compatSetting compatibilityMode`: 15 = Word 2013+). From 15 on,
   * justified lines may shrink their spaces to fit one more word (see line/breaker.ts).
   */
  compatibilityMode?: number;
}

/** Extra fields buildIR sets on an {@link IRParagraph}. */
export interface IRParagraphExtras {
  /**
   * For each entry of `bookmarks`, the index in `inlines` before which the bookmark starts (so
   * layoutParagraph can report it on the right line). Missing: every bookmark is on line 1.
   */
  bookmarkAt?: number[];
  compat?: LineCompat;
}

/** Positional tab (`w:ptab`) details on a tab inline. */
export interface PtabInfo {
  alignment: 'left' | 'center' | 'right';
  relativeTo: 'margin' | 'indent';
  leader?: 'dot' | 'hyphen' | 'underscore' | 'middleDot';
}

/** Extra fields on inlines. */
export interface IRInlineExtras {
  /** Tab inlines from `w:ptab`. */
  ptab?: PtabInfo;
  /** Image crop in fractions (a:srcRect / 100000). */
  crop?: { left: number; top: number; right: number; bottom: number };
}

/** Extra fields on anchors. */
export interface IRAnchorExtras {
  /** Text box insets (wps:bodyPr lIns/tIns/rIns/bIns), points. */
  insets?: { top: number; right: number; bottom: number; left: number };
  /** Image crop in fractions. */
  crop?: { left: number; top: number; right: number; bottom: number };
}

/** The extras of a paragraph (empty object when it has none). */
export function paragraphExtras(p: IRParagraph): IRParagraphExtras {
  return p as IRParagraph & IRParagraphExtras;
}

/** The extras of an inline. */
export function inlineExtras(i: IRInline): IRInlineExtras {
  return i as IRInline & IRInlineExtras;
}

/** The extras of an anchor. */
export function anchorExtras(a: IRAnchor): IRAnchorExtras {
  return a as IRAnchor & IRAnchorExtras;
}
