// `w:updateFields` in the settings part ("update fields when the document opens"), and the
// creation of the settings and styles parts when a document lacks them.
//
// Settings children are an XSD sequence (CT_Settings, ECMA-376 Part 1 §17.15.1.78): a new
// `w:updateFields` goes at its schema position, never blindly at the end.

import type { Part } from '../opc/part.js';
import { CT, REL } from '../opc/rel-types.js';
import { relativeTarget } from '../opc/uri.js';
import type { WordDocument } from '../wml/document.js';
import { insertOrdered, makeOrder, setWAttr, wElement } from '../wml/order.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { parseXml } from '../xml/parse.js';

/** CT_Settings child order (elements up to and around `updateFields`). */
const SETTINGS_ORDER = /* @__PURE__ */ makeOrder([
  'writeProtection',
  'view',
  'zoom',
  'removePersonalInformation',
  'removeDateAndTime',
  'doNotDisplayPageBoundaries',
  'displayBackgroundShape',
  'printPostScriptOverText',
  'printFractionalCharacterWidth',
  'printFormsData',
  'embedTrueTypeFonts',
  'embedSystemFonts',
  'saveSubsetFonts',
  'saveFormsData',
  'mirrorMargins',
  'alignBordersAndEdges',
  'bordersDoNotSurroundHeader',
  'bordersDoNotSurroundFooter',
  'gutterAtTop',
  'hideSpellingErrors',
  'hideGrammaticalErrors',
  'activeWritingStyle',
  'proofState',
  'formsDesign',
  'attachedTemplate',
  'linkStyles',
  'stylePaneFormatFilter',
  'stylePaneSortMethod',
  'documentType',
  'mailMerge',
  'revisionView',
  'trackRevisions',
  'doNotTrackMoves',
  'doNotTrackFormatting',
  'documentProtection',
  'autoFormatOverride',
  'styleLockTheme',
  'styleLockQFSet',
  'defaultTabStop',
  'autoHyphenation',
  'consecutiveHyphenLimit',
  'hyphenationZone',
  'doNotHyphenateCaps',
  'showEnvelope',
  'summaryLength',
  'clickAndTypeStyle',
  'defaultTableStyle',
  'evenAndOddHeaders',
  'bookFoldRevPrinting',
  'bookFoldPrinting',
  'bookFoldPrintingSheets',
  'drawingGridHorizontalSpacing',
  'drawingGridVerticalSpacing',
  'displayHorizontalDrawingGridEvery',
  'displayVerticalDrawingGridEvery',
  'doNotUseMarginsForDrawingGridOrigin',
  'drawingGridHorizontalOrigin',
  'drawingGridVerticalOrigin',
  'doNotShadeFormData',
  'noPunctuationKerning',
  'characterSpacingControl',
  'printTwoOnOne',
  'strictFirstAndLastChars',
  'noLineBreaksAfter',
  'noLineBreaksBefore',
  'savePreviewPicture',
  'doNotValidateAgainstSchema',
  'saveInvalidXml',
  'ignoreMixedContent',
  'alwaysShowPlaceholderText',
  'doNotDemarcateInvalidXml',
  'saveXmlDataOnly',
  'useXSLTWhenSaving',
  'saveThroughXslt',
  'showXMLTags',
  'alwaysMergeEmptyNamespace',
  'updateFields',
  'hdrShapeDefaults',
  'footnotePr',
  'endnotePr',
  'compat',
  'docVars',
  'rsids',
  'm:mathPr',
  'attachedSchema',
  'themeFontLang',
  'clrSchemeMapping',
  'doNotIncludeSubdocsInStats',
  'doNotAutoCompressPictures',
  'forceUpgrade',
  'captions',
  'readModeInkLockDown',
  'smartTagType',
  'schemaLibrary',
  'shapeDefaults',
  'doNotEmbedSmartTags',
  'decimalSymbol',
  'listSeparator',
]);

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** The main part's related part of `kind`, created (empty root `local`) when missing. */
export function ensurePart(doc: WordDocument, kind: 'settings' | 'styles', local: string): Part {
  const found = doc.mainPart.relatedByType(kind)[0];
  if (found !== undefined) return found;
  const pkg = doc.package;
  const name =
    pkg.part(`word/${kind}.xml`) === undefined
      ? `word/${kind}.xml`
      : pkg.uniqueName(`word/${kind}{n}.xml`);
  const xml = parseXml(`${DECL}<w:${local} xmlns:w="${NS.w}"/>`);
  const part = pkg.addPart(name, kind === 'settings' ? CT.settings : CT.styles, xml);
  part.xml().markDirty();
  doc.mainPart.rels.add(
    kind === 'settings' ? REL.settings : REL.styles,
    relativeTarget(doc.mainPart.name, name),
  );
  return part;
}

function rootOf(part: Part): XElement {
  return part.xml().root;
}

/**
 * Sets (or with `on` false removes) `w:updateFields` in the settings part, creating the part
 * when needed, so Word recalculates every field when it opens the document.
 */
export function setUpdateFieldsOnOpen(doc: WordDocument, on = true): void {
  if (!on && doc.mainPart.relatedByType('settings').length === 0) return;
  const root = rootOf(ensurePart(doc, 'settings', 'settings'));
  const el = root.element(NS.w, 'updateFields');
  if (!on) {
    el?.remove();
    return;
  }
  if (el) setWAttr(el, 'val', 'true');
  else insertOrdered(root, wElement(root, 'updateFields', { val: 'true' }), SETTINGS_ORDER);
}
