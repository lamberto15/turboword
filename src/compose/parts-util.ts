// Helpers shared by the compose "parts" steps (notes, comments, bookmarks, sections, settings).
//
// Those steps all do the same few things to the target package: find a part related from the
// main document, create one when it is missing (named, typed and declared the way Word writes
// it), make copied elements' namespace prefixes resolvable in the part they land in, and find
// everything that was imported (blocks plus note/comment content) to rewrite references in it.

import type { Part } from '../opc/part.js';
import { CT, REL } from '../opc/rel-types.js';
import { relativeTarget } from '../opc/uri.js';
import { XML_DECLARATION } from '../opc/xml-util.js';
import type { WordDocument } from '../wml/document.js';
import { type ChildOrder, insertOrdered, makeOrder } from '../wml/order.js';
import { XDocument, XElement } from '../xml/dom.js';
import { ignorablePrefixes, setIgnorable } from '../xml/mc.js';
import { canonicalNs, conventionalPrefix, IGNORABLE_NS, NS, prefixOf } from '../xml/ns.js';
import type { ImportContext } from './types.js';

/** Relationship types of the Word 2010+ comment parts (no Strict variants exist). */
export const REL_COMMENTS_EXTENDED =
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
export const REL_COMMENTS_IDS =
  'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds';
export const REL_COMMENTS_EXTENSIBLE =
  'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible';
export const REL_PEOPLE = 'http://schemas.microsoft.com/office/2011/relationships/people';

const CT_BASE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
export const CT_COMMENTS_EXTENDED: string = `${CT_BASE}commentsExtended+xml`;
export const CT_COMMENTS_IDS: string = `${CT_BASE}commentsIds+xml`;
export const CT_COMMENTS_EXTENSIBLE: string = `${CT_BASE}commentsExtensible+xml`;
export const CT_PEOPLE: string = `${CT_BASE}people+xml`;

/** `w:settings` children in schema order (CT_Settings, ECMA-376 Part 1 §17.15.1.78) + w14/w15. */
const SETTINGS_ORDER: ChildOrder = /* @__PURE__ */ makeOrder([
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
  'w14:docId',
  'w15:docId',
  'w14:conflictMode',
  'w14:discardImageEditingData',
  'w14:defaultImageDpi',
  'w15:chartTrackingRefBased',
]);

/** Root element of a part (undefined for an empty or unparsable part). */
export function partRoot(part: Part): XElement | undefined {
  if (!part.isXml) return undefined;
  try {
    for (const k of part.xml().children) if (k.nodeType === 1) return k;
  } catch {
    return undefined;
  }
  return undefined;
}

/** First part related from `doc`'s main part by relationship `kind` (last segment or full URI). */
export function mainRelated(doc: WordDocument, kind: string): Part | undefined {
  return doc.mainPart.relatedByType(kind)[0];
}

const STRICT_REL_BASE = 'http://purl.oclc.org/ooxml/officeDocument/relationships/';
const TRANSITIONAL_REL_BASE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/** `uri` in the target's flavour: the Strict twin when the target package is Strict. */
export function relTypeFor(doc: WordDocument, uri: string): string {
  if (!uri.startsWith(TRANSITIONAL_REL_BASE)) return uri;
  const strict = doc.package.rels.all().some((r) => r.type.startsWith(STRICT_REL_BASE));
  return strict ? STRICT_REL_BASE + uri.slice(TRANSITIONAL_REL_BASE.length) : uri;
}

/**
 * A new detached root element `{uri}local` carrying the namespace declarations and
 * `mc:Ignorable` of the target's main document root (so content copied in later finds the usual
 * prefixes, in the target's own flavour), plus a declaration of `uri` when the main root lacks it.
 */
export function newRoot(doc: WordDocument, uri: string, local: string): XElement {
  const main = partRoot(doc.mainPart);
  const attrs: string[] = [];
  if (main !== undefined) {
    const a = main._attrs;
    for (let i = 0; i < a.length; i += 3) {
      const isDecl = a[i + 1] === NS.xmlns;
      const isIgnorable = a[i + 1] === NS.mc && a[i] !== undefined && a[i]!.endsWith(':Ignorable');
      if (isDecl || isIgnorable) attrs.push(a[i]!, a[i + 1]!, a[i + 2]!);
    }
  }
  const root = new XElement('tmp', uri, attrs);
  let prefix = root.lookupPrefix(uri);
  if (prefix === undefined) {
    prefix = conventionalPrefix(uri) ?? 'ns0';
    root._attrs.push(`xmlns:${prefix}`, NS.xmlns, uri);
  }
  root.name = prefix === '' ? local : `${prefix}:${local}`;
  root.local = local;
  return root;
}

/**
 * Adds a new part named `preferred` (or the first free name from `template`) with `root` as its
 * content, relates it from the target main part and records it in `ctx.addedParts`.
 */
export function addMainPart(
  ctx: ImportContext,
  preferred: string,
  template: string,
  contentType: string,
  relType: string,
  root: XElement,
): Part {
  const doc = ctx.target;
  const pkg = doc.package;
  const name = pkg.part(preferred) === undefined ? preferred : pkg.uniqueName(template);
  const part = pkg.addPart(name, contentType, new XDocument(XML_DECLARATION, [root]));
  doc.mainPart.rels.add(relTypeFor(doc, relType), relativeTarget(doc.mainPart.name, name));
  ctx.addedParts.push(name);
  return part;
}

/** The id of a relationship from `from` to `to`, added when there is none yet. */
export function relationshipTo(from: Part, to: Part, relType: string): string {
  const existing = from.rels.toPart(to.name);
  if (existing !== undefined) return existing.id;
  return from.rels.add(relType, relativeTarget(from.name, to.name)).id;
}

/** The target settings part, created (empty `w:settings`) when the target has none. */
export function ensureSettings(ctx: ImportContext): XElement {
  const existing = mainRelated(ctx.target, 'settings');
  const root = existing === undefined ? undefined : partRoot(existing);
  if (root !== undefined) return root;
  const created = newRoot(ctx.target, NS.w, 'settings');
  addMainPart(ctx, 'word/settings.xml', 'word/settings{n}.xml', CT.settings, REL.settings, created);
  return created;
}

/** The settings `w:<local>` child of `settings`, created at its schema position when missing. */
export function ensureSettingsChild(settings: XElement, local: string): XElement {
  const found = settings.element(NS.w, local);
  if (found !== undefined) return found;
  return insertOrdered(settings, wChild(settings, local), SETTINGS_ORDER);
}

/** New detached `w:<local>` named with `scope`'s prefix for WordprocessingML. */
export function wChild(
  scope: XElement,
  local: string,
  attrs?: Readonly<Record<string, string>>,
): XElement {
  const p = scope.lookupPrefix(NS.w) ?? 'w';
  const el = new XElement(p === '' ? local : `${p}:${local}`, NS.w);
  const ap = p === '' ? 'w' : p;
  if (attrs) for (const k in attrs) el._attrs.push(`${ap}:${k}`, NS.w, attrs[k]!);
  return el;
}

/** ST_OnOff value of a toggle element: present without `w:val` = on. */
export function onOff(el: XElement | undefined): boolean {
  if (el === undefined) return false;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined || !(v === 'false' || v === '0' || v === 'off');
}

/** True when the document-wide `w:evenAndOddHeaders` setting of `doc` is on. */
export function evenAndOddHeaders(doc: WordDocument): boolean {
  const settings = mainRelated(doc, 'settings');
  const root = settings === undefined ? undefined : partRoot(settings);
  return onOff(root?.element(NS.w, 'evenAndOddHeaders'));
}

/** Content imported by the parts steps besides the blocks (note and comment bodies), per context. */
const extraContent = new WeakMap<ImportContext, XElement[]>();

/** Elements other than the blocks that were imported into the target (note/comment content). */
export function importedExtras(ctx: ImportContext): XElement[] {
  let list = extraContent.get(ctx);
  if (list === undefined) {
    list = [];
    extraContent.set(ctx, list);
  }
  return list;
}

/** Target parts whose content went through importContent already (memo per context). */
const processedParts = new WeakMap<ImportContext, Set<Part>>();

/** True the first time `part` is seen for `ctx` (and remembers it). */
export function firstVisit(ctx: ImportContext, part: Part): boolean {
  let set = processedParts.get(ctx);
  if (set === undefined) {
    set = new Set();
    processedParts.set(ctx, set);
  }
  if (set.has(part)) return false;
  set.add(part);
  return true;
}

/**
 * Makes every namespace prefix used by `el` (element and attribute names) resolvable at
 * `targetRoot`: a prefix bound to the same URI is fine; an undeclared one is declared on the
 * root (and listed in `mc:Ignorable` when it is a Word extension namespace and the root already
 * has an Ignorable list); a prefix bound to another URI is renamed to the root's prefix for the
 * URI. For the small comment-extension parts, whose content needs no other import work.
 */
export function adoptNamespaces(el: XElement, targetRoot: XElement): void {
  const ignorable = ignorablePrefixes(targetRoot);
  let ignorableChanged = false;
  const fix = (qname: string, uri: string): string => {
    const p = prefixOf(qname);
    if (p === '' || p === 'xml' || p === 'xmlns' || uri === '') return qname;
    const bound = targetRoot.lookupNamespace(p);
    if (bound !== undefined && canonicalNs(bound) === uri) return qname;
    const local = qname.slice(p.length + 1);
    if (bound === undefined) {
      targetRoot.setAttr(`xmlns:${p}`, uri);
      if (IGNORABLE_NS.has(uri) && ignorable.length > 0 && !ignorable.includes(p)) {
        ignorable.push(p);
        ignorableChanged = true;
      }
      return qname;
    }
    let other = targetRoot.lookupPrefix(uri);
    if (other === undefined) {
      other = `${p}x`;
      for (let n = 1; targetRoot.lookupNamespace(other) !== undefined; n++) other = `${p}x${n}`;
      targetRoot.setAttr(`xmlns:${other}`, uri);
    }
    return other === '' ? local : `${other}:${local}`;
  };
  let renamed = false;
  for (const e of [el, ...el.descendants()]) {
    const name = fix(e.name, e.uri);
    if (name !== e.name) e.setName(name, e.uri);
    const a = e._attrs;
    for (let i = 0; i < a.length; i += 3) {
      if (a[i + 1] === NS.xmlns) continue;
      const fixed = fix(a[i]!, a[i + 1]!);
      if (fixed !== a[i]) {
        a[i] = fixed;
        renamed = true;
      }
    }
  }
  if (ignorableChanged) setIgnorable(targetRoot, ignorable);
  // Attribute names were edited in place (no DOM method): mark the document modified.
  if (renamed) (el.ownerDocument() ?? targetRoot.ownerDocument())?.markDirty();
}

/** Runs of `w:r` left with nothing but their properties after a marker was removed. */
export function removeIfEmptyRun(run: XElement | undefined): void {
  if (run === undefined || !run.is(NS.w, 'r')) return;
  for (const k of run.elements()) if (!k.is(NS.w, 'rPr')) return;
  run.remove();
}

/**
 * Detaches `el`; when it is one of the top-level `blocks`, it is also taken out of that array
 * (kept in sync because the orchestrator returns it as the imported blocks).
 */
export function removeElement(blocks: XElement[], el: XElement): void {
  const i = blocks.indexOf(el);
  if (i >= 0) blocks.splice(i, 1);
  el.remove();
}

/** The elements of `roots` and all their descendants, in document order. */
export function allElements(roots: readonly XElement[]): XElement[] {
  const out: XElement[] = [];
  for (const r of roots) {
    out.push(r);
    for (const d of r.descendants()) out.push(d);
  }
  return out;
}
