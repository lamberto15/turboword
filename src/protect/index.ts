// Document protection ("Restrict Editing"): w:documentProtection in the settings part.
//
// Protection is a UI restriction Word enforces, not encryption: the content stays readable and
// any tool can remove it (that is what unprotectDocument does). The optional password is stored
// as Word's salted, iterated hash of its legacy password key (see hash.ts).
//
// Attribute form. By default the ISO/IEC 29500 attributes are written (w:algorithmName
// w:hashValue w:saltValue w:spinCount): the current standard form, and the only one the Strict
// schema allows. `attributes: 'legacy'` writes the ECMA-376 Part 4 transitional attributes
// instead (w:cryptProviderType ... w:cryptAlgorithmSid w:cryptSpinCount w:hash w:salt), which is
// what Word itself writes unless its UseIsoPasswordVerifier registry key is set (MS-OI29500
// 2.1.402 note f); Word ignores the ISO attributes when both forms are present (2.1.1775), so
// only one form is ever written. Both are hashed the same way (Word's legacy key pre-processing,
// then salted, iterated SHA), and verifyProtectionPassword reads both.
//
// Placement: CT_Settings is an XSD sequence, so w:documentProtection (and w:trackRevisions for
// the trackedChanges mode) are inserted at their schema position. Elements of other namespaces
// (w14:/w15: extensions) always sort after the w: sequence.

import { DocxFormatError, UsageError } from '../core/errors.js';
import { CT, REL } from '../opc/rel-types.js';
import type { WordDocument } from '../wml/document.js';
import { type ChildOrder, makeOrder, wElement } from '../wml/order.js';
import { XDocument, XElement } from '../xml/dom.js';
import { isStrictNs, NS } from '../xml/ns.js';
import { fromBase64, protectionHash, protectionHashFn, SIDS, spunHash, toBase64 } from './hash.js';

export type ProtectionMode = 'readOnly' | 'comments' | 'trackedChanges' | 'forms';

/** Hash algorithms accepted by {@link ProtectOptions.algorithm}. */
export type ProtectionHashAlgorithm = 'SHA-1' | 'SHA-256' | 'SHA-384' | 'SHA-512';

export interface ProtectOptions {
  mode: ProtectionMode;
  /**
   * Optional password (hashed with salt + spin count, as Word 2013+). Word only uses the first
   * 15 characters, one byte each. The empty string means no password.
   */
  password?: string;
  /** Default 100000 (what Word writes). 0 to 10,000,000. */
  spinCount?: number;
  /** 16-byte salt; default from crypto.getRandomValues (pass one for deterministic output). */
  salt?: Uint8Array;
  /** Hash algorithm. Default 'SHA-512' (Word 2013+); Word 2007/2010 wrote 'SHA-1'. */
  algorithm?: ProtectionHashAlgorithm;
  /**
   * Attribute form for the password hash: 'iso' (default; w:algorithmName, w:hashValue,
   * w:saltValue, w:spinCount) or 'legacy' (Word's transitional w:cryptProviderType ... w:hash,
   * w:salt). Strict documents always get 'iso', the only form their schema allows.
   */
  attributes?: 'iso' | 'legacy';
}

const MODES: readonly string[] = ['readOnly', 'comments', 'trackedChanges', 'forms'];
const MAX_SPIN = 10_000_000; // Word's limit (MS-OI29500 2.1.402 note d)

/** CT_Settings children in schema order (ECMA-376 Part 1 §17.15.1.78; wml.xsd). */
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
]);

/** Schema rank of a settings child: known → its position; other namespaces → last. */
function rank(el: XElement): number | undefined {
  const r = SETTINGS_ORDER.get(`${el.uri} ${el.local}`);
  if (r !== undefined) return r;
  // sl:schemaLibrary has its own namespace; the list above names it by local name only.
  if (el.local === 'schemaLibrary') return SETTINGS_ORDER.get(`${NS.w} schemaLibrary`);
  return el.uri === NS.w ? undefined : Number.POSITIVE_INFINITY;
}

/** Inserts `child` before the first settings child the schema puts after it. */
function insertSetting(settings: XElement, child: XElement): void {
  const mine = rank(child)!;
  for (const k of settings.children) {
    if (k.nodeType !== 1) continue;
    const r = rank(k);
    if (r !== undefined && r > mine) {
      settings.insertBefore(child, k);
      return;
    }
  }
  settings.append(child);
}

/** The w:settings root of the document's settings part, created when asked and missing. */
function settingsOf(doc: WordDocument, create: boolean): XElement | undefined {
  const part = doc.mainPart.relatedByType('settings')[0];
  if (part !== undefined) {
    const root = part.xml().children.find((k): k is XElement => k.nodeType === 1);
    if (root?.is(NS.w, 'settings')) return root;
    if (!create) return undefined;
    throw new DocxFormatError(`Settings part ${part.name} has no <w:settings> root`);
  }
  if (!create) return undefined;
  // Same namespace URI (transitional or Strict) and prefix as the main document.
  const main = doc.mainPart.xml().root;
  const declared = main.lookupPrefix(NS.w);
  const wUri = (declared === undefined ? undefined : main.lookupNamespace(declared)) ?? NS.w;
  const prefix = declared || 'w';
  const root = new XElement(`${prefix}:settings`, NS.w, [`xmlns:${prefix}`, NS.xmlns, wUri]);
  const dir = doc.mainPart.name.includes('/')
    ? doc.mainPart.name.slice(0, doc.mainPart.name.lastIndexOf('/') + 1)
    : '';
  const pkg = doc.package;
  const name = pkg.part(`${dir}settings.xml`)
    ? pkg.uniqueName(`${dir}settings{n}.xml`)
    : `${dir}settings.xml`;
  pkg.addPart(
    name,
    CT.settings,
    new XDocument('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', [root]),
  );
  const type = isStrictNs(wUri)
    ? 'http://purl.oclc.org/ooxml/officeDocument/relationships/settings'
    : REL.settings;
  doc.mainPart.rels.add(type, name.slice(dir.length));
  return root;
}

/** True when the settings part uses the Strict WordprocessingML namespace. */
function isStrict(settings: XElement): boolean {
  return isStrictNs(settings.lookupNamespace(settings.lookupPrefix(NS.w) ?? 'w') ?? '');
}

/** ST_OnOff: false for 0/false/off, true otherwise. */
const onOff = (v: string): boolean => !/^(0|false|off)$/i.test(v.trim());

function randomSalt(): Uint8Array {
  const c = globalThis.crypto;
  if (c?.getRandomValues === undefined) {
    throw new UsageError('crypto.getRandomValues is not available: pass ProtectOptions.salt');
  }
  return c.getRandomValues(new Uint8Array(16));
}

/**
 * Sets w:documentProtection (enforced) in settings, creating the settings part when the document
 * has none. `trackedChanges` also turns w:trackRevisions on. An existing w:formatting
 * restriction is kept. Replaces any previous editing restriction and password.
 */
export function protectDocument(doc: WordDocument, o: ProtectOptions): void {
  if (!MODES.includes(o.mode)) throw new UsageError(`Unknown protection mode: ${String(o.mode)}`);
  const spin = o.spinCount ?? 100_000;
  if (!Number.isInteger(spin) || spin < 0 || spin > MAX_SPIN) {
    throw new UsageError(`spinCount must be an integer from 0 to ${MAX_SPIN}`);
  }
  const algorithm = o.algorithm ?? 'SHA-512';
  const fn = protectionHashFn(algorithm);
  if (fn === undefined || SIDS[algorithm] === undefined) {
    throw new UsageError(`Unsupported protection hash algorithm: ${String(algorithm)}`);
  }
  if (o.salt !== undefined && o.salt.length === 0) throw new UsageError('salt must not be empty');
  if (o.attributes !== undefined && o.attributes !== 'iso' && o.attributes !== 'legacy') {
    throw new UsageError(`Unknown protection attribute form: ${String(o.attributes)}`);
  }

  const settings = settingsOf(doc, true)!;
  const old = settings.element(NS.w, 'documentProtection');
  const attrs: Record<string, string> = { edit: o.mode };
  const formatting = old?.attrNS(NS.w, 'formatting');
  if (formatting !== undefined) attrs.formatting = formatting;
  attrs.enforcement = '1';
  if (o.password !== undefined && o.password !== '') {
    const salt = o.salt ?? randomSalt();
    const hash = toBase64(protectionHash(fn, salt, o.password, spin));
    if (o.attributes !== 'legacy' || isStrict(settings)) {
      attrs.algorithmName = algorithm;
      attrs.hashValue = hash;
      attrs.saltValue = toBase64(salt);
      attrs.spinCount = String(spin);
    } else {
      attrs.cryptProviderType = algorithm === 'SHA-1' ? 'rsaFull' : 'rsaAES';
      attrs.cryptAlgorithmClass = 'hash';
      attrs.cryptAlgorithmType = 'typeAny';
      attrs.cryptAlgorithmSid = String(SIDS[algorithm]);
      attrs.cryptSpinCount = String(spin);
      attrs.hash = hash;
      attrs.salt = toBase64(salt);
    }
  }
  const el = wElement(settings, 'documentProtection', attrs);
  if (old) old.replaceWith(el);
  else insertSetting(settings, el);

  if (o.mode === 'trackedChanges') {
    const track = settings.element(NS.w, 'trackRevisions');
    if (track === undefined) insertSetting(settings, wElement(settings, 'trackRevisions'));
    else track.removeAttrNS(NS.w, 'val'); // w:val="0"/"false" would leave tracking off
  }
}

/** Removes w:documentProtection (w:trackRevisions, if on, stays on: Word does the same). */
export function unprotectDocument(doc: WordDocument): void {
  settingsOf(doc, false)?.element(NS.w, 'documentProtection')?.remove();
}

/**
 * Current protection, if enforced. Like Word (MS-OI29500 2.1.402 note c), a w:documentProtection
 * without w:enforcement counts as enforced. Formatting-only restrictions (no w:edit) and
 * `w:edit="none"` give undefined.
 */
export function protectionOf(
  doc: WordDocument,
): { mode: ProtectionMode; hasPassword: boolean } | undefined {
  const el = settingsOf(doc, false)?.element(NS.w, 'documentProtection');
  if (el === undefined) return undefined;
  const enforcement = el.attrNS(NS.w, 'enforcement');
  if (enforcement !== undefined && !onOff(enforcement)) return undefined;
  const mode = el.attrNS(NS.w, 'edit');
  if (mode === undefined || !MODES.includes(mode)) return undefined;
  const hash = el.attrNS(NS.w, 'hash') ?? el.attrNS(NS.w, 'hashValue');
  return { mode: mode as ProtectionMode, hasPassword: hash !== undefined && hash.trim() !== '' };
}

/**
 * True when `password` matches the document's protection hash. Reads Word's transitional form
 * (w:cryptAlgorithmSid w:cryptSpinCount w:hash w:salt; SHA-1/256/384/512) and the ISO form
 * (w:algorithmName w:hashValue w:saltValue w:spinCount; tried with Word's legacy-key
 * pre-processing first, then on the raw password as some producers write it). When both are
 * present the transitional form wins, as in Word. A document with no protection password
 * accepts only the empty password. MD2/MD4/MD5 hashes and spin counts above 10,000,000 never
 * match.
 */
export function verifyProtectionPassword(doc: WordDocument, password: string): boolean {
  const el = settingsOf(doc, false)?.element(NS.w, 'documentProtection');
  const a = (local: string): string | undefined => el?.attrNS(NS.w, local);
  const legacy = a('hash') !== undefined;
  const hashText = legacy ? a('hash') : a('hashValue');
  if (hashText === undefined || hashText.trim() === '') return password === '';
  const want = fromBase64(hashText);
  const salt = fromBase64((legacy ? a('salt') : a('saltValue')) ?? '');
  const spin = Number((legacy ? a('cryptSpinCount') : a('spinCount')) ?? '0');
  const fn = protectionHashFn(
    (legacy ? (a('cryptAlgorithmSid') ?? '4') : a('algorithmName')) ?? '',
  );
  if (!want || !salt || !fn || !Number.isInteger(spin) || spin < 0 || spin > MAX_SPIN) return false;
  const same = (got: Uint8Array): boolean => {
    let d = got.length ^ want.length;
    for (let i = 0; i < got.length && i < want.length; i++) d |= got[i]! ^ want[i]!;
    return d === 0;
  };
  if (same(protectionHash(fn, salt, password, spin))) return true;
  return !legacy && same(spunHash(fn, salt, password, spin));
}
