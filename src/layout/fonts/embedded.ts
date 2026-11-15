// Fonts embedded in a .docx and the font table (word/fontTable.xml).
//
// Word embeds fonts as `word/fonts/fontN.odttf` parts referenced from `w:font/w:embedRegular`
// (embedBold, embedItalic, embedBoldItalic) with an `r:id` and a `w:fontKey` GUID. The part is the
// TrueType file with its first 32 bytes XORed with the key (ECMA-376 Part 1 §17.8.1, "Embedded
// Font Obfuscation"): the GUID's 16 bytes, taken from its hex string in reverse order, repeated
// twice. The same XOR undoes it. Parts without a key, or whose bytes already start with an sfnt
// signature, are used as they are.
//
// The font table also tells what kind of face each family is (w:family roman/swiss/modern,
// w:pitch fixed, w:altName, w:charset), which the font manager uses to pick generic fallbacks.

import type { WarningHandler } from '../../core/warnings.js';
import { warning } from '../../core/warnings.js';
import type { Part } from '../../opc/part.js';
import type { WordDocument } from '../../wml/document.js';
import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import type { Font } from '../types.js';
import { loadFont } from './sfnt.js';

/** The style slot a face is embedded for. */
export type EmbedStyle = 'regular' | 'bold' | 'italic' | 'boldItalic';

/** One embedded face. */
export interface EmbeddedFace {
  /** `w:font/@w:name` it is embedded under. */
  name: string;
  style: EmbedStyle;
  font: Font;
  /** Part name of the (obfuscated) font part. */
  part: string;
  /** `w:subsetted` was set: only the characters used in the document are present. */
  subsetted: boolean;
}

/** What the font table says about a family. */
export interface FontTableEntry {
  name: string;
  altName?: string;
  /** w:family: roman (serif), swiss (sans), modern (monospace), script, decorative, auto. */
  family?: string;
  /** w:pitch: fixed, variable, default. */
  pitch?: string;
  /** w:charset (hex), e.g. 02 = SYMBOL_CHARSET. */
  charset?: string;
}

const EMBED: readonly [string, EmbedStyle][] = [
  ['embedRegular', 'regular'],
  ['embedBold', 'bold'],
  ['embedItalic', 'italic'],
  ['embedBoldItalic', 'boldItalic'],
];

/** The 16 XOR key bytes of a `w:fontKey` GUID, or undefined when it is not a GUID. */
export function fontKeyBytes(fontKey: string): Uint8Array | undefined {
  const hex = fontKey.replace(/[{}-]/g, '');
  if (!/^[0-9a-fA-F]{32}$/.test(hex)) return undefined;
  const key = new Uint8Array(16);
  for (let i = 0; i < 16; i++) key[i] = Number.parseInt(hex.slice(30 - i * 2, 32 - i * 2), 16);
  return key;
}

/**
 * Undoes (or applies — XOR is its own inverse) the ECMA-376 font obfuscation: returns a copy of
 * `data` with its first 32 bytes XORed with the key derived from the `w:fontKey` GUID.
 */
export function deobfuscateFont(data: Uint8Array, fontKey: string): Uint8Array {
  const key = fontKeyBytes(fontKey);
  if (key === undefined) throw new TypeError(`Not a font key GUID: ${fontKey}`);
  const out = data.slice();
  const n = Math.min(32, out.length);
  for (let i = 0; i < n; i++) out[i] = out[i]! ^ key[i % 16]!;
  return out;
}

/** Obfuscates font bytes for embedding (the same XOR as deobfuscateFont). */
export const obfuscateFont: (data: Uint8Array, fontKey: string) => Uint8Array = deobfuscateFont;

function fontTablePart(doc: WordDocument): Part | undefined {
  return doc.mainPart.relatedByType('fontTable')[0];
}

/** The font table entries by lower-cased family name. */
export function fontTable(doc: WordDocument): Map<string, FontTableEntry> {
  const out = new Map<string, FontTableEntry>();
  const part = fontTablePart(doc);
  if (part === undefined) return out;
  let root: XElement;
  try {
    root = part.xml().root;
  } catch {
    return out;
  }
  for (const el of root.elements(NS.w, 'font')) {
    const name = el.attrNS(NS.w, 'name');
    if (name === undefined || name === '') continue;
    const entry: FontTableEntry = { name };
    const val = (local: string): string | undefined => el.element(NS.w, local)?.attrNS(NS.w, 'val');
    const alt = val('altName');
    if (alt) entry.altName = alt;
    const family = val('family');
    if (family) entry.family = family;
    const pitch = val('pitch');
    if (pitch) entry.pitch = pitch;
    const charset = val('charset');
    if (charset) entry.charset = charset;
    const key = name.toLowerCase();
    if (!out.has(key)) out.set(key, entry);
  }
  return out;
}

const SFNT_SIGNATURES = new Set([0x00010000, 0x4f54544f, 0x74727565, 0x74746366]);

function hasSfntSignature(b: Uint8Array): boolean {
  if (b.length < 4) return false;
  return SFNT_SIGNATURES.has(((b[0]! << 24) | (b[1]! << 16) | (b[2]! << 8) | b[3]!) >>> 0);
}

/**
 * Every embedded face of the document, in font-table order. Faces that cannot be read (missing
 * part, bad key, not a font) are skipped and reported through `onWarning`.
 */
export function embeddedFaces(doc: WordDocument, onWarning?: WarningHandler): EmbeddedFace[] {
  const out: EmbeddedFace[] = [];
  const part = fontTablePart(doc);
  if (part === undefined) return out;
  const warn = (code: string, message: string): void =>
    onWarning?.(warning(code, message, part.name));
  let root: XElement;
  try {
    root = part.xml().root;
  } catch {
    return out;
  }
  for (const el of root.elements(NS.w, 'font')) {
    const name = el.attrNS(NS.w, 'name');
    if (name === undefined || name === '') continue;
    for (const [local, style] of EMBED) {
      const embed = el.element(NS.w, local);
      if (embed === undefined) continue;
      const id = embed.attrNS(NS.r, 'id');
      const target = id === undefined ? undefined : part.related(id);
      if (target === undefined) {
        warn('FONT_EMBED_MISSING', `Embedded font ${name} (${style}): part not found`);
        continue;
      }
      let data = target.bytes();
      const key = embed.attrNS(NS.w, 'fontKey');
      if (key !== undefined && key !== '' && !hasSfntSignature(data)) {
        if (fontKeyBytes(key) === undefined) {
          warn('FONT_EMBED_KEY', `Embedded font ${name} (${style}): bad fontKey ${key}`);
          continue;
        }
        data = deobfuscateFont(data, key);
      }
      let font: Font;
      try {
        font = loadFont(data);
      } catch (err) {
        const why = err instanceof Error ? err.message : String(err);
        warn('FONT_EMBED_INVALID', `Embedded font ${name} (${style}) is not usable: ${why}`);
        continue;
      }
      const sub = embed.attrNS(NS.w, 'subsetted');
      out.push({
        name,
        style,
        font,
        part: target.name,
        subsetted: sub === '1' || sub === 'true' || sub === 'on',
      });
    }
  }
  return out;
}

/** Fonts embedded in the docx (word/fonts/*.odttf, deobfuscated via fontTable keys). */
export function embeddedFonts(doc: WordDocument): Font[] {
  return embeddedFaces(doc).map((f) => f.font);
}
