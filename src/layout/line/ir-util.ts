// Small helpers shared by the IR builder (src/layout/ir.ts and line/ir-*.ts): unit conversion,
// ST_OnOff, the document's settings part, the `mc:AlternateContent` branch layout reads and the
// pixel size of embedded images.
//
// Layout reads the *Choice* branch of `mc:AlternateContent` when it only requires namespaces the
// layout understands (DrawingML shapes, groups, canvases, Word 2010+ extensions): those carry real
// geometry and text boxes, while the Fallback is usually legacy VML. text.ts picks the Fallback for
// visible text, which is fine there because both branches show the same characters.

import type { Part } from '../../opc/part.js';
import type { WordDocument } from '../../wml/document.js';
import { ELEMENT_NODE, type XElement } from '../../xml/dom.js';
import { canonicalNs, NS } from '../../xml/ns.js';

/** Points per twip. */
export const PT_PER_TWIP: number = 1 / 20;
/** English Metric Units per point. */
export const EMU_PER_PT: number = 12700;

/** Twips → points. */
export function twipsToPt(v: number): number {
  return v / 20;
}

/** EMU → points. */
export function emuToPt(v: number): number {
  return v / EMU_PER_PT;
}

/** Integer attribute (or `fallback` when absent/invalid). */
export function intAttr(
  el: XElement | undefined,
  uri: string,
  local: string,
  fallback: number,
): number {
  const v = el?.attrNS(uri, local);
  if (v === undefined) return fallback;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : fallback;
}

/** Numeric `w:val` of a child element (twips etc.), or `fallback`. */
export function wVal(el: XElement | undefined, fallback: number): number {
  return intAttr(el, NS.w, 'val', fallback);
}

/** ST_OnOff of an element: present without `w:val` = on; `false`/`0`/`off` = off. */
export function onOff(el: XElement | undefined): boolean {
  if (el === undefined) return false;
  return onOffValue(el.attrNS(NS.w, 'val'), true);
}

/** ST_OnOff of an attribute value (`dflt` when absent). */
export function onOffValue(v: string | undefined, dflt: boolean): boolean {
  if (v === undefined) return dflt;
  return !(v === 'false' || v === '0' || v === 'off' || v === 'f');
}

/** First element child of a part's DOM (undefined for empty or non-XML parts). */
export function partRoot(part: Part | undefined): XElement | undefined {
  if (part === undefined || !part.isXml) return undefined;
  for (const k of part.xml().children) if (k.nodeType === ELEMENT_NODE) return k;
  return undefined;
}

/** The `w:settings` element of the document, if any. */
export function settingsRoot(doc: WordDocument): XElement | undefined {
  return partRoot(doc.mainPart.relatedByType('settings')[0]);
}

/** The `w:styles` element of the document, if any. */
export function stylesRoot(doc: WordDocument): XElement | undefined {
  return partRoot(doc.mainPart.relatedByType('styles')[0]);
}

/** Extension namespaces whose `mc:Choice` content the layout can read. */
const UNDERSTOOD: ReadonlySet<string> = /* @__PURE__ */ new Set<string>([
  NS.w,
  NS.wp,
  NS.a,
  NS.pic,
  NS.m,
  NS.wps,
  NS.wpg,
  NS.wpc,
  NS.wpi,
  NS.wp14,
  NS.w14,
  NS.w15,
  NS.w16se,
  NS.v,
  NS.o,
  NS.w10,
  'http://schemas.microsoft.com/office/drawing/2010/main', // a14
  'http://schemas.microsoft.com/office/drawing/2014/main', // a16
]);

/**
 * The branch of an `mc:AlternateContent` the layout reads: the first `mc:Choice` whose
 * `Requires` prefixes all map to understood namespaces, else the `mc:Fallback`, else nothing.
 */
export function layoutAlternate(ac: XElement): XElement | undefined {
  for (const choice of ac.elements(NS.mc, 'Choice')) {
    const req = choice.attr('Requires') ?? '';
    const prefixes = req.split(/\s+/).filter((s) => s !== '');
    let ok = prefixes.length > 0;
    for (const p of prefixes) {
      const uri = choice.lookupNamespace(p);
      if (uri === undefined || !UNDERSTOOD.has(canonicalNs(uri))) {
        ok = false;
        break;
      }
    }
    if (ok) return choice;
  }
  return ac.element(NS.mc, 'Fallback');
}

/** Pixel size of a PNG, JPEG, GIF or BMP image from its header (undefined when unknown). */
export function imagePixelSize(b: Uint8Array): { width: number; height: number } | undefined {
  const n = b.length;
  // PNG: signature + IHDR (width, height big-endian at 16 and 20).
  if (n >= 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
    return { width: be32(b, 16), height: be32(b, 20) };
  }
  // GIF: "GIF8", logical screen size little-endian at 6 and 8.
  if (n >= 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) {
    return { width: b[6]! | (b[7]! << 8), height: b[8]! | (b[9]! << 8) };
  }
  // BMP: "BM", BITMAPINFOHEADER width/height (int32 LE) at 18 and 22.
  if (n >= 26 && b[0] === 0x42 && b[1] === 0x4d) {
    const w = le32(b, 18);
    const h = le32(b, 22);
    return { width: Math.abs(w), height: Math.abs(h) };
  }
  // JPEG: walk the markers to the first SOFn (not DHT C4, JPG C8, DAC CC).
  if (n >= 4 && b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i + 9 < n) {
      if (b[i] !== 0xff) {
        i++;
        continue;
      }
      const marker = b[i + 1]!;
      if (marker === 0xff) {
        i++;
        continue;
      }
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        i += 2;
        continue;
      }
      const len = (b[i + 2]! << 8) | b[i + 3]!;
      if (
        marker >= 0xc0 &&
        marker <= 0xcf &&
        marker !== 0xc4 &&
        marker !== 0xc8 &&
        marker !== 0xcc
      ) {
        return { height: (b[i + 5]! << 8) | b[i + 6]!, width: (b[i + 7]! << 8) | b[i + 8]! };
      }
      if (len < 2) return undefined;
      i += 2 + len;
    }
  }
  return undefined;
}

function be32(b: Uint8Array, i: number): number {
  return ((b[i]! << 24) | (b[i + 1]! << 16) | (b[i + 2]! << 8) | b[i + 3]!) >>> 0;
}

function le32(b: Uint8Array, i: number): number {
  return b[i]! | (b[i + 1]! << 8) | (b[i + 2]! << 16) | (b[i + 3]! << 24);
}

/** CSS length (VML `style` values: pt, in, cm, mm, pc, px, em; bare numbers = px) → points. */
export function cssLengthPt(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const m = /^\s*(-?\d*\.?\d+(?:e-?\d+)?)\s*(pt|in|cm|mm|pc|px|em)?\s*$/i.exec(v);
  if (m === null) return undefined;
  const n = Number(m[1]);
  switch ((m[2] ?? 'px').toLowerCase()) {
    case 'pt':
      return n;
    case 'in':
      return n * 72;
    case 'cm':
      return (n * 72) / 2.54;
    case 'mm':
      return (n * 72) / 25.4;
    case 'pc':
      return n * 12;
    case 'em':
      return n * 12;
    default:
      return n * 0.75;
  }
}

/** Parses a VML/CSS `style` attribute into lower-cased property → value. */
export function parseCssStyle(style: string | undefined): Map<string, string> {
  const out = new Map<string, string>();
  if (style === undefined) return out;
  for (const decl of style.split(';')) {
    const i = decl.indexOf(':');
    if (i <= 0) continue;
    out.set(decl.slice(0, i).trim().toLowerCase(), decl.slice(i + 1).trim());
  }
  return out;
}
