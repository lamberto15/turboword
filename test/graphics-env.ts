// Test helper for the graphics tests: a fake FontManager/Shaper (fixed-advance faces, glyph id =
// code point) and a LayoutEnv over a WordDocument, so graphics can be tested without real fonts.

import type { FontManager, LayoutEnv } from '../src/layout/engine.js';
import type { Font, ImageResource, ShapedGlyph, Shaper } from '../src/layout/types.js';
import type { Part } from '../src/opc/part.js';
import type { ResolvedRunProps } from '../src/resolve/types.js';
import type { WordDocument } from '../src/wml/document.js';
import type { XElement } from '../src/xml/dom.js';

/** A fake face: 1000 units/em, every glyph 500 units wide (digits 556, 'W'/'M' 833). */
export function fakeFont(family: string, bold = false, italic = false): Font {
  return {
    family,
    subfamily: bold ? (italic ? 'Bold Italic' : 'Bold') : italic ? 'Italic' : 'Regular',
    postscriptName: family.replace(/\s/g, ''),
    bold,
    italic,
    metrics: {
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      lineGap: 0,
      winAscent: 900,
      winDescent: 250,
      capHeight: 700,
      xHeight: 500,
      underlinePosition: -100,
      underlineThickness: 50,
      strikeoutPosition: 250,
      strikeoutSize: 50,
      italicAngle: italic ? -12 : 0,
    },
    outlines: 'standard14',
    glyphId: (cp) => cp,
    advance: (g) => (g >= 48 && g <= 57 ? 556 : g === 87 || g === 77 ? 833 : 500),
    kerning: () => 0,
    bytes: new Uint8Array(0),
    numGlyphs: 0x110000,
  };
}

/** Shaper for fake (or real) faces: one glyph per code point, no kerning. */
export const fakeShaper: Shaper = {
  shape(font, text): ShapedGlyph[] {
    const out: ShapedGlyph[] = [];
    let i = 0;
    for (const ch of text) {
      const id = font.glyphId(ch.codePointAt(0) ?? 0);
      out.push({ id, advance: font.advance(id), dx: 0, dy: 0, cluster: i });
      i += ch.length;
    }
    return out;
  },
};

/** A FontManager handing out fake faces per (family, bold, italic). */
export function fakeFontManager(): FontManager & { requested: string[] } {
  const faces = new Map<string, Font>();
  const order: Font[] = [];
  const requested: string[] = [];
  return {
    requested,
    want() {},
    async load() {},
    fontFor(props: ResolvedRunProps): Font {
      const family = props.fonts.ascii ?? props.fonts.hAnsi ?? 'Fake';
      const key = `${family}|${props.bold}|${props.italic}`;
      let f = faces.get(key);
      if (f === undefined) {
        f = fakeFont(family, props.bold, props.italic);
        faces.set(key, f);
        requested.push(key);
      }
      if (!order.includes(f)) order.push(f);
      return f;
    },
    fonts: () => order.slice(),
    indexOf: (font) => order.indexOf(font),
  };
}

/** The part whose DOM contains `el`. */
export function partFinder(doc: WordDocument): (el: XElement) => Part | undefined {
  return (el) => {
    const owner = el.ownerDocument();
    if (owner === undefined) return doc.mainPart;
    for (const p of doc.package.parts()) {
      if (p.isXml && p._doc === owner) return p;
    }
    return doc.mainPart;
  };
}

/** A LayoutEnv with fake fonts over `doc`; warnings are collected in `warnings`. */
export function fakeEnv(
  doc: WordDocument,
  fonts: FontManager = fakeFontManager(),
  shaper: Shaper = fakeShaper,
): LayoutEnv & { warnings: string[] } {
  const warnings: string[] = [];
  const images = new Map<string, ImageResource>();
  return {
    doc,
    partOf: partFinder(doc),
    fonts,
    shaper,
    defaultTabStop: 36,
    images,
    warn(code, message) {
      warnings.push(`${code}: ${message}`);
    },
    warnings,
  };
}
