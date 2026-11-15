// PDF fonts for the faces a layout references.
//
// - TrueType faces become Type0 fonts over a CIDFontType2 with Identity-H encoding and
//   CIDToGIDMap /Identity: the content stream writes the layout's glyph ids as 2-byte codes, and
//   the embedded program is a GID-preserving subset (sfnt.ts) named with a 6-letter tag derived
//   from a hash of the face and the glyph set.
// - CFF-flavoured OpenType faces are embedded whole as FontFile3 /OpenType under a CIDFontType0
//   (not subset; CID = GID holds for the common name-keyed CFF).
// - Standard-14 faces are simple Type1 fonts with WinAnsiEncoding; by agreement with the font
//   loader their glyph ids ARE WinAnsi codes, written as single bytes.
// Faces are grouped by baseFont(): metric-substitute wrappers share their base face's single font
// dictionary and subset. Every embedded font gets a /W widths array for the used glyphs and a
// ToUnicode CMap built from the TextItems' text and clusters (a ligature glyph maps to all the
// characters of its cluster), so text extraction and search work.

import { baseFont } from '../../layout/fonts/index.js';
import type { Font, TextItem } from '../../layout/types.js';
import { md5 } from './md5.js';
import { faceInfo, subsetTrueType } from './sfnt.js';
import { latin1, name, num, type PdfWriter, ref } from './writer.js';

/** Reports a non-fatal problem. */
export type Warn = (code: string, message: string) => void;

/** A face as used by the PDF: its resource name and per-glyph data. */
export interface PdfFont {
  readonly font: Font;
  /** Object number of the font dictionary (allocated on first use). */
  readonly id: number;
  /** Resource name, without the slash. */
  readonly name: string;
  /** Glyph codes are 2 bytes (Type0) rather than 1 (standard-14 WinAnsi). */
  readonly twoByte: boolean;
  /** Glyph id → the text it stands for ('' when it continues a cluster). */
  readonly glyphs: Map<number, string>;
  /** Width in 1/1000 em as written to the PDF. */
  width(gid: number): number;
}

/** Collects the faces used by the pages, then writes their font objects. */
export class FontRegistry {
  /** Layout font index → its PDF font. */
  private readonly byIndex = new Map<number, PdfFont>();
  /** One PDF font per embedded face: keyed by the face's bytes (or the face itself if none). */
  private readonly faces = new Map<Font | Uint8Array, PdfFont>();

  constructor(
    private readonly fonts: readonly Font[],
    private readonly w: PdfWriter,
    private readonly compress: boolean,
    private readonly warn: Warn,
  ) {}

  /**
   * Registers the glyphs of a text item; undefined when its font index is invalid. Layout fonts
   * that share a face (SubstituteFont wrappers over one base, or the base itself) share one PDF
   * font: one dictionary, one subset built from the union of their glyphs. The wrappers only
   * change vertical metrics, which matter to line layout, not to the PDF.
   */
  use(item: TextItem): PdfFont | undefined {
    let u = this.byIndex.get(item.font);
    if (!u) {
      const wrapped = this.fonts[item.font];
      if (!wrapped) {
        this.warn('PDF_FONT_MISSING', `text item references unknown font ${item.font}`);
        return undefined;
      }
      const font = baseFont(wrapped);
      const key = font.bytes.length ? font.bytes : font;
      u = this.faces.get(key);
      if (!u) {
        const widths = new Map<number, number>();
        const upm = font.metrics.unitsPerEm || 1000;
        u = {
          font,
          id: this.w.alloc(),
          name: `F${item.font}`,
          twoByte: font.outlines !== 'standard14',
          glyphs: new Map(),
          width(gid: number): number {
            let v = widths.get(gid);
            if (v === undefined) {
              v = Math.round((font.advance(gid) * 1000) / upm);
              widths.set(gid, v);
            }
            return v;
          },
        };
        this.faces.set(key, u);
      }
      this.byIndex.set(item.font, u);
    }
    mapClusters(item, u.glyphs);
    return u;
  }

  /** The /Font resource entries (`/F0 5 0 R ...`). */
  resources(): string {
    return [...this.faces.values()].map((u) => `/${u.name} ${ref(u.id)}`).join(' ');
  }

  /** Writes every used face (call once, after all pages). */
  write(): void {
    for (const u of this.faces.values()) {
      if (u.font.outlines === 'standard14') {
        const ps = psName(u.font);
        const sym = ps === 'Symbol' || ps === 'ZapfDingbats';
        this.w.set(
          u.id,
          `<< /Type /Font /Subtype /Type1 /BaseFont ${name(ps)}${sym ? '' : ' /Encoding /WinAnsiEncoding'} >>`,
        );
      } else this.writeType0(u);
    }
  }

  private writeType0(u: PdfFont): void {
    const f = u.font;
    const w = this.w;
    const info = faceInfo(f.bytes);
    const m = f.metrics;
    const scale = 1000 / (m.unitsPerEm || info.unitsPerEm);
    const gids = [...u.glyphs.keys()].sort((a, b) => a - b);
    const cff = f.outlines === 'cff';
    const embed = f.bytes.length > 0;
    if (!embed) this.warn('PDF_FONT_NOT_EMBEDDED', `font ${f.family} has no bytes to embed`);
    const subset = embed && !cff;
    const base = `${subset ? `${subsetTag(f, gids)}+` : ''}${psName(f)}`;
    let file = '';
    if (embed) {
      const program = subset ? subsetTrueType(f.bytes, gids) : f.bytes;
      const id = w.addStream(
        cff ? '/Subtype /OpenType' : `/Length1 ${program.length}`,
        program,
        this.compress,
      );
      file = ` /${cff ? 'FontFile3' : 'FontFile2'} ${ref(id)}`;
    }
    let flags = 4; // symbolic: glyphs are addressed by id, not by a standard encoding
    if (info.fixedPitch) flags |= 1;
    if (info.serif) flags |= 2;
    if (f.italic || m.italicAngle !== 0) flags |= 64;
    const [x0, y0, x1, y1] = info.bbox;
    const ascent = m.ascender || y1;
    const descent = m.descender || y0;
    const weight = f.bold && info.weight < 600 ? 700 : info.weight;
    const stemV = Math.round(10 + 220 * ((weight - 50) / 900) ** 2);
    const fd = w.add(
      `<< /Type /FontDescriptor /FontName ${name(base)} /Flags ${flags} ` +
        `/FontBBox [${[x0, y0, x1, y1].map((v) => num(v * scale)).join(' ')}] ` +
        `/ItalicAngle ${num(m.italicAngle)} /Ascent ${num(ascent * scale)} ` +
        `/Descent ${num(descent * scale)} /CapHeight ${num((m.capHeight || ascent) * scale)} ` +
        `/StemV ${stemV}${file} >>`,
    );
    const cid = w.add(
      `<< /Type /Font /Subtype /${cff ? 'CIDFontType0' : 'CIDFontType2'} /BaseFont ${name(base)} ` +
        '/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> ' +
        `/FontDescriptor ${ref(fd)} /W [${widthsArray(gids, (g) => u.width(g))}]` +
        `${cff ? '' : ' /CIDToGIDMap /Identity'} >>`,
    );
    const tu = w.addStream('', latin1(toUnicodeCMap(u.glyphs)), this.compress);
    w.set(
      u.id,
      `<< /Type /Font /Subtype /Type0 /BaseFont ${name(cff ? `${base}-Identity-H` : base)} ` +
        `/Encoding /Identity-H /DescendantFonts [${ref(cid)}] /ToUnicode ${ref(tu)} >>`,
    );
  }
}

/**
 * Records, for each glyph of `item`, the text of its cluster: the first glyph of a cluster gets
 * text[cluster .. next cluster start) (several characters for a ligature), later glyphs of the same
 * cluster get ''. A glyph keeps the first non-empty mapping it was given.
 */
function mapClusters(item: TextItem, out: Map<number, string>): void {
  const { glyphs, clusters, text } = item;
  const starts = [...new Set(clusters)].sort((a, b) => a - b);
  const next = new Map<number, number>();
  for (let i = 0; i < starts.length; i++) next.set(starts[i]!, starts[i + 1] ?? text.length);
  const seen = new Set<number>();
  for (let i = 0; i < glyphs.length; i++) {
    const gid = glyphs[i]!;
    if (out.get(gid)) continue;
    const c = clusters[i];
    let s: string;
    if (c === undefined) s = text[i] ?? '';
    else if (seen.has(c)) s = '';
    else s = text.slice(c, Math.max(c, next.get(c) ?? text.length));
    if (c !== undefined) seen.add(c);
    out.set(gid, s);
  }
}

/** A ToUnicode CMap (bfchar entries, 100 per block) for 2-byte glyph codes. */
export function toUnicodeCMap(glyphs: ReadonlyMap<number, string>): string {
  const entries: string[] = [];
  for (const gid of [...glyphs.keys()].sort((a, b) => a - b)) {
    const s = glyphs.get(gid)!.slice(0, 256);
    if (!s) continue;
    let u = '';
    for (let i = 0; i < s.length; i++) u += s.charCodeAt(i).toString(16).padStart(4, '0');
    entries.push(`<${gid.toString(16).padStart(4, '0')}> <${u}>`);
  }
  let body = '';
  for (let i = 0; i < entries.length; i += 100) {
    const block = entries.slice(i, i + 100);
    body += `${block.length} beginbfchar\n${block.join('\n')}\nendbfchar\n`;
  }
  return (
    '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n' +
    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n' +
    '/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n' +
    '1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n' +
    `${body}endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n`
  );
}

/** The /W array body: runs of consecutive glyph ids grouped as `first [w1 w2 ...]`. */
function widthsArray(gids: readonly number[], width: (g: number) => number): string {
  const out: string[] = [];
  let i = 0;
  while (i < gids.length) {
    const start = gids[i]!;
    const run: number[] = [];
    while (i < gids.length && gids[i] === start + run.length) run.push(width(gids[i++]!));
    out.push(`${start} [${run.join(' ')}]`);
  }
  return out.join(' ');
}

/** The face's PostScript name, stripped of characters that do not belong in a font name. */
function psName(f: Font): string {
  const s = (f.postscriptName || f.family).replace(/[\s()<>[\]{}/%#]/g, '');
  return s || 'Font';
}

/** Six upper-case letters derived from the face and its glyph set. */
function subsetTag(f: Font, gids: readonly number[]): string {
  const h = md5(new TextEncoder().encode(`${psName(f)}:${f.numGlyphs}:${gids.join(',')}`));
  let tag = '';
  for (let i = 0; i < 6; i++) tag += String.fromCharCode(65 + (h[i]! % 26));
  return tag;
}
