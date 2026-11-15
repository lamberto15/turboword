// Shared helpers for the PDF renderer tests: hand-built PageModels, the OFL test fonts, and
// wrappers around external validators (qpdf, poppler's pdftotext/pdfinfo/pdfimages/pdftoppm).
// The validators are optional: tests that need them are skipped when the tool is not installed.

import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'fflate';
import { loadFont, standardFont } from '../src/layout/fonts/index.js';
import type {
  DisplayItem,
  DocumentLayout,
  Font,
  ImageResource,
  PageLayout,
  TextItem,
} from '../src/layout/types.js';

const tool = (cmd: string, args: string[]): boolean => {
  const r = spawnSync(cmd, args, { encoding: 'utf8' });
  return !r.error;
};
export const HAS_QPDF: boolean = tool('qpdf', ['--version']);
/**
 * Poppler's pdftotext/pdfinfo/pdfimages/pdftoppm, used as independent readers of our PDFs. Off on
 * Windows: the poppler builds found on Windows runners lack pdftoppm and `pdfinfo -dests`, and the
 * PDF bytes are platform-independent (deterministic), so the Linux and macOS runs cover these.
 */
export const HAS_POPPLER: boolean =
  process.platform !== 'win32' &&
  ['pdftotext', 'pdfinfo', 'pdfimages', 'pdftoppm'].every((t) => tool(t, ['-v']));

/** Runs `cmd args... <file>` on the PDF bytes (written to a temp file). */
export function runOn(
  pdf: Uint8Array,
  cmd: string,
  args: string[],
  after: string[] = [],
): { status: number; stdout: string; stderr: string } {
  const dir = mkdtempSync(join(tmpdir(), 'turboword-pdf-'));
  try {
    const file = join(dir, 'in.pdf');
    writeFileSync(file, pdf);
    const r = spawnSync(cmd, [...args, file, ...after], { encoding: 'utf8', cwd: dir });
    return { status: r.status ?? -1, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** Writes `pdf` to $TURBOWORD_PDF_OUT/<name> when that variable is set (manual inspection). */
export function dump(name: string, pdf: Uint8Array): Uint8Array {
  const dir = process.env.TURBOWORD_PDF_OUT;
  if (dir) writeFileSync(join(dir, name), pdf);
  return pdf;
}

/** `qpdf --check`: exit 0 means no errors and no warnings. */
export function qpdfCheck(pdf: Uint8Array): { status: number; output: string } {
  const r = runOn(pdf, 'qpdf', ['--check']);
  return { status: r.status, output: r.stdout + r.stderr };
}

/** Text extracted by poppler (`pdftotext -raw`, UTF-8), trailing form feeds removed. */
export function pdfText(pdf: Uint8Array): string {
  return runOn(pdf, 'pdftotext', ['-raw', '-enc', 'UTF-8'], ['-'])
    .stdout.replace(/\r\n/g, '\n')
    .replace(/\n?\f/g, '\n');
}

/** Renders page 1 with pdftoppm and returns RGB pixels. */
export function rasterize(
  pdf: Uint8Array,
  dpi = 72,
): {
  width: number;
  height: number;
  px: (x: number, y: number) => [number, number, number];
  stderr: string;
} {
  const dir = mkdtempSync(join(tmpdir(), 'turboword-pdf-'));
  try {
    const file = join(dir, 'in.pdf');
    writeFileSync(file, pdf);
    const r = spawnSync(
      'pdftoppm',
      ['-r', String(dpi), '-singlefile', '-f', '1', '-l', '1', file, join(dir, 'out')],
      {
        encoding: 'utf8',
      },
    );
    const ppm = readFileSync(join(dir, 'out.ppm'));
    // P6\n<w> <h>\n255\n<data>
    let p = 0;
    const fields: string[] = [];
    while (fields.length < 4) {
      while (/\s/.test(String.fromCharCode(ppm[p]!))) p++;
      let s = '';
      while (!/\s/.test(String.fromCharCode(ppm[p]!))) s += String.fromCharCode(ppm[p++]!);
      fields.push(s);
    }
    p++;
    const width = Number(fields[1]);
    const height = Number(fields[2]);
    const data = ppm.subarray(p);
    return {
      width,
      height,
      px: (x, y) => {
        const o = (y * width + x) * 3;
        return [data[o]!, data[o + 1]!, data[o + 2]!];
      },
      stderr: r.stderr ?? '',
    };
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** Bytes of an OFL test font from test/fixtures/fonts (stored gzipped). */
export function fontBytes(file: string): Uint8Array {
  const p = fileURLToPath(new URL(`./fixtures/fonts/${file}.ttf.gz`, import.meta.url));
  return gunzipSync(readFileSync(p));
}

/** A parsed face: the fonts module's loadFont, or a minimal local parser while it is a stub. */
export function face(file: string): Font {
  const bytes = fontBytes(file);
  try {
    return loadFont(bytes);
  } catch (e) {
    if (!String((e as Error).message).includes('not implemented')) throw e;
    return miniFont(bytes);
  }
}

/** A standard-14 face (metrics only); a tiny local stand-in while the fonts module is a stub. */
export function standard(name: string): Font {
  try {
    return standardFont(name);
  } catch (e) {
    if (!String((e as Error).message).includes('not implemented')) throw e;
    const metrics = {
      unitsPerEm: 1000,
      ascender: 718,
      descender: -207,
      lineGap: 0,
      winAscent: 718,
      winDescent: 207,
      capHeight: 718,
      xHeight: 523,
      underlinePosition: -100,
      underlineThickness: 50,
      strikeoutPosition: 250,
      strikeoutSize: 50,
      italicAngle: 0,
    };
    return {
      family: name,
      subfamily: 'Regular',
      postscriptName: name,
      bold: false,
      italic: false,
      metrics,
      outlines: 'standard14',
      glyphId: (cp) => (cp >= 32 && cp < 127 ? cp : 0),
      advance: () => 556,
      kerning: () => 0,
      bytes: new Uint8Array(0),
      numGlyphs: 256,
    };
  }
}

/** One glyph per code point, unkerned (enough for renderer tests). */
export function textItem(
  f: Font,
  fontIndex: number,
  text: string,
  x: number,
  y: number,
  size = 12,
  extra: Partial<TextItem> = {},
): TextItem {
  const glyphs: number[] = [];
  const advances: number[] = [];
  const clusters: number[] = [];
  let i = 0;
  for (const ch of text) {
    const g = f.glyphId(ch.codePointAt(0)!);
    glyphs.push(g);
    // Advances are final page advances: the horizontal scale is included (as layout does).
    advances.push(((f.advance(g) * size) / f.metrics.unitsPerEm) * ((extra.scale ?? 100) / 100));
    clusters.push(i);
    i += ch.length;
  }
  return {
    kind: 'text',
    x,
    y,
    font: fontIndex,
    size,
    color: '000000',
    glyphs,
    advances,
    text,
    clusters,
    ...extra,
  };
}

/** A DocumentLayout around hand-built pages (US Letter by default). */
export function layoutOf(
  pages: (DisplayItem[] | Partial<PageLayout>)[],
  fonts: Font[] = [],
  images: ImageResource[] = [],
  extra: Partial<DocumentLayout> = {},
): DocumentLayout {
  return {
    pages: pages.map((p, index) => {
      const base = Array.isArray(p) ? { items: p } : p;
      return {
        index,
        number: String(index + 1),
        section: 0,
        width: 612,
        height: 792,
        items: [],
        ...base,
      };
    }),
    fonts,
    images: new Map(images.map((im) => [im.id, im])),
    outline: [],
    destinations: new Map(),
    metadata: {},
    paragraphPages: new Map(),
    ...extra,
  };
}

/** A minimal TrueType Font (cmap 4/12, hmtx, head/hhea/OS/2/post/name) for tests. */
function miniFont(bytes: Uint8Array): Font {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const t = new Map<string, number>();
  for (let i = 0, n = dv.getUint16(4); i < n; i++) {
    const e = 12 + i * 16;
    t.set(String.fromCharCode(...bytes.subarray(e, e + 4)), dv.getUint32(e + 8));
  }
  const at = (tag: string): number => t.get(tag) ?? -1;
  const upm = dv.getUint16(at('head') + 18);
  const hhea = at('hhea');
  const nhm = dv.getUint16(hhea + 34);
  const numGlyphs = dv.getUint16(at('maxp') + 4);
  const hmtx = at('hmtx');
  const os2 = at('OS/2');
  const post = at('post');
  const cmap = at('cmap');
  let sub = -1;
  for (let i = 0, n = dv.getUint16(cmap + 2); i < n; i++) {
    const e = cmap + 4 + i * 8;
    const o = cmap + dv.getUint32(e + 4);
    if (dv.getUint16(e) === 3 && dv.getUint16(e + 2) === 1 && dv.getUint16(o) === 4) sub = o;
  }
  const seg = dv.getUint16(sub + 6) / 2;
  const glyphId = (cp: number): number => {
    for (let s = 0; s < seg; s++) {
      const end = dv.getUint16(sub + 14 + s * 2);
      if (cp > end) continue;
      const start = dv.getUint16(sub + 16 + seg * 2 + s * 2);
      if (cp < start) return 0;
      const delta = dv.getInt16(sub + 16 + seg * 4 + s * 2);
      const ro = sub + 16 + seg * 6 + s * 2;
      const range = dv.getUint16(ro);
      if (!range) return (cp + delta) & 0xffff;
      const g = dv.getUint16(ro + range + (cp - start) * 2);
      return g ? (g + delta) & 0xffff : 0;
    }
    return 0;
  };
  const name = at('name');
  let ps = 'Test';
  for (let i = 0, n = dv.getUint16(name + 2); i < n; i++) {
    const r = name + 6 + i * 12;
    if (dv.getUint16(r) === 3 && dv.getUint16(r + 6) === 6) {
      const len = dv.getUint16(r + 8);
      const off = name + dv.getUint16(name + 4) + dv.getUint16(r + 10);
      ps = '';
      for (let k = 0; k < len; k += 2) ps += String.fromCharCode(dv.getUint16(off + k));
    }
  }
  return {
    family: ps,
    subfamily: 'Regular',
    postscriptName: ps,
    bold: false,
    italic: false,
    metrics: {
      unitsPerEm: upm,
      ascender: dv.getInt16(hhea + 4),
      descender: dv.getInt16(hhea + 6),
      lineGap: dv.getInt16(hhea + 8),
      winAscent: dv.getUint16(os2 + 74),
      winDescent: dv.getUint16(os2 + 76),
      capHeight: dv.getInt16(os2 + 88),
      xHeight: dv.getInt16(os2 + 86),
      underlinePosition: dv.getInt16(post + 8),
      underlineThickness: dv.getInt16(post + 10),
      strikeoutPosition: dv.getInt16(os2 + 28),
      strikeoutSize: dv.getInt16(os2 + 26),
      italicAngle: dv.getInt32(post + 4) / 65536,
    },
    outlines: 'truetype',
    glyphId,
    advance: (g) => dv.getUint16(hmtx + Math.min(g, nhm - 1) * 4),
    kerning: () => 0,
    bytes,
    numGlyphs,
  };
}
