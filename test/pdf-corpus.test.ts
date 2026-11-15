// toPdf over the whole fixture corpus: every document must become a PDF that qpdf accepts, and
// poppler must extract most of the document's words from it. Opt-in (TURBOWORD_PDF_CORPUS=1)
// because it exercises the whole layout engine and takes a while; the default suite covers the
// renderer with hand-built PageModels.

import { describe, expect, it } from 'vitest';
import { readDocx } from '../src/index.js';
import { fixtures } from './helpers/fixtures.js';
import { dump, fontBytes, HAS_POPPLER, HAS_QPDF, pdfText, qpdfCheck } from './pdf-helpers.js';

const ON = process.env.TURBOWORD_PDF_CORPUS === '1';

describe.skipIf(!ON || !HAS_QPDF)('toPdf over the fixture corpus', () => {
  const fonts = ['LiberationSans', 'LiberationSerif', 'LiberationMono', 'Carlito', 'Caladea'];
  const files = new Map<string, Uint8Array>();
  const resolver = (req: { family: string; bold: boolean }): Uint8Array | undefined => {
    const fam = req.family.replace(/\s+/g, '').toLowerCase();
    const hit = fonts.find((f) => f.toLowerCase() === fam);
    if (!hit) return undefined;
    const key = `${hit}-${req.bold ? 'Bold' : 'Regular'}`;
    if (!files.has(key)) files.set(key, fontBytes(key));
    return files.get(key);
  };

  for (const fx of fixtures().filter((f) => !f.synthetic)) {
    it(fx.name, async () => {
      const { toPdf } = await import('../src/render/pdf/index.js');
      const doc = readDocx(fx.bytes);
      const warnings: string[] = [];
      const pdf = await toPdf(doc, { fonts: resolver, onWarning: (w) => warnings.push(w.code) });
      dump(`${fx.name.replace(/[/\\]/g, '__')}.pdf`, pdf);
      const check = qpdfCheck(pdf);
      expect(check.status, check.output).toBe(0);
      if (HAS_POPPLER) {
        const words = (s: string) => s.split(/\s+/).filter((w) => /\w/.test(w));
        const expected = new Set(words(doc.text()));
        const got = new Set(words(pdfText(pdf)));
        if (expected.size >= 10) {
          let hit = 0;
          for (const w of expected) if (got.has(w)) hit++;
          expect(hit / expected.size).toBeGreaterThan(0.6);
        }
      }
    });
  }
});
