// Layout fidelity scoreboard (plan D8): turboword's pagination against free oracles.
//
// Corpus: test/fixtures/layout/corpus.json — every Word-saved fixture (deduplicated) whose
// docProps/app.xml records Word's page count. Oracles:
// - Word's page count (`<Pages>`): pages exactly equal;
// - Word's `w:lastRenderedPageBreak` markers: the share of Word's (soft) page breaks we make at
//   the same paragraph (before it, or inside it when Word broke inside);
// - LibreOffice (test/fixtures/layout/libreoffice.json, recorded by record-libreoffice.mjs with
//   the same metric-compatible fonts): text rows per page, compared as hashes of their letters
//   and digits — matched rows / max(our rows, LibreOffice's rows). The "line match".
// Files are tiered by content: C = DrawingML/VML shapes, charts, SmartArt, text boxes, math,
// OLE; B = tables, pictures, floats, frames, columns; A = the rest (prose).
//
// The measured baseline is asserted (with a small margin) so a regression fails CI; the numbers
// are published in docs/fidelity.md. FIDELITY_REPORT=1 prints every miss. Diagnostics for one
// file: FX=<path under test/fixtures> npx vitest run test/layout-oracle.test.ts -t dumps (our
// rows per page) or -t "page breaks" (Word's breaks against ours).

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { readDocx } from '../src/index.js';
import type { DocumentLayout } from '../src/layout/index.js';
import { paginate } from '../src/layout/index.js';
import type { WordDocument } from '../src/wml/document.js';
import type { XElement } from '../src/xml/dom.js';
import { NS } from '../src/xml/ns.js';
import { FIXTURES_DIR } from './helpers/fixtures.js';
import { testFontResolver } from './helpers/test-fonts.js';

const LAYOUT_DIR = join(FIXTURES_DIR, 'layout');

interface CorpusEntry {
  name: string;
  sha1: string;
  pages: number;
}
interface LibreOfficeEntry {
  pages: number;
  rows: string[][];
}

const corpus = JSON.parse(readFileSync(join(LAYOUT_DIR, 'corpus.json'), 'utf8')) as CorpusEntry[];
const libre = JSON.parse(readFileSync(join(LAYOUT_DIR, 'libreoffice.json'), 'utf8')) as Record<
  string,
  LibreOfficeEntry
>;

/** Same as record-libreoffice.mjs. */
function fnv(s: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, '0');
}
const normalize = (s: string): string => s.normalize('NFC').replace(/[^\p{L}\p{N}]/gu, '');

/** Text rows of each page (text items grouped by baseline, x order), hashed like the recorder. */
function ourRows(l: DocumentLayout): string[][] {
  return l.pages.map((p) => {
    const ws: { x: number; y: number; t: string }[] = [];
    for (const it of p.items) if (it.kind === 'text') ws.push({ x: it.x, y: it.y, t: it.text });
    ws.sort((a, b) => a.y - b.y || a.x - b.x);
    const rows: { y: number; ws: typeof ws }[] = [];
    for (const w of ws) {
      const cur = rows[rows.length - 1];
      if (cur === undefined || Math.abs(w.y - cur.y) > 2) rows.push({ y: w.y, ws: [w] });
      else cur.ws.push(w);
    }
    return rows
      .map((r) =>
        normalize(
          r.ws
            .sort((a, b) => a.x - b.x)
            .map((w) => w.t)
            .join(''),
        ),
      )
      .filter((t) => t !== '')
      .map(fnv);
  });
}

/** Rows of `ours` found on the same page of `theirs` (multiset intersection). */
function matchedRows(ours: string[][], theirs: string[][]): { matched: number; total: number } {
  let matched = 0;
  let a = 0;
  let b = 0;
  const pages = Math.max(ours.length, theirs.length);
  for (let i = 0; i < pages; i++) {
    const mine = ours[i] ?? [];
    const other = theirs[i] ?? [];
    a += mine.length;
    b += other.length;
    const count = new Map<string, number>();
    for (const h of other) count.set(h, (count.get(h) ?? 0) + 1);
    for (const h of mine) {
      const n = count.get(h) ?? 0;
      if (n > 0) {
        matched++;
        count.set(h, n - 1);
      }
    }
  }
  return { matched, total: Math.max(a, b) };
}

const onOff = (el: XElement | undefined): boolean => {
  if (el === undefined) return false;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined || !(v === 'false' || v === '0' || v === 'off');
};

/** A page break Word recorded (w:lastRenderedPageBreak not right after a hard break). */
interface WordBreak {
  p: XElement;
  prev: XElement | undefined;
  next: XElement | undefined;
  /** The page starts inside `p` (text before the marker), not at its start. */
  mid: boolean;
}

/**
 * Word's soft page breaks. Word writes the marker only into runs, so a page that starts with an
 * empty paragraph has none: markers say where pages start, not every page start.
 */
function wordBreaks(doc: WordDocument): WordBreak[] {
  const ps = doc.body.element
    .descendants(NS.w, 'p')
    .filter((p) => p.closest(NS.w, 'txbxContent') === undefined);
  const out: WordBreak[] = [];
  let fresh = true; // a hard break with no text since
  ps.forEach((p, i) => {
    const pPr = p.element(NS.w, 'pPr');
    if (onOff(pPr?.element(NS.w, 'pageBreakBefore'))) fresh = true;
    let text = false;
    for (const el of p.descendants()) {
      if (el.closest(NS.w, 'txbxContent') !== undefined) continue;
      if (el.is(NS.w, 'lastRenderedPageBreak')) {
        if (!fresh) out.push({ p, prev: ps[i - 1], next: ps[i + 1], mid: text });
        fresh = true;
      } else if (el.is(NS.w, 'br') && el.attrNS(NS.w, 'type') === 'page') fresh = true;
      else if (el.is(NS.w, 't') && el.text !== '') {
        text = true;
        fresh = false;
      }
    }
    if (pPr?.element(NS.w, 'sectPr') !== undefined) fresh = true;
  });
  return out;
}

/** Does our layout break the page where Word did? */
function sameBreak(b: WordBreak, pages: ReadonlyMap<XElement, number>): boolean {
  const at = pages.get(b.p);
  if (at === undefined) return false;
  if (b.mid) {
    const next = b.next === undefined ? undefined : pages.get(b.next);
    const prev = b.prev === undefined ? undefined : pages.get(b.prev);
    return next !== undefined && next > at && (prev === undefined || prev === at);
  }
  const prev = b.prev === undefined ? undefined : pages.get(b.prev);
  return prev !== undefined && at > prev;
}

type Tier = 'A' | 'B' | 'C';

/** VML elements that belong to a plain picture (v:shape + v:imagedata), not to a drawing. */
const VML_PICTURE = new Set(['imagedata', 'shapetype', 'formulas', 'f', 'path', 'stroke', 'lock']);

/** Content tier of a document (body, headers and footers). */
function tierOf(doc: WordDocument): Tier {
  const roots = [doc.body.element, ...doc.headersFooters().map((h) => h.story.element)];
  let tier: Tier = 'A';
  for (const root of roots) {
    for (const el of root.descendants()) {
      const u = el.uri;
      if (
        u === NS.wps ||
        u === NS.wpg ||
        u === NS.wpc ||
        u === NS.dgm ||
        u === NS.c ||
        u === NS.m ||
        (u === NS.v && !VML_PICTURE.has(el.local) && el.element(NS.v, 'imagedata') === undefined) ||
        el.is(NS.w, 'txbxContent') ||
        el.is(NS.w, 'object') ||
        el.local === 'chart'
      ) {
        return 'C';
      }
      if (
        el.is(NS.w, 'tbl') ||
        el.is(NS.wp, 'anchor') ||
        el.is(NS.pic, 'pic') ||
        el.is(NS.v, 'imagedata') ||
        el.is(NS.w, 'framePr') ||
        (el.is(NS.w, 'cols') && Number(el.attrNS(NS.w, 'num') ?? '1') > 1)
      ) {
        tier = 'B';
      }
    }
  }
  return tier;
}

interface Score {
  docs: number;
  pagesExact: number;
  breaks: number;
  breaksAgree: number;
  rows: number;
  rowsMatched: number;
  loPagesExact: number;
}
const empty = (): Score => ({
  docs: 0,
  pagesExact: 0,
  breaks: 0,
  breaksAgree: 0,
  rows: 0,
  rowsMatched: 0,
  loPagesExact: 0,
});
const pct = (a: number, b: number): number => (b === 0 ? 100 : Math.round((1000 * a) / b) / 10);

async function scoreboard(): Promise<{
  tiers: Record<Tier, Score>;
  misses: string[];
  skipped: string[];
}> {
  const tiers: Record<Tier, Score> = { A: empty(), B: empty(), C: empty() };
  const misses: string[] = [];
  const skipped: string[] = [];
  const resolver = testFontResolver();
  for (const entry of corpus) {
    let bytes: Uint8Array;
    try {
      bytes = new Uint8Array(readFileSync(join(FIXTURES_DIR, entry.name)));
    } catch {
      skipped.push(`${entry.name} (missing)`);
      continue;
    }
    const doc = readDocx(bytes);
    const tier = tierOf(doc);
    const l = await paginate(doc, { fonts: resolver });
    const s = tiers[tier];
    s.docs++;
    if (l.pages.length === entry.pages) s.pagesExact++;
    else misses.push(`[${tier}] ${entry.name}: ${l.pages.length} pages, Word ${entry.pages}`);
    const breaks = wordBreaks(doc);
    if (breaks.length > 0) {
      const agree = breaks.filter((b) => sameBreak(b, l.paragraphPages)).length;
      s.breaks += breaks.length;
      s.breaksAgree += agree;
      if (agree < breaks.length) {
        misses.push(`[${tier}] ${entry.name}: ${agree}/${breaks.length} of Word's page breaks`);
      }
    }
    const lo = libre[entry.name];
    if (lo !== undefined) {
      if (lo.pages === entry.pages) s.loPagesExact++;
      const m = matchedRows(ourRows(l), lo.rows);
      s.rows += m.total;
      s.rowsMatched += m.matched;
      if (m.matched < m.total) {
        misses.push(`[${tier}] ${entry.name}: ${m.matched}/${m.total} rows as LibreOffice`);
      }
    }
  }
  return { tiers, misses, skipped };
}

describe.runIf(process.env.FX)('debug', () => {
  it('dumps', async () => {
    const doc = readDocx(new Uint8Array(readFileSync(join(FIXTURES_DIR, process.env.FX!))));
    const l = await paginate(doc, { fonts: testFontResolver() });
    console.log(l.fonts.map((f) => `${f.postscriptName} ${f.outlines}`).join(', '));
    for (const p of l.pages) {
      console.log('--- page', p.index, p.width, p.height);
      const rows = new Map<number, string>();
      for (const it of p.items) {
        if (it.kind !== 'text') continue;
        const y = Math.round(it.y * 10) / 10;
        rows.set(y, `${rows.get(y) ?? `${it.x.toFixed(1)} `}${it.text}`);
      }
      for (const [y, t] of [...rows].sort((a, b) => a[0] - b[0])) {
        console.log(y.toFixed(1).padStart(6), t.slice(0, 100));
      }
    }
  });
  it('page breaks', async () => {
    const doc = readDocx(new Uint8Array(readFileSync(join(FIXTURES_DIR, process.env.FX!))));
    const l = await paginate(doc, { fonts: testFontResolver() });
    for (const b of wordBreaks(doc)) {
      const pg = (e: XElement | undefined) => (e === undefined ? '-' : l.paragraphPages.get(e));
      console.log(
        sameBreak(b, l.paragraphPages) ? 'ok ' : 'BAD',
        b.mid ? 'mid' : 'top',
        pg(b.prev),
        pg(b.p),
        pg(b.next),
        b.p.text.slice(0, 70),
      );
    }
  });
});

describe('layout fidelity scoreboard', () => {
  it('meets the measured baseline against Word and LibreOffice', async () => {
    const { tiers, misses, skipped } = await scoreboard();
    const table = (['A', 'B', 'C'] as const).map((t) => {
      const s = tiers[t];
      return {
        tier: t,
        docs: s.docs,
        pagesExact: pct(s.pagesExact, s.docs),
        wordBreaks: pct(s.breaksAgree, s.breaks),
        breaks: s.breaks,
        lineMatch: pct(s.rowsMatched, s.rows),
        rows: s.rows,
        libreOfficePagesExact: pct(s.loPagesExact, s.docs),
      };
    });
    console.log(
      table
        .map(
          (r) =>
            `Tier ${r.tier}: ${r.docs} docs, pages exact ${r.pagesExact}%, Word's page breaks ${r.wordBreaks}% (${r.breaks}), line match vs LibreOffice ${r.lineMatch}% (${r.rows} rows); LibreOffice pages exact ${r.libreOfficePagesExact}%`,
        )
        .join('\n'),
    );
    if (process.env.FIDELITY_REPORT) console.log([...misses, ...skipped].join('\n'));
    expect(skipped.length).toBeLessThanOrEqual(10);
    const [a, b, c] = table;
    // baseline (docs/fidelity.md); raise these when fidelity improves
    expect(a!.docs).toBeGreaterThanOrEqual(80);
    expect(a!.pagesExact).toBeGreaterThanOrEqual(BASELINE.A.pages);
    expect(a!.wordBreaks).toBeGreaterThanOrEqual(BASELINE.A.breaks);
    expect(a!.lineMatch).toBeGreaterThanOrEqual(BASELINE.A.lines);
    expect(b!.pagesExact).toBeGreaterThanOrEqual(BASELINE.B.pages);
    expect(b!.lineMatch).toBeGreaterThanOrEqual(BASELINE.B.lines);
    expect(c!.pagesExact).toBeGreaterThanOrEqual(BASELINE.C.pages);
    expect(c!.lineMatch).toBeGreaterThanOrEqual(BASELINE.C.lines);
  }, 300_000);
});

/**
 * Measured baseline (docs/fidelity.md) minus a small margin, percent. Tier A has only 3 soft page
 * breaks recorded by Word, so its break score moves in steps of a third.
 */
const BASELINE = {
  A: { pages: 97, breaks: 66, lines: 91 },
  B: { pages: 95, lines: 87 },
  C: { pages: 90, lines: 74 },
};
