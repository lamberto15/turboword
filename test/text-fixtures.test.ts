// The text engine on the real fixture corpus (and its synthetic Strict copies): flatten agrees
// with the visible text everywhere, find/replace works across every story (text boxes
// included) with exact expected texts, outputs stay valid and reopen identically, and tracked
// replacements accept to the plain result and reject to the original.

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  NS,
  paragraphText,
  readDocx,
  validatePackage,
  type WordDocument,
  writeDocx,
} from '../src/index.js';
import {
  acceptAllRevisions,
  findText,
  flattenElement,
  rejectAllRevisions,
  replaceText,
  type TextMatch,
} from '../src/text/index.js';
import { storyTargets } from '../src/text/stories.js';
import { fixture, fixtures, strictFixtures } from './helpers/fixtures.js';
import { OUT_DIR, originalText, revisionsLeft } from './text-helpers.js';

const CORPUS = [...fixtures(), ...strictFixtures()];

/** Every searched paragraph of every story (text boxes included), as elements. */
function paragraphs(doc: WordDocument) {
  return doc.stories().flatMap((s) => storyTargets(s, true));
}

const errorsOf = (doc: WordDocument): string[] =>
  validatePackage(doc)
    .filter((i) => i.severity === 'error')
    .map((i) => `${i.code} ${i.part ?? ''}`)
    .sort();

function save(doc: WordDocument, name: string): Uint8Array {
  const bytes = writeDocx(doc);
  const file = join(OUT_DIR, 'fixtures', name.replace(/[\\/]/g, '__'));
  mkdirSync(join(OUT_DIR, 'fixtures'), { recursive: true });
  writeFileSync(file, bytes);
  return bytes;
}

/** Applies recorded replacements (start, end, value) to a string. */
function applyTo(text: string, edits: { start: number; end: number; value: string }[]): string {
  let out = text;
  for (const e of [...edits].sort((a, b) => b.start - a.start)) {
    out = out.slice(0, e.start) + e.value + out.slice(e.end);
  }
  return out;
}

describe.each(CORPUS.map((f) => [f.name, f] as const))('%s', (name, f) => {
  it('flatten() agrees with the visible text of every paragraph', () => {
    const doc = readDocx(f.bytes);
    for (const t of paragraphs(doc)) {
      const flat = flattenElement(t.el);
      expect(flat.text).toBe(paragraphText(t.el));
      let at = 0;
      for (const s of flat.segments) {
        expect(s.start).toBe(at);
        at = s.end;
      }
      expect(at).toBe(flat.text.length);
    }
  });

  it("replaces across all stories with exact results (in: 'all')", () => {
    const doc = readDocx(f.bytes);
    const before = errorsOf(doc);
    const targets = paragraphs(doc);
    const texts = new Map(targets.map((t) => [t.el, paragraphText(t.el)]));
    // Swap the first and last letters of every word of 3+ letters: matches span runs.
    const re = /\b(\p{L})(\p{L}+)(\p{L})\b/u;
    const found = findText(doc, re, { in: 'all' });
    const edits = new Map<unknown, { start: number; end: number; value: string }[]>();
    let warned = 0;
    const n = replaceText(
      doc,
      re,
      (m: TextMatch) => {
        const value = `${m.groups[2]}${m.groups[1]}${m.groups[0]}`;
        const list = edits.get(m.paragraph.element) ?? [];
        list.push({ start: m.start, end: m.end, value });
        edits.set(m.paragraph.element, list);
        return value;
      },
      { in: 'all', onWarning: () => warned++ },
    );
    expect(n + warned).toBe(found.length);
    for (const t of targets) {
      const want = applyTo(texts.get(t.el)!, edits.get(t.el) ?? []);
      expect(paragraphText(t.el)).toBe(want);
    }
    expect(errorsOf(doc)).toEqual(before);
    const back = readDocx(save(doc, `swap-${name}`));
    expect(paragraphs(back).map((t) => paragraphText(t.el))).toEqual(
      targets.map((t) => paragraphText(t.el)),
    );
  });

  it('a single character everywhere (many edits per run)', () => {
    const doc = readDocx(f.bytes);
    const targets = paragraphs(doc);
    const texts = targets.map((t) => paragraphText(t.el));
    const total = texts.reduce((a, s) => a + (s.match(/e/g)?.length ?? 0), 0);
    let warned = 0;
    const n = replaceText(doc, 'e', 'É', { in: 'all', onWarning: () => warned++ });
    expect(n + warned).toBe(total);
    if (warned === 0) {
      expect(targets.map((t) => paragraphText(t.el))).toEqual(
        texts.map((s) => s.replaceAll('e', 'É')),
      );
    }
    readDocx(save(doc, `e-${name}`));
  });

  it('tracked replace accepts to the plain result and rejects to the original', () => {
    const src = readDocx(f.bytes);
    if (revisionsLeft(src).length > 0) return; // revisions.test.ts covers documents with revisions
    const plain = readDocx(f.bytes);
    const re = /\b(\p{L})(\p{L}+)(\p{L})\b/u;
    replaceText(plain, re, '$3$2$1', { in: 'all' });
    const tracked = readDocx(f.bytes);
    const original = paragraphs(tracked).map((t) => paragraphText(t.el));
    replaceText(tracked, re, '$3$2$1', { in: 'all', track: { author: 'turboword' } });
    expect(paragraphs(tracked).map((t) => paragraphText(t.el))).toEqual(
      paragraphs(plain).map((t) => paragraphText(t.el)),
    );
    expect(paragraphs(tracked).map((t) => originalText(t.el))).toEqual(original);
    const back = readDocx(save(tracked, `tracked-${name}`));
    const accepted = back.fork();
    acceptAllRevisions(accepted);
    expect(revisionsLeft(accepted)).toEqual([]);
    expect(paragraphs(accepted).map((t) => paragraphText(t.el))).toEqual(
      paragraphs(plain).map((t) => paragraphText(t.el)),
    );
    const rejected = back.fork();
    rejectAllRevisions(rejected);
    expect(revisionsLeft(rejected)).toEqual([]);
    expect(paragraphs(rejected).map((t) => paragraphText(t.el))).toEqual(original);
    expect(errorsOf(accepted)).toEqual(errorsOf(src));
    expect(errorsOf(rejected)).toEqual(errorsOf(src));
    save(accepted, `tracked-accepted-${name}`);
    save(rejected, `tracked-rejected-${name}`);
  });
});

describe('fields in a real document (open-xml-sdk/TOC2.docx)', () => {
  it('skips matches crossing PAGEREF fields; replaces inside TOC and PAGEREF results', () => {
    const doc = readDocx(fixture('open-xml-sdk/TOC2.docx').bytes);
    const codes: string[] = [];
    // TOC entries read "<title>\t<page>": the page number is a PAGEREF result.
    expect(replaceText(doc, /\w\t1/, 'X', { onWarning: (w) => codes.push(w.code) })).toBe(0);
    expect(codes.length).toBeGreaterThan(0);
    expect(new Set(codes)).toEqual(new Set(['TEXT_FIELD_CROSSING']));
    const instr = doc.body.element.descendants(NS.w, 'instrText').map((e) => e.text);
    expect(replaceText(doc, /(?<=\t)1$/, '2')).toBe(6);
    expect(replaceText(doc, 'Asfasdf', 'Title')).toBe(2);
    expect(doc.body.element.descendants(NS.w, 'instrText').map((e) => e.text)).toEqual(instr);
    expect(doc.text()).toContain('Titlesfsadfasdf\t2');
    const back = readDocx(save(doc, 'toc-fields.docx'));
    expect(back.text()).toBe(doc.text());
  });
});
