// Every resolver entry point over every paragraph, run, table and cell of every curated fixture,
// their synthetic Strict variants and the numbering corpus of test/fixtures/resolve: nothing may
// throw, and the results must be sane (finite numbers, labels only on list paragraphs). Layout
// and compose call these on whatever documents users bring, so a crash here is a crash there.

import { readdirSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  listInfo,
  listLabel,
  resolveCell,
  resolveParagraph,
  resolveRun,
  resolveTable,
  themeOf,
} from '../src/resolve/index.js';
import { readDocx } from '../src/wml/document.js';
import type { Table } from '../src/wml/table.js';
import { FIXTURES_DIR, fixture, fixtures, strictFixtures } from './helpers/fixtures.js';

/** The numbering corpus of test/fixtures/resolve (not part of the curated corpus). */
function resolveFixtures(): ReturnType<typeof fixture>[] {
  return readdirSync(join(FIXTURES_DIR, 'resolve'))
    .filter((f) => f.endsWith('.docx'))
    .sort()
    .map((f) => fixture(`resolve/${f}`));
}

function allTables(tables: Table[]): Table[] {
  const out: Table[] = [];
  const stack = [...tables];
  while (stack.length > 0) {
    const t = stack.pop()!;
    out.push(t);
    for (const row of t.rows()) for (const cell of row.cells()) stack.push(...cell.tables());
  }
  return out;
}

describe('resolver sweep over the corpus', () => {
  for (const f of [...fixtures(), ...strictFixtures(), ...resolveFixtures()]) {
    it(f.name, () => {
      const doc = readDocx(f.bytes);
      expect(() => themeOf(doc)).not.toThrow();
      for (const story of doc.stories()) {
        for (const p of story.paragraphs()) {
          const props = resolveParagraph(p);
          expect(Number.isFinite(props.indent.left), 'indent').toBe(true);
          expect(Number.isFinite(props.spacing.line), 'line').toBe(true);
          expect(props.markRun.size).toBeGreaterThan(0);
          const info = listInfo(p);
          expect(listLabel(p)).toBe(info?.label);
          if (info !== undefined) {
            expect(props.numbering).toEqual({ numId: info.numId, level: info.level });
            expect(info.counters.length).toBe(info.level + 1);
          }
          for (const r of p.runs()) {
            const rp = resolveRun(r, props);
            expect(rp.size).toBeGreaterThan(0);
            expect(resolveRun(r)).toEqual(rp);
          }
        }
        for (const t of allTables(story.tables())) {
          const tp = resolveTable(t);
          expect(Number.isFinite(tp.cellMargins.left)).toBe(true);
          for (const row of t.rows()) {
            for (const cell of row.cells()) {
              const cp = resolveCell(cell);
              expect(Number.isFinite(cp.margins.top)).toBe(true);
            }
          }
        }
      }
    });
  }
});
