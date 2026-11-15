// DOM benchmark: parse + serialize with turboword's xml layer vs @xmldom/xmldom.
//
// The plan's exit criterion for v0.1 (decision D1, own DOM instead of xmldom) is ≥ 2× xmldom.
// Inputs: a large synthetic WordprocessingML document.xml (~8 MB, deterministic), and every XML
// part of the fixture corpus (Word-saved docx files) when available.
//
// Run: pnpm bench   (vitest bench --run). Uses the Vitest 5 benchmark API (`bench` fixture of the
// test context + `bench.compare`); each comparison prints the speed-up over xmldom.

import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { strFromU8, unzipSync } from 'fflate';
import { expect, test } from 'vitest';
import { NS, parseXml, serializeXml } from '../src/xml/index.js';

// ---- inputs -------------------------------------------------------------------------------

const ROOT_NS = [
  `xmlns:wpc="${NS.wpc}"`,
  `xmlns:mc="${NS.mc}"`,
  `xmlns:o="${NS.o}"`,
  `xmlns:r="${NS.r}"`,
  `xmlns:m="${NS.m}"`,
  `xmlns:v="${NS.v}"`,
  `xmlns:wp14="${NS.wp14}"`,
  `xmlns:wp="${NS.wp}"`,
  `xmlns:w10="${NS.w10}"`,
  `xmlns:w="${NS.w}"`,
  `xmlns:w14="${NS.w14}"`,
  `xmlns:w15="${NS.w15}"`,
  `xmlns:wpg="${NS.wpg}"`,
  `xmlns:wps="${NS.wps}"`,
  'mc:Ignorable="w14 w15 wp14"',
].join(' ');

const WORDS =
  'lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua'.split(
    ' ',
  );

/** Deterministic Word-like document.xml of roughly `targetBytes` characters. */
function syntheticDocument(targetBytes: number): string {
  let seed = 1;
  const next = (): number => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return seed;
  };
  const hex = (): string => (next() & 0xffffffff).toString(16).toUpperCase().padStart(8, '0');
  // Revision-session ids repeat (one per editing session); paragraph ids are unique.
  const rsids = Array.from({ length: 64 }, hex);
  const rsid = (): string => rsids[next() % rsids.length]!;
  const words = (n: number): string => {
    let s = '';
    for (let i = 0; i < n; i++) s += (i ? ' ' : '') + WORDS[next() % WORDS.length];
    return s;
  };
  const run = (): string => {
    const bold = next() % 4 === 0 ? '<w:b/><w:bCs/>' : '';
    const amp = next() % 10 === 0 ? ' &amp; &lt;tag&gt;' : '';
    return `<w:r w:rsidRPr="${rsid()}"><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/>${bold}<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr><w:t xml:space="preserve">${words(3 + (next() % 12))}${amp} </w:t></w:r>`;
  };
  const para = (): string => {
    let runs = '';
    const n = 1 + (next() % 5);
    for (let i = 0; i < n; i++) runs += run();
    return `<w:p w14:paraId="${hex()}" w14:textId="${hex()}" w:rsidR="${rsid()}" w:rsidRDefault="${rsid()}"><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:after="120" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>${runs}</w:p>`;
  };
  const table = (): string => {
    let rows = '';
    for (let r = 0; r < 4; r++) {
      let cells = '';
      for (let c = 0; c < 3; c++) {
        cells += `<w:tc><w:tcPr><w:tcW w:w="3116" w:type="dxa"/></w:tcPr>${para()}</w:tc>`;
      }
      rows += `<w:tr w:rsidR="${rsid()}">${cells}</w:tr>`;
    }
    return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3116"/><w:gridCol w:w="3117"/><w:gridCol w:w="3117"/></w:tblGrid>${rows}</w:tbl>`;
  };
  const chunks: string[] = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document ${ROOT_NS}><w:body>`,
  ];
  let size = chunks[0]!.length;
  for (let i = 0; size < targetBytes; i++) {
    const c = i % 25 === 24 ? table() : para();
    chunks.push(c);
    size += c.length;
  }
  chunks.push(
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
  );
  return chunks.join('');
}

function fixtureTexts(): string[] {
  const dirs = [
    fileURLToPath(new URL('../test/fixtures/', import.meta.url)),
    fileURLToPath(new URL('../../turbomerge/test/fixtures/', import.meta.url)),
  ];
  const files: string[] = [];
  const walk = (d: string): void => {
    if (!existsSync(d)) return;
    for (const f of readdirSync(d)) {
      const p = join(d, f);
      if (statSync(p).isDirectory()) walk(p);
      else if (f.endsWith('.docx')) files.push(p);
    }
  };
  for (const d of dirs) walk(d);
  const out: string[] = [];
  for (const f of files) {
    try {
      for (const [name, bytes] of Object.entries(unzipSync(readFileSync(f)))) {
        if (/\.(xml|rels)$/i.test(name))
          out.push(strFromU8(bytes.subarray(bytes[0] === 0xef ? 3 : 0)));
      }
    } catch {
      // not a zip: skip
    }
  }
  return out;
}

// ---- runners ------------------------------------------------------------------------------

const xmldomParser = new DOMParser({ onError: () => {} });
const xmldomSerializer = new XMLSerializer();

function turboword(x: string): number {
  return serializeXml(parseXml(x)).length;
}

function xmldom(x: string): number {
  return xmldomSerializer.serializeToString(xmldomParser.parseFromString(x, 'text/xml')).length;
}

// ---- benchmarks ---------------------------------------------------------------------------

const big = syntheticDocument(8 * 1024 * 1024);
const mb = (n: number): string => `${(n / 1024 / 1024).toFixed(1)} MB`;
// tinybench defaults to >= 64 iterations, far too many for an 8 MB document through xmldom.
const RUN = { time: 2000, iterations: 5, warmupTime: 0, warmupIterations: 1 };

/** Mean-latency ratio xmldom / turboword, printed and returned. */
function speedup(
  label: string,
  ours: { latency: { mean: number } },
  theirs: { latency: { mean: number } },
): number {
  const ratio = theirs.latency.mean / ours.latency.mean;
  console.log(
    `[bench] ${label}: turboword ${ours.latency.mean.toFixed(1)} ms, xmldom ${theirs.latency.mean.toFixed(1)} ms → ${ratio.toFixed(2)}×`,
  );
  return ratio;
}

test(`document.xml ${mb(big.length)}: parse + serialize`, async ({ bench }) => {
  const r = await bench.compare(
    bench('turboword', () => void turboword(big)),
    bench('@xmldom/xmldom', () => void xmldom(big)),
    RUN,
  );
  const ratio = speedup(
    `document.xml ${mb(big.length)} parse+serialize`,
    r.get('turboword'),
    r.get('@xmldom/xmldom'),
  );
  expect(ratio).toBeGreaterThanOrEqual(2); // plan exit criterion (D1)
});

test(`document.xml ${mb(big.length)}: parse only`, async ({ bench }) => {
  const r = await bench.compare(
    bench('turboword', () => void parseXml(big)),
    bench('@xmldom/xmldom', () => void xmldomParser.parseFromString(big, 'text/xml')),
    RUN,
  );
  speedup(`document.xml ${mb(big.length)} parse`, r.get('turboword'), r.get('@xmldom/xmldom'));
});

const parts = fixtureTexts();
const total = parts.reduce((n, p) => n + p.length, 0);

test.skipIf(parts.length === 0)(
  `fixture XML parts (${parts.length} parts, ${mb(total)}): parse + serialize`,
  async ({ bench }) => {
    const r = await bench.compare(
      bench('turboword', () => {
        for (const p of parts) turboword(p);
      }),
      bench('@xmldom/xmldom', () => {
        for (const p of parts) xmldom(p);
      }),
      { time: 2000, iterations: 10 },
    );
    speedup(
      `${parts.length} fixture parts parse+serialize`,
      r.get('turboword'),
      r.get('@xmldom/xmldom'),
    );
  },
);
