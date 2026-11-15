// Charts drawn from cached values: real Word charts (Apache POI test data: 61745.docx with a
// horizontal and a vertical clustered bar chart, chartex.docx with column, line, radar, stock and
// two chartex charts) and synthetic chart parts for the other chart types and options.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { createDocx, NS, readDocx } from '../src/index.js';
import { chartItems, chartRunProps } from '../src/layout/graphics/chart.js';
import { graphicsTheme } from '../src/layout/graphics/color.js';
import type { DisplayItem, PathItem, TextItem } from '../src/layout/types.js';
import type { WordDocument } from '../src/wml/document.js';
import type { XElement } from '../src/xml/dom.js';
import { fakeEnv } from './graphics-env.js';

const FIX = (name: string): Uint8Array =>
  new Uint8Array(
    readFileSync(fileURLToPath(new URL(`./fixtures/graphics/apache-poi/${name}`, import.meta.url))),
  );

const W = 432;
const H = 252;

function chartRefs(doc: WordDocument): XElement[] {
  return doc.mainPart.xml().root.descendants(NS.c, 'chart');
}

function partName(doc: WordDocument, ref: XElement): string {
  return doc.mainPart.related(ref.attrNS(NS.r, 'id') ?? '')?.name ?? '';
}

const texts = (items: DisplayItem[]): string[] =>
  items.filter((i): i is TextItem => i.kind === 'text').map((t) => t.text);

const nums = (d: string): number[] => (d.match(/-?[\d.]+(?:e-?\d+)?/g) ?? []).map(Number);

function bbox(d: string): [number, number, number, number] {
  const v = nums(d);
  const xs = v.filter((_, i) => i % 2 === 0);
  const ys = v.filter((_, i) => i % 2 === 1);
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}

/** Filled axis-aligned rectangles (M + 3 L + Z), excluding 7×7 legend swatches. */
function bars(items: DisplayItem[], fill?: string): { item: PathItem; box: number[] }[] {
  return items
    .filter((i): i is PathItem => i.kind === 'path' && i.fill !== undefined)
    .filter((p) => /^M [^A-Z]+ L [^A-Z]+ L [^A-Z]+ L [^A-Z]+ Z$/.test(p.d))
    .filter((p) => fill === undefined || p.fill === fill)
    .map((p) => ({ item: p, box: bbox(p.d) }))
    .filter(({ box }) => {
      const w = (box[2] ?? 0) - (box[0] ?? 0);
      const h = (box[3] ?? 0) - (box[1] ?? 0);
      return !(Math.abs(w - 7) < 1e-6 && Math.abs(h - 7) < 1e-6) && !(w >= W - 1 && h >= H - 1);
    });
}

function allFinite(items: DisplayItem[]): boolean {
  const walk = (v: unknown): boolean => {
    if (typeof v === 'number') return Number.isFinite(v);
    if (typeof v === 'string') return true;
    if (Array.isArray(v)) return v.every(walk);
    if (v !== null && typeof v === 'object') return Object.values(v).every(walk);
    return true;
  };
  return (
    walk(items) &&
    items.every(
      (i) => i.kind !== 'path' || (!/NaN|Infinity/.test(i.d) && nums(i.d).every(Number.isFinite)),
    )
  );
}

// ---- synthetic chart parts ----------------------------------------------------------------

const CHART_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function strCache(vals: string[]): string {
  return `<c:strCache><c:ptCount val="${vals.length}"/>${vals
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:strCache>`;
}

function numCache(vals: (number | null)[], fmt = 'General'): string {
  return `<c:numCache><c:formatCode>${fmt}</c:formatCode><c:ptCount val="${vals.length}"/>${vals
    .map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`))
    .join('')}</c:numCache>`;
}

function ser(
  idx: number,
  name: string,
  cats: string[],
  vals: (number | null)[],
  extra = '',
): string {
  return `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:strRef><c:f>S!A1</c:f>${strCache([name])}</c:strRef></c:tx>${extra}<c:cat><c:strRef><c:f>S!A2</c:f>${strCache(cats)}</c:strRef></c:cat><c:val><c:numRef><c:f>S!B2</c:f>${numCache(vals)}</c:numRef></c:val></c:ser>`;
}

const AXES = (cat = '1', val = '2', valExtra = '', catExtra = ''): string =>
  `<c:catAx><c:axId val="${cat}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="${val}"/><c:crosses val="autoZero"/>${catExtra}</c:catAx>` +
  `<c:valAx><c:axId val="${val}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cat}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>${valExtra}</c:valAx>`;

function chartXml(
  plot: string,
  opts: { title?: string; legend?: string; layout?: string } = {},
): string {
  const title =
    opts.title === undefined
      ? '<c:autoTitleDeleted val="1"/>'
      : `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${opts.title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
  const legend =
    opts.legend === undefined
      ? ''
      : `<c:legend><c:legendPos val="${opts.legend}"/><c:overlay val="0"/></c:legend>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace ${CHART_NS}><c:chart>${title}<c:plotArea><c:layout>${opts.layout ?? ''}</c:layout>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}

let chartNo = 0;
function addChart(doc: WordDocument, xml: string): XElement {
  const name = `word/charts/chart${++chartNo}.xml`;
  doc.package.addPart(
    name,
    'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    new TextEncoder().encode(xml),
  );
  const rel = doc.mainPart.rels.add(
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
    name.replace('word/', ''),
  );
  const ref = doc.mainPart.xml().createElement(NS.c, 'chart');
  ref.setAttrNS(NS.r, 'id', rel.id);
  return ref;
}

const CATS = ['A', 'B', 'C'];

// ---- real charts --------------------------------------------------------------------------

describe('charts from real documents', () => {
  it('draws the horizontal clustered bar chart of 61745.docx', () => {
    const doc = readDocx(FIX('61745.docx'));
    const env = fakeEnv(doc);
    const ref = chartRefs(doc).find((r) => partName(doc, r).endsWith('chart1.xml'));
    expect(ref).toBeDefined();
    const items = chartItems(ref as XElement, W, H, env) ?? [];
    expect(items.length).toBeGreaterThan(20);
    expect(allFinite(items)).toBe(true);
    const theme = graphicsTheme(doc);
    const accents = [theme.colors.accent1, theme.colors.accent2, theme.colors.accent3];
    const all = accents.flatMap((c) => bars(items, c));
    expect(all).toHaveLength(12); // 3 series × 4 categories
    for (const { box } of all) {
      expect(box[0]).toBeGreaterThanOrEqual(0);
      expect(box[1]).toBeGreaterThanOrEqual(0);
      expect(box[2]).toBeLessThanOrEqual(W);
      expect(box[3]).toBeLessThanOrEqual(H);
    }
    // Horizontal bars all start at the same x (value 0) and their lengths follow the values.
    const lefts = new Set(all.map(({ box }) => (box[0] ?? 0).toFixed(2)));
    expect(lefts.size).toBe(1);
    const s1 = bars(items, accents[0]).map(({ box }) => (box[2] ?? 0) - (box[0] ?? 0));
    const s3 = bars(items, accents[2]).map(({ box }) => (box[2] ?? 0) - (box[0] ?? 0));
    // Series 1 = 4.3, 2.5, 3.5, 4.5 and series 3 = 2, 2, 3, 5 (first category at the bottom).
    const len = (vals: number[]) => vals.map((v) => v.toFixed(1));
    const k = (s3[3] ?? 0) / 5;
    expect(len(s1.map((l) => l / k)).sort()).toEqual(len([4.3, 2.5, 3.5, 4.5]).sort());
    const t = texts(items);
    for (const c of ['Category 1', 'Category 2', 'Category 3', 'Category 4'])
      expect(t).toContain(c);
    for (const s of ['Series 1', 'Series 2', 'Series 3']) expect(t).toContain(s);
    expect(t).toContain('Chart Title');
    // Category labels are left of the bars, first category lowest.
    const cat1 = items.find((i): i is TextItem => i.kind === 'text' && i.text === 'Category 1');
    const cat4 = items.find((i): i is TextItem => i.kind === 'text' && i.text === 'Category 4');
    expect((cat1?.y ?? 0) > (cat4?.y ?? 0)).toBe(true);
    expect((cat1?.x ?? W) < Number(lefts.values().next().value)).toBe(true);
    expect(env.warnings).toEqual([]);
  });

  it('draws the clustered column chart of 61745.docx', () => {
    const doc = readDocx(FIX('61745.docx'));
    const env = fakeEnv(doc);
    const ref = chartRefs(doc).find((r) => partName(doc, r).endsWith('chart2.xml')) as XElement;
    const items = chartItems(ref, W, H, env) ?? [];
    const theme = graphicsTheme(doc);
    const cols = [theme.colors.accent1, theme.colors.accent2, theme.colors.accent3].flatMap((c) =>
      bars(items, c),
    );
    expect(cols).toHaveLength(12);
    // Columns share a baseline (bottom edge).
    expect(new Set(cols.map(({ box }) => (box[3] ?? 0).toFixed(2))).size).toBe(1);
    // Value axis labels: 0 .. something, as numbers.
    const t = texts(items);
    expect(t).toContain('0');
    expect(t.some((s) => /^[1-9]\d*(\.\d+)?$/.test(s))).toBe(true);
  });

  it('draws the chartex.docx charts and skips chartex (cx) parts', () => {
    const doc = readDocx(FIX('chartex.docx'));
    const env = fakeEnv(doc);
    const byPart = new Map(chartRefs(doc).map((r) => [partName(doc, r), r]));
    const draw = (n: number) =>
      chartItems(byPart.get(`word/charts/chart${n}.xml`) as XElement, W, H, env);
    // chart1: clustered column with a rich-text title.
    const c1 = draw(1) ?? [];
    const xml1 = new TextDecoder().decode(doc.package.part('word/charts/chart1.xml')?.bytes());
    const titleXml = /<c:title>.*?<\/c:title>/.exec(xml1)?.[0] ?? '';
    const title = [...titleXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('');
    expect(title).toBe('my chart looks nice'); // two runs
    expect(texts(c1)).toContain(title);
    expect(bars(c1).length).toBeGreaterThanOrEqual(9);
    // chart2: line chart, three stroked series, no markers (symbol none).
    const c2 = draw(2) ?? [];
    const theme = graphicsTheme(doc);
    const lines = c2.filter(
      (i): i is PathItem =>
        i.kind === 'path' &&
        i.fill === undefined &&
        (i.strokeWidth ?? 0) > 2 &&
        i.stroke !== undefined,
    );
    // Three series polylines (+ three 2-point legend swatches).
    expect(lines.filter((l) => nums(l.d).length >= 6)).toHaveLength(3);
    expect(lines.filter((l) => nums(l.d).length === 4)).toHaveLength(3);
    expect(lines.map((l) => l.stroke)).toContain(theme.colors.accent1);
    // chart3 radar and chart5 stock are drawn (radar as polygons, stock as hi-low lines + markers).
    expect((draw(3) ?? []).length).toBeGreaterThan(10);
    const c5 = draw(5) ?? [];
    expect(c5.length).toBeGreaterThan(10);
    expect(texts(c5).some((s) => /^\d+\/\d+\/\d{4}$/.test(s))).toBe(true); // date axis labels
    // chart4 / chart6 are cx:chartSpace (chartex): left to the fallback.
    expect(draw(4)).toBeUndefined();
    expect(draw(6)).toBeUndefined();
    for (const c of [c1, c2, c5]) expect(allFinite(c)).toBe(true);
  });

  it('is deterministic and reports the run props it uses', () => {
    const doc = readDocx(FIX('61745.docx'));
    const ref = chartRefs(doc)[0] as XElement;
    const a = JSON.stringify(chartItems(ref, W, H, fakeEnv(doc)));
    const b = JSON.stringify(chartItems(ref, W, H, fakeEnv(doc)));
    expect(a).toBe(b);
    const props = chartRunProps(ref, fakeEnv(doc));
    expect(props.length).toBeGreaterThan(0);
    expect(props.map((p) => p.size)).toContain(14); // title
    expect(props.map((p) => p.size)).toContain(9); // axis labels (sz=900)
    const minor = graphicsTheme(doc).fonts.minor;
    expect(props.every((p) => p.fonts.ascii === minor)).toBe(true);
  });
});

// ---- synthetic charts ---------------------------------------------------------------------

describe('synthetic charts', () => {
  it('stacks columns', () => {
    const doc = createDocx();
    const plot =
      `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>` +
      ser(0, 'S1', CATS, [1, 2, 3]) +
      ser(1, 'S2', CATS, [3, 2, 1]) +
      `<c:gapWidth val="150"/><c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${AXES()}`;
    const items =
      chartItems(addChart(doc, chartXml(plot, { legend: 'r' })), W, H, fakeEnv(doc)) ?? [];
    const theme = graphicsTheme(doc);
    const s1 = bars(items, theme.colors.accent1).map((b) => b.box);
    const s2 = bars(items, theme.colors.accent2).map((b) => b.box);
    expect(s1).toHaveLength(3);
    expect(s2).toHaveLength(3);
    for (let i = 0; i < 3; i++) {
      expect(s2[i]?.[3]).toBeCloseTo(s1[i]?.[1] ?? 0, 6); // S2 sits on S1
      expect(s2[i]?.[1]).toBeCloseTo(s2[0]?.[1] ?? 0, 6); // every stack totals 4
    }
    expect(texts(items)).toEqual(expect.arrayContaining(['S1', 'S2', 'A', 'B', 'C']));
  });

  it('normalises percent-stacked bars to 100%', () => {
    const doc = createDocx();
    const plot =
      `<c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/>` +
      ser(0, 'S1', CATS, [1, 5, 9]) +
      ser(1, 'S2', CATS, [1, 15, 1]) +
      `<c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${AXES()}`;
    const items = chartItems(addChart(doc, chartXml(plot)), W, H, fakeEnv(doc)) ?? [];
    const theme = graphicsTheme(doc);
    const right = bars(items, theme.colors.accent2).map((b) => (b.box[2] ?? 0).toFixed(3));
    expect(new Set(right).size).toBe(1);
    expect(texts(items)).toContain('100%');
    expect(texts(items)).toContain('0%');
  });

  it('draws pie and doughnut slices with varied colours and category legend', () => {
    const doc = createDocx();
    const pie = `<c:pieChart><c:varyColors val="1"/>${ser(0, 'Share', CATS, [50, 30, 20])}<c:firstSliceAng val="0"/></c:pieChart>`;
    const items =
      chartItems(addChart(doc, chartXml(pie, { legend: 'r', title: 'Pie' })), W, H, fakeEnv(doc)) ??
      [];
    const slices = items.filter(
      (i): i is PathItem => i.kind === 'path' && i.fill !== undefined && i.d.includes('C'),
    );
    expect(slices).toHaveLength(3);
    expect(new Set(slices.map((s) => s.fill)).size).toBe(3);
    // First slice starts at 12 o'clock: centre, then a point straight above it.
    const v = nums(slices[0]?.d ?? '');
    expect(v[2]).toBeCloseTo(v[0] ?? 0, 3);
    expect((v[3] ?? 0) < (v[1] ?? 0)).toBe(true);
    expect(texts(items)).toEqual(expect.arrayContaining(['Pie', 'A', 'B', 'C']));
    expect(allFinite(items)).toBe(true);

    const dn = `<c:doughnutChart><c:varyColors val="1"/>${ser(0, 'D', CATS, [1, 1, 2])}<c:firstSliceAng val="90"/><c:holeSize val="50"/></c:doughnutChart>`;
    const d = chartItems(addChart(doc, chartXml(dn)), W, H, fakeEnv(doc)) ?? [];
    const rings = d.filter(
      (i): i is PathItem => i.kind === 'path' && i.fill !== undefined && i.d.includes('C'),
    );
    expect(rings).toHaveLength(3);
    // Rings never touch the centre.
    const cx = W / 2;
    const cy = H / 2;
    for (const r of rings) {
      const p = nums(r.d);
      for (let i = 0; i + 1 < p.length; i += 2) {
        expect(Math.hypot((p[i] ?? 0) - cx, (p[i + 1] ?? 0) - cy)).toBeGreaterThan(20);
      }
    }
  });

  it('draws areas, scatter points and blank gaps', () => {
    const doc = createDocx();
    const area = `<c:areaChart><c:grouping val="standard"/>${ser(0, 'Ar', CATS, [1, 3, 2])}<c:axId val="1"/><c:axId val="2"/></c:areaChart>${AXES()}`;
    const a = chartItems(addChart(doc, chartXml(area)), W, H, fakeEnv(doc)) ?? [];
    const theme = graphicsTheme(doc);
    const poly = a.filter(
      (i): i is PathItem => i.kind === 'path' && i.fill === theme.colors.accent1,
    );
    expect(poly).toHaveLength(1);
    expect(poly[0]?.d.endsWith('Z')).toBe(true);

    const xs = (vals: number[]) =>
      `<c:xVal><c:numRef><c:f>S!A2</c:f>${numCache(vals)}</c:numRef></c:xVal>`;
    const ys = (vals: number[]) =>
      `<c:yVal><c:numRef><c:f>S!B2</c:f>${numCache(vals)}</c:numRef></c:yVal>`;
    const scatter =
      `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:marker><c:symbol val="square"/><c:size val="6"/></c:marker>${xs([1, 4, 10])}${ys([2, 8, 5])}<c:smooth val="0"/></c:ser><c:axId val="10"/><c:axId val="20"/></c:scatterChart>` +
      `<c:valAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="20"/><c:crosses val="autoZero"/></c:valAx>` +
      `<c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:tickLblPos val="nextTo"/><c:crossAx val="10"/><c:crosses val="autoZero"/></c:valAx>`;
    const s = chartItems(addChart(doc, chartXml(scatter)), W, H, fakeEnv(doc)) ?? [];
    const markers = bars(s, theme.colors.accent1).map((b) => b.box);
    expect(markers).toHaveLength(3);
    const cx = markers.map((m) => ((m[0] ?? 0) + (m[2] ?? 0)) / 2);
    // x positions proportional to 1, 4, 10 on a linear axis.
    expect((cx[1]! - cx[0]!) / (cx[2]! - cx[0]!)).toBeCloseTo(3 / 9, 3);
    expect(
      s.some((i) => i.kind === 'path' && i.stroke === theme.colors.accent1 && i.fill === undefined),
    ).toBe(true);

    const gap = `<c:lineChart><c:grouping val="standard"/>${ser(0, 'L', ['a', 'b', 'c', 'd', 'e'], [1, 2, null, 4, 5], '<c:marker><c:symbol val="none"/></c:marker>')}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${AXES()}`;
    const g = chartItems(addChart(doc, chartXml(gap)), W, H, fakeEnv(doc)) ?? [];
    const line = g.find(
      (i): i is PathItem => i.kind === 'path' && i.stroke === theme.colors.accent1,
    );
    expect(line?.d.match(/M/g)).toHaveLength(2);
  });

  it('combines bars and a line on a secondary axis', () => {
    const doc = createDocx();
    const plot =
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, 'Bars', CATS, [10, 20, 30])}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      `<c:lineChart><c:grouping val="standard"/>${ser(1, 'Line', CATS, [0.1, 0.5, 0.9])}<c:marker val="1"/><c:axId val="3"/><c:axId val="4"/></c:lineChart>` +
      AXES() +
      `<c:catAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="4"/></c:catAx>` +
      `<c:valAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="0%" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="3"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>`;
    const items =
      chartItems(addChart(doc, chartXml(plot, { legend: 'b' })), W, H, fakeEnv(doc)) ?? [];
    const theme = graphicsTheme(doc);
    const cols = bars(items, theme.colors.accent1);
    expect(cols).toHaveLength(3);
    const line = items.find(
      (i): i is PathItem =>
        i.kind === 'path' &&
        i.stroke === theme.colors.accent2 &&
        i.fill === undefined &&
        (i.strokeWidth ?? 0) > 2,
    );
    expect(line).toBeDefined();
    // The line's last point (0.9 on the secondary axis) is above the bars' baseline.
    const right = Math.max(...cols.map((c) => c.box[2] ?? 0));
    const pct = items.filter((i): i is TextItem => i.kind === 'text' && /%$/.test(i.text));
    expect(pct.length).toBeGreaterThan(2);
    for (const p of pct) expect(p.x).toBeGreaterThan(right);
    expect(texts(items)).toEqual(expect.arrayContaining(['Bars', 'Line']));
  });

  it('handles empty caches, deleted axes, manual layout and data labels', () => {
    const doc = createDocx();
    const empty = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>S!B2</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>${AXES()}`;
    const e = chartItems(addChart(doc, chartXml(empty)), W, H, fakeEnv(doc));
    expect(e).toBeDefined();
    expect(bars(e ?? [])).toHaveLength(0);
    expect(allFinite(e ?? [])).toBe(true);

    const deleted = AXES().replace(/<c:delete val="0"\/>/g, '<c:delete val="1"/>');
    const plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, 'S', CATS, [1, 2, 3])}<c:axId val="1"/><c:axId val="2"/></c:barChart>`;
    const d = chartItems(addChart(doc, chartXml(plot + deleted)), W, H, fakeEnv(doc)) ?? [];
    expect(texts(d)).toEqual([]);
    expect(bars(d)).toHaveLength(3);

    const layout =
      '<c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.5"/><c:y val="0.1"/><c:w val="0.4"/><c:h val="0.5"/></c:manualLayout>';
    const m =
      chartItems(addChart(doc, chartXml(plot + AXES(), { layout })), W, H, fakeEnv(doc)) ?? [];
    for (const { box } of bars(m)) {
      expect(box[0]).toBeGreaterThanOrEqual(0.5 * W - 1e-6);
      expect(box[2]).toBeLessThanOrEqual(0.9 * W + 1e-6);
      expect(box[3]).toBeLessThanOrEqual(0.6 * H + 1e-6);
      expect(box[1]).toBeGreaterThanOrEqual(0.1 * H - 1e-6);
    }

    const lbl = plot.replace(
      '<c:axId val="1"/>',
      '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls><c:axId val="1"/>',
    );
    const l = chartItems(addChart(doc, chartXml(lbl + deleted)), W, H, fakeEnv(doc)) ?? [];
    expect(texts(l)).toEqual(['1', '2', '3']);
  });

  it('returns undefined for missing or foreign parts', () => {
    const doc = createDocx();
    const ref = doc.mainPart.xml().createElement(NS.c, 'chart');
    expect(chartItems(ref, W, H, fakeEnv(doc))).toBeUndefined();
    ref.setAttrNS(NS.r, 'id', 'rId999');
    expect(chartItems(ref, W, H, fakeEnv(doc))).toBeUndefined();
    const cx = addChart(
      doc,
      '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart/></cx:chartSpace>',
    );
    expect(chartItems(cx, W, H, fakeEnv(doc))).toBeUndefined();
    expect(chartRunProps(cx, fakeEnv(doc))).toEqual([]);
    const noPlot = addChart(doc, chartXml('<c:surfaceChart/>'));
    expect(chartItems(noPlot, W, H, fakeEnv(doc))).toBeUndefined();
  });
});
