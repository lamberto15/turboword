// SVG renderer: well-formed standalone documents (parsed back with our own parser), per-glyph
// text positions, every display item kind, escaping, determinism, and an end-to-end drawing.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { readDocx } from '../src/index.js';
import { drawingItems } from '../src/layout/graphics/index.js';
import type { DisplayItem, DocumentLayout, PageLayout, TextItem } from '../src/layout/types.js';
import { renderSvg, renderSvgPages } from '../src/render/svg/index.js';
import type { XElement } from '../src/xml/dom.js';
import { NS } from '../src/xml/ns.js';
import { parseXml } from '../src/xml/parse.js';
import { fakeEnv, fakeFont } from './graphics-env.js';

const SVG = 'http://www.w3.org/2000/svg';
// 1×1 transparent PNG.
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

function layoutOf(items: DisplayItem[], extraPages = 0): DocumentLayout {
  const page = (index: number, its: DisplayItem[]): PageLayout => ({
    index,
    number: String(index + 1),
    section: 0,
    width: 612,
    height: 792,
    items: its,
  });
  return {
    pages: [page(0, items), ...Array.from({ length: extraPages }, (_, i) => page(i + 1, []))],
    fonts: [fakeFont('Carlito'), fakeFont('Caladea', true, true)],
    images: new Map([
      ['img', { id: 'img', contentType: 'image/png', bytes: PNG, pixelWidth: 1, pixelHeight: 1 }],
      ['emf', { id: 'emf', contentType: 'image/x-emf', bytes: new Uint8Array(4) }],
    ]),
    outline: [],
    destinations: new Map(),
    metadata: {},
    paragraphPages: new Map(),
  };
}

const text = (o: Partial<TextItem> = {}): TextItem => ({
  kind: 'text',
  x: 72,
  y: 100,
  font: 0,
  size: 12,
  color: '1F3864',
  glyphs: [72, 105],
  advances: [8, 4],
  text: 'Hi',
  clusters: [0, 1],
  ...o,
});

function parse(svg: string): XElement {
  const root = parseXml(svg).root;
  expect(root.is(SVG, 'svg')).toBe(true);
  return root;
}

describe('renderSvg', () => {
  it('writes a standalone SVG in points', () => {
    const layout = layoutOf([]);
    const svg = renderSvg(layout.pages[0] as PageLayout, layout);
    const root = parse(svg);
    expect(root.attr('width')).toBe('612pt');
    expect(root.attr('height')).toBe('792pt');
    expect(root.attr('viewBox')).toBe('0 0 612 792');
    expect(svg.startsWith('<?xml version="1.0" encoding="UTF-8"?>\n<svg ')).toBe(true);
  });

  it('positions every character from the glyph advances', () => {
    const layout = layoutOf([
      text(),
      text({
        font: 1,
        text: 'ffi!',
        glyphs: [1, 33],
        advances: [9, 3],
        clusters: [0, 3],
        color: 'auto',
      }),
      text({ text: 'ab', dy: [0, -2], scale: 80 }),
    ]);
    const root = parse(renderSvg(layout.pages[0] as PageLayout, layout));
    const [a, b, c] = root.elements(SVG, 'text');
    expect(a?.attr('x')).toBe('72 80');
    expect(a?.attr('y')).toBe('100');
    expect(a?.attr('font-family')).toBe('Carlito');
    expect(a?.attr('font-size')).toBe('12');
    expect(a?.attr('fill')).toBe('#1F3864');
    expect(a?.attr('font-weight')).toBeUndefined();
    expect(a?.text).toBe('Hi');
    // A ligature cluster: the characters inside it are spread over the glyph's advance.
    expect(b?.attr('x')).toBe('72 75 78 81');
    expect(b?.attr('font-weight')).toBe('bold');
    expect(b?.attr('font-style')).toBe('italic');
    expect(b?.attr('fill')).toBe('#000000');
    expect(c?.attr('y')).toBe('100 98');
    expect(c?.attr('transform')).toBe('matrix(0.8 0 0 1 14.4 0)');
    expect(c?.attrNS(NS.xml, 'space')).toBe('preserve');
  });

  it('draws images, rects, lines, paths and links', () => {
    const layout = layoutOf([
      { kind: 'image', x: 10, y: 20, width: 30, height: 40, image: 'img' },
      {
        kind: 'image',
        x: 10,
        y: 20,
        width: 30,
        height: 40,
        image: 'img',
        rotation: 90,
        crop: { left: 0.5, top: 0, right: 0, bottom: 0.5 },
      },
      { kind: 'image', x: 0, y: 0, width: 5, height: 5, image: 'emf' },
      { kind: 'rect', x: 1, y: 2, width: 3, height: 4, fill: 'ff0000' },
      { kind: 'rect', x: 1, y: 2, width: 3, height: 4, stroke: '00FF00', strokeWidth: 0.5 },
      { kind: 'line', x1: 0, y1: 0, x2: 10, y2: 10, color: '0000FF', width: 2, dash: [3, 1] },
      { kind: 'path', d: 'M 0 0 L 10 0 L 10 10 Z', fill: '123456', evenOdd: true },
      { kind: 'path', d: 'M 0 0 L 10 10', stroke: '654321', strokeWidth: 1.5 },
      { kind: 'link', x: 5, y: 6, width: 7, height: 8, url: 'https://example.com/?a=1&b="2"' },
      { kind: 'link', x: 5, y: 6, width: 7, height: 8, dest: '_Toc1' },
    ]);
    const svg = renderSvg(layout.pages[0] as PageLayout, layout);
    const root = parse(svg);
    const images = root.descendants(SVG, 'image');
    expect(images).toHaveLength(2);
    expect(
      images[0]
        ?.attrNS('http://www.w3.org/1999/xlink', 'href')
        ?.startsWith('data:image/png;base64,iVBOR'),
    ).toBe(true);
    expect(images[0]?.attr('preserveAspectRatio')).toBe('none');
    // Cropped: drawn twice as large, shifted left, clipped by a nested viewport.
    expect(images[1]?.attr('x')).toBe('-30');
    expect(images[1]?.attr('width')).toBe('60');
    expect(images[1]?.parent).toMatchObject({ name: 'svg' });
    expect(root.descendants(SVG, 'g')[0]?.attr('transform')).toBe('rotate(90 25 40)');
    const rects = root.elements(SVG, 'rect');
    expect(rects[1]?.attr('stroke')).toBe('#999999'); // EMF placeholder
    expect(rects[2]?.attr('fill')).toBe('#FF0000');
    expect(rects[3]?.attr('fill')).toBe('none');
    expect(rects[3]?.attr('stroke-width')).toBe('0.5');
    const line = root.element(SVG, 'line');
    expect(line?.attr('stroke-dasharray')).toBe('3 1');
    const [p1, p2] = root.elements(SVG, 'path');
    expect(p1?.attr('fill-rule')).toBe('evenodd');
    expect(p2?.attr('fill')).toBe('none');
    expect(p2?.attr('stroke')).toBe('#654321');
    const links = root.elements(SVG, 'a');
    expect(links[0]?.attrNS('http://www.w3.org/1999/xlink', 'href')).toBe(
      'https://example.com/?a=1&b="2"',
    );
    expect(links[1]?.attrNS('http://www.w3.org/1999/xlink', 'href')).toBe('#_Toc1');
  });

  it('escapes text and drops characters XML cannot hold', () => {
    const layout = layoutOf([
      text({
        text: 'a<&>"\u0001',
        glyphs: [1, 2, 3, 4, 5, 6],
        advances: [1, 1, 1, 1, 1, 1],
        clusters: [0, 1, 2, 3, 4, 5],
      }),
    ]);
    const root = parse(renderSvg(layout.pages[0] as PageLayout, layout));
    expect(root.element(SVG, 'text')?.text).toBe('a<&>"');
  });

  it('is deterministic and renders every page', () => {
    const layout = layoutOf(
      [text(), { kind: 'path', d: 'M 0 0 L 1 1', stroke: '000000', strokeWidth: 1 }],
      2,
    );
    const a = renderSvgPages(layout);
    const b = renderSvgPages(layout);
    expect(a).toHaveLength(3);
    expect(a).toEqual(b);
    for (const s of a) parse(s);
  });

  it('renders real drawings end to end', () => {
    const bytes = readFileSync(
      fileURLToPath(new URL('./fixtures/graphics/apache-poi/61745.docx', import.meta.url)),
    );
    const doc = readDocx(new Uint8Array(bytes));
    const env = fakeEnv(doc);
    const items: DisplayItem[] = [];
    for (const d of doc.mainPart.xml().root.descendants(NS.w, 'drawing')) {
      items.push(...(drawingItems(d, 400, 250, env) ?? []));
    }
    expect(items.some((i) => i.kind === 'text')).toBe(true);
    expect(items.some((i) => i.kind === 'path')).toBe(true);
    const layout = { ...layoutOf(items), fonts: env.fonts.fonts(), images: env.images };
    const svg = renderSvg(layout.pages[0] as PageLayout, layout);
    const root = parse(svg);
    expect(root.descendants(SVG, 'text').length).toBeGreaterThan(3);
  });
});
