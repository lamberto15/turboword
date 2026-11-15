// DrawingML shapes (every preset, custom geometry, groups, style matrix, colours, lines and
// arrowheads), legacy VML, SmartArt drawings and the drawingItems dispatcher.

import { describe, expect, it } from 'vitest';
import { createDocx, readDocx } from '../src/index.js';
import { colorOf, colorRgba, parseTheme, rgbaToHex } from '../src/layout/graphics/color.js';
import { evalFormula, presetGeometry } from '../src/layout/graphics/geometry.js';
import { drawingItems, presetNames, wantGraphicsFonts } from '../src/layout/graphics/index.js';
import { dashPath, Path } from '../src/layout/graphics/path.js';
import { drawingMLItems } from '../src/layout/graphics/shape.js';
import { runProps } from '../src/layout/graphics/text.js';
import { vmlItems } from '../src/layout/graphics/vml.js';
import type { DisplayItem, PathItem, TextItem } from '../src/layout/types.js';
import type { XElement } from '../src/xml/dom.js';
import { NS } from '../src/xml/ns.js';
import { parseXml } from '../src/xml/parse.js';
import { fakeEnv, fakeFontManager } from './graphics-env.js';
import { fixture } from './helpers/fixtures.js';

const DECL = [
  `xmlns:w="${NS.w}"`,
  `xmlns:wp="${NS.wp}"`,
  `xmlns:a="${NS.a}"`,
  `xmlns:pic="${NS.pic}"`,
  `xmlns:r="${NS.r}"`,
  `xmlns:wps="${NS.wps}"`,
  `xmlns:wpg="${NS.wpg}"`,
  `xmlns:wpc="${NS.wpc}"`,
  `xmlns:mc="${NS.mc}"`,
  `xmlns:v="${NS.v}"`,
  `xmlns:o="${NS.o}"`,
  `xmlns:m="${NS.m}"`,
].join(' ');

/** Parses one element (namespaces declared on a wrapper). */
function el(xml: string): XElement {
  const root = parseXml(`<root ${DECL}>${xml}</root>`).root;
  const first = root.firstElement();
  if (first === undefined) throw new Error('no element');
  return first;
}

const env = (): ReturnType<typeof fakeEnv> => fakeEnv(createDocx());
const paths = (items: DisplayItem[]): PathItem[] =>
  items.filter((i): i is PathItem => i.kind === 'path');
const numbers = (d: string): number[] => (d.match(/-?[\d.]+(?:e-?\d+)?/g) ?? []).map(Number);
/** True when every subpath of `d` is closed. */
const closedPath = (d: string): boolean =>
  d
    .split('M')
    .slice(1)
    .every((s) => s.trim().endsWith('Z'));

function wsp(
  prst: string,
  o: { cx?: number; cy?: number; xfrm?: string; extra?: string; av?: string } = {},
) {
  const cx = o.cx ?? 1828800;
  const cy = o.cy ?? 914400;
  return `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm${o.xfrm ?? ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst>${o.av ?? ''}</a:avLst></a:prstGeom>${
    o.extra ??
    '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>'
  }</wps:spPr><wps:bodyPr/></wps:wsp>`;
}

describe('preset geometry', () => {
  const names = presetNames();

  it('supports the common presets', () => {
    expect(names.length).toBeGreaterThanOrEqual(45);
    for (const n of ['rect', 'roundRect', 'star5', 'rightArrow', 'cloud', 'heart', 'can', 'cube']) {
      expect(names).toContain(n);
    }
  });

  it.each(names)('%s produces finite paths, filled ones closed', (name) => {
    const e = env();
    const items = drawingMLItems(el(wsp(name)), 144, 72, e, {});
    expect(items.length).toBeGreaterThan(0);
    for (const p of paths(items)) {
      const ns = numbers(p.d);
      expect(ns.length).toBeGreaterThan(0);
      expect(ns.every(Number.isFinite)).toBe(true);
      // Callout tails and arrowheads may leave the box, but not by much.
      expect(ns.every((v) => v > -200 && v < 400)).toBe(true);
      if (p.fill !== undefined) expect(closedPath(p.d), p.d).toBe(true);
    }
    // Lines and connectors are outlines only; everything else has a fill.
    const open = [
      'line',
      'straightConnector1',
      'bentConnector2',
      'bentConnector3',
      'curvedConnector2',
      'curvedConnector3',
      'arc',
      'leftBracket',
      'rightBracket',
      'leftBrace',
      'rightBrace',
    ];
    if (!open.includes(name)) expect(paths(items).some((p) => p.fill !== undefined)).toBe(true);
    expect(e.warnings).toEqual([]);
  });

  it('draws a rectangle exactly', () => {
    const [p] = paths(drawingMLItems(el(wsp('rect')), 144, 72, env(), {}));
    expect(p).toEqual({
      kind: 'path',
      d: 'M 0 0 L 144 0 L 144 72 L 0 72 Z',
      fill: 'FF0000',
      stroke: '0000FF',
      strokeWidth: 1,
    });
  });

  it('keeps an ellipse inside its box', () => {
    const [p] = paths(drawingMLItems(el(wsp('ellipse')), 144, 72, env(), {}));
    const ns = numbers(p?.d ?? '');
    const xs = ns.filter((_, i) => i % 2 === 0);
    const ys = ns.filter((_, i) => i % 2 === 1);
    expect(Math.min(...xs)).toBeCloseTo(0, 2);
    expect(Math.max(...xs)).toBeCloseTo(144, 2);
    expect(Math.min(...ys)).toBeCloseTo(0, 2);
    expect(Math.max(...ys)).toBeCloseTo(72, 2);
  });

  it('honours adjust values', () => {
    const [p] = paths(
      drawingMLItems(
        el(wsp('roundRect', { av: '<a:gd name="adj" fmla="val 50000"/>' })),
        144,
        72,
        env(),
        {},
      ),
    );
    expect(p?.d.startsWith('M 0 36 ')).toBe(true);
    const g = presetGeometry('roundRect', 100, 100, new Map([['adj', 'val 0']]));
    expect(g?.[0]?.path.toString().startsWith('M 0 0 ')).toBe(true);
  });

  it('applies rotation and flips to the path coordinates', () => {
    const [r] = paths(
      drawingMLItems(el(wsp('rect', { xfrm: ' rot="5400000"' })), 144, 72, env(), {}),
    );
    const ns = numbers(r?.d ?? '');
    const xs = ns.filter((_, i) => i % 2 === 0);
    const ys = ns.filter((_, i) => i % 2 === 1);
    expect([Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)]).toEqual([
      36, 108, -36, 108,
    ]);
    const [t] = paths(
      drawingMLItems(el(wsp('rtTriangle', { xfrm: ' flipH="1"' })), 144, 72, env(), {}),
    );
    expect(t?.d).toBe('M 144 72 L 144 0 L 0 72 Z');
    const [v] = paths(
      drawingMLItems(el(wsp('rtTriangle', { xfrm: ' flipV="1"' })), 144, 72, env(), {}),
    );
    expect(v?.d).toBe('M 0 0 L 0 72 L 144 0 Z');
  });

  it('evaluates every guide operator', () => {
    const vars = new Map([
      ['w', 100],
      ['h', 50],
    ]);
    expect(evalFormula('*/ w 3 4', vars)).toBe(75);
    expect(evalFormula('+- w h 10', vars)).toBe(140);
    expect(evalFormula('+/ w h 3', vars)).toBe(50);
    expect(evalFormula('?: -1 w h', vars)).toBe(50);
    expect(evalFormula('abs -7', vars)).toBe(7);
    expect(evalFormula('at2 0 10', vars)).toBe(5400000);
    expect(evalFormula('cos 10 0', vars)).toBe(10);
    expect(evalFormula('sin 10 5400000', vars)).toBeCloseTo(10, 9);
    expect(evalFormula('tan 10 2700000', vars)).toBeCloseTo(10, 9);
    expect(evalFormula('cat2 10 1 0', vars)).toBe(10);
    expect(evalFormula('sat2 10 0 1', vars)).toBe(10);
    expect(evalFormula('max w h', vars)).toBe(100);
    expect(evalFormula('min w h', vars)).toBe(50);
    expect(evalFormula('mod 3 4 0', vars)).toBe(5);
    expect(evalFormula('pin 0 120 w', vars)).toBe(100);
    expect(evalFormula('pin 0 -5 w', vars)).toBe(0);
    expect(evalFormula('sqrt 49', vars)).toBe(7);
    expect(evalFormula('val h', vars)).toBe(50);
    expect(evalFormula('*/ w 1 0', vars)).toBe(0);
  });
});

describe('custom geometry', () => {
  const cust = (paths: string, gd = '') =>
    `<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="1270000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst>${gd}</a:gdLst><a:pathLst>${paths}</a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></wps:spPr></wps:wsp>`;

  it('scales path coordinates by the path w/h', () => {
    const items = drawingMLItems(
      el(
        cust(
          '<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path>',
        ),
      ),
      200,
      100,
      env(),
      {},
    );
    expect(items).toEqual([{ kind: 'path', d: 'M 0 100 L 100 0 L 200 100 Z', fill: '00FF00' }]);
  });

  it('uses guides, arcs and Béziers in shape coordinates', () => {
    const items = drawingMLItems(
      el(
        cust(
          '<a:path><a:moveTo><a:pt x="0" y="vc2"/></a:moveTo><a:arcTo wR="hw" hR="vc2" stAng="10800000" swAng="10800000"/><a:quadBezTo><a:pt x="hw" y="1270000"/><a:pt x="0" y="vc2"/></a:quadBezTo><a:cubicBezTo><a:pt x="0" y="0"/><a:pt x="0" y="0"/><a:pt x="0" y="vc2"/></a:cubicBezTo><a:close/></a:path>',
          '<a:gd name="hw" fmla="*/ w 1 2"/><a:gd name="vc2" fmla="val vc"/>',
        ),
      ),
      200,
      100,
      env(),
      {},
    );
    const [p] = paths(items);
    expect(p?.d.startsWith('M 0 50 C')).toBe(true);
    expect(p?.d).toContain('200 50 C'); // the arc ends at the right edge
    const ys = numbers(p?.d ?? '').filter((_, i) => i % 2 === 1);
    expect(Math.min(...ys)).toBeCloseTo(0, 6); // half ellipse reaches the top
    expect(closedPath(p?.d ?? '')).toBe(true);
  });

  it('honours path fill="none" and stroke="false"', () => {
    const items = drawingMLItems(
      el(
        cust(
          '<a:path fill="none" w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="10"/></a:lnTo></a:path><a:path stroke="false" w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="0"/></a:lnTo><a:lnTo><a:pt x="10" y="10"/></a:lnTo><a:close/></a:path>',
        ).replace(
          '</a:custGeom>',
          '</a:custGeom><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
        ),
      ),
      100,
      100,
      env(),
      {},
    );
    expect(items).toEqual([
      { kind: 'path', d: 'M 0 0 L 100 100', stroke: '000000', strokeWidth: 2 },
      { kind: 'path', d: 'M 0 0 L 100 0 L 100 100 Z', fill: '00FF00' },
    ]);
  });
});

describe('groups', () => {
  const rect = (x: number, y: number, c: number) =>
    `<wps:wsp><wps:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${c}" cy="${c}"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="123456"/></a:solidFill></wps:spPr></wps:wsp>`;
  const group = (body: string, xfrm: string) =>
    `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm${xfrm}</a:xfrm></wpg:grpSpPr>${body}</wpg:wgp>`;
  const X = (off: number, ext: number, chOff: number, chExt: number) =>
    `><a:off x="${off}" y="${off}"/><a:ext cx="${ext}" cy="${ext}"/><a:chOff x="${chOff}" y="${chOff}"/><a:chExt cx="${chExt}" cy="${chExt}"/>`;

  it('maps child coordinates through chOff/chExt', () => {
    const g = el(group(rect(1270000, 1270000, 1270000), X(0, 1270000, 0, 2540000)));
    expect(paths(drawingMLItems(g, 100, 100, env(), {}))[0]?.d).toBe(
      'M 50 50 L 100 50 L 100 100 L 50 100 Z',
    );
    // The box scales the whole group.
    expect(paths(drawingMLItems(g, 200, 200, env(), {}))[0]?.d).toBe(
      'M 100 100 L 200 100 L 200 200 L 100 200 Z',
    );
  });

  it('nests groups with child offsets and keeps document order', () => {
    const inner = group(rect(1270000, 1270000, 635000), X(0, 1270000, 1270000, 1270000));
    const g = el(group(rect(0, 0, 254000) + inner, X(0, 1270000, 0, 1270000)));
    const ds = paths(drawingMLItems(g, 100, 100, env(), {})).map((p) => p.d);
    expect(ds).toEqual(['M 0 0 L 20 0 L 20 20 L 0 20 Z', 'M 0 0 L 50 0 L 50 50 L 0 50 Z']);
  });

  it('rotates a group about its centre', () => {
    const g = el(group(rect(0, 0, 635000), ` rot="10800000"${X(0, 1270000, 0, 1270000)}`));
    expect(paths(drawingMLItems(g, 100, 100, env(), {}))[0]?.d).toBe(
      'M 100 100 L 50 100 L 50 50 L 100 50 Z',
    );
  });

  it('inherits the group fill through a:grpFill', () => {
    const child = rect(0, 0, 635000).replace(
      '<a:solidFill><a:srgbClr val="123456"/></a:solidFill>',
      '<a:grpFill/>',
    );
    const g = el(
      group(
        child,
        `${X(0, 1270000, 0, 1270000)}</a:xfrm><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:xfrm>`,
      ).replace('<a:xfrm></a:xfrm>', ''),
    );
    expect(paths(drawingMLItems(g, 100, 100, env(), {}))[0]?.fill).toBe('ABCDEF');
  });

  it('draws deep nesting without recursion', () => {
    let body = rect(0, 0, 1270000);
    for (let i = 0; i < 300; i++) body = group(body, X(0, 1270000, 0, 1270000));
    const items = drawingMLItems(el(body), 100, 100, env(), {});
    expect(paths(items)[0]?.d).toBe('M 0 0 L 100 0 L 100 100 L 0 100 Z');
  });
});

describe('fills, lines and colours', () => {
  const theme = parseTheme(undefined);

  it('resolves colour transforms', () => {
    const c = (xml: string) => rgbaToHex(colorRgba(el(xml), theme) ?? { r: 0, g: 0, b: 0, a: 0 });
    expect(
      c('<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'),
    ).toBe('595959');
    expect(c('<a:schemeClr val="accent1"/>')).toBe('4472C4');
    expect(c('<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>')).toBe('FF8080');
    expect(c('<a:srgbClr val="808080"><a:tint val="50000"/></a:srgbClr>')).toBe('C0C0C0');
    expect(c('<a:srgbClr val="808080"><a:shade val="50000"/></a:srgbClr>')).toBe('404040');
    expect(c('<a:prstClr val="red"/>')).toBe('FF0000');
    expect(c('<a:sysClr val="windowText" lastClr="111111"/>')).toBe('111111');
    expect(c('<a:sysClr val="window"/>')).toBe('FFFFFF');
    expect(c('<a:scrgbClr r="100000" g="0" b="0"/>')).toBe('FF0000');
    expect(c('<a:hslClr hue="0" sat="100000" lum="50000"/>')).toBe('FF0000');
    expect(c('<a:srgbClr val="FF0000"><a:gray/></a:srgbClr>')).toBe('4C4C4C');
    expect(colorOf(el('<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'), theme)).toBe(
      'FFFFFF',
    );
  });

  it('reads theme colours, line widths, fonts and the settings mapping', () => {
    const t = parseTheme(
      el(
        '<a:theme><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="010101"/></a:dk1><a:lt1><a:srgbClr val="FEFEFE"/></a:lt1><a:accent1><a:srgbClr val="AA0000"/></a:accent1></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Major"/></a:majorFont><a:minorFont><a:latin typeface="Minor"/></a:minorFont></a:fontScheme><a:fmtScheme name="s"><a:lnStyleLst><a:ln w="9525"/><a:ln w="25400"/><a:ln w="38100"/></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>',
      ),
      el(`<w:clrSchemeMapping w:bg1="dark1" w:t1="light1"/>`),
    );
    expect(t.colors.accent1).toBe('AA0000');
    expect(t.colors.bg1).toBe('010101');
    expect(t.colors.tx1).toBe('FEFEFE');
    expect(t.lineWidths).toEqual([0.75, 2, 3]);
    expect(t.fonts).toEqual({ major: 'Major', minor: 'Minor' });
  });

  it('falls back to the style matrix', () => {
    const style =
      '<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>';
    const x = wsp('rect', { extra: '' }).replace('<wps:bodyPr/>', `${style}<wps:bodyPr/>`);
    expect(paths(drawingMLItems(el(x), 144, 72, env(), {}))[0]).toMatchObject({
      fill: '4472C4',
      stroke: '223962',
      strokeWidth: 1,
    });
    const noFill = x.replace('</a:prstGeom>', '</a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>');
    expect(drawingMLItems(el(noFill), 144, 72, env(), {})).toEqual([]);
  });

  it('approximates gradients and patterns', () => {
    const grad = wsp('rect', {
      extra:
        '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill>',
    });
    expect(paths(drawingMLItems(el(grad), 10, 10, env(), {}))[0]?.fill).toBe('808080');
    const patt = wsp('rect', {
      extra:
        '<a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="0000FF"/></a:bgClr></a:pattFill>',
    });
    expect(paths(drawingMLItems(el(patt), 10, 10, env(), {}))[0]?.fill).toBe('800080');
  });

  it('materialises dashes and draws arrowheads', () => {
    const ln = (inner: string) =>
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>${inner}</a:ln>`;
    const dashed = drawingMLItems(
      el(wsp('line', { cy: 0, extra: ln('<a:prstDash val="dash"/>') })),
      144,
      0,
      env(),
      {},
    );
    const d = paths(dashed)[0]?.d ?? '';
    expect(d.split('M').length - 1).toBe(Math.ceil(144 / 7)); // 4 on, 3 off, line width 1
    const arrows = paths(
      drawingMLItems(
        el(
          wsp('straightConnector1', {
            extra: ln('<a:headEnd type="oval"/><a:tailEnd type="triangle" w="lg" len="lg"/>'),
          }),
        ),
        144,
        72,
        env(),
        {},
      ),
    );
    expect(arrows).toHaveLength(3);
    const tri = arrows[2] as PathItem;
    expect(tri.fill).toBe('000000');
    expect(tri.d.startsWith('M 144 72 ')).toBe(true);
    const open = paths(
      drawingMLItems(
        el(wsp('line', { extra: ln('<a:tailEnd type="arrow"/>') })),
        100,
        0,
        env(),
        {},
      ),
    );
    expect(open[1]).toMatchObject({ stroke: '000000' });
    expect(open[1]?.fill).toBeUndefined();
  });

  it('dashes curves and closed paths', () => {
    const p = dashPath(new Path().ellipse(0, 0, 100, 100), [10, 10]);
    expect(p.cmds.filter((c) => c.op === 'M').length).toBeGreaterThan(10);
  });
});

describe('VML', () => {
  const vml = (xml: string, w = 100, h = 50) =>
    vmlItems(el(`<w:pict>${xml}</w:pict>`), w, h, env(), {}) ?? [];

  it('draws rect, oval and roundrect with fill and stroke attributes', () => {
    expect(
      vml(
        '<v:rect style="width:100pt;height:50pt" fillcolor="#00ff00" strokecolor="red" strokeweight="2pt"/>',
      ),
    ).toEqual([
      {
        kind: 'path',
        d: 'M 0 0 L 100 0 L 100 50 L 0 50 Z',
        fill: '00FF00',
        stroke: 'FF0000',
        strokeWidth: 2,
      },
    ]);
    const [o] = paths(vml('<v:oval style="width:100pt;height:50pt" filled="f" stroked="t"/>'));
    expect(o?.fill).toBeUndefined();
    expect(o?.stroke).toBe('000000');
    const [r] = paths(vml('<v:roundrect style="width:100pt;height:50pt" arcsize="10923f"/>'));
    expect(r?.fill).toBe('FFFFFF');
    expect(closedPath(r?.d ?? '')).toBe(true);
  });

  it('maps a v:shape path through coordsize', () => {
    const [p] = paths(
      vml(
        '<v:shape coordsize="21600,21600" path="m10800,l,10800,10800,21600,21600,10800xe" style="width:100pt;height:100pt" stroked="f"/>',
        100,
        100,
      ),
    );
    expect(p?.d).toBe('M 50 0 L 0 50 L 50 100 L 100 50 Z');
  });

  it('evaluates shapetype formulas and adjust values', () => {
    const type =
      '<v:shapetype id="_x0000_t5" coordsize="21600,21600" adj="10800" path="m@0,l,21600r21600,xe"><v:formulas><v:f eqn="val #0"/></v:formulas></v:shapetype>';
    const [a] = paths(
      vml(`${type}<v:shape type="#_x0000_t5" style="width:100pt;height:100pt"/>`, 100, 100),
    );
    expect(a?.d).toBe('M 50 0 L 0 100 L 100 100 Z');
    const [b] = paths(
      vml(
        `${type}<v:shape type="#_x0000_t5" adj="5400" style="width:100pt;height:100pt"/>`,
        100,
        100,
      ),
    );
    expect(b?.d).toBe('M 25 0 L 0 100 L 100 100 Z');
  });

  it('places group children in the group coordinate space', () => {
    const items = vml(
      '<v:group coordsize="1000,1000" style="width:200pt;height:100pt"><v:rect style="left:500;top:0;width:500;height:500" fillcolor="blue" stroked="f"/></v:group>',
      200,
      100,
    );
    expect(items).toEqual([
      { kind: 'path', d: 'M 100 0 L 200 0 L 200 50 L 100 50 Z', fill: '0000FF' },
    ]);
  });

  it('draws lines with dash styles and arrows', () => {
    const items = vml(
      '<v:line from="0,0" to="100pt,50pt"><v:stroke dashstyle="dash" endarrow="block"/></v:line>',
    );
    expect(paths(items).length).toBe(2);
    expect(paths(items)[0]?.d.split('M').length).toBeGreaterThan(3);
    expect(paths(items)[1]?.fill).toBe('000000');
    expect(paths(vml('<v:line from="0,0" to="100pt,50pt"/>'))[0]?.d).toBe('M 0 0 L 100 50');
  });

  it('supports arcs, curves and quadrants in paths', () => {
    const [p] = paths(
      vml(
        '<v:shape coordsize="100,100" path="m0,50qx50,0qy100,50wa0,0,100,100,100,50,0,50c0,60,10,70,20,80qb30,90,40,100xe" style="width:100pt;height:100pt"/>',
        100,
        100,
      ),
    );
    expect(numbers(p?.d ?? '').every(Number.isFinite)).toBe(true);
    expect(closedPath(p?.d ?? '')).toBe(true);
  });

  it('renders the VML fallback of a real file', () => {
    const doc = readDocx(fixture('apache-poi/shapes-with-text.docx').bytes);
    const picts = doc.mainPart.xml().root.descendants(NS.w, 'pict');
    expect(picts.length).toBeGreaterThan(0);
    const e = fakeEnv(doc);
    let drawn = 0;
    for (const p of picts) {
      const items = vmlItems(p, 200, 100, e, {}) ?? [];
      for (const it of paths(items)) expect(numbers(it.d).every(Number.isFinite)).toBe(true);
      if (items.length > 0) drawn++;
    }
    expect(drawn).toBeGreaterThan(0);
  });
});

describe('drawingItems', () => {
  it('draws the DrawingML shapes and groups of a real file', () => {
    const doc = readDocx(fixture('apache-poi/shapes-with-text.docx').bytes);
    const e = fakeEnv(doc);
    const drawings = doc.mainPart.xml().root.descendants(NS.w, 'drawing');
    expect(drawings.length).toBeGreaterThan(0);
    let drawn = 0;
    for (const d of drawings) {
      const ext = d.descendants(NS.wp, 'extent')[0];
      const w = Number(ext?.attr('cx') ?? 0) / 12700;
      const h = Number(ext?.attr('cy') ?? 0) / 12700;
      const items = drawingItems(d, w, h, e);
      if (items === undefined) continue;
      drawn++;
      for (const p of paths(items)) {
        const ns = numbers(p.d);
        expect(ns.every(Number.isFinite)).toBe(true);
        expect(ns.every((v) => v > -w - 10 && v < 2 * Math.max(w, h) + 10)).toBe(true);
      }
    }
    expect(drawn).toBeGreaterThan(0);
    expect(e.warnings).toEqual([]);
  });

  it('reads mc:AlternateContent, preferring the DrawingML branch', () => {
    const doc = readDocx(fixture('apache-poi/shapes-with-text.docx').bytes);
    const ac = doc.mainPart.xml().root.descendants(NS.mc, 'AlternateContent')[0];
    expect(ac).toBeDefined();
    const items = drawingItems(ac as XElement, 100, 100, fakeEnv(doc));
    expect(items?.length).toBeGreaterThan(0);
  });

  it('draws SmartArt from the diagram drawing part, with its text', () => {
    const doc = readDocx(fixture('docxcompose/smart_art.docx').bytes);
    const e = fakeEnv(doc);
    const d = doc.mainPart
      .xml()
      .root.descendants(NS.w, 'drawing')
      .find((x) => x.descendants(NS.dgm, 'relIds').length > 0);
    expect(d).toBeDefined();
    const ext = d?.descendants(NS.wp, 'extent')[0];
    const items =
      drawingItems(
        d as XElement,
        Number(ext?.attr('cx')) / 12700,
        Number(ext?.attr('cy')) / 12700,
        e,
      ) ?? [];
    expect(paths(items).length).toBeGreaterThan(2);
    const text = items
      .filter((i): i is TextItem => i.kind === 'text')
      .map((t) => t.text)
      .join('');
    const dataText = doc.package
      .parts()
      .filter((p) => /diagrams\/drawing/.test(p.name))
      .flatMap((p) => p.xml().root.descendants(NS.a, 't'))
      .map((t) => t.text)
      .join('')
      .replace(/\s/g, '');
    // This diagram's nodes are empty: no text, and none invented.
    expect(text.replace(/\s/g, '')).toBe(dataText);
  });

  it('lays out DrawingML text bodies (wrapped, aligned, anchored)', () => {
    const sp = el(
      '<dsp:sp xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="1270000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></dsp:spPr><dsp:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1000" b="1"/><a:t>Hello wide world of shapes</a:t></a:r></a:p><a:p><a:r><a:rPr sz="1000"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Courier"/></a:rPr><a:t>red</a:t></a:r></a:p></dsp:txBody></dsp:sp>',
    );
    const items = drawingMLItems(sp, 100, 100, env(), {});
    const texts = items.filter((i): i is TextItem => i.kind === 'text');
    expect(texts.map((t) => t.text).join('')).toBe(
      'Hello wide world of shapes red'.replace(' red', 'red'),
    );
    // 10 pt fake glyphs are 5 pt wide: "Hello wide world of " is 100 pt, so it wraps.
    const ys = [...new Set(texts.map((t) => t.y))];
    expect(ys.length).toBe(3);
    expect(texts[0]?.color).toBe('FFFFFF'); // fontRef colour
    expect(texts.at(-1)?.color).toBe('FF0000');
    // Centred block: the middle line's baseline sits near the box centre.
    expect(Math.abs((ys[1] ?? 0) - 50 - 3)).toBeLessThan(10);
    for (const t of texts) expect(t.x).toBeGreaterThanOrEqual(0);
  });

  it('returns undefined for plain pictures and unknown content, never throws', () => {
    const doc = readDocx(fixture('docxcompose/header_with_image.docx').bytes);
    const e = fakeEnv(doc);
    const header = doc.package.parts().find((p) => /header\d*\.xml$/.test(p.name));
    const pic = header?.xml().root.descendants(NS.w, 'drawing')[0];
    expect(pic).toBeDefined();
    expect(drawingItems(pic as XElement, 100, 100, e)).toBeUndefined();
    expect(drawingItems(el('<w:p/>'), 10, 10, e)).toBeUndefined();
    expect(
      drawingItems(
        el(
          '<w:drawing><wp:inline><a:graphic><a:graphicData uri="x"><foo/></a:graphicData></a:graphic></wp:inline></w:drawing>',
        ),
        10,
        10,
        e,
      ),
    ).toBeUndefined();
    expect(
      drawingItems(
        el(
          '<w:drawing><wp:inline><a:graphic><a:graphicData><wps:wsp><wps:spPr><a:xfrm><a:ext cx="abc"/></a:xfrm><a:prstGeom prst="nope"/></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>',
        ),
        10,
        10,
        e,
      ),
    ).toBeDefined();
    expect(e.warnings.some((w) => w.startsWith('UNSUPPORTED_PRESET_SHAPE'))).toBe(true);
  });

  it('draws pictures inside groups as image items', () => {
    const doc = readDocx(fixture('docxcompose/header_with_image.docx').bytes);
    const header = doc.package.parts().find((p) => /header\d*\.xml$/.test(p.name));
    const pic = header?.xml().root.descendants(NS.pic, 'pic')[0] as XElement;
    const group = header?.xml().createElement(NS.wpg, 'wgp') as XElement;
    group.append(header?.xml().createElement(NS.wpg, 'grpSpPr') as XElement, pic.clone());
    // The clone must still resolve r:embed through the header part.
    pic.parent?.insertAfter(group, pic);
    const e = fakeEnv(doc);
    const items = drawingItems(group, 100, 50, e) ?? [];
    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({ kind: 'image' });
    const id = (items[0] as { image: string }).image;
    expect(e.images.get(id)?.contentType).toMatch(/^image\//);
    group.remove();
  });

  it('registers the fonts drawings and math will need', () => {
    const doc = readDocx(fixture('docxcompose/smart_art.docx').bytes);
    const fm = fakeFontManager();
    const wanted: string[] = [];
    fm.want = (p) => wanted.push(p.fonts.ascii ?? '');
    const e = fakeEnv(doc, fm);
    const d = doc.mainPart.xml().root.descendants(NS.w, 'drawing')[0] as XElement;
    wantGraphicsFonts(d, runProps({ size: 11 }), e);
    wantGraphicsFonts(el('<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>'), runProps({ size: 11 }), e);
    expect(wanted).toContain('Cambria Math');
  });
});
