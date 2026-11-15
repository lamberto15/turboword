// Graphics for layout: turns a w:drawing / w:pict that is not a plain picture into display items
// (DrawingML shapes and groups, SmartArt drawings, charts from cached values, EMF/WMF pictures,
// legacy VML), and lays out OMML math. Entry points for the layout engine; the work happens in
// the sibling modules (shape, geometry, vml, chart, emf/wmf, math).
//
// drawingItems never throws: graphics are best effort, so an unexpected input is reported through
// env.warn and the caller falls back to whatever it does for unsupported drawings.

import type { ResolvedRunProps } from '../../resolve/types.js';
import type { XElement } from '../../xml/dom.js';
import { localOf, NS } from '../../xml/ns.js';
import type { LayoutEnv, MathBox } from '../engine.js';
import type { DisplayItem, ImageResource } from '../types.js';
import { chartItems, chartRunProps } from './chart.js';
import { graphicsTheme } from './color.js';
import { emfItems } from './emf.js';
import { blipPart, imageContentType } from './images.js';
import { mathFontRequests, layoutMath as mathLayout } from './math.js';
import { drawingMLItems, kid, pickBranch, type ShapeHooks } from './shape.js';
import { runProps } from './text.js';
import { vmlItems } from './vml.js';
import { wmfItems } from './wmf.js';

const DSP_REL = 'diagramDrawing';

/** EMF/WMF bytes → items in a (width × height) box. */
function metafileItems(
  bytes: Uint8Array,
  contentType: string,
  width: number,
  height: number,
  env: LayoutEnv,
  id: string,
): DisplayItem[] | undefined {
  return contentType === 'image/x-emf'
    ? emfItems(bytes, width, height, env, id)
    : wmfItems(bytes, width, height, env, id);
}

const HOOKS: ShapeHooks = { chart: chartItems, metafile: metafileItems };

/** The SmartArt drawing (dsp:spTree) of a dgm:relIds, via the data part's dataModelExt relId. */
function smartArtTree(relIds: XElement, env: LayoutEnv): XElement | undefined {
  const part = env.partOf(relIds);
  if (part === undefined) return undefined;
  const dm = relIds.attrNS(NS.r, 'dm');
  const data = dm === undefined ? undefined : part.related(dm);
  let drawing = undefined as ReturnType<typeof part.related>;
  if (data !== undefined) {
    const ext = data
      .xml()
      .root.descendants()
      .find((e) => localOf(e.name) === 'dataModelExt');
    const rid = ext?.attr('relId');
    if (rid !== undefined) drawing = part.related(rid);
  }
  if (drawing === undefined) {
    const all = part.relatedByType(DSP_REL);
    if (all.length === 1) drawing = all[0];
  }
  return kid(drawing?.xml().root, 'spTree');
}

/** Items for an a:graphicData. */
function graphicDataItems(
  data: XElement,
  width: number,
  height: number,
  env: LayoutEnv,
): DisplayItem[] | undefined {
  for (const child of data.elements()) {
    const local = localOf(child.name);
    if (child.is(NS.pic, 'pic')) {
      const blip = kid(kid(child, 'blipFill'), 'blip');
      const part = blip === undefined ? undefined : blipPart(blip, env);
      if (part === undefined) return undefined;
      const ct = imageContentType(part);
      if (ct !== 'image/x-emf' && ct !== 'image/x-wmf') return undefined; // plain picture
      return metafileItems(part.bytes(), ct, width, height, env, part.name);
    }
    if (child.is(NS.c, 'chart')) return chartItems(child, width, height, env);
    if (child.is(NS.dgm, 'relIds')) {
      const tree = smartArtTree(child, env);
      return tree === undefined ? undefined : drawingMLItems(tree, width, height, env, HOOKS);
    }
    if (local === 'wsp' || local === 'wgp' || local === 'wpc') {
      return drawingMLItems(child, width, height, env, HOOKS);
    }
  }
  return undefined;
}

function dispatch(
  el: XElement,
  width: number,
  height: number,
  env: LayoutEnv,
): DisplayItem[] | undefined {
  let cur: XElement | undefined = el;
  for (let guard = 0; cur !== undefined && guard < 16; guard++) {
    const local = localOf(cur.name);
    if (cur.uri === NS.v) return vmlItems(cur, width, height, env, HOOKS);
    switch (local) {
      case 'AlternateContent': {
        const branch: XElement | undefined = cur.uri === NS.mc ? pickBranch(cur) : undefined;
        cur = branch?.elements()[0];
        continue;
      }
      case 'r':
      case 'drawing':
      case 'Choice':
      case 'Fallback':
        cur = cur
          .elements()
          .find((e) =>
            ['inline', 'anchor', 'drawing', 'pict', 'object', 'AlternateContent'].includes(
              localOf(e.name),
            ),
          );
        continue;
      case 'inline':
      case 'anchor':
        cur = kid(cur, 'graphic');
        continue;
      case 'graphic':
        cur = kid(cur, 'graphicData');
        continue;
      case 'graphicData':
        return graphicDataItems(cur, width, height, env);
      case 'pict':
      case 'object':
        return vmlItems(cur, width, height, env, HOOKS);
      case 'wsp':
      case 'wgp':
      case 'wpc':
      case 'pic':
        return drawingMLItems(cur, width, height, env, HOOKS);
      default:
        return undefined;
    }
  }
  return undefined;
}

/**
 * Display items for a w:drawing / w:pict that is not a plain picture (DrawingML preset/custom
 * shapes with fill/line/text, charts drawn from their cached values, EMF/WMF pictures converted to
 * paths, group shapes), positioned relative to the box top-left; undefined when unsupported.
 */
export function drawingItems(
  el: XElement,
  width: number,
  height: number,
  env: LayoutEnv,
): DisplayItem[] | undefined {
  try {
    return dispatch(el, width, height, env);
  } catch (e) {
    env.warn(
      'GRAPHICS_FAILED',
      `Drawing not rendered: ${e instanceof Error ? e.message : String(e)}`,
    );
    return undefined;
  }
}

/**
 * Vector items for an EMF/WMF image resource drawn in a (width × height) box (undefined for
 * other image types, which renderers draw as they are). For pictures the layout already turned
 * into ImageItems: PDF and SVG cannot show metafiles directly.
 */
export function metafileImageItems(
  image: ImageResource,
  width: number,
  height: number,
  env: LayoutEnv,
): DisplayItem[] | undefined {
  const ct = image.contentType;
  if (ct !== 'image/x-emf' && ct !== 'image/x-wmf') return undefined;
  try {
    return metafileItems(image.bytes, ct, width, height, env, image.id);
  } catch (e) {
    env.warn(
      'GRAPHICS_FAILED',
      `Metafile not rendered: ${e instanceof Error ? e.message : String(e)}`,
    );
    return undefined;
  }
}

/** Lays out an m:oMath / m:oMathPara (fractions, scripts, radicals, n-ary, delimiters, matrices). */
export function layoutMath(el: XElement, props: ResolvedRunProps, env: LayoutEnv): MathBox {
  return mathLayout(el, props, env);
}

/**
 * Registers with `env.fonts` the faces a drawing or math zone will ask for (chart text, DrawingML
 * text bodies, Cambria Math), so they are loaded before layout. Call during font collection, for
 * `drawing` and `math` inlines/anchors; harmless for anything else.
 */
export function wantGraphicsFonts(
  el: XElement,
  props: ResolvedRunProps,
  env: Pick<LayoutEnv, 'doc' | 'partOf' | 'fonts'>,
): void {
  try {
    if (el.is(NS.m, 'oMath') || el.is(NS.m, 'oMathPara')) {
      for (const p of mathFontRequests(props)) env.fonts.want(p);
      return;
    }
    for (const c of el.descendants(NS.c, 'chart')) {
      for (const p of chartRunProps(c, env)) env.fonts.want(p);
    }
    const theme = graphicsTheme(env.doc);
    const fams = new Set<string>();
    for (const r of el.descendants(NS.a, 'latin')) {
      const f = r.attr('typeface');
      if (f === '+mn-lt' && theme.fonts.minor) fams.add(theme.fonts.minor);
      else if (f === '+mj-lt' && theme.fonts.major) fams.add(theme.fonts.major);
      else if (f !== undefined && !f.startsWith('+')) fams.add(f);
    }
    if (el.descendants(NS.a, 't').length > 0 && theme.fonts.minor) fams.add(theme.fonts.minor);
    for (const f of [...fams].sort()) {
      for (const bold of [false, true]) env.fonts.want(runProps({ font: f, bold }));
    }
  } catch {
    // Best effort: missing faces fall back at layout time.
  }
}

export { presetNames } from './geometry.js';
