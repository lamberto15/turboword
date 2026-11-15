// Drawings in the IR: `w:drawing` (DrawingML), `w:pict` and `w:object` (VML, OLE previews).
//
// - `wp:inline` with a picture (`pic:pic` → `a:blip r:embed`) → an image inline; any other inline
//   graphic (chart, SmartArt, shape, group, canvas) → a `drawing` inline the graphics module draws.
// - `wp:anchor` → an IRAnchor (position, wrap, distances, z-order), holding an image, a text box
//   (`wps:txbx/w:txbxContent`, whose blocks are built by the caller) or a drawing.
// - VML: a shape with `v:imagedata` is a picture, one with `v:textbox` a text box; `position:absolute`
//   in its style makes it floating. Sizes come from the CSS-like `style` attribute.
// Sizes are converted to points (EMU / 12700).

import type { ResolvedRunProps } from '../../resolve/types.js';
import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import type { IRAnchor, IRBlock, IRInline, IRLink } from '../types.js';
import type { IRAnchorExtras, IRInlineExtras } from './extras.js';
import { cssLengthPt, emuToPt, intAttr, onOffValue, parseCssStyle } from './ir-util.js';

/** What the drawing reader needs from the IR builder. */
export interface DrawingContext {
  /** Registers the image a relationship id (of the current part) points to; returns its IR id. */
  image(relId: string): string | undefined;
  /** Schedules a text box content element; returns the block array the builder will fill. */
  textbox(content: XElement): IRBlock[];
}

export type DrawingResult = { inline: IRInline } | { anchor: IRAnchor } | undefined;

const WPS_URI = NS.wps;

/** `a:srcRect` crop as fractions, or undefined when absent/zero. */
function cropOf(blipFill: XElement | undefined): IRInlineExtras['crop'] {
  const r = blipFill?.element(NS.a, 'srcRect');
  if (r === undefined) return undefined;
  const f = (k: string): number => intAttr(r, '', k, 0) / 100000;
  const c = { left: f('l'), top: f('t'), right: f('r'), bottom: f('b') };
  return c.left || c.top || c.right || c.bottom ? c : undefined;
}

/** The picture blip and its fill of an `a:graphicData`, if it is a picture. */
function pictureOf(gd: XElement | undefined): { embed: string; blipFill: XElement } | undefined {
  const pic = gd?.element(NS.pic, 'pic');
  const blipFill = pic?.element(NS.pic, 'blipFill');
  const embed = blipFill?.element(NS.a, 'blip')?.attrNS(NS.r, 'embed');
  if (blipFill === undefined || embed === undefined || embed === '') return undefined;
  return { embed, blipFill };
}

function extentOf(container: XElement): { width: number; height: number } {
  const ext = container.element(NS.wp, 'extent');
  return {
    width: emuToPt(Math.max(0, intAttr(ext, '', 'cx', 0))),
    height: emuToPt(Math.max(0, intAttr(ext, '', 'cy', 0))),
  };
}

/** Text box content of a DrawingML shape (`wps:wsp/wps:txbx/w:txbxContent`). */
function txbxContentOf(gd: XElement | undefined): XElement | undefined {
  return gd?.element(WPS_URI, 'wsp')?.element(WPS_URI, 'txbx')?.element(NS.w, 'txbxContent');
}

const WRAPS: readonly [string, IRAnchor['wrap']][] = [
  ['wrapNone', 'none'],
  ['wrapSquare', 'square'],
  ['wrapTight', 'tight'],
  ['wrapThrough', 'through'],
  ['wrapTopAndBottom', 'topAndBottom'],
];

function positionOf(el: XElement | undefined, dflt: string): IRAnchor['horizontal'] {
  const out: IRAnchor['horizontal'] = { relativeTo: el?.attr('relativeFrom') ?? dflt };
  const off = el?.element(NS.wp, 'posOffset')?.text;
  if (off !== undefined && /^\s*-?\d+\s*$/.test(off))
    out.offset = emuToPt(Number.parseInt(off, 10));
  const align = el?.element(NS.wp, 'align')?.text.trim();
  if (align !== undefined && align !== '') out.align = align;
  if (out.offset === undefined && out.align === undefined) out.offset = 0;
  return out;
}

function anchorFromDrawingML(drawing: XElement, anchor: XElement, ctx: DrawingContext): IRAnchor {
  const { width, height } = extentOf(anchor);
  const dist = (k: string): number => emuToPt(Math.max(0, intAttr(anchor, '', k, 0)));
  let wrap: IRAnchor['wrap'] = 'none';
  for (const [local, w] of WRAPS) {
    if (anchor.element(NS.wp, local) !== undefined) {
      wrap = w;
      break;
    }
  }
  const simple = onOffValue(anchor.attr('simplePos'), false);
  const sp = anchor.element(NS.wp, 'simplePos');
  const horizontal = simple
    ? { relativeTo: 'page', offset: emuToPt(intAttr(sp, '', 'x', 0)) }
    : positionOf(anchor.element(NS.wp, 'positionH'), 'column');
  const vertical = simple
    ? { relativeTo: 'page', offset: emuToPt(intAttr(sp, '', 'y', 0)) }
    : positionOf(anchor.element(NS.wp, 'positionV'), 'paragraph');
  const out: IRAnchor & IRAnchorExtras = {
    kind: 'drawing',
    source: drawing,
    width,
    height,
    horizontal,
    vertical,
    wrap,
    behindText: onOffValue(anchor.attr('behindDoc'), false),
    distance: {
      top: dist('distT'),
      right: dist('distR'),
      bottom: dist('distB'),
      left: dist('distL'),
    },
    zIndex: intAttr(anchor, '', 'relativeHeight', 0),
  };
  const gd = anchor.element(NS.a, 'graphic')?.element(NS.a, 'graphicData');
  const pic = pictureOf(gd);
  if (pic !== undefined) {
    const id = ctx.image(pic.embed);
    if (id !== undefined) {
      out.kind = 'image';
      out.image = id;
      const crop = cropOf(pic.blipFill);
      if (crop !== undefined) out.crop = crop;
    }
    return out;
  }
  const content = txbxContentOf(gd);
  if (content !== undefined) {
    out.kind = 'textbox';
    out.blocks = ctx.textbox(content);
    const body = gd?.element(WPS_URI, 'wsp')?.element(WPS_URI, 'bodyPr');
    const ins = (k: string, d: number): number => emuToPt(intAttr(body, '', k, d));
    out.insets = {
      left: ins('lIns', 91440),
      top: ins('tIns', 45720),
      right: ins('rIns', 91440),
      bottom: ins('bIns', 45720),
    };
  }
  return out;
}

function fromDrawing(
  drawing: XElement,
  props: ResolvedRunProps,
  link: IRLink | undefined,
  ctx: DrawingContext,
): DrawingResult {
  const inline = drawing.element(NS.wp, 'inline');
  if (inline !== undefined) {
    const { width, height } = extentOf(inline);
    const gd = inline.element(NS.a, 'graphic')?.element(NS.a, 'graphicData');
    const pic = pictureOf(gd);
    if (pic !== undefined) {
      const id = ctx.image(pic.embed);
      if (id !== undefined) {
        const img: IRInline & IRInlineExtras = { kind: 'image', image: id, width, height, props };
        if (link !== undefined) img.link = link;
        const crop = cropOf(pic.blipFill);
        if (crop !== undefined) img.crop = crop;
        return { inline: img };
      }
    }
    return { inline: { kind: 'drawing', source: drawing, width, height, props } };
  }
  const anchor = drawing.element(NS.wp, 'anchor');
  if (anchor !== undefined) return { anchor: anchorFromDrawingML(drawing, anchor, ctx) };
  return undefined;
}

/** Shape elements of VML that can carry a picture or a text box. */
const VML_SHAPES: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'shape',
  'rect',
  'roundrect',
  'oval',
  'line',
  'polyline',
  'arc',
  'curve',
  'image',
  'group',
]);

function vmlRelative(v: string | undefined, horizontal: boolean): string {
  switch (v) {
    case 'margin':
      return 'margin';
    case 'page':
      return 'page';
    case 'char':
      return horizontal ? 'character' : 'line';
    case 'line':
      return 'line';
    default:
      return horizontal ? 'column' : 'paragraph';
  }
}

function fromVml(
  container: XElement,
  props: ResolvedRunProps,
  link: IRLink | undefined,
  ctx: DrawingContext,
): DrawingResult {
  const shape = container.elements(NS.v).find((e) => VML_SHAPES.has(e.local));
  if (shape === undefined) return undefined;
  const style = parseCssStyle(shape.attr('style'));
  const width = Math.max(0, cssLengthPt(style.get('width')) ?? 0);
  const height = Math.max(0, cssLengthPt(style.get('height')) ?? 0);
  const imagedata = shape.element(NS.v, 'imagedata');
  const relId = imagedata?.attrNS(NS.r, 'id') ?? imagedata?.attrNS(NS.o, 'relid');
  const image = relId !== undefined && relId !== '' ? ctx.image(relId) : undefined;
  const content = shape.element(NS.v, 'textbox')?.element(NS.w, 'txbxContent');
  const floating = style.get('position') === 'absolute';
  if (!floating) {
    if (image !== undefined) {
      const img: IRInline = { kind: 'image', image, width, height, props };
      if (link !== undefined) img.link = link;
      return { inline: img };
    }
    if (width <= 0 && height <= 0 && content === undefined) return undefined;
    return { inline: { kind: 'drawing', source: container, width, height, props } };
  }
  const z = Number.parseInt(style.get('z-index') ?? '0', 10);
  const wrapType = container.descendants(NS.w10, 'wrap')[0]?.attr('type');
  const wrap: IRAnchor['wrap'] =
    wrapType === 'square' || wrapType === 'tight' || wrapType === 'through'
      ? wrapType
      : wrapType === 'topAndBottom'
        ? 'topAndBottom'
        : 'none';
  const pos = (k: string): string | undefined => {
    const v = style.get(k);
    return v === undefined || v === 'absolute' ? undefined : v;
  };
  const horizontal: IRAnchor['horizontal'] = {
    relativeTo: vmlRelative(style.get('mso-position-horizontal-relative'), true),
    offset: cssLengthPt(style.get('margin-left') ?? style.get('left')) ?? 0,
  };
  const hAlign = pos('mso-position-horizontal');
  if (hAlign !== undefined) horizontal.align = hAlign;
  const vertical: IRAnchor['vertical'] = {
    relativeTo: vmlRelative(style.get('mso-position-vertical-relative'), false),
    offset: cssLengthPt(style.get('margin-top') ?? style.get('top')) ?? 0,
  };
  const vAlign = pos('mso-position-vertical');
  if (vAlign !== undefined) vertical.align = vAlign;
  const d = (k: string): number => Math.max(0, cssLengthPt(style.get(k)) ?? 0);
  const out: IRAnchor = {
    kind: image !== undefined ? 'image' : content !== undefined ? 'textbox' : 'drawing',
    source: container,
    width,
    height,
    horizontal,
    vertical,
    wrap,
    behindText: Number.isFinite(z) && z < 0,
    distance: {
      top: d('mso-wrap-distance-top'),
      right: d('mso-wrap-distance-right'),
      bottom: d('mso-wrap-distance-bottom'),
      left: d('mso-wrap-distance-left'),
    },
    zIndex: Number.isFinite(z) ? z : 0,
  };
  if (image !== undefined) out.image = image;
  else if (content !== undefined) out.blocks = ctx.textbox(content);
  return { anchor: out };
}

/** IR for a `w:drawing`, `w:pict` or `w:object` met in a run. */
export function drawingToIR(
  el: XElement,
  props: ResolvedRunProps,
  link: IRLink | undefined,
  ctx: DrawingContext,
): DrawingResult {
  if (el.local === 'drawing') return fromDrawing(el, props, link, ctx);
  if (el.local === 'object') {
    const inner = el.element(NS.w, 'drawing');
    if (inner !== undefined) return fromDrawing(inner, props, link, ctx);
  }
  return fromVml(el, props, link, ctx);
}
