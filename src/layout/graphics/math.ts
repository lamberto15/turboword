// OMML (Office Math) layout: m:oMathPara / m:oMath trees → a MathBox of display items placed
// relative to (left edge, baseline), y growing downward.
//
// The model is TeX-like boxes (width, ascent, descent) plus an estimated *ink* extent: font
// ascent/descent are much larger than the visible glyphs of lowercase letters, and Word places
// fraction bars, accents, radicals and scripts against the ink, not the em box. Ink is estimated
// per character from the face's x-height and cap height (no glyph outlines are read here).
//
// Without a MATH table, constants are fixed fractions of the em: math axis 0.25 em, rule
// thickness em/18, script sizes 70% and 50%. Stretchy delimiters, radicals, horizontal braces and
// most accents are drawn as paths so they fit any content and any (fallback) font; text goes
// through the layout's FontManager with 'Cambria Math' requested.
//
// Recursion is bounded: past MAX_DEPTH nested structures the remaining content is laid out flat
// (its text in one run), so arbitrarily deep documents cannot overflow the stack.

import type { ResolvedRunProps } from '../../resolve/types.js';
import type { XElement } from '../../xml/dom.js';
import { NS } from '../../xml/ns.js';
import type { LayoutEnv, MathBox } from '../engine.js';
import type { DisplayItem, Font } from '../types.js';
import { Path } from './path.js';
import { offsetItems, runProps, shapeText } from './text.js';

const MATH_FONT = 'Cambria Math';
/** Structure nesting beyond which content is laid out flat. */
const MAX_DEPTH = 48;
/** Font size factor per script level (0 = base, 1 = script, 2 = scriptscript). */
const LEVEL_SCALE = [1, 0.7, 0.5];

/** A laid-out box; `ink`/`inkD` estimate the visible extent above/below the baseline. */
interface Box {
  w: number;
  a: number;
  d: number;
  ink: number;
  inkD: number;
  items: DisplayItem[];
  /** x of the first alignment point (m:aln) in a row (equation arrays). */
  alignX?: number;
}

interface Ctx {
  env: LayoutEnv;
  /** Base run props (level-0 size). */
  props: ResolvedRunProps;
  level: number;
  /** Display style (m:oMathPara): big operators, full-size fractions. */
  display: boolean;
  depth: number;
  /** Letters upright by default (function names). */
  upright: boolean;
}

/** Props for math text: 'Cambria Math' in every slot, size and colour from `props`. */
export function mathRunProps(props: ResolvedRunProps, italic = false): ResolvedRunProps {
  return runProps(
    { font: MATH_FONT, size: props.size, bold: props.bold, italic, color: props.color },
    props,
  );
}

/** The run props layoutMath may request (regular/italic × base/script/scriptscript sizes). */
export function mathFontRequests(props: ResolvedRunProps): ResolvedRunProps[] {
  const out: ResolvedRunProps[] = [];
  for (const s of LEVEL_SCALE) {
    for (const italic of [false, true]) {
      out.push(mathRunProps({ ...props, size: props.size * s }, italic));
    }
  }
  return out;
}

/** Lays out an m:oMath / m:oMathPara (fractions, scripts, radicals, n-ary, delimiters, matrices). */
export function layoutMath(el: XElement, props: ResolvedRunProps, env: LayoutEnv): MathBox {
  const parent = el.parent;
  const display =
    el.is(NS.m, 'oMathPara') ||
    (parent !== null && parent.nodeType === 1 && parent.is(NS.m, 'oMathPara'));
  const ctx: Ctx = { env, props, level: 0, display, depth: 0, upright: false };
  let box: Box;
  if (el.is(NS.m, 'oMathPara')) box = para(el, ctx);
  else if (el.is(NS.m, 'oMath')) box = row(el, ctx);
  else box = node(el, ctx);
  return { width: box.w, ascent: box.a, descent: box.d, items: box.items };
}

// ---- helpers ----------------------------------------------------------------------------------

const localOf = (el: XElement): string => el.name.slice(el.name.indexOf(':') + 1);

function empty(): Box {
  return { w: 0, a: 0, d: 0, ink: 0, inkD: 0, items: [] };
}

/** Places `b` at (x, baseline y) inside `into`, growing its bounds. */
function put(into: Box, b: Box, x: number, y: number): void {
  offsetItems(b.items, x, y);
  for (const it of b.items) into.items.push(it);
  into.w = Math.max(into.w, x + b.w);
  into.a = Math.max(into.a, b.a - y);
  into.d = Math.max(into.d, b.d + y);
  into.ink = Math.max(into.ink, b.ink - y);
  into.inkD = Math.max(into.inkD, b.inkD + y);
}

/** Grows a box to cover the vertical band [top, bottom] (y down). */
function cover(b: Box, top: number, bottom: number): void {
  b.a = Math.max(b.a, -top);
  b.ink = Math.max(b.ink, -top);
  b.d = Math.max(b.d, bottom);
  b.inkD = Math.max(b.inkD, bottom);
}

/** OMML on/off property: present without val, or val not 0/off/false. */
function onOff(el: XElement | undefined): boolean {
  if (el === undefined) return false;
  const v = el.attrNS(NS.m, 'val') ?? el.attrNS(NS.w, 'val') ?? el.attr('val');
  return v === undefined || !/^(0|off|false)$/i.test(v);
}

function prVal(pr: XElement | undefined, name: string): string | undefined {
  const el = pr?.element(NS.m, name);
  return el?.attrNS(NS.m, 'val') ?? el?.attr('val');
}

/** A character property: default when absent, '' (no character) when given empty. */
function chrVal(pr: XElement | undefined, name: string, dflt: string): string {
  const el = pr?.element(NS.m, name);
  if (el === undefined) return dflt;
  return el.attrNS(NS.m, 'val') ?? el.attr('val') ?? dflt;
}

const arg = (el: XElement, name: string): XElement | undefined => el.element(NS.m, name);

function sizeAt(ctx: Ctx, base = ctx.props.size): number {
  return base * (LEVEL_SCALE[Math.min(ctx.level, 2)] ?? 0.5);
}

const em = (ctx: Ctx): number => sizeAt(ctx);
const axis = (ctx: Ctx): number => 0.25 * em(ctx);
const rule = (ctx: Ctx): number => em(ctx) / 18;

function hexColor(c: string): string {
  return /^[0-9A-Fa-f]{6}$/.test(c) ? c.toUpperCase() : '000000';
}

const sub = (ctx: Ctx): Ctx => ({ ...ctx, level: ctx.level + 1, display: false, upright: false });

/** x-height and cap height of the math face at the context size. */
function faceHeights(ctx: Ctx): { xh: number; cap: number } {
  const size = em(ctx);
  const font = ctx.env.fonts.fontFor(mathRunProps({ ...ctx.props, size }), 120);
  return heightsOf(font, size);
}

function heightsOf(font: Font, size: number): { xh: number; cap: number; desc: number } {
  const m = font.metrics;
  const upm = m.unitsPerEm || 1000;
  return {
    xh: (m.xHeight > 0 ? m.xHeight / upm : 0.5) * size,
    cap: (m.capHeight > 0 ? m.capHeight / upm : 0.7) * size,
    desc: Math.max(0.15, -m.descender / upm) * size * 0.85,
  };
}

/** Estimated ink extent of `text` set in `font` at `size`. */
function inkOf(text: string, font: Font, size: number): { ink: number; inkD: number } {
  const { xh, cap, desc } = heightsOf(font, size);
  let ink = 0;
  let inkD = 0;
  for (const ch of text) {
    let top = cap;
    let bot = 0;
    if (/[acemnorsuvwxzıαεικνοπστυω]/.test(ch)) top = xh;
    else if (/[gpqyγημρφχψ]/.test(ch)) {
      top = xh;
      bot = desc;
    } else if (ch === 'j' || ch === 'β' || ch === 'ζ' || ch === 'ξ') bot = desc;
    else if (/[+−\-±∓×÷·∙∘*]/.test(ch)) top = 0.55 * size;
    else if (/[=<>≤≥≠≈≡∼≃≅∝∈∉∋⊂⊃⊆⊇→←↔⇒⇐⇔]/.test(ch)) top = 0.5 * size;
    else if (/[()[\]{}|‖⟨⟩]/.test(ch)) {
      top = Math.max(cap, 0.75 * size);
      bot = desc;
    } else if (ch === ',' || ch === ';') {
      top = ch === ',' ? 0.1 * size : xh;
      bot = desc * 0.8;
    } else if (ch === '.' || ch === ' ') top = ch === '.' ? 0.1 * size : 0;
    if (top > ink) ink = top;
    if (bot > inkD) inkD = bot;
  }
  return { ink, inkD };
}

function textBox(text: string, props: ResolvedRunProps, ctx: Ctx): Box {
  const s = shapeText(text, props, ctx.env);
  const font = ctx.env.fonts.fontFor(props, text.codePointAt(0) ?? 32);
  const { ink, inkD } = inkOf(text, font, props.size);
  return { w: s.width, a: s.ascent, d: s.descent, ink, inkD, items: s.items };
}

/** Upright math text at the context size. */
function mathText(text: string, ctx: Ctx, factor = 1): Box {
  return textBox(text, mathRunProps({ ...ctx.props, size: em(ctx) * factor }), ctx);
}

/** Text of a subtree in document order (flat fallback for over-deep content). */
function flat(el: XElement, ctx: Ctx): Box {
  const text = el
    .descendants()
    .filter((d) => d.is(NS.m, 't') || d.is(NS.w, 't'))
    .map((t) => t.text)
    .join('');
  return mathText(text, ctx);
}

// ---- rows and runs ------------------------------------------------------------------------------

/** Elements that are transparent wrappers inside math (their children are laid out in place). */
function isWrapper(el: XElement): boolean {
  if (el.uri !== NS.w) return false;
  const l = localOf(el);
  return (
    l === 'ins' ||
    l === 'moveTo' ||
    l === 'smartTag' ||
    l === 'customXml' ||
    l === 'sdt' ||
    l === 'sdtContent' ||
    l === 'hyperlink'
  );
}

function isSkipped(el: XElement): boolean {
  const l = localOf(el);
  if (el.uri === NS.m) return l.endsWith('Pr') || l === 'ctrlPr';
  if (el.uri === NS.w) {
    return (
      l === 'del' ||
      l === 'moveFrom' ||
      l === 'rPr' ||
      l === 'proofErr' ||
      l.startsWith('bookmark') ||
      l.startsWith('commentRange') ||
      l.startsWith('permStart') ||
      l.startsWith('permEnd')
    );
  }
  return true;
}

/** A horizontal row of the children of a container (m:oMath, m:e, m:num, ...). */
function row(el: XElement | undefined, ctx: Ctx): Box {
  const out = empty();
  if (el === undefined) return out;
  const list: XElement[] = [];
  const stack = el.elements().reverse();
  while (stack.length > 0) {
    const k = stack.pop() as XElement;
    if (isWrapper(k)) {
      const kids = k.elements();
      for (let i = kids.length - 1; i >= 0; i--) stack.push(kids[i] as XElement);
    } else if (k.is(NS.m, 'r') || k.is(NS.w, 'r') || !isSkipped(k)) list.push(k);
  }
  let x = 0;
  for (const k of list) {
    let b: Box;
    if (k.is(NS.m, 'r') || k.is(NS.w, 'r')) {
      b = runBox(k, ctx, out.items.length === 0);
      if (out.alignX === undefined && onOff(k.element(NS.m, 'rPr')?.element(NS.m, 'aln'))) {
        out.alignX = x;
      }
    } else b = node(k, ctx);
    put(out, b, x, 0);
    x += b.w;
  }
  out.w = x;
  return out;
}

const BINARY = new Set([...'+−±∓×÷·∙∘∗⊕⊖⊗⊙∪∩∧∨*']);
const RELATION = new Set([...'=<>≤≥≠≈≡∼≃≅∝→←↔⇒⇐⇔∈∉∋⊂⊃⊆⊇≪≫≺≻∣∥:']);
const FUNCTIONS = new Set(
  (
    'sin cos tan cot sec csc arcsin arccos arctan sinh cosh tanh coth sech csch log ln lg exp ' +
    'lim liminf limsup max min sup inf det dim gcd lcm ker deg arg Pr hom mod'
  ).split(' '),
);

/** m:scr letter alphabets: [capital A, small a, exceptions]. */
const SCRIPTS: Record<string, [number, number, Record<string, number>]> = {
  'double-struck': [
    0x1d538,
    0x1d552,
    { C: 0x2102, H: 0x210d, N: 0x2115, P: 0x2119, Q: 0x211a, R: 0x211d, Z: 0x2124 },
  ],
  script: [
    0x1d49c,
    0x1d4b6,
    {
      B: 0x212c,
      E: 0x2130,
      F: 0x2131,
      H: 0x210b,
      I: 0x2110,
      L: 0x2112,
      M: 0x2133,
      R: 0x211b,
      e: 0x212f,
      g: 0x210a,
      o: 0x2134,
    },
  ],
  fraktur: [0x1d504, 0x1d51e, { C: 0x212d, H: 0x210c, I: 0x2111, R: 0x211c, Z: 0x2128 }],
  'sans-serif': [0x1d5a0, 0x1d5ba, {}],
  monospace: [0x1d670, 0x1d68a, {}],
};

function mapScript(text: string, scr: string | undefined): string {
  const alpha = scr === undefined ? undefined : SCRIPTS[scr];
  if (alpha === undefined) return text;
  let out = '';
  for (const ch of text) {
    const c = ch.charCodeAt(0);
    const exc = alpha[2][ch];
    if (exc !== undefined) out += String.fromCodePoint(exc);
    else if (c >= 65 && c <= 90) out += String.fromCodePoint(alpha[0] + c - 65);
    else if (c >= 97 && c <= 122) out += String.fromCodePoint(alpha[1] + c - 97);
    else out += ch;
  }
  return out;
}

const isLetter = (ch: string): boolean => /^[A-Za-zΑ-Ωα-ωϑϕϵ]$/.test(ch);

function wOn(el: XElement | undefined): boolean | undefined {
  if (el === undefined) return undefined;
  const v = el.attrNS(NS.w, 'val');
  return v === undefined || !/^(0|off|false|none)$/i.test(v);
}

/** An m:r (or a w:r used as normal text inside math). */
function runBox(r: XElement, ctx: Ctx, atStart: boolean): Box {
  let text = '';
  for (const t of r.elements()) if (t.is(NS.m, 't') || t.is(NS.w, 't')) text += t.text;
  const mrPr = r.element(NS.m, 'rPr');
  const wrPr = r.element(NS.w, 'rPr');
  let base = ctx.props;
  if (wrPr !== undefined) {
    base = { ...base, fonts: { ...base.fonts } };
    const sz = Number(wrPr.element(NS.w, 'sz')?.attrNS(NS.w, 'val'));
    if (Number.isFinite(sz) && sz > 0) {
      base.size = sz / 2;
      base.sizeCs = sz / 2;
    }
    const color = wrPr.element(NS.w, 'color')?.attrNS(NS.w, 'val');
    if (color !== undefined && /^[0-9A-Fa-f]{6}$/.test(color)) base.color = color;
    const b = wOn(wrPr.element(NS.w, 'b'));
    if (b !== undefined) base.bold = b;
    const i = wOn(wrPr.element(NS.w, 'i'));
    if (i !== undefined) base.italic = i;
    const font = wrPr.element(NS.w, 'rFonts')?.attrNS(NS.w, 'ascii');
    if (font !== undefined) base.fonts = { ascii: font, hAnsi: font, eastAsia: font, cs: font };
  }
  const size = sizeAt(ctx, base.size);
  const nor = r.is(NS.w, 'r') || onOff(mrPr?.element(NS.m, 'nor'));
  if (nor) {
    // Normal text: the run's own (non-math) font, upright unless w:i.
    const props = runProps({ size, italic: base.italic, bold: base.bold, color: base.color }, base);
    return textBox(text, props, ctx);
  }
  const sty = prVal(mrPr, 'sty');
  const lit = onOff(mrPr?.element(NS.m, 'lit'));
  text = mapScript(text.replace(/-/g, '−'), prVal(mrPr, 'scr'));
  const bold = sty === 'b' || sty === 'bi' ? true : base.bold;
  const upright = ctx.upright || FUNCTIONS.has(text);
  const e = size;
  const spacing = ctx.level === 0 && !lit;
  // Segments: runs of same italic flag; spaced operators on their own.
  const segs: { text: string; italic: boolean; before: number; after: number }[] = [];
  let first = atStart;
  for (const ch of text) {
    const italic =
      sty !== undefined ? sty === 'i' || sty === 'bi' : !upright && isLetter(ch) && ch.length === 1;
    let space = 0;
    if (spacing && BINARY.has(ch)) space = first ? 0 : (4 / 18) * e;
    else if (spacing && RELATION.has(ch)) space = (5 / 18) * e;
    const last = segs[segs.length - 1];
    if (space === 0 && last !== undefined && last.italic === italic && last.after === 0) {
      last.text += ch;
    } else segs.push({ text: ch, italic, before: space, after: space });
    first = false;
  }
  const out = empty();
  let x = 0;
  for (const s of segs) {
    const props = mathRunProps({ ...base, size, bold }, s.italic);
    const b = textBox(s.text, props, ctx);
    put(out, b, x + s.before, 0);
    x += s.before + b.w + s.after;
  }
  if (segs.length === 0) {
    const b = textBox('', mathRunProps({ ...base, size, bold }), ctx);
    out.a = b.a;
    out.d = b.d;
  }
  out.w = x;
  return out;
}

// ---- structures -------------------------------------------------------------------------------

function node(el: XElement, ctx: Ctx): Box {
  if (ctx.depth >= MAX_DEPTH) return flat(el, ctx);
  const c: Ctx = { ...ctx, depth: ctx.depth + 1 };
  if (el.uri !== NS.m) return row(el, c);
  const l = localOf(el);
  if (l !== 'e' && l !== 'num' && l !== 'den' && l !== 'fName' && l !== 'oMath') c.upright = false;
  switch (l) {
    case 'f':
      return frac(el, c);
    case 'sSup':
    case 'sSub':
    case 'sSubSup':
      return scripts(el, c, l);
    case 'sPre':
      return preScripts(el, c);
    case 'rad':
      return radical(el, c);
    case 'nary':
      return nary(el, c);
    case 'd':
      return delimited(el, c);
    case 'm':
      return matrix(el, c);
    case 'func':
      return func(el, c);
    case 'acc':
      return accent(el, c);
    case 'bar':
      return bar(el, c);
    case 'box':
      return row(arg(el, 'e'), c);
    case 'borderBox':
      return borderBox(el, c);
    case 'groupChr':
      return groupChr(el, c);
    case 'limLow':
    case 'limUpp':
      return limit(el, c, l === 'limLow');
    case 'eqArr':
      return eqArr(el, c);
    case 'phant':
      return phantom(el, c);
    default:
      return row(el, c);
  }
}

function frac(el: XElement, ctx: Ctx): Box {
  const type = prVal(el.element(NS.m, 'fPr'), 'type') ?? 'bar';
  const e = em(ctx);
  if (type === 'lin') {
    const out = empty();
    const num = row(arg(el, 'num'), ctx);
    const slash = mathText('/', ctx);
    const den = row(arg(el, 'den'), ctx);
    put(out, num, 0, 0);
    put(out, slash, num.w, 0);
    put(out, den, num.w + slash.w, 0);
    return out;
  }
  const inner: Ctx = ctx.display ? { ...ctx, display: false } : sub(ctx);
  const num = row(arg(el, 'num'), inner);
  const den = row(arg(el, 'den'), inner);
  const out = empty();
  if (type === 'skw') {
    const slash = mathText('/', ctx, 1.4);
    const ax = axis(ctx);
    put(out, num, 0, -ax - 0.1 * e);
    put(out, slash, num.w, ax * 0.6);
    put(out, den, num.w + slash.w, ax + 0.15 * e);
    return out;
  }
  const ax = axis(ctx);
  const t = rule(ctx);
  const gap = ctx.display ? 0.12 * e : 0.08 * e;
  const pad = 0.1 * e;
  const w = Math.max(num.w, den.w) + 2 * pad;
  const numY = -ax - t / 2 - gap - num.inkD;
  const denY = -ax + t / 2 + gap + den.ink;
  put(out, num, (w - num.w) / 2, numY);
  put(out, den, (w - den.w) / 2, denY);
  if (type !== 'noBar') {
    out.items.push({
      kind: 'rect',
      x: pad / 2,
      y: -ax - t / 2,
      width: w - pad,
      height: t,
      fill: hexColor(ctx.props.color),
    });
  }
  out.w = w;
  return out;
}

function scripts(el: XElement, ctx: Ctx, kind: string): Box {
  const base = row(arg(el, 'e'), ctx);
  const sc = sub(ctx);
  const e = em(ctx);
  const t = rule(ctx);
  const sup = kind !== 'sSub' ? row(arg(el, 'sup'), sc) : undefined;
  const sb = kind !== 'sSup' ? row(arg(el, 'sub'), sc) : undefined;
  const u = Math.max(0.42 * e, base.ink - 0.27 * e);
  let v = Math.max(0.2 * e, base.inkD + 0.035 * e);
  if (sup !== undefined && sb !== undefined) {
    const clearance = v - sb.ink - (-u + sup.inkD);
    if (clearance < 4 * t) v += 4 * t - clearance;
  }
  const out = empty();
  put(out, base, 0, 0);
  const x = base.w + 0.03 * e;
  if (sup !== undefined) put(out, sup, x, -u);
  if (sb !== undefined) put(out, sb, x, v);
  return out;
}

function preScripts(el: XElement, ctx: Ctx): Box {
  const sc = sub(ctx);
  const e = em(ctx);
  const sb = row(arg(el, 'sub'), sc);
  const sup = row(arg(el, 'sup'), sc);
  const base = row(arg(el, 'e'), ctx);
  const u = Math.max(0.42 * e, base.ink - 0.27 * e);
  const v = Math.max(0.25 * e, base.inkD + 0.035 * e);
  const sw = Math.max(sb.w, sup.w);
  const out = empty();
  put(out, sup, sw - sup.w, -u);
  put(out, sb, sw - sb.w, v);
  put(out, base, sw + 0.03 * e, 0);
  return out;
}

function stroke(p: Path, color: string, width: number): DisplayItem {
  return { kind: 'path', d: p.toString(), stroke: color, strokeWidth: width };
}

function radical(el: XElement, ctx: Ctx): Box {
  const pr = el.element(NS.m, 'radPr');
  const hide = onOff(pr?.element(NS.m, 'degHide'));
  const body = row(arg(el, 'e'), ctx);
  const degEl = arg(el, 'deg');
  const deg =
    !hide && degEl !== undefined && degEl.descendants(NS.m, 't').length > 0
      ? row(degEl, { ...ctx, level: Math.min(ctx.level + 2, 2), display: false })
      : undefined;
  const e = em(ctx);
  const t = rule(ctx);
  const gap = ctx.display ? 0.15 * e : 0.1 * e;
  const { cap } = faceHeights(ctx);
  const top = Math.max(body.ink, cap) + gap + t;
  const yT = -top + t / 2;
  const yB = Math.max(body.inkD, 0) + 0.05 * e;
  const h = yB - yT;
  const sw = Math.min(0.55 * e + 0.1 * h, e);
  const shift = deg !== undefined ? Math.max(0, deg.w - 0.5 * sw) : 0;
  const x0 = shift;
  const yHook = yB - Math.min(0.5 * h, 0.55 * e);
  const bodyX = x0 + sw + 0.08 * e;
  const end = bodyX + body.w + 0.08 * e;
  const p = new Path()
    .moveTo(x0, yHook)
    .lineTo(x0 + 0.18 * sw, yHook - 0.06 * e)
    .lineTo(x0 + 0.45 * sw, yB)
    .lineTo(x0 + sw, yT)
    .lineTo(end, yT);
  const out = empty();
  put(out, body, bodyX, 0);
  out.items.push(stroke(p, hexColor(ctx.props.color), t));
  cover(out, yT - t / 2, yB + t / 2);
  if (deg !== undefined) put(out, deg, Math.max(0, x0 + 0.5 * sw - deg.w), yB - 0.6 * h - deg.inkD);
  out.w = end;
  return out;
}

const isIntegral = (ch: string): boolean => /[∫-∳⨋-⨜]/.test(ch);

function nary(el: XElement, ctx: Ctx): Box {
  const pr = el.element(NS.m, 'naryPr');
  const chr = chrVal(pr, 'chr', '∫') || '∫';
  const integral = isIntegral(chr);
  const limLoc = prVal(pr, 'limLoc') ?? (integral ? 'subSup' : 'undOvr');
  const e = em(ctx);
  const k = ctx.display ? 1.6 : 1.2;
  const opSize = e * k;
  const op = textBox(chr, mathRunProps({ ...ctx.props, size: opSize }), ctx);
  const font = ctx.env.fonts.fontFor(mathRunProps({ ...ctx.props, size: opSize }), 0x2211);
  const { cap } = heightsOf(font, opSize);
  const inkTop = integral ? 0.85 * opSize : cap;
  const inkBot = integral ? 0.25 * opSize : 0;
  const opY = -axis(ctx) + (inkTop - inkBot) / 2;
  op.ink = inkTop;
  op.inkD = inkBot;
  const sc = sub(ctx);
  const sb = onOff(pr?.element(NS.m, 'subHide')) ? undefined : row(arg(el, 'sub'), sc);
  const sup = onOff(pr?.element(NS.m, 'supHide')) ? undefined : row(arg(el, 'sup'), sc);
  const out = empty();
  const gap = 0.1 * e;
  let colW: number;
  if (limLoc === 'undOvr') {
    colW = Math.max(op.w, sb?.w ?? 0, sup?.w ?? 0);
    put(out, op, (colW - op.w) / 2, opY);
    if (sup !== undefined) put(out, sup, (colW - sup.w) / 2, opY - inkTop - gap - sup.inkD);
    if (sb !== undefined) put(out, sb, (colW - sb.w) / 2, opY + inkBot + gap + sb.ink);
  } else {
    put(out, op, 0, opY);
    colW = op.w;
    if (sup !== undefined) {
      const x = op.w + 0.03 * e;
      put(out, sup, x, opY - inkTop + 0.6 * sup.ink);
      colW = Math.max(colW, x + sup.w);
    }
    if (sb !== undefined) {
      const x = integral ? op.w * 0.65 : op.w + 0.03 * e;
      put(out, sb, x, opY + inkBot + 0.4 * sb.ink);
      colW = Math.max(colW, x + sb.w);
    }
  }
  const body = row(arg(el, 'e'), ctx);
  put(out, body, colW + e / 6, 0);
  return out;
}

/** Characters drawn as stretchy paths. */
const PATH_DELIMS = new Set([...'()[]{}|‖⟨⟩〈〉⌊⌋⌈⌉']);

/** A delimiter stretched over [yT, yB] as a stroked path. */
function delimPath(ch: string, yT: number, yB: number, ctx: Ctx): Box {
  const e = em(ctx);
  const h = yB - yT;
  const dw = Math.min(0.3 * e + 0.04 * h, 0.7 * e);
  const m = 0.07 * e;
  const closing = ')]}⟩〉⌋⌉'.includes(ch);
  const L = m;
  const R = dw - m;
  const X = (u: number): number => (closing ? L + R - u : u);
  const mid = (L + R) / 2;
  const ym = (yT + yB) / 2;
  const p = new Path();
  switch (ch) {
    case '(':
    case ')':
      p.moveTo(X(R), yT).cubicTo(X(L), yT + h * 0.25, X(L), yB - h * 0.25, X(R), yB);
      break;
    case '[':
    case ']':
      p.moveTo(X(R), yT).lineTo(X(L), yT).lineTo(X(L), yB).lineTo(X(R), yB);
      break;
    case '{':
    case '}': {
      const c = Math.min(0.08 * h, 0.3 * e);
      p.moveTo(X(R), yT)
        .quadTo(X(mid), yT, X(mid), yT + c)
        .lineTo(X(mid), ym - c)
        .quadTo(X(mid), ym, X(L), ym)
        .quadTo(X(mid), ym, X(mid), ym + c)
        .lineTo(X(mid), yB - c)
        .quadTo(X(mid), yB, X(R), yB);
      break;
    }
    case '‖':
      p.moveTo(L + (R - L) * 0.3, yT).lineTo(L + (R - L) * 0.3, yB);
      p.moveTo(R - (R - L) * 0.3, yT).lineTo(R - (R - L) * 0.3, yB);
      break;
    case '⟨':
    case '⟩':
    case '〈':
    case '〉':
      p.moveTo(X(R), yT).lineTo(X(L), ym).lineTo(X(R), yB);
      break;
    case '⌊':
    case '⌋':
      p.moveTo(X(L), yT).lineTo(X(L), yB).lineTo(X(R), yB);
      break;
    case '⌈':
    case '⌉':
      p.moveTo(X(R), yT).lineTo(X(L), yT).lineTo(X(L), yB);
      break;
    default:
      p.moveTo(mid, yT).lineTo(mid, yB);
  }
  return {
    w: dw,
    a: -yT,
    d: yB,
    ink: -yT,
    inkD: yB,
    items: [stroke(p, hexColor(ctx.props.color), rule(ctx) * 1.3)],
  };
}

/** A delimiter character covering half-height `hh` around the axis (glyph when it fits). */
function delimBox(ch: string, hh: number, stretch: boolean, ctx: Ctx): Box {
  const e = em(ctx);
  const ax = axis(ctx);
  if (!stretch) return mathText(ch, ctx);
  if (PATH_DELIMS.has(ch)) return delimPath(ch, -ax - hh, -ax + hh, ctx);
  // Other characters: scale the glyph up, narrowing it so it does not get fat.
  const k = (2 * hh) / (0.95 * e);
  const props = mathRunProps({ ...ctx.props, size: e * k });
  props.scale = Math.round((100 * (1 + 0.25 * (k - 1))) / k);
  const b = textBox(ch, props, ctx);
  const shifted = empty();
  put(shifted, b, 0, -ax + (b.a - b.d) / 2);
  return shifted;
}

function delimited(el: XElement, ctx: Ctx): Box {
  const pr = el.element(NS.m, 'dPr');
  const beg = chrVal(pr, 'begChr', '(');
  const end = chrVal(pr, 'endChr', ')');
  const sep = chrVal(pr, 'sepChr', '|');
  const growEl = pr?.element(NS.m, 'grow');
  const grow = growEl === undefined || onOff(growEl);
  const parts = el.elements(NS.m, 'e').map((x) => row(x, ctx));
  const e = em(ctx);
  const ax = axis(ctx);
  let hh = 0;
  for (const b of parts) hh = Math.max(hh, b.ink - ax, b.inkD + ax);
  hh += 0.08 * e;
  const stretch = grow && hh > 0.55 * e;
  const out = empty();
  let x = 0;
  const add = (b: Box): void => {
    put(out, b, x, 0);
    x += b.w;
  };
  if (beg !== '') add(delimBox(beg, hh, stretch, ctx));
  parts.forEach((b, i) => {
    if (i > 0 && sep !== '') add(delimBox(sep, hh, stretch, ctx));
    add(b);
  });
  if (end !== '') add(delimBox(end, hh, stretch, ctx));
  out.w = x;
  return out;
}

function matrix(el: XElement, ctx: Ctx): Box {
  const rows = el.elements(NS.m, 'mr').map((r) => r.elements(NS.m, 'e').map((c) => row(c, ctx)));
  const jcs: string[] = [];
  for (const mc of el.element(NS.m, 'mPr')?.element(NS.m, 'mcs')?.elements(NS.m, 'mc') ?? []) {
    const mcPr = mc.element(NS.m, 'mcPr');
    const count = Math.max(1, Math.min(64, Number(prVal(mcPr, 'count') ?? 1) || 1));
    for (let i = 0; i < count; i++) jcs.push(prVal(mcPr, 'mcJc') ?? 'center');
  }
  const e = em(ctx);
  const ncol = rows.reduce((n, r) => Math.max(n, r.length), 0);
  const colW = new Array<number>(ncol).fill(0);
  const rowA = rows.map((r) => r.reduce((n, c) => Math.max(n, c.a), 0));
  const rowD = rows.map((r) => r.reduce((n, c) => Math.max(n, c.d), 0));
  for (const r of rows) {
    r.forEach((c, j) => {
      colW[j] = Math.max(colW[j] ?? 0, c.w);
    });
  }
  const cGap = e;
  const rGap = 0.2 * e;
  let total = rGap * Math.max(0, rows.length - 1);
  for (let i = 0; i < rows.length; i++) total += (rowA[i] ?? 0) + (rowD[i] ?? 0);
  const out = empty();
  let y = -axis(ctx) - total / 2;
  rows.forEach((r, i) => {
    const baseline = y + (rowA[i] ?? 0);
    let x = 0;
    for (let j = 0; j < ncol; j++) {
      const c = r[j];
      const cw = colW[j] ?? 0;
      if (c !== undefined) {
        const jc = jcs[j] ?? 'center';
        const dx = jc === 'left' ? 0 : jc === 'right' ? cw - c.w : (cw - c.w) / 2;
        put(out, c, x + dx, baseline);
      }
      x += cw + (j < ncol - 1 ? cGap : 0);
    }
    out.w = Math.max(out.w, x);
    y = baseline + (rowD[i] ?? 0) + rGap;
  });
  return out;
}

function eqArr(el: XElement, ctx: Ctx): Box {
  const rows = el.elements(NS.m, 'e').map((r) => row(r, ctx));
  const e = em(ctx);
  const gap = 0.2 * e;
  const maxAlign = rows.reduce((n, r) => Math.max(n, r.alignX ?? 0), 0);
  const offs = rows.map((r) => (r.alignX !== undefined ? maxAlign - r.alignX : 0));
  const width = rows.reduce((n, r, i) => Math.max(n, (offs[i] ?? 0) + r.w), 0);
  let total = gap * Math.max(0, rows.length - 1);
  for (const r of rows) total += r.a + r.d;
  const out = empty();
  let y = -axis(ctx) - total / 2;
  rows.forEach((r, i) => {
    const baseline = y + r.a;
    const x = r.alignX !== undefined ? (offs[i] ?? 0) : (width - r.w) / 2;
    put(out, r, x, baseline);
    y = baseline + r.d + gap;
  });
  out.w = width;
  return out;
}

function func(el: XElement, ctx: Ctx): Box {
  const name = row(arg(el, 'fName'), { ...ctx, upright: true });
  const body = row(arg(el, 'e'), { ...ctx, upright: false });
  const out = empty();
  put(out, name, 0, 0);
  put(out, body, name.w + em(ctx) / 6, 0);
  return out;
}

function accent(el: XElement, ctx: Ctx): Box {
  const chr = chrVal(el.element(NS.m, 'accPr'), 'chr', '̂') || '̂';
  const base = row(arg(el, 'e'), ctx);
  const e = em(ctx);
  const t = rule(ctx) * 1.1;
  const color = hexColor(ctx.props.color);
  const yb = -(base.ink + 0.08 * e);
  const w = base.w;
  const cx = w / 2;
  const aw = Math.min(Math.max(w * 0.8, 0.3 * e), 0.6 * e);
  const p = new Path();
  let ah = 0.15 * e;
  let filled = false;
  const arrow = (x: number, y: number, dir: number): void => {
    p.moveTo(x - dir * 0.12 * e, y - 0.07 * e)
      .lineTo(x, y)
      .lineTo(x - dir * 0.12 * e, y + 0.07 * e);
  };
  switch (chr) {
    case '̂':
    case '^':
    case 'ˆ':
      p.moveTo(cx - aw / 2, yb)
        .lineTo(cx, yb - ah)
        .lineTo(cx + aw / 2, yb);
      break;
    case '̌':
    case 'ˇ':
      p.moveTo(cx - aw / 2, yb - ah)
        .lineTo(cx, yb)
        .lineTo(cx + aw / 2, yb - ah);
      break;
    case '̃':
    case '~':
    case '˜':
      ah = 0.1 * e;
      p.moveTo(cx - aw / 2, yb - ah * 0.2).cubicTo(
        cx - aw * 0.2,
        yb - ah * 1.4,
        cx + aw * 0.2,
        yb + ah * 0.4,
        cx + aw / 2,
        yb - ah * 0.8,
      );
      break;
    case '̄':
    case '̅':
    case '¯':
    case '‾':
      ah = t;
      p.moveTo(0, yb - t / 2).lineTo(w, yb - t / 2);
      break;
    case '̇':
    case '˙':
    case '̈':
    case '¨':
    case '⃛': {
      const r = 0.055 * e;
      const n = chr === '⃛' ? 3 : chr === '̈' || chr === '¨' ? 2 : 1;
      for (let i = 0; i < n; i++) {
        const dx = (i - (n - 1) / 2) * 0.2 * e;
        p.ellipse(cx + dx - r, yb - 2 * r, 2 * r, 2 * r);
      }
      ah = 2 * r;
      filled = true;
      break;
    }
    case '́':
    case '´':
      p.moveTo(cx - aw * 0.15, yb).lineTo(cx + aw * 0.2, yb - ah);
      break;
    case '̀':
    case '`':
      p.moveTo(cx + aw * 0.15, yb).lineTo(cx - aw * 0.2, yb - ah);
      break;
    case '̆':
    case '˘':
      ah = 0.12 * e;
      p.moveTo(cx - aw / 2, yb - ah).cubicTo(
        cx - aw / 2,
        yb + ah * 0.33,
        cx + aw / 2,
        yb + ah * 0.33,
        cx + aw / 2,
        yb - ah,
      );
      break;
    case '⃗':
    case '→':
    case '⃖':
    case '←':
    case '⃡':
    case '↔': {
      ah = 0.16 * e;
      const y = yb - ah / 2;
      p.moveTo(0, y).lineTo(w, y);
      if ('⃗→⃡↔'.includes(chr)) arrow(w, y, 1);
      if ('⃖←⃡↔'.includes(chr)) arrow(0, y, -1);
      break;
    }
    default: {
      const g = mathText(chr, ctx);
      const { xh } = faceHeights(ctx);
      const out = empty();
      put(out, base, 0, 0);
      put(out, g, (w - g.w) / 2, yb + xh);
      out.w = Math.max(w, g.w);
      return out;
    }
  }
  const out = empty();
  put(out, base, 0, 0);
  out.items.push(filled ? { kind: 'path', d: p.toString(), fill: color } : stroke(p, color, t));
  cover(out, yb - ah - t, 0);
  out.w = Math.max(w, cx + aw / 2);
  return out;
}

function bar(el: XElement, ctx: Ctx): Box {
  const pos = prVal(el.element(NS.m, 'barPr'), 'pos') ?? 'bot';
  const base = row(arg(el, 'e'), ctx);
  const e = em(ctx);
  const t = rule(ctx);
  const gap = 0.1 * e;
  const y = pos === 'top' ? -(base.ink + gap + t) : base.inkD + gap;
  const out = empty();
  put(out, base, 0, 0);
  out.items.push({
    kind: 'rect',
    x: 0,
    y,
    width: base.w,
    height: t,
    fill: hexColor(ctx.props.color),
  });
  cover(out, y, y + t);
  return out;
}

function borderBox(el: XElement, ctx: Ctx): Box {
  const pr = el.element(NS.m, 'borderBoxPr');
  const on = (n: string): boolean => onOff(pr?.element(NS.m, n));
  const inner = row(arg(el, 'e'), ctx);
  const e = em(ctx);
  const t = rule(ctx);
  const p = 0.12 * e;
  const { cap } = faceHeights(ctx);
  const x1 = inner.w + 2 * p;
  const yT = -(Math.max(inner.ink, cap) + p);
  const yB = inner.inkD + p;
  const color = hexColor(ctx.props.color);
  const out = empty();
  put(out, inner, p, 0);
  const line = (xa: number, ya: number, xb: number, yb: number): void => {
    out.items.push({ kind: 'line', x1: xa, y1: ya, x2: xb, y2: yb, color, width: t });
  };
  if (!on('hideTop')) line(0, yT, x1, yT);
  if (!on('hideBot')) line(0, yB, x1, yB);
  if (!on('hideLeft')) line(0, yT, 0, yB);
  if (!on('hideRight')) line(x1, yT, x1, yB);
  if (on('strikeH')) line(0, -axis(ctx), x1, -axis(ctx));
  if (on('strikeV')) line(x1 / 2, yT, x1 / 2, yB);
  if (on('strikeBLTR')) line(0, yB, x1, yT);
  if (on('strikeTLBR')) line(0, yT, x1, yB);
  cover(out, yT - t / 2, yB + t / 2);
  out.w = x1;
  return out;
}

function groupChr(el: XElement, ctx: Ctx): Box {
  const pr = el.element(NS.m, 'groupChrPr');
  const chr = chrVal(pr, 'chr', '⏟') || '⏟';
  const pos = prVal(pr, 'pos') ?? 'bot';
  const base = row(arg(el, 'e'), ctx);
  const e = em(ctx);
  const gap = 0.08 * e;
  const w = Math.max(base.w, 0.5 * e);
  const arrows = '←→↔⇐⇒⇔';
  const hb = arrows.includes(chr) ? 0.18 * e : 0.22 * e;
  const bandTop = pos === 'top' ? -(base.ink + gap + hb) : base.inkD + gap;
  const under = '⏟⏝⎵'.includes(chr);
  const yEdge = under ? bandTop : bandTop + hb;
  const yTip = under ? bandTop + hb : bandTop;
  const out = empty();
  put(out, base, (w - base.w) / 2, 0);
  const p = new Path();
  if (chr === '⏟' || chr === '⏞') {
    const r = Math.min(0.15 * e, w / 4);
    const yM = (yEdge + yTip) / 2;
    const mid = w / 2;
    p.moveTo(0, yEdge)
      .quadTo(0, yM, r, yM)
      .lineTo(mid - r, yM)
      .quadTo(mid, yM, mid, yTip)
      .quadTo(mid, yM, mid + r, yM)
      .lineTo(w - r, yM)
      .quadTo(w, yM, w, yEdge);
  } else if (chr === '⏝' || chr === '⏜') {
    p.moveTo(0, yEdge).cubicTo(0, yTip, w, yTip, w, yEdge);
  } else if (chr === '⎵' || chr === '⎴') {
    p.moveTo(0, yEdge).lineTo(0, yTip).lineTo(w, yTip).lineTo(w, yEdge);
  } else if (arrows.includes(chr)) {
    const y = bandTop + hb / 2;
    p.moveTo(0, y).lineTo(w, y);
    const head = (x: number, dir: number): void => {
      p.moveTo(x - dir * 0.12 * e, y - 0.07 * e)
        .lineTo(x, y)
        .lineTo(x - dir * 0.12 * e, y + 0.07 * e);
    };
    if ('→↔⇒⇔'.includes(chr)) head(w, 1);
    if ('←↔⇐⇔'.includes(chr)) head(0, -1);
  } else {
    const g = mathText(chr, ctx);
    put(out, g, (w - g.w) / 2, bandTop + hb);
    out.w = w;
    return out;
  }
  out.items.push(stroke(p, hexColor(ctx.props.color), rule(ctx) * 1.1));
  cover(out, bandTop, bandTop + hb);
  out.w = w;
  return out;
}

function limit(el: XElement, ctx: Ctx, low: boolean): Box {
  const base = row(arg(el, 'e'), ctx);
  const lim = row(arg(el, 'lim'), sub(ctx));
  const gap = 0.08 * em(ctx);
  const w = Math.max(base.w, lim.w);
  const out = empty();
  put(out, base, (w - base.w) / 2, 0);
  const y = low ? base.inkD + gap + lim.ink : -(base.ink + gap + lim.inkD);
  put(out, lim, (w - lim.w) / 2, y);
  out.w = w;
  return out;
}

function phantom(el: XElement, ctx: Ctx): Box {
  const pr = el.element(NS.m, 'phantPr');
  const showEl = pr?.element(NS.m, 'show');
  const b = row(arg(el, 'e'), ctx);
  if (showEl !== undefined && !onOff(showEl)) b.items = [];
  if (onOff(pr?.element(NS.m, 'zeroWid'))) b.w = 0;
  if (onOff(pr?.element(NS.m, 'zeroAsc'))) {
    b.a = 0;
    b.ink = 0;
  }
  if (onOff(pr?.element(NS.m, 'zeroDesc'))) {
    b.d = 0;
    b.inkD = 0;
  }
  return b;
}

/** m:oMathPara: each m:oMath on its own line, justified per m:oMathParaPr/m:jc. */
function para(el: XElement, ctx: Ctx): Box {
  const jc = prVal(el.element(NS.m, 'oMathParaPr'), 'jc') ?? 'centerGroup';
  const lines = el.elements(NS.m, 'oMath').map((m) => row(m, { ...ctx, display: true }));
  const width = lines.reduce((n, l) => Math.max(n, l.w), 0);
  const gap = 0.3 * em(ctx);
  const out = empty();
  let y = 0;
  lines.forEach((l, i) => {
    const prev = lines[i - 1];
    if (prev !== undefined) y += prev.d + gap + l.a;
    const x = jc === 'left' ? 0 : jc === 'right' ? width - l.w : (width - l.w) / 2;
    put(out, l, x, y);
  });
  out.w = width;
  return out;
}
