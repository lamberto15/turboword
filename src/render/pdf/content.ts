// One page's content stream from its display items. The PageModel is in points with a top-left
// origin and y growing downward; PDF user space has a bottom-left origin, so every y is flipped
// against the page height here (and in path.ts).
//
// Text: one BT/ET per TextItem. Glyphs are shown with a TJ array whose numbers make the PDF's
// advance (the font's /W width) match the layout's advance for each glyph; the correction is
// computed against the running pen position so rounding never accumulates. Text state that
// persists across text objects (font, colour, Tz, Tr, Ts, line width) is tracked to avoid
// repeating it. Vectors and images are drawn inside q/Q, so they never disturb that state.

import type {
  ImageItem,
  LineItem,
  LinkItem,
  PageLayout,
  PathItem,
  RectItem,
  TextItem,
} from '../../layout/types.js';
import type { FontRegistry } from './fonts.js';
import type { ImageRegistry } from './images/index.js';
import { svgPathToPdf } from './path.js';
import { num } from './writer.js';

/** What a page needs while being rendered. */
export interface PageContext {
  fonts: FontRegistry;
  images: ImageRegistry;
  /** Wrap text in marked content with MCIDs (tagged PDF). */
  tagged: boolean;
}

/** A rendered page: its content operators, links and how many MCIDs it used. */
export interface PageContent {
  ops: string;
  links: LinkItem[];
  mcids: number;
}

interface TextState {
  font: string;
  size: number;
  fill: string;
  stroke: string;
  lw: number;
  tz: number;
  tr: number;
  rise: number;
}

/** tan(12°): the shear of synthesized italics. */
const ITALIC_SKEW = 0.2126;

/** Content operators for one page. */
export function pageContent(page: PageLayout, ctx: PageContext): PageContent {
  const H = page.height;
  const out: string[] = [];
  const links: LinkItem[] = [];
  const st: TextState = {
    font: '',
    size: -1,
    fill: '0 0 0',
    stroke: '0 0 0',
    lw: 1,
    tz: 100,
    tr: 0,
    rise: 0,
  };
  let mcid = 0;
  for (const it of page.items) {
    if (it.kind === 'link') {
      links.push(it);
      continue;
    }
    const start = out.length;
    if (it.kind === 'text') text(it, H, ctx, st, out);
    else if (it.kind === 'image') image(it, H, ctx, out);
    else if (it.kind === 'rect') rect(it, H, out);
    else if (it.kind === 'line') line(it, H, out);
    else if (it.kind === 'path') path(it, H, out);
    if (ctx.tagged && out.length > start) {
      out.splice(start, 0, it.kind === 'text' ? `/P << /MCID ${mcid++} >> BDC` : '/Artifact BMC');
      out.push('EMC');
    }
  }
  return { ops: out.join('\n'), links, mcids: mcid };
}

/** `r g b` (0-1) for an RRGGBB colour; undefined for none/transparent/empty. */
export function rgb(c: string | undefined): string | undefined {
  if (!c || c === 'none' || c === 'transparent') return undefined;
  const h = c.replace(/^#/, '');
  if (!/^[0-9a-fA-F]{6}/.test(h)) return '0 0 0';
  return [0, 2, 4].map((i) => num(Number.parseInt(h.slice(i, i + 2), 16) / 255)).join(' ');
}

function text(it: TextItem, H: number, ctx: PageContext, st: TextState, out: string[]): void {
  const n = it.glyphs.length;
  const th = (it.scale ?? 100) / 100;
  const k = (it.size * th) / 1000; // points per 1/1000 em
  if (!n || !(k > 0)) return;
  const f = ctx.fonts.use(it);
  if (!f) return;
  out.push('BT');
  const set = <K extends keyof TextState>(key: K, v: TextState[K], op: string): void => {
    if (st[key] === v) return;
    st[key] = v;
    out.push(op);
  };
  if (st.font !== f.name || st.size !== it.size) {
    out.push(`/${f.name} ${num(it.size)} Tf`);
    st.font = f.name;
    st.size = it.size;
  }
  const fill = rgb(it.color) ?? '0 0 0';
  set('fill', fill, `${fill} rg`);
  set('tz', th * 100, `${num(th * 100)} Tz`);
  if (it.fakeBold) {
    set('stroke', fill, `${fill} RG`);
    set('lw', it.size / 30, `${num(it.size / 30)} w`);
  }
  const tr = it.fakeBold ? 2 : 0;
  set('tr', tr, `${tr} Tr`);
  out.push(`1 0 ${it.fakeItalic ? ITALIC_SKEW : 0} 1 ${num(it.x)} ${num(H - it.y)} Tm`);

  let tj = '';
  let run = '';
  const flushRun = (): void => {
    if (run) tj += `<${run}>`;
    run = '';
  };
  const flushTj = (): void => {
    flushRun();
    if (tj) out.push(`[${tj}] TJ`);
    tj = '';
  };
  let pen = 0; // where the PDF has placed the pen, points from the item origin
  let target = 0; // where the layout wants the next glyph
  for (let i = 0; i < n; i++) {
    const gid = it.glyphs[i]!;
    const rise = -(it.dy?.[i] ?? 0);
    if (rise !== st.rise) {
      flushTj();
      set('rise', rise, `${num(rise)} Ts`);
    }
    const adj = Math.round((pen - target) / k);
    if (adj !== 0) {
      flushRun();
      tj += String(adj);
      pen -= adj * k;
    }
    run += f.twoByte
      ? gid.toString(16).padStart(4, '0')
      : (gid & 0xff).toString(16).padStart(2, '0');
    const w = f.width(gid) * k;
    const a = it.advances[i];
    pen += w;
    target += a !== undefined && Number.isFinite(a) ? a : w;
  }
  flushTj();
  out.push('ET');
}

function image(it: ImageItem, H: number, ctx: PageContext, out: string[]): void {
  const { width: w, height: h } = it;
  if (!(w > 0 && h > 0)) return;
  const name = ctx.images.use(it.image);
  const bx = it.x;
  const by = H - it.y - h;
  out.push('q');
  if (it.rotation) {
    const t = (-it.rotation * Math.PI) / 180;
    const c = Math.cos(t);
    const s = Math.sin(t);
    const cx = bx + w / 2;
    const cy = by + h / 2;
    out.push(
      `${num(c, 6)} ${num(s, 6)} ${num(-s, 6)} ${num(c, 6)} ${num(cx - c * cx + s * cy)} ${num(cy - s * cx - c * cy)} cm`,
    );
  }
  const box = `${num(bx)} ${num(by)} ${num(w)} ${num(h)} re`;
  if (!name) {
    out.push(`0.85 g ${box} f 0.6 G 0.5 w ${box} S Q`);
    return;
  }
  const cr = it.crop;
  const sx = cr ? 1 - cr.left - cr.right : 1;
  const sy = cr ? 1 - cr.top - cr.bottom : 1;
  if (cr && sx > 0 && sy > 0 && (cr.left || cr.right || cr.top || cr.bottom)) {
    const fw = w / sx;
    const fh = h / sy;
    out.push(
      `${box} W n`,
      `${num(fw)} 0 0 ${num(fh)} ${num(bx - cr.left * fw)} ${num(by - cr.bottom * fh)} cm`,
    );
  } else out.push(`${num(w)} 0 0 ${num(h)} ${num(bx)} ${num(by)} cm`);
  out.push(`/${name} Do`, 'Q');
}

/** fill/stroke operators: `f`, `S`, `B` (or their even-odd forms), after setting colours. */
function paint(
  fill: string | undefined,
  stroke: string | undefined,
  width: number | undefined,
  evenOdd: boolean,
  out: string[],
): string | undefined {
  const f = rgb(fill);
  const s = rgb(stroke);
  if (!f && !s) return undefined;
  if (f) out.push(`${f} rg`);
  if (s) out.push(`${s} RG ${num(width ?? 1)} w`);
  const star = evenOdd ? '*' : '';
  return f && s ? `B${star}` : f ? `f${star}` : 'S';
}

function rect(it: RectItem, H: number, out: string[]): void {
  const ops: string[] = ['q'];
  const op = paint(it.fill, it.stroke, it.strokeWidth, false, ops);
  if (!op) return;
  ops.push(
    `${num(it.x)} ${num(H - it.y - it.height)} ${num(it.width)} ${num(it.height)} re ${op}`,
    'Q',
  );
  out.push(...ops);
}

function line(it: LineItem, H: number, out: string[]): void {
  const c = rgb(it.color);
  if (!c || !(it.width > 0)) return;
  const dash = it.dash?.length ? `[${it.dash.map((v) => num(v)).join(' ')}] 0 d ` : '';
  out.push(
    `q ${c} RG ${num(it.width)} w ${dash}${num(it.x1)} ${num(H - it.y1)} m ${num(it.x2)} ${num(H - it.y2)} l S Q`,
  );
}

function path(it: PathItem, H: number, out: string[]): void {
  const ops: string[] = ['q'];
  const op = paint(it.fill, it.stroke, it.strokeWidth, it.evenOdd === true, ops);
  if (!op) return;
  const d = svgPathToPdf(it.d, H);
  if (!d) return;
  ops.push(d, op, 'Q');
  out.push(...ops);
}
