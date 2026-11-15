// SVG renderer for the PageModel: one standalone SVG document per page, in points (1 user unit =
// 1 pt). A debugging aid and a quick way to look at a layout in a browser, not a print format:
// text is emitted as <text> with one x position per character (from the glyph advances), naming
// the font family instead of embedding glyph outlines, so what you see depends on the fonts the
// viewer has. Images are inlined as data URIs (EMF/WMF, which browsers cannot show, become
// outlined placeholders). Output is deterministic: no ids from counters shared across calls,
// numbers rounded to 1/1000 pt, and attributes in a fixed order.

import type {
  DisplayItem,
  DocumentLayout,
  ImageResource,
  PageLayout,
  TextItem,
} from '../../layout/types.js';
import { escapeAttr, escapeText } from '../../xml/escape.js';

const n = (v: number): string => {
  const r = Math.round((Number.isFinite(v) ? v : 0) * 1000) / 1000;
  return r === 0 ? '0' : String(r);
};

const color = (c: string | undefined): string =>
  c !== undefined && /^[0-9A-Fa-f]{6}$/.test(c) ? `#${c.toUpperCase()}` : '#000000';

function base64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(s);
}

/** Per UTF-16 unit x positions of a text item (cluster starts from the advances, the rest interpolated). */
function charX(t: TextItem): number[] {
  const len = t.text.length;
  const xs: (number | undefined)[] = new Array(len).fill(undefined);
  let pos = t.x;
  t.glyphs.forEach((_, i) => {
    const c = t.clusters[i] ?? 0;
    if (c < len && xs[c] === undefined) xs[c] = pos;
    pos += t.advances[i] ?? 0;
  });
  const out: number[] = [];
  for (let k = 0; k < len; k++) {
    const v = xs[k];
    if (v !== undefined) {
      out.push(v);
      continue;
    }
    let p = k - 1;
    while (p >= 0 && xs[p] === undefined) p--;
    let q = k + 1;
    while (q < len && xs[q] === undefined) q++;
    const xp = p >= 0 ? (xs[p] as number) : t.x;
    const xq = q < len ? (xs[q] as number) : pos;
    out.push(xp + ((xq - xp) * (k - Math.max(p, 0))) / Math.max(1, q - Math.max(p, 0)));
  }
  return out;
}

function text(t: TextItem, layout: DocumentLayout): string {
  const font = layout.fonts[t.font];
  const xs = charX(t);
  let attrs = `x="${xs.map(n).join(' ')}" y="`;
  if (t.dy !== undefined && t.dy.length > 0) {
    // Absolute y per character: the dy of the glyph whose cluster covers it.
    const ys: number[] = new Array(t.text.length).fill(t.y);
    t.clusters.forEach((c, i) => {
      if (c < ys.length) ys[c] = t.y + (t.dy?.[i] ?? 0);
    });
    attrs += `${ys.map(n).join(' ')}"`;
  } else attrs += `${n(t.y)}"`;
  attrs += ` font-family="${escapeAttr(font?.family ?? 'sans-serif')}" font-size="${n(t.size)}"`;
  if (font?.bold || t.fakeBold) attrs += ' font-weight="bold"';
  if (font?.italic || t.fakeItalic) attrs += ' font-style="italic"';
  attrs += ` fill="${color(t.color)}"`;
  const sx = (t.scale ?? 100) / 100;
  if (sx !== 1) attrs += ` transform="matrix(${n(sx)} 0 0 1 ${n(t.x * (1 - sx))} 0)"`;
  return `<text ${attrs} xml:space="preserve">${escapeText(t.text)}</text>`;
}

function image(
  it: Extract<DisplayItem, { kind: 'image' }>,
  res: ImageResource | undefined,
): string {
  const rot =
    it.rotation === undefined || it.rotation === 0
      ? ''
      : ` transform="rotate(${n(it.rotation)} ${n(it.x + it.width / 2)} ${n(it.y + it.height / 2)})"`;
  const box = `x="${n(it.x)}" y="${n(it.y)}" width="${n(it.width)}" height="${n(it.height)}"`;
  if (res === undefined || /x-emf|x-wmf/.test(res.contentType)) {
    return `<rect ${box} fill="none" stroke="#999999" stroke-width="0.5"${rot}/>`;
  }
  const href = `data:${escapeAttr(res.contentType)};base64,${base64(res.bytes)}`;
  const c = it.crop;
  if (c === undefined || (c.left === 0 && c.top === 0 && c.right === 0 && c.bottom === 0)) {
    return `<image ${box} preserveAspectRatio="none"${rot} xlink:href="${href}"/>`;
  }
  // Crop: the full image is drawn larger, shifted, and clipped by a nested viewport.
  const w = it.width / Math.max(1e-6, 1 - c.left - c.right);
  const h = it.height / Math.max(1e-6, 1 - c.top - c.bottom);
  return `<g${rot}><svg ${box} overflow="hidden"><image x="${n(-c.left * w)}" y="${n(-c.top * h)}" width="${n(w)}" height="${n(h)}" preserveAspectRatio="none" xlink:href="${href}"/></svg></g>`;
}

function item(it: DisplayItem, layout: DocumentLayout): string {
  switch (it.kind) {
    case 'text':
      return text(it, layout);
    case 'image':
      return image(it, layout.images.get(it.image));
    case 'rect': {
      const fill = it.fill === undefined ? 'none' : color(it.fill);
      const stroke =
        it.stroke === undefined
          ? ''
          : ` stroke="${color(it.stroke)}" stroke-width="${n(it.strokeWidth ?? 1)}"`;
      return `<rect x="${n(it.x)}" y="${n(it.y)}" width="${n(it.width)}" height="${n(it.height)}" fill="${fill}"${stroke}/>`;
    }
    case 'line': {
      const dash =
        it.dash === undefined || it.dash.length === 0
          ? ''
          : ` stroke-dasharray="${it.dash.map(n).join(' ')}"`;
      return `<line x1="${n(it.x1)}" y1="${n(it.y1)}" x2="${n(it.x2)}" y2="${n(it.y2)}" stroke="${color(it.color)}" stroke-width="${n(it.width)}"${dash}/>`;
    }
    case 'path': {
      const fill = it.fill === undefined ? 'none' : color(it.fill);
      const rule = it.evenOdd ? ' fill-rule="evenodd"' : '';
      const stroke =
        it.stroke === undefined
          ? ''
          : ` stroke="${color(it.stroke)}" stroke-width="${n(it.strokeWidth ?? 1)}"`;
      return `<path d="${escapeAttr(it.d)}" fill="${fill}"${rule}${stroke}/>`;
    }
    case 'link': {
      const href = it.url ?? (it.dest !== undefined ? `#${it.dest}` : '');
      return `<a xlink:href="${escapeAttr(href)}"><rect x="${n(it.x)}" y="${n(it.y)}" width="${n(it.width)}" height="${n(it.height)}" fill="#000000" fill-opacity="0"/></a>`;
    }
  }
}

/** One page as a standalone SVG document (debugging aid; text as glyph runs with font-family). */
export function renderSvg(page: PageLayout, layout: DocumentLayout): string {
  const w = n(page.width);
  const h = n(page.height);
  const parts = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" version="1.1" width="${w}pt" height="${h}pt" viewBox="0 0 ${w} ${h}">`,
    `<rect width="${w}" height="${h}" fill="#FFFFFF"/>`,
  ];
  for (const it of page.items) parts.push(item(it, layout));
  parts.push('</svg>');
  return `${parts.join('\n')}\n`;
}

/** Every page as SVG strings. */
export function renderSvgPages(layout: DocumentLayout): string[] {
  return layout.pages.map((p) => renderSvg(p, layout));
}
