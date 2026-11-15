// renderPdf: a PageModel (DocumentLayout) → PDF 1.7 bytes. Depends on the layout *types* only, so
// it can be used (and tested) without the layout engine.
//
// Object order is fixed (catalog, page tree, info, shared resources, then per page: page,
// content, annotations; then fonts, images, destinations, outline, structure), output carries no
// clock unless `date` is given, and the trailer /ID hashes the body, so the same layout always
// produces the same bytes.

import { type WarningHandler, warning } from '../../core/warnings.js';
import type { DocumentLayout, OutlineEntry } from '../../layout/types.js';
import { pageContent } from './content.js';
import { FontRegistry, type Warn } from './fonts.js';
import { ImageRegistry } from './images/index.js';
import { latin1, num, PdfWriter, ref, str } from './writer.js';

export interface PdfOptions {
  /** Deflate content streams (default true). */
  compress?: boolean;
  /** Override metadata from the layout. */
  title?: string;
  author?: string;
  /** Creation date written to /Info (default: none, for deterministic output). */
  date?: Date;
  /** Tagged PDF structure (basic, default false). */
  tagged?: boolean;
  /** Non-fatal problems (missing fonts, unsupported image formats drawn as placeholders). */
  onWarning?: WarningHandler;
}

/** Renders a layout to PDF bytes (deterministic). */
export function renderPdf(layout: DocumentLayout, o: PdfOptions = {}): Uint8Array {
  const compress = o.compress ?? true;
  const warn: Warn = (code, message) => o.onWarning?.(warning(code, message));
  const w = new PdfWriter();
  const catalog = w.alloc();
  const pagesId = w.alloc();
  const info = w.alloc();
  const resources = w.alloc();
  const fonts = new FontRegistry(layout.fonts, w, compress, warn);
  const images = new ImageRegistry(layout.images, w, warn);
  const pages = layout.pages.length
    ? layout.pages
    : [{ index: 0, number: '1', section: 0, width: 612, height: 792, items: [] }];
  const pageIds = pages.map(() => w.alloc());
  const tagged = o.tagged === true;
  const docElem = tagged ? w.alloc() : 0;
  const structElems: number[][] = [];

  // Destinations: [page /XYZ left top zoom], null keeps the viewer's current value.
  const dest = (page: number, y: number): string => {
    const i = Math.max(0, Math.min(pages.length - 1, Math.trunc(page) || 0));
    return `[${ref(pageIds[i]!)} /XYZ null ${num(pages[i]!.height - y)} null]`;
  };

  pages.forEach((page, i) => {
    const c = pageContent(page, { fonts, images, tagged });
    const content = w.addStream('', latin1(c.ops), compress);
    const H = page.height;
    const annots: number[] = [];
    for (const l of c.links) {
      let action = '';
      if (l.url && /^\s*(javascript|vbscript|data):/i.test(l.url)) {
        warn('PDF_LINK_UNSAFE', `link to ${l.url.slice(0, 40)} dropped (script URI)`);
      } else if (l.url) action = `/A << /S /URI /URI ${str(asciiUri(l.url))} >>`;
      else if (l.dest !== undefined && layout.destinations.has(l.dest))
        action = `/Dest ${str(l.dest)}`;
      if (!action) continue;
      const rect = [l.x, H - l.y - l.height, l.x + l.width, H - l.y].map((v) => num(v)).join(' ');
      annots.push(
        w.add(`<< /Type /Annot /Subtype /Link /Rect [${rect}] /Border [0 0 0] ${action} >>`),
      );
    }
    if (tagged) {
      const elems: number[] = [];
      for (let m = 0; m < c.mcids; m++) {
        elems.push(
          w.add(`<< /Type /StructElem /S /P /P ${ref(docElem)} /Pg ${ref(pageIds[i]!)} /K ${m} >>`),
        );
      }
      structElems.push(elems);
    }
    w.set(
      pageIds[i]!,
      `<< /Type /Page /Parent ${ref(pagesId)} /MediaBox [0 0 ${num(page.width)} ${num(H)}] ` +
        `/Resources ${ref(resources)} /Contents ${ref(content)}` +
        `${annots.length ? ` /Annots [${annots.map(ref).join(' ')}]` : ''}` +
        `${tagged ? ` /StructParents ${i} /Tabs /S` : ''} >>`,
    );
  });

  fonts.write();
  const fontRes = fonts.resources();
  const imageRes = images.resources();
  w.set(
    resources,
    `<< /ProcSet [/PDF /Text /ImageB /ImageC /ImageI]${fontRes ? ` /Font << ${fontRes} >>` : ''}` +
      `${imageRes ? ` /XObject << ${imageRes} >>` : ''} >>`,
  );
  w.set(
    pagesId,
    `<< /Type /Pages /Kids [${pageIds.map(ref).join(' ')}] /Count ${pageIds.length} >>`,
  );

  // Catalog extras.
  const md = layout.metadata;
  let extra = '';
  if (md.lang) extra += ` /Lang ${str(md.lang)}`;
  if (layout.destinations.size) {
    const names = [...layout.destinations.keys()].sort(compareKeys);
    const body = names
      .map((n) => {
        const d = layout.destinations.get(n)!;
        return `${str(n)} ${dest(d.page, d.y)}`;
      })
      .join(' ');
    extra += ` /Names << /Dests ${ref(w.add(`<< /Names [${body}] >>`))} >>`;
  }
  const outline = writeOutline(w, layout.outline, dest);
  if (outline) extra += ` /Outlines ${ref(outline)} /PageMode /UseOutlines`;
  if (tagged) {
    const root = w.alloc();
    const all = structElems.flat();
    w.set(
      docElem,
      `<< /Type /StructElem /S /Document /P ${ref(root)} /K [${all.map(ref).join(' ')}] >>`,
    );
    const nums = structElems.map((e, i) => `${i} [${e.map(ref).join(' ')}]`).join(' ');
    w.set(
      root,
      `<< /Type /StructTreeRoot /K ${ref(docElem)} /ParentTree << /Nums [${nums}] >> ` +
        `/ParentTreeNextKey ${pages.length} >>`,
    );
    extra += ` /MarkInfo << /Marked true >> /StructTreeRoot ${ref(root)}`;
  }
  const title = o.title ?? md.title;
  if (title) extra += ' /ViewerPreferences << /DisplayDocTitle true >>';
  w.set(catalog, `<< /Type /Catalog /Pages ${ref(pagesId)}${extra} >>`);

  let inf = '/Producer (turboword)';
  const meta: [string, string | undefined][] = [
    ['Title', title],
    ['Author', o.author ?? md.author],
    ['Subject', md.subject],
    ['Keywords', md.keywords],
    ['Creator', md.creator],
  ];
  for (const [k, v] of meta) if (v) inf += ` /${k} ${str(v)}`;
  if (o.date && !Number.isNaN(o.date.getTime())) {
    const d = pdfDate(o.date);
    inf += ` /CreationDate ${str(d)} /ModDate ${str(d)}`;
  }
  w.set(info, `<< ${inf} >>`);
  return w.finish(catalog, info);
}

/** Writes the /Outlines tree (levels may skip; an entry nests under the nearest shallower one). */
function writeOutline(
  w: PdfWriter,
  entries: readonly OutlineEntry[],
  dest: (page: number, y: number) => string,
): number | undefined {
  if (!entries.length) return undefined;
  interface Node {
    id: number;
    entry?: OutlineEntry;
    parent?: Node;
    children: Node[];
    count: number;
  }
  const root: Node = { id: w.alloc(), children: [], count: 0 };
  const order: Node[] = [];
  const stack: { level: number; node: Node }[] = [{ level: -Infinity, node: root }];
  for (const entry of entries) {
    while (stack.length > 1 && stack[stack.length - 1]!.level >= entry.level) stack.pop();
    const parent = stack[stack.length - 1]!.node;
    const node: Node = { id: w.alloc(), entry, parent, children: [], count: 0 };
    parent.children.push(node);
    order.push(node);
    stack.push({ level: entry.level, node });
  }
  // Preorder reversed = children before parents: every item is open, so /Count = all descendants.
  for (let i = order.length - 1; i >= 0; i--) {
    const n = order[i]!;
    n.parent!.count += 1 + n.count;
  }
  const kids = (n: Node): string =>
    n.children.length
      ? ` /First ${ref(n.children[0]!.id)} /Last ${ref(n.children[n.children.length - 1]!.id)} /Count ${n.count}`
      : '';
  for (const n of order) {
    const sib = n.parent!.children;
    const i = sib.indexOf(n);
    const e = n.entry!;
    w.set(
      n.id,
      `<< /Title ${str(e.title)} /Parent ${ref(n.parent!.id)}` +
        `${i > 0 ? ` /Prev ${ref(sib[i - 1]!.id)}` : ''}` +
        `${i < sib.length - 1 ? ` /Next ${ref(sib[i + 1]!.id)}` : ''}` +
        `${kids(n)} /Dest ${dest(e.page, e.y)} >>`,
    );
  }
  w.set(root.id, `<< /Type /Outlines${kids(root)} >>`);
  return root.id;
}

/** Name-tree keys sort by the bytes of their PDF string encoding (see writer.str). */
function compareKeys(a: string, b: string): number {
  const ka = keyBytes(a);
  const kb = keyBytes(b);
  for (let i = 0; i < Math.min(ka.length, kb.length); i++) {
    if (ka[i] !== kb[i]) return ka[i]! - kb[i]!;
  }
  return ka.length - kb.length;
}

function keyBytes(s: string): number[] {
  if (/^[\x20-\x7e]*$/.test(s)) return [...s].map((c) => c.charCodeAt(0));
  const out = [0xfe, 0xff];
  for (let i = 0; i < s.length; i++) out.push(s.charCodeAt(i) >> 8, s.charCodeAt(i) & 0xff);
  return out;
}

/** URI actions take 7-bit ASCII: percent-encode anything else as UTF-8. */
function asciiUri(url: string): string {
  return url.replace(/[^\x20-\x7e]+/g, (s) => {
    try {
      return encodeURIComponent(s);
    } catch {
      return '';
    }
  });
}

/** D:YYYYMMDDHHmmSSZ in UTC. */
function pdfDate(d: Date): string {
  const p = (n: number): string => String(n).padStart(2, '0');
  return (
    `D:${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}` +
    `${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`
  );
}
