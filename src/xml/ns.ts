// Namespace URIs, Strict OOXML aliases and well-known prefixes.
//
// Elements store a *canonical* namespace URI: the Strict OOXML (ISO/IEC 29500 Strict) URIs are
// mapped to their transitional twins, so `el.is(NS.w, 'p')` matches in both flavours. The URI
// actually declared in the file stays in its `xmlns:*` attribute, so nothing changes on save.

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  w16: 'http://schemas.microsoft.com/office/word/2018/wordml',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
  w16du: 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wne: 'http://schemas.microsoft.com/office/word/2006/wordml',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  w10: 'urn:schemas-microsoft-com:office:word',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
  ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  custom: 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
  vt: 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  xml: 'http://www.w3.org/XML/1998/namespace',
  xmlns: 'http://www.w3.org/2000/xmlns/',
} as const;

/** Strict OOXML namespace URI → transitional URI. */
const STRICT: ReadonlyMap<string, string> = /* @__PURE__ */ new Map([
  ['http://purl.oclc.org/ooxml/wordprocessingml/main', NS.w],
  ['http://purl.oclc.org/ooxml/officeDocument/relationships', NS.r],
  ['http://purl.oclc.org/ooxml/officeDocument/math', NS.m],
  ['http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing', NS.wp],
  ['http://purl.oclc.org/ooxml/drawingml/main', NS.a],
  ['http://purl.oclc.org/ooxml/drawingml/picture', NS.pic],
  ['http://purl.oclc.org/ooxml/drawingml/chart', NS.c],
  ['http://purl.oclc.org/ooxml/drawingml/diagram', NS.dgm],
  [
    'http://purl.oclc.org/ooxml/drawingml/chartDrawing',
    'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing',
  ],
  [
    'http://purl.oclc.org/ooxml/drawingml/lockedCanvas',
    'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas',
  ],
  [
    'http://purl.oclc.org/ooxml/drawingml/compatibility',
    'http://schemas.openxmlformats.org/drawingml/2006/compatibility',
  ],
  ['http://purl.oclc.org/ooxml/officeDocument/extendedProperties', NS.ep],
  ['http://purl.oclc.org/ooxml/officeDocument/customProperties', NS.custom],
  ['http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes', NS.vt],
  [
    'http://purl.oclc.org/ooxml/officeDocument/sharedTypes',
    'http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes',
  ],
  [
    'http://purl.oclc.org/ooxml/officeDocument/bibliography',
    'http://schemas.openxmlformats.org/officeDocument/2006/bibliography',
  ],
  [
    'http://purl.oclc.org/ooxml/officeDocument/customXml',
    'http://schemas.openxmlformats.org/officeDocument/2006/customXml',
  ],
]);

/** Canonical form of a namespace URI (Strict OOXML URIs map to their transitional twins). */
export function canonicalNs(uri: string): string {
  return STRICT.get(uri) ?? uri;
}

/** True for a Strict OOXML namespace URI. */
export function isStrictNs(uri: string): boolean {
  return STRICT.has(uri);
}

/** Conventional prefix for a (canonical) namespace URI → used when a new prefix is needed. */
const PREFIX_OF: ReadonlyMap<string, string> = /* @__PURE__ */ new Map(
  Object.entries(NS)
    .filter(([p]) => p !== 'xmlns' && p !== 'rel' && p !== 'ct' && p !== 'ep' && p !== 'custom')
    .map(([p, u]) => [u, p]),
);

/** Conventional prefix → URI, for qualified names whose prefix is not declared in scope. */
const URI_OF: ReadonlyMap<string, string> = /* @__PURE__ */ new Map(
  [...PREFIX_OF].map(([u, p]) => [p, u]),
);

/** Conventional prefix for a namespace URI, or undefined when it has none. */
export function conventionalPrefix(uri: string): string | undefined {
  return PREFIX_OF.get(canonicalNs(uri));
}

/** URI conventionally bound to `prefix` (`w`, `r`, `wp`, ...), or undefined. */
export function conventionalUri(prefix: string): string | undefined {
  return URI_OF.get(prefix);
}

/**
 * Namespaces Word only understands as extensions: when one of them is declared on a part's root
 * because content uses it, its prefix must also be listed in `mc:Ignorable` (or old consumers
 * reject the part).
 */
export const IGNORABLE_NS: ReadonlySet<string> = /* @__PURE__ */ new Set<string>([
  NS.w14,
  NS.w15,
  NS.w16,
  NS.w16se,
  NS.w16cid,
  NS.w16cex,
  NS.w16sdtdh,
  NS.w16du,
  NS.wp14,
]);

/** Prefix of a qualified name (`'w:p'` → `'w'`, `'p'` → `''`). */
export function prefixOf(qname: string): string {
  const i = qname.indexOf(':');
  return i < 0 ? '' : qname.slice(0, i);
}

/** Local part of a qualified name (`'w:p'` → `'p'`). */
export function localOf(qname: string): string {
  const i = qname.indexOf(':');
  return i < 0 ? qname : qname.slice(i + 1);
}

/** True when qualified name `qname` has local part `local` (no allocation). */
export function hasLocal(qname: string, local: string): boolean {
  const d = qname.length - local.length;
  if (d === 0) return qname === local;
  return d > 1 && qname.charCodeAt(d - 1) === 58 /* : */ && qname.endsWith(local);
}
