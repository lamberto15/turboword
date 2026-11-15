// Well-known relationship types and content types of WordprocessingML packages.
//
// Strict OOXML (ISO/IEC 29500 Strict) uses different relationship type URIs
// (`http://purl.oclc.org/ooxml/officeDocument/relationships/...`, a couple of them camel-cased),
// so types are compared through `relTypeIs`, which maps Strict URIs to their transitional twins.

/** Transitional relationship type URIs (Strict ones are matched by {@link relTypeIs}). */
export const REL = {
  officeDocument:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  coreProperties:
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  customProperties:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
  thumbnail: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  webSettings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
  fontTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
} as const;

/** Content types. */
export const CT = {
  relationships: 'application/vnd.openxmlformats-package.relationships+xml',
  coreProperties: 'application/vnd.openxmlformats-package.core-properties+xml',
  extendedProperties: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  customProperties: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
  xml: 'application/xml',
  document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  documentMacro: 'application/vnd.ms-word.document.macroEnabled.main+xml',
  template: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
  templateMacro: 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  webSettings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml',
  fontTable: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
} as const;

const STRICT_REL_BASE = 'http://purl.oclc.org/ooxml/officeDocument/relationships/';
const TRANSITIONAL_REL_BASE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
/** Strict relationship names that differ from the transitional ones beyond the base URI. */
const STRICT_RENAMED: ReadonlyMap<string, string> = /* @__PURE__ */ new Map([
  ['extendedProperties', 'extended-properties'],
  ['customProperties', 'custom-properties'],
]);

/** Transitional form of a relationship type URI (Strict URIs mapped; anything else unchanged). */
function canonicalRelType(type: string): string {
  if (!type.startsWith(STRICT_REL_BASE)) return type;
  const name = type.slice(STRICT_REL_BASE.length);
  return TRANSITIONAL_REL_BASE + (STRICT_RENAMED.get(name) ?? name);
}

/**
 * True when relationship type `type` ends with segment `kind` (matches transitional and Strict URIs).
 * `kind` may also be a full URI (e.g. `REL.header`): then both are compared in transitional form.
 */
export function relTypeIs(type: string, kind: string): boolean {
  const t = canonicalRelType(type);
  if (kind.includes('/')) return t === canonicalRelType(kind);
  return t.slice(t.lastIndexOf('/') + 1) === kind;
}
