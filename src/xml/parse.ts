// XML parser: one pass over the whole string, building the DOM of dom.ts directly.
//
// Why a hand-written parser: it is the hot path of every edit (document.xml can be tens of MB),
// and we need things generic parsers drop or mangle: comments, processing instructions, whitespace
// text, attribute order, namespace declarations as plain attributes, and the raw whitespace around
// the root (Word writes `<?xml …?>\r\n<w:document …>`, kept raw so re-serialization is identical).
//
// Derived from the turbosheet/exceljs streaming tokenizer (`exceljs/src/xml/tokenizer.ts`):
// - only the 5 predefined entities and numeric character references are decoded; unknown entities
//   stay literal; `<!DOCTYPE …>` (internal subset included) is skipped, never expanded;
// - line endings are normalized (CRLF / CR → LF) and attribute whitespace (TAB/LF/CR → space) as
//   the XML spec requires; character references such as `&#10;` survive;
// - CDATA becomes plain text (losing the delimiters is by design).
// Unlike the tokenizer it is whole-string (no streaming), checks well-formedness of the structure
// and resolves namespaces with a copy-on-declare scope stack. An undeclared prefix is resolved
// leniently to its conventional OOXML URI (`w` → WordprocessingML, see ns.ts) or '' when it has
// none, instead of failing: real-world files from odd producers do this, and the serializer
// declares such prefixes on save.
//
// Nodes are constructed with their arrays and `.parent` set directly (no `append()` checks), so
// the resulting document is not dirty. No recursion: open elements live on explicit stacks, and an
// element is built when it closes, with an exact-size child array. GC is the main cost of parsing
// a large part, hence also the per-parse interning of names and short attribute values.

import { XmlParseError } from '../core/errors.js';
import { DEFAULT_LIMITS } from '../core/limits.js';
import { XComment, XDocument, XElement, type XNode, XProcessingInstruction, XText } from './dom.js';
import { canonicalNs, conventionalUri, NS } from './ns.js';

export interface ParseOptions {
  /** Part name, used in error messages. */
  part?: string;
  /** Max element nesting depth (default DEFAULT_LIMITS.maxXmlDepth). */
  maxDepth?: number;
}

/**
 * Parses a complete XML string into an XDocument (lossless; see dom.ts). Throws XmlParseError.
 *
 * - The XML declaration is kept verbatim in `doc.declaration`; a leading U+FEFF sets `doc.bom`.
 * - Comments and processing instructions become nodes; CDATA becomes text; all whitespace text
 *   is kept. Whitespace outside the root is kept raw (no line-ending normalization).
 * - Element `uri` and prefixed attribute URIs are canonical (Strict OOXML URIs alias their
 *   transitional twins); unprefixed attributes get '' and `xmlns*` declarations `NS.xmlns`.
 *   An undeclared prefix resolves leniently to `conventionalUri(prefix) ?? ''`.
 */
export function parseXml(text: string, options?: ParseOptions): XDocument {
  return parse(text, options?.part, options?.maxDepth ?? DEFAULT_LIMITS.maxXmlDepth);
}

/**
 * Decodes part bytes (UTF-8, or UTF-16 by BOM) and parses them; records `bom` on the result.
 *
 * `bom` means "re-emit a UTF-8 BOM on save", so it is true only for a UTF-8 BOM: UTF-16 input is
 * re-encoded as UTF-8 by the serializer. Without a BOM, UTF-16 is detected from the first `<`,
 * and a declared single-byte encoding (`encoding="windows-1252"`) is honoured when the runtime's
 * TextDecoder supports it; everything else decodes as UTF-8.
 */
export function parseXmlBytes(bytes: Uint8Array, options?: ParseOptions): XDocument {
  const b0 = bytes[0];
  const b1 = bytes[1];
  let text: string;
  let bom = false;
  if (b0 === 0xef && b1 === 0xbb && bytes[2] === 0xbf) {
    bom = true;
    text = utf8().decode(bytes.subarray(3));
  } else if ((b0 === 0xff && b1 === 0xfe) || (b0 === 0x3c && b1 === 0)) {
    text = new TextDecoder('utf-16le').decode(bytes); // the decoder drops the BOM
  } else if ((b0 === 0xfe && b1 === 0xff) || (b0 === 0 && b1 === 0x3c)) {
    text = new TextDecoder('utf-16be').decode(bytes);
  } else {
    const enc = declaredEncoding(bytes);
    text = (enc === undefined ? undefined : decodeAs(enc, bytes)) ?? utf8().decode(bytes);
  }
  const doc = parseXml(text, options);
  doc.bom = bom;
  return doc;
}

// ---- decoding -----------------------------------------------------------------------------

let utf8Decoder: TextDecoder | undefined;

function utf8(): TextDecoder {
  // ignoreBOM: a BOM was already stripped above; any later U+FEFF is content.
  utf8Decoder ??= new TextDecoder('utf-8', { ignoreBOM: true });
  return utf8Decoder;
}

/** `encoding="…"` of an ASCII XML declaration when it is not UTF-8, else undefined. */
function declaredEncoding(bytes: Uint8Array): string | undefined {
  if (bytes[0] !== 0x3c || bytes[1] !== 0x3f) return undefined; // "<?"
  let end = 2;
  const max = Math.min(bytes.length, 200);
  while (end < max && bytes[end] !== 0x3e) end++;
  let head = '';
  for (let i = 0; i < end; i++) head += String.fromCharCode(bytes[i]!);
  const m = /^<\?xml\s[^>]*?\bencoding\s*=\s*["']([A-Za-z0-9._-]+)["']/.exec(head);
  if (m === null) return undefined;
  const enc = m[1]!.toLowerCase();
  // UTF-16 is detected from the bytes themselves; an ASCII-compatible file claiming it is wrong.
  return enc === 'utf-8' || enc === 'utf8' || enc.startsWith('utf-16') ? undefined : enc;
}

function decodeAs(encoding: string, bytes: Uint8Array): string | undefined {
  try {
    return new TextDecoder(encoding).decode(bytes);
  } catch {
    return undefined; // unknown label: keep the UTF-8 decoding
  }
}

// ---- parser -------------------------------------------------------------------------------

const XMLNS: string = NS.xmlns;
const LT = 60; // <
const GT = 62; // >
const SLASH = 47; // /
const BANG = 33; // !
const QMARK = 63; // ?
const EQ = 61; // =
const DQ = 34; // "
const SQ = 39; // '
const AMP = 38; // &
const COLON = 58; // :
const LBRACK = 91; // [
const RBRACK = 93; // ]

/** Namespace scope: shared by reference until an element declares something (copy-on-declare). */
interface Scope {
  /** prefix → canonical URI. */
  readonly ns: Map<string, string>;
  /** Canonical default namespace ('' when none). */
  readonly def: string;
  /** Memo: prefixed qualified name → canonical URI of its prefix in this scope. */
  readonly names: Map<string, string>;
}

function fail(message: string, offset: number, part: string | undefined): never {
  throw new XmlParseError(message, offset, part);
}

function isWs(c: number): boolean {
  return c === 32 || c === 10 || c === 9 || c === 13;
}

/**
 * Per-parse string interning: element/attribute names and short attribute values repeat
 * massively (`w:r`, `w:val`, `"22"`), so sharing one string per distinct value cuts allocation
 * and the heap the DOM keeps alive. Direct-mapped table indexed by a hash computed while
 * scanning; a collision just allocates.
 */
interface Interner {
  readonly s: string;
  readonly mask: number;
  readonly table: (string | undefined)[];
}

function interner(s: string): Interner {
  const size = s.length < 16384 ? 256 : 4096;
  return { s, mask: size - 1, table: new Array<string | undefined>(size).fill(undefined) };
}

/** `s.slice(a, b)`, shared when seen before; `h` is the hash of those characters. */
function intern(t: Interner, a: number, b: number, h: number): string {
  const idx = (h ^ (h >>> 13)) & t.mask;
  const c = t.table[idx];
  const s = t.s;
  if (c !== undefined && c.length === b - a) {
    let k = 0;
    while (a + k < b && c.charCodeAt(k) === s.charCodeAt(a + k)) k++;
    if (a + k === b) return c;
  }
  const v = s.slice(a, b);
  t.table[idx] = v;
  return v;
}

/**
 * Longest attribute value worth interning: covers `w:val` keywords, numbers and rsids (8 hex
 * digits, repeated across a document); longer values are mostly unique.
 */
const INTERN_VALUE_MAX = 12;

function parse(s: string, part: string | undefined, maxDepth: number): XDocument {
  const n = s.length;
  const doc = new XDocument();
  let p = 0;
  if (n > 0 && s.charCodeAt(0) === 0xfeff) {
    doc.bom = true;
    p = 1;
  }
  if (s.startsWith('<?xml', p) && (isWs(s.charCodeAt(p + 5)) || s.charCodeAt(p + 5) === QMARK)) {
    const e = s.indexOf('?>', p + 5);
    if (e < 0) fail('Unterminated XML declaration', p, part);
    doc.declaration = s.slice(p, e + 2);
    p = e + 2;
  }

  const strings = interner(s);
  const locals = new Map<string, string>(); // interned qualified name → interned local name
  // Children of every open element, concatenated: an element is built when it closes, with an
  // exact-size child array (growing arrays by push would leave most of them 17 slots long).
  const buf: XNode[] = [];
  let top = 0;
  // Open elements (parallel stacks).
  const oName: string[] = [];
  const oUri: string[] = [];
  const oAttrs: string[][] = [];
  const oStart: number[] = [];
  const oScope: Scope[] = []; // scope to restore when the element closes
  let scope: Scope = { ns: new Map([['xml', NS.xml]]), def: '', names: new Map() };
  const ab: string[] = []; // attribute triples of the current start tag
  let rootSeen = false;
  // Pending character data, merged across text and CDATA (decoded inside the root, raw outside).
  let text = '';

  while (p < n) {
    if (s.charCodeAt(p) !== LT) {
      let lt = s.indexOf('<', p);
      if (lt < 0) lt = n;
      if (oName.length > 0) {
        const t = decodeText(s.slice(p, lt));
        text = text === '' ? t : text + t;
      } else {
        for (let i = p; i < lt; i++) {
          if (!isWs(s.charCodeAt(i))) fail('Text outside the root element', i, part);
        }
        text += s.slice(p, lt);
      }
      p = lt;
      continue;
    }

    const c = s.charCodeAt(p + 1);

    // ---- end tag
    if (c === SLASH) {
      const e = s.indexOf('>', p + 2);
      if (e < 0) fail('Unterminated end tag', p, part);
      let ne = e;
      while (ne > p + 2 && isWs(s.charCodeAt(ne - 1))) ne--;
      const name = oName.pop();
      if (name === undefined) fail('End tag without a matching start tag', p, part);
      if (name.length !== ne - p - 2 || !s.startsWith(name, p + 2)) {
        fail(`Mismatched end tag </${s.slice(p + 2, ne)}>, expected </${name}>`, p, part);
      }
      if (text !== '') {
        buf[top++] = new XText(text);
        text = '';
      }
      const start = oStart.pop()!;
      const kids = buf.slice(start, top);
      const el = new XElement(name, oUri.pop()!, oAttrs.pop()!, kids);
      el.local = locals.get(name)!;
      for (let k = 0; k < kids.length; k++) kids[k]!.parent = el;
      top = start;
      buf[top++] = el;
      scope = oScope.pop()!;
      p = e + 1;
      continue;
    }

    // ---- processing instruction
    if (c === QMARK) {
      const e = s.indexOf('?>', p + 2);
      if (e < 0) fail('Unterminated processing instruction', p, part);
      let i = p + 2;
      while (i < e && !isWs(s.charCodeAt(i))) i++;
      const target = s.slice(p + 2, i);
      if (target === '') fail('Processing instruction without a target', p, part);
      if (target.toLowerCase() === 'xml') fail('XML declaration not at the start', p, part);
      while (i < e && isWs(s.charCodeAt(i))) i++;
      if (text !== '') {
        buf[top++] = new XText(text);
        text = '';
      }
      buf[top++] = new XProcessingInstruction(target, normalizeEol(s.slice(i, e)));
      p = e + 2;
      continue;
    }

    // ---- comment, CDATA, DOCTYPE
    if (c === BANG) {
      if (s.startsWith('--', p + 2)) {
        const e = s.indexOf('-->', p + 4);
        if (e < 0) fail('Unterminated comment', p, part);
        if (text !== '') {
          buf[top++] = new XText(text);
          text = '';
        }
        buf[top++] = new XComment(normalizeEol(s.slice(p + 4, e)));
        p = e + 3;
      } else if (s.startsWith('[CDATA[', p + 2)) {
        if (oName.length === 0) fail('CDATA outside the root element', p, part);
        const e = s.indexOf(']]>', p + 9);
        if (e < 0) fail('Unterminated CDATA section', p, part);
        const t = normalizeEol(s.slice(p + 9, e));
        text = text === '' ? t : text + t;
        p = e + 3;
      } else if (s.startsWith('DOCTYPE', p + 2)) {
        if (rootSeen) fail('DOCTYPE after the root element', p, part);
        const e = skipDoctype(s, p, n);
        if (e < 0) fail('Unterminated DOCTYPE', p, part);
        p = e;
      } else {
        fail('Invalid markup', p, part);
      }
      continue;
    }

    // ---- start tag
    let i = p + 1;
    let colon = -1;
    let h = 0;
    for (; i < n; i++) {
      const ch = s.charCodeAt(i);
      if (ch <= 32 || ch === GT || ch === SLASH) break;
      if (ch === COLON && colon < 0) colon = i;
      h = (Math.imul(h, 31) + ch) | 0;
    }
    if (i === p + 1) fail('Invalid markup', p, part);
    if (i >= n) fail('Unterminated start tag', p, part);
    const name = intern(strings, p + 1, i, h);
    let na = 0; // ab.length in use
    let declares = false;
    let empty = false;
    for (;;) {
      let ch = s.charCodeAt(i);
      while (ch <= 32) ch = s.charCodeAt(++i); // NaN past the end stops the loop
      if (i >= n) fail('Unterminated start tag', p, part);
      if (ch === GT) {
        i++;
        break;
      }
      if (ch === SLASH) {
        if (s.charCodeAt(i + 1) !== GT) fail("Expected '>' after '/'", i, part);
        i += 2;
        empty = true;
        break;
      }
      const as = i;
      h = 0;
      while (i < n) {
        ch = s.charCodeAt(i);
        if (ch === EQ || ch <= 32 || ch === GT || ch === SLASH) break;
        h = (Math.imul(h, 31) + ch) | 0;
        i++;
      }
      if (i === as) fail('Invalid attribute name', i, part);
      const an = intern(strings, as, i, h);
      while (isWs(ch)) ch = s.charCodeAt(++i);
      if (ch !== EQ) fail(`Attribute ${an} has no value`, as, part);
      ch = s.charCodeAt(++i);
      while (isWs(ch)) ch = s.charCodeAt(++i);
      if (ch !== DQ && ch !== SQ) fail(`Attribute ${an} value is not quoted`, i, part);
      const ve = s.indexOf(ch === DQ ? '"' : "'", i + 1);
      if (ve < 0) fail('Unterminated attribute value', i, part);
      for (let k = 0; k < na; k += 3) {
        if (ab[k] === an) fail(`Duplicate attribute ${an}`, as, part);
      }
      const value = attrValue(s, i + 1, ve, strings);
      if (
        an.charCodeAt(0) === 120 /* x */ &&
        an.startsWith('xmlns') &&
        (an.length === 5 || an.charCodeAt(5) === COLON)
      ) {
        // Only the default namespace can be undeclared (`xmlns=""`); `xmlns:p=""` is illegal.
        if (value === '' && an.length > 5) fail(`Empty namespace declaration ${an}`, as, part);
        declares = true;
      }
      ab[na] = an;
      ab[na + 1] = '';
      ab[na + 2] = value;
      na += 3;
      i = ve + 1;
    }

    if (oName.length === 0) {
      if (rootSeen) fail('More than one root element', p, part);
      rootSeen = true;
    }
    if (oName.length >= maxDepth) fail(`Element nesting deeper than ${maxDepth}`, p, part);

    const outer = scope;
    if (declares) scope = declare(scope, ab, na);
    let uri: string;
    if (colon < 0) uri = scope.def;
    else uri = scope.names.get(name) ?? resolve(scope, name, colon - p - 1);
    for (let k = 0; k < na; k += 3) {
      const an = ab[k]!;
      const ci = an.indexOf(':');
      if (ci < 0) ab[k + 1] = an === 'xmlns' ? XMLNS : '';
      else ab[k + 1] = scope.names.get(an) ?? resolve(scope, an, ci);
    }
    const attrs = na === 0 ? [] : ab.slice(0, na);
    let local = locals.get(name);
    if (local === undefined) {
      local = colon < 0 ? name : name.slice(colon - p);
      locals.set(name, local);
    }

    if (text !== '') {
      buf[top++] = new XText(text);
      text = '';
    }
    if (empty) {
      const el = new XElement(name, uri, attrs);
      el.local = local;
      buf[top++] = el;
      scope = outer;
    } else {
      oName.push(name);
      oUri.push(uri);
      oAttrs.push(attrs);
      oStart.push(top);
      oScope.push(outer);
    }
    p = i;
  }

  if (oName.length > 0) fail(`Unclosed element <${oName[oName.length - 1]}>`, n, part);
  if (!rootSeen) fail('No root element', n, part);
  if (text !== '') buf[top++] = new XText(text);
  const kids = buf.slice(0, top);
  for (const k of kids) k.parent = doc;
  doc._children = kids;
  return doc;
}

/** Attribute value between `a` and `b`: decoded; short plain values interned. */
function attrValue(s: string, a: number, b: number, strings: Interner): string {
  if (b - a > INTERN_VALUE_MAX) return decodeAttr(s.slice(a, b));
  let h = 0;
  for (let k = a; k < b; k++) {
    const c = s.charCodeAt(k);
    if (c === AMP || c === 9 || c === 10 || c === 13) return decodeSlow(s.slice(a, b), k - a, true);
    h = (Math.imul(h, 31) + c) | 0;
  }
  return intern(strings, a, b, h);
}

/** New scope with the `xmlns*` declarations among the first `na` of `attrs` applied over `outer`. */
function declare(outer: Scope, attrs: readonly string[], na: number): Scope {
  const ns = new Map(outer.ns);
  let def = outer.def;
  for (let k = 0; k < na; k += 3) {
    const an = attrs[k]!;
    if (an === 'xmlns') def = canonicalNs(attrs[k + 2]!);
    else if (an.startsWith('xmlns:')) {
      const prefix = an.slice(6);
      if (prefix === 'xml' || prefix === 'xmlns') continue;
      ns.set(prefix, canonicalNs(attrs[k + 2]!)); // never '' (rejected while parsing the tag)
    }
  }
  return { ns, def, names: new Map() };
}

/** Canonical URI of the prefix of qualified name `qname` (prefix length `ci`); memoized. */
function resolve(scope: Scope, qname: string, ci: number): string {
  const prefix = qname.slice(0, ci);
  const uri =
    prefix === 'xmlns'
      ? XMLNS
      : (scope.ns.get(prefix) ?? canonicalNs(conventionalUri(prefix) ?? ''));
  scope.names.set(qname, uri);
  return uri;
}

/** Skips `<!DOCTYPE …>` including an internal subset `[…]`; returns the end index or -1. */
function skipDoctype(s: string, p: number, n: number): number {
  let i = p + 2;
  let depth = 0;
  while (i < n) {
    const c = s.charCodeAt(i);
    if (c === DQ || c === SQ) {
      const e = s.indexOf(c === DQ ? '"' : "'", i + 1);
      if (e < 0) return -1;
      i = e + 1;
      continue;
    }
    if (c === LT && depth > 0) {
      if (s.startsWith('<!--', i)) {
        const e = s.indexOf('-->', i + 4);
        if (e < 0) return -1;
        i = e + 3;
        continue;
      }
      if (s.charCodeAt(i + 1) === QMARK) {
        const e = s.indexOf('?>', i + 2);
        if (e < 0) return -1;
        i = e + 2;
        continue;
      }
    } else if (c === LBRACK) depth++;
    else if (c === RBRACK) depth--;
    else if (c === GT && depth <= 0) return i + 1;
    i++;
  }
  return -1;
}

// ---- character data -----------------------------------------------------------------------

/** CRLF / CR → LF. */
function normalizeEol(s: string): string {
  let i = s.indexOf('\r');
  if (i < 0) return s;
  let out = '';
  let last = 0;
  while (i >= 0) {
    out += `${s.slice(last, i)}\n`;
    last = s.charCodeAt(i + 1) === 10 ? i + 2 : i + 1;
    i = s.indexOf('\r', last);
  }
  return out + s.slice(last);
}

/** Text content: entities decoded, line endings normalized. */
function decodeText(s: string): string {
  if (s.indexOf('&') < 0) return normalizeEol(s);
  return decodeSlow(s, 0, false);
}

/** Attribute value: entities decoded, literal TAB/LF/CR (CRLF as one) → space. */
function decodeAttr(s: string): string {
  const n = s.length;
  for (let i = 0; i < n; i++) {
    const c = s.charCodeAt(i);
    if (c === AMP || c === 9 || c === 10 || c === 13) return decodeSlow(s, i, true);
  }
  return s;
}

function decodeSlow(s: string, start: number, attr: boolean): string {
  const n = s.length;
  let out = '';
  let last = 0;
  for (let i = start; i < n; i++) {
    const c = s.charCodeAt(i);
    if (c === AMP) {
      const semi = s.indexOf(';', i + 1);
      if (semi < 0 || semi - i > 12) continue;
      const r = resolveEntity(s, i + 1, semi);
      if (r === undefined) continue; // unknown entity: kept literally
      out += s.slice(last, i) + r;
      last = semi + 1;
      i = semi;
    } else if (c === 13) {
      out += s.slice(last, i) + (attr ? ' ' : '\n');
      if (s.charCodeAt(i + 1) === 10) i++;
      last = i + 1;
    } else if (attr && (c === 10 || c === 9)) {
      out += `${s.slice(last, i)} `;
      last = i + 1;
    }
  }
  return last === 0 ? s : out + s.slice(last);
}

const HEX = /^[0-9a-fA-F]+$/;
const DEC = /^[0-9]+$/;

/** The 5 predefined entities and numeric character references; undefined otherwise. */
function resolveEntity(s: string, a: number, b: number): string | undefined {
  const len = b - a;
  const c0 = s.charCodeAt(a);
  if (c0 === 35 /* # */) {
    let cp: number;
    if (len > 2 && (s.charCodeAt(a + 1) | 0x20) === 120 /* x */) {
      const hex = s.slice(a + 2, b);
      if (!HEX.test(hex)) return undefined;
      cp = Number.parseInt(hex, 16);
    } else {
      const dec = s.slice(a + 1, b);
      if (!DEC.test(dec)) return undefined;
      cp = Number.parseInt(dec, 10);
    }
    if (cp === 0 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff)) return undefined;
    return String.fromCodePoint(cp);
  }
  switch (len) {
    case 2:
      if (s.charCodeAt(a + 1) === 116 /* t */) {
        if (c0 === 108 /* l */) return '<';
        if (c0 === 103 /* g */) return '>';
      }
      return undefined;
    case 3:
      return s.startsWith('amp', a) ? '&' : undefined;
    case 4:
      if (s.startsWith('quot', a)) return '"';
      if (s.startsWith('apos', a)) return "'";
      return undefined;
    default:
      return undefined;
  }
}
