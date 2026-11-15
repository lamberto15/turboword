// Part-name helpers for OPC packages (ECMA-376 Part 2).
//
// A part name is a zip entry name without its leading slash ("word/document.xml"). Relationship
// targets are URIs relative to the source part's directory (or absolute from the package root),
// may be percent-encoded and may carry a fragment, so they are resolved here into part names.
// Part names compare case-insensitively (ASCII) and after percent-decoding: `partKey` gives the
// comparison key used by the package's part map and by content-type lookups.

/** Normalized part name: leading '/' stripped. */
export function normalizePartName(name: string): string {
  let i = 0;
  while (i < name.length && name.charCodeAt(i) === 47) i++;
  return i === 0 ? name : name.slice(i);
}

/** Percent-decodes `s`, returning it unchanged when it holds a malformed escape. */
function safeDecode(s: string): string {
  if (s.indexOf('%') < 0) return s;
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/** @internal Comparison key of a part name (no leading '/', percent-decoded, lower-case). */
export function partKey(name: string): string {
  return safeDecode(normalizePartName(name)).toLowerCase();
}

/** @internal Lower-case extension of a part name ('_rels/.rels' → 'rels'), or undefined. */
export function partExtension(name: string): string | undefined {
  const slash = name.lastIndexOf('/');
  const dot = name.lastIndexOf('.');
  return dot <= slash ? undefined : name.slice(dot + 1).toLowerCase();
}

/** Directory of a part name ("word/document.xml" → "word/"; "a.xml" → ""). */
export function partDir(name: string): string {
  const i = name.lastIndexOf('/');
  return i < 0 ? '' : name.slice(0, i + 1);
}

/** Relationship part of a part ("word/document.xml" → "word/_rels/document.xml.rels"; "" → "_rels/.rels"). */
export function relsPathFor(name: string): string {
  const n = normalizePartName(name);
  const i = n.lastIndexOf('/');
  return `${n.slice(0, i + 1)}_rels/${n.slice(i + 1)}.rels`;
}

/** Source part of a relationship part ("word/_rels/document.xml.rels" → "word/document.xml"; "_rels/.rels" → ""), or undefined if not a rels path. */
export function sourceOfRels(relsName: string): string | undefined {
  const n = normalizePartName(relsName);
  const slash = n.lastIndexOf('/');
  const file = n.slice(slash + 1);
  const dir = n.slice(0, slash + 1);
  if (!file.toLowerCase().endsWith('.rels')) return undefined;
  const lowerDir = dir.toLowerCase();
  if (lowerDir !== '_rels/' && !lowerDir.endsWith('/_rels/')) return undefined;
  return dir.slice(0, -'_rels/'.length) + file.slice(0, -'.rels'.length);
}

/** Resolves a relationship target against a base directory into a part name (no leading slash, dot segments collapsed, %-decoded, fragment dropped). */
export function resolveTarget(baseDir: string, target: string): string {
  let t = target.replace(/\\/g, '/');
  const hash = t.indexOf('#');
  if (hash >= 0) t = t.slice(0, hash);
  t = safeDecode(t);
  let base = baseDir;
  if (base !== '' && !base.endsWith('/')) base += '/';
  const full = t.charCodeAt(0) === 47 ? t : base + t;
  const out: string[] = [];
  for (const seg of full.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}

/** Characters that would change a relative URI's meaning if written literally. */
const URI_UNSAFE = /[%#? ]/g;

/** Relative target from part `from` to part `to` ("word/document.xml", "word/media/a.png" → "media/a.png"). */
export function relativeTarget(from: string, to: string): string {
  const fromDir = partDir(normalizePartName(from)).split('/');
  fromDir.pop(); // partDir ends with '/' (or is ''), so the last segment is always empty
  const toSegs = normalizePartName(to).split('/');
  let common = 0;
  while (
    common < fromDir.length &&
    common < toSegs.length - 1 &&
    fromDir[common] === toSegs[common]
  ) {
    common++;
  }
  const rel = '../'.repeat(fromDir.length - common) + toSegs.slice(common).join('/');
  return rel.replace(URI_UNSAFE, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`);
}
