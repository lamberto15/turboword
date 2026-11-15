// PartCopier: brings source parts into the target package, and remapRelationships, which
// rewrites every relationship reference (`r:*` attributes) of imported markup through it.
//
// Copying rules:
// - A part is copied once (memoized per source part) together with everything it references,
//   each copied part keeping its OWN relationship ids: its rels part is new, so the ids are free,
//   and the copied XML therefore needs no rewriting. Only markup merged into an existing target
//   part (body, header, notes, comments content) gets new ids, through copyRelationship.
// - New names come from `pkg.uniqueName` with the source's naming pattern (`word/media/image7.png`
//   → `word/media/image{n}.png`); the content type is carried over (as a Default when the source
//   used one and the target has none for the extension, else as an Override).
// - Media (images, audio, video) are deduped against the target: same size and CRC-32 (from the
//   zip directory when available), confirmed by a byte compare.
// - External relationships are re-created as external.
// - Relationships to document-wide singletons (styles, numbering, settings, theme, notes,
//   comments, ...) are never followed into a copy: they point at the target's own part, which the
//   compose steps merge. Glossary documents are never copied.
// - customXml items are not duplicated: an item whose itemProps storeItemID the target already
//   has maps to the target's item.
// - Known exception: SmartArt. The diagram data part names the diagramDrawing part through
//   `dsp:dataModelExt/@relId`, an id of the HOST part's relationships (not its own), so that id
//   is copied host-to-host and rewritten in the copied data part.

import type { OpcPackage } from '../opc/package.js';
import type { Part } from '../opc/part.js';
import { relTypeIs } from '../opc/rel-types.js';
import { partExtension, relativeTarget } from '../opc/uri.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import { crc32 } from '../zip/crc32.js';
import type { ImportContext, PartCopier } from './types.js';

/** Office 2008 diagram drawing namespace (`dsp:`). */
const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
/** customXml datastore item properties (`ds:`). */
const DS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';

/**
 * Relationship kinds (last type segment) of parts that exist once per document: never copied,
 * mapped to the target's own part instead.
 */
const SINGLETONS: readonly string[] = [
  'officeDocument',
  'styles',
  'stylesWithEffects',
  'numbering',
  'settings',
  'webSettings',
  'fontTable',
  'theme',
  'footnotes',
  'endnotes',
  'comments',
  'commentsExtended',
  'commentsIds',
  'commentsExtensible',
  'people',
  'core-properties',
  'extended-properties',
  'custom-properties',
];

/** Media relationship kinds deduped by content. */
const MEDIA_TYPES = /^(?:image|audio|video)\//i;

function singletonKind(type: string): string | undefined {
  return SINGLETONS.find((k) => relTypeIs(type, k));
}

/** `word/media/image7.png` → `word/media/image{n}.png` (trailing digits of the stem dropped). */
export function namePattern(name: string): string {
  const slash = name.lastIndexOf('/');
  const dir = name.slice(0, slash + 1);
  const file = name.slice(slash + 1);
  const dot = file.lastIndexOf('.');
  let stem = dot > 0 ? file.slice(0, dot) : file;
  const ext = dot > 0 ? file.slice(dot) : '';
  stem = stem.replace(/\d+$/, '').replaceAll('{n}', '');
  if (stem === '') stem = 'part';
  return `${dir}${stem}{n}${ext}`;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** Size and CRC-32 of a part, from the zip directory when its content is the untouched entry. */
function fingerprint(part: Part): string {
  const e = part._entry;
  if (e !== undefined && part._data === undefined && !part.dirty) return `${e.size}:${e.crc32}`;
  const b = part.bytes();
  return `${b.length}:${crc32(b)}`;
}

function rootOf(part: Part): XElement | undefined {
  try {
    for (const k of part.xml().children) if (k.nodeType === 1) return k;
  } catch {
    // not XML / malformed: no root
  }
  return undefined;
}

/** storeItemID of a customXml item (from its itemProps part), upper-cased; undefined if none. */
function storeItemId(item: Part): string | undefined {
  for (const props of item.relatedByType('customXmlProps')) {
    const id = rootOf(props)?.attrNS(DS_NS, 'itemID');
    if (id !== undefined) return id.toUpperCase();
  }
  return undefined;
}

export interface CopierOptions {
  /** Receives the names of parts added to the target. */
  added: string[];
  warn(code: string, message: string, part?: string): void;
  sourceMain: Part;
  targetMain: Part;
}

/** The {@link PartCopier} of one import (source package → target package). */
export class PackageCopier implements PartCopier {
  readonly #src: OpcPackage;
  readonly #dst: OpcPackage;
  readonly #o: CopierOptions;
  /** Source part → its copy (or the deduped / mapped target part). */
  readonly #parts = new Map<Part, Part>();
  /** `sourcePart\0relId\0targetPart` → relationship id in the target part. */
  readonly #rels = new Map<string, string>();
  /** Target media by fingerprint (built on first use). */
  #media: Map<string, Part[]> | undefined;
  /** Copied diagram data parts whose dataModelExt was already rewritten. */
  readonly #diagrams = new Set<Part>();

  constructor(source: OpcPackage, target: OpcPackage, o: CopierOptions) {
    this.#src = source;
    this.#dst = target;
    this.#o = o;
  }

  copyPart(sourcePart: Part): Part {
    if (sourcePart.package === this.#dst) return sourcePart;
    const done = this.#parts.get(sourcePart);
    if (done !== undefined) return done;
    const ct = sourcePart.contentType ?? 'application/octet-stream';
    const bytes = sourcePart.bytes();
    if (!sourcePart.isXml && (MEDIA_TYPES.test(ct) || sourcePart.name.includes('/media/'))) {
      const same = this.#findMedia(sourcePart, bytes);
      if (same !== undefined) {
        this.#parts.set(sourcePart, same);
        return same;
      }
    }
    const name = this.#dst.uniqueName(namePattern(sourcePart.name));
    const ext = partExtension(name);
    const fromDefault = this.#src.contentTypes.getOverride(sourcePart.name) === undefined;
    const types = this.#dst.contentTypes;
    if (ext !== undefined && fromDefault && types.getDefault(ext) === undefined) {
      types.setDefault(ext, ct);
    }
    const copy = this.#dst.addPart(name, ct, bytes.slice());
    this.#o.added.push(name);
    this.#parts.set(sourcePart, copy);
    if (!sourcePart.isXml) this.#indexMedia(copy);
    for (const rel of sourcePart.rels.all()) {
      if (rel.external) {
        copy.rels.add(rel.type, rel.target, { external: true, id: rel.id });
        continue;
      }
      const target = this.#resolveInternal(sourcePart, rel.type, rel.partName);
      if (target === undefined) continue;
      copy.rels.add(rel.type, relativeTarget(copy.name, target.name), { id: rel.id });
    }
    return copy;
  }

  copyRelationship(sourcePart: Part, relId: string, targetPart: Part): string {
    const key = `${sourcePart.name}\0${relId}\0${targetPart.name}`;
    const memo = this.#rels.get(key);
    if (memo !== undefined) return memo;
    const id = this.#copyRelationship(sourcePart, relId, targetPart);
    this.#rels.set(key, id);
    return id;
  }

  /**
   * Makes the source customXml item with this storeItemID available to the target (copied
   * unless the target has an item with the same id). Returns false when the source has none.
   */
  copyCustomXmlItem(storeItemID: string): boolean {
    const want = storeItemID.toUpperCase();
    const main = this.#o.sourceMain;
    for (const rel of main.rels.byType('customXml')) {
      const item = main.related(rel.id);
      if (item === undefined || storeItemId(item) !== want) continue;
      this.copyRelationship(main, rel.id, this.#o.targetMain);
      return true;
    }
    return false;
  }

  #copyRelationship(sourcePart: Part, relId: string, targetPart: Part): string {
    const rel = sourcePart.rels.get(relId);
    if (rel === undefined) {
      this.#o.warn(
        'COMPOSE_DANGLING_REL',
        `Relationship ${relId} of ${sourcePart.name} does not exist; reference dropped`,
        sourcePart.name,
      );
      return '';
    }
    if (rel.external) {
      if (sourcePart === targetPart) return relId;
      return targetPart.rels.add(rel.type, rel.target, { external: true }).id;
    }
    if (sourcePart.package === this.#dst) {
      // Same package (importBlocks within one document): no copy, just a relationship.
      if (sourcePart === targetPart) return relId;
      const part = rel.partName === undefined ? undefined : this.#dst.part(rel.partName);
      return part === undefined ? '' : this.#relate(targetPart, rel.type, part);
    }
    if (relTypeIs(rel.type, 'glossaryDocument')) return '';
    let part: Part | undefined;
    if (relTypeIs(rel.type, 'customXml')) {
      const src = rel.partName === undefined ? undefined : this.#src.part(rel.partName);
      part = src === undefined ? undefined : (this.#existingCustomXml(src) ?? this.copyPart(src));
    } else {
      part = this.#resolveInternal(sourcePart, rel.type, rel.partName);
    }
    if (part === undefined) return '';
    const id = this.#relate(targetPart, rel.type, part);
    if (relTypeIs(rel.type, 'diagramData')) this.#fixDiagramData(part, sourcePart, targetPart);
    return id;
  }

  /**
   * The target part an internal source relationship should point at: the target's own part for
   * singletons, else a copy. Undefined (with a warning) when there is none.
   */
  #resolveInternal(from: Part, type: string, partName: string | undefined): Part | undefined {
    const src = partName === undefined ? undefined : this.#src.part(partName);
    if (src === undefined) {
      this.#o.warn(
        'COMPOSE_MISSING_PART',
        `${from.name} references missing part ${partName ?? '?'}; relationship dropped`,
        from.name,
      );
      return undefined;
    }
    if (relTypeIs(type, 'glossaryDocument')) return undefined;
    const kind = singletonKind(type);
    if (kind !== undefined || src === this.#o.sourceMain) {
      if (src === this.#o.sourceMain || kind === 'officeDocument') return this.#o.targetMain;
      const own = this.#o.targetMain.relatedByType(kind!)[0];
      if (own === undefined) {
        this.#o.warn(
          'COMPOSE_SINGLETON_MISSING',
          `${from.name} references a ${kind} part the target does not have; relationship dropped`,
          from.name,
        );
      }
      return own;
    }
    return this.copyPart(src);
  }

  /** Id of a relationship of `type` from `from` to `part` (an existing one is reused). */
  #relate(from: Part, type: string, part: Part): string {
    for (const r of from.rels.all()) {
      if (!r.external && r.partName === part.name && relTypeIs(r.type, type)) return r.id;
    }
    return from.rels.add(type, relativeTarget(from.name, part.name)).id;
  }

  #existingCustomXml(src: Part): Part | undefined {
    const id = storeItemId(src);
    if (id === undefined) return undefined;
    for (const item of this.#o.targetMain.relatedByType('customXml')) {
      if (storeItemId(item) === id) {
        this.#parts.set(src, item);
        return item;
      }
    }
    return undefined;
  }

  /** Rewrites `dsp:dataModelExt/@relId` of a copied diagram data part (host-relative id). */
  #fixDiagramData(data: Part, sourceHost: Part, targetHost: Part): void {
    if (this.#diagrams.has(data) || data.package !== this.#dst) return;
    this.#diagrams.add(data);
    const root = rootOf(data);
    if (root === undefined) return;
    for (const ext of root.descendants(DSP_NS, 'dataModelExt')) {
      const old = ext.attrNS('', 'relId');
      if (old === undefined || old === '') continue;
      const id = this.copyRelationship(sourceHost, old, targetHost);
      if (id !== old) ext.setAttr('relId', id);
    }
  }

  #findMedia(source: Part, bytes: Uint8Array): Part | undefined {
    const index = this.#mediaIndex();
    const fp = `${bytes.length}:${crc32(bytes)}`;
    for (const cand of index.get(fp) ?? []) {
      if (cand.contentType !== source.contentType) continue;
      if (sameBytes(cand.bytes(), bytes)) return cand;
    }
    return undefined;
  }

  #mediaIndex(): Map<string, Part[]> {
    if (this.#media !== undefined) return this.#media;
    this.#media = new Map();
    for (const part of this.#dst.parts()) {
      if (part.isXml || part.name.endsWith('/')) continue;
      const ct = part.contentType ?? '';
      if (MEDIA_TYPES.test(ct) || part.name.includes('/media/')) this.#indexMedia(part);
    }
    return this.#media;
  }

  #indexMedia(part: Part): void {
    if (this.#media === undefined) return; // indexed lazily on first lookup
    const fp = fingerprint(part);
    const list = this.#media.get(fp);
    if (list === undefined) this.#media.set(fp, [part]);
    else list.push(part);
  }
}

/** True for `w:headerReference`/`w:footerReference` (left to importSections). */
function isHeaderFooterRef(el: XElement): boolean {
  return el.uri === NS.w && (el.local === 'headerReference' || el.local === 'footerReference');
}

/**
 * Rewrites every relationship reference in `elements` (and their descendants) from ids of
 * `sourcePart` to ids of `targetPart`, copying the referenced parts through ctx.copier: every
 * attribute in the relationships namespace (`r:id`, `r:embed`, `r:link`, `r:dm`/`lo`/`qs`/`cs`,
 * `r:pict`, ...) and VML's `o:relid`. Header/footer references inside section properties are
 * left alone: importSections copies those parts itself (with their content imported).
 */
export function remapRelationships(
  ctx: ImportContext,
  sourcePart: Part,
  targetPart: Part,
  elements: XElement[],
): void {
  const stack: XElement[] = [];
  for (let i = elements.length - 1; i >= 0; i--) stack.push(elements[i]!);
  while (stack.length > 0) {
    const el = stack.pop()!;
    const kids = el._children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
    const a = el._attrs;
    if (a.length === 0 || isHeaderFooterRef(el)) continue;
    for (let i = 0; i < a.length; i += 3) {
      const uri = a[i + 1];
      const isRel = uri === NS.r || (uri === NS.o && a[i]!.endsWith(':relid'));
      if (!isRel) continue;
      const old = a[i + 2]!;
      if (old === '') continue;
      const id = ctx.copier.copyRelationship(sourcePart, old, targetPart);
      if (id !== old) el.setAttr(a[i]!, id);
    }
  }
}
