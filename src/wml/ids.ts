// Document-wide id registry.
//
// Word requires several kinds of ids to be unique across the main document and every story part
// (headers, footers, notes, comments); duplicates are a classic cause of "unreadable content"
// after merging or editing. The registry scans those parts lazily, on first use, collecting every
// kind in one pass per part, and hands out fresh ids deterministically:
// - numeric kinds: one more than the largest id in use (Word itself counts up), falling back to
//   the smallest free value when that would overflow a signed 32-bit integer;
// - 8-hex-digit kinds (`w14:paraId`, ...): a seeded PRNG (never Math.random), values in
//   [1, 0x7FFFFFFF] as Word requires, skipping used ones;
// - VML shape ids: `_x0000_s<n>` with n above every number already used by `_x0000_?<n>` ids.
//
// Staleness: every scan remembers its part's XDocument and `version` (a mutation counter) and the
// version of the main part's relationships (which say which story parts exist). Any change
// re-scans just the changed parts. Handed-out ids stay reserved across re-scans and
// `invalidate()`, so an id given out but not yet written is never handed out twice.
// `hold()` suspends the staleness checks for a batch of edits that only write ids obtained here
// (compose uses it so that thousands of writes do not cost thousands of re-scans).

import type { OpcPackage } from '../opc/package.js';
import type { Part } from '../opc/part.js';
import { relTypeIs } from '../opc/rel-types.js';
import { relsPathFor } from '../opc/uri.js';
import type { XDocument, XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import type { WordDocument } from './document.js';

/** Numeric id kinds. */
export type NumericIdKind =
  | 'docPr' // wp:docPr/@id (drawings), also pic:cNvPr
  | 'bookmark' // w:bookmarkStart/@w:id
  | 'comment' // w:comment/@w:id (+ commentRangeStart/End, commentReference)
  | 'footnote' // w:footnote/@w:id
  | 'endnote' // w:endnote/@w:id
  | 'revision' // w:ins/w:del/... /@w:id
  | 'sdt'; // w:sdtPr/w:id/@w:val (signed 32-bit)

/** 8-hex-digit id kinds (values < 0x80000000, as Word requires). */
export type HexIdKind = 'paraId' | 'textId' | 'durableId';

const HEX: readonly HexIdKind[] = ['paraId', 'textId', 'durableId'];

/** Smallest id handed out per numeric kind (footnote/endnote 0 and -1 are separators). */
const MIN: Readonly<Record<NumericIdKind, number>> = {
  docPr: 1,
  bookmark: 0,
  comment: 0,
  footnote: 1,
  endnote: 1,
  revision: 0,
  sdt: 1,
};
const MAX_INT = 0x7fffffff;
/** First VML shape number Word uses in the main document. */
const VML_BASE = 1024;

/**
 * `w:` elements whose `w:id` is a revision id (ECMA-376 Part 1 §17.13). Range starts and ends
 * (`moveFromRangeStart`/`End`, `customXml*RangeStart`/`End`) share the id of their pair.
 */
export const REVISION_ELEMENTS: ReadonlySet<string> = /* @__PURE__ */ new Set([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
  'rPrChange',
  'pPrChange',
  'sectPrChange',
  'tblPrChange',
  'tblPrExChange',
  'trPrChange',
  'tcPrChange',
  'tblGridChange',
  'numberingChange',
  'cellIns',
  'cellDel',
  'cellMerge',
  'customXmlInsRangeStart',
  'customXmlInsRangeEnd',
  'customXmlDelRangeStart',
  'customXmlDelRangeEnd',
  'customXmlMoveFromRangeStart',
  'customXmlMoveFromRangeEnd',
  'customXmlMoveToRangeStart',
  'customXmlMoveToRangeEnd',
]);

/** Relationship kinds (from the main part) of the parts that hold ids. */
const STORY_RELS = [
  'header',
  'footer',
  'footnotes',
  'endnotes',
  'comments',
  'commentsExtended',
  'commentsIds',
  'commentsExtensible',
];

/** `_x0000_s1025`-style VML ids → their number. */
const VML_ID = /^_x0000_[a-z](\d{1,9})$/;
const HEX_ID = /^[0-9A-Fa-f]{1,8}$/;

/** Normalized 8-digit upper-case hex, or undefined when `v` is not a hex id. */
function normHex(v: string): string | undefined {
  return HEX_ID.test(v) ? v.toUpperCase().padStart(8, '0') : undefined;
}

function toInt(v: string | number): number | undefined {
  if (typeof v === 'number') return Number.isInteger(v) ? v : undefined;
  if (!/^\s*[-+]?\d{1,10}\s*$/.test(v)) return undefined;
  return Number.parseInt(v, 10);
}

/** Ids of every kind found in one part. */
class IdSets {
  readonly num: Record<NumericIdKind, Set<number>> = {
    docPr: new Set(),
    bookmark: new Set(),
    comment: new Set(),
    footnote: new Set(),
    endnote: new Set(),
    revision: new Set(),
    sdt: new Set(),
  };
  readonly max: Record<NumericIdKind, number> = {
    docPr: -Infinity,
    bookmark: -Infinity,
    comment: -Infinity,
    footnote: -Infinity,
    endnote: -Infinity,
    revision: -Infinity,
    sdt: -Infinity,
  };
  readonly hex: Record<HexIdKind, Set<string>> = {
    paraId: new Set(),
    textId: new Set(),
    durableId: new Set(),
  };
  readonly vml = new Set<number>();
  vmlMax = -Infinity;

  addNum(kind: NumericIdKind, raw: string | number | undefined): void {
    if (raw === undefined) return;
    const n = toInt(raw);
    if (n === undefined) return;
    this.num[kind].add(n);
    if (n > this.max[kind]) this.max[kind] = n;
  }

  addHex(kind: HexIdKind, raw: string | undefined): void {
    if (raw === undefined) return;
    const h = normHex(raw);
    if (h !== undefined) this.hex[kind].add(h);
  }

  addVml(raw: string | undefined): void {
    if (raw === undefined) return;
    const m = VML_ID.exec(raw);
    if (m === null) return;
    const n = Number(m[1]);
    this.vml.add(n);
    if (n > this.vmlMax) this.vmlMax = n;
  }
}

/** One scanned part. */
interface Scan {
  readonly part: Part;
  /** The DOM scanned (undefined when the part does not parse). */
  readonly doc: XDocument | undefined;
  readonly version: number;
  readonly ids: IdSets;
}

/** Collects every id of `root`'s tree into a new IdSets (explicit stack, single pass). */
function scanTree(root: XElement): IdSets {
  const ids = new IdSets();
  const stack: XElement[] = [root];
  while (stack.length > 0) {
    const el = stack.pop()!;
    const kids = el._children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
    const a = el._attrs;
    if (a.length === 0) continue;
    const uri = el.uri;
    if (uri === NS.w) {
      switch (el.local) {
        case 'bookmarkStart':
        case 'bookmarkEnd':
          ids.addNum('bookmark', el.attrNS(NS.w, 'id'));
          break;
        case 'comment':
        case 'commentRangeStart':
        case 'commentRangeEnd':
        case 'commentReference':
          ids.addNum('comment', el.attrNS(NS.w, 'id'));
          break;
        case 'footnote':
        case 'footnoteReference':
          ids.addNum('footnote', el.attrNS(NS.w, 'id'));
          break;
        case 'endnote':
        case 'endnoteReference':
          ids.addNum('endnote', el.attrNS(NS.w, 'id'));
          break;
        case 'id':
          if (el.parent !== null && (el.parent as XElement).is(NS.w, 'sdtPr')) {
            ids.addNum('sdt', el.attrNS(NS.w, 'val'));
          }
          break;
        default:
          if (REVISION_ELEMENTS.has(el.local)) ids.addNum('revision', el.attrNS(NS.w, 'id'));
      }
    } else if (uri === NS.wp) {
      if (el.local === 'docPr') ids.addNum('docPr', el.attrNS('', 'id'));
    } else if (uri === NS.pic) {
      if (el.local === 'cNvPr') ids.addNum('docPr', el.attrNS('', 'id'));
    } else if (uri === NS.v || uri === NS.o) {
      ids.addVml(el.attrNS('', 'id'));
    }
    // Attributes that can sit on any element.
    for (let i = 0; i < a.length; i += 3) {
      const au = a[i + 1];
      if (au === NS.w14) {
        const name = a[i]!;
        if (name.endsWith('paraId')) ids.addHex('paraId', a[i + 2]);
        else if (name.endsWith('textId')) ids.addHex('textId', a[i + 2]);
      } else if (au === NS.w15) {
        // w15:commentEx/@w15:paraId (+ paraIdParent) name comment paragraphs.
        if (a[i]!.includes('paraId')) ids.addHex('paraId', a[i + 2]);
      } else if (au === NS.w16cid || au === NS.w16cex) {
        const name = a[i]!;
        if (name.endsWith('durableId')) ids.addHex('durableId', a[i + 2]);
        else if (name.endsWith('paraId')) ids.addHex('paraId', a[i + 2]);
      } else if (au === NS.o && a[i]!.endsWith('spid')) {
        ids.addVml(a[i + 2]);
      }
    }
  }
  return ids;
}

/** Deterministic 32-bit PRNG (mulberry32). */
function prng(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = s;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return (t ^ (t >>> 14)) >>> 0;
  };
}

const registries = new WeakMap<OpcPackage, DocumentIds>();

/** The document-wide id registry of a document (see the file header). */
export class DocumentIds {
  readonly #doc: WordDocument;
  /** Scans by part, or undefined before the first scan / after invalidate(). */
  #scans: Scan[] | undefined;
  /** The main part's rels DOM and version at the last scan. */
  #relsDoc: XDocument | undefined;
  #relsVersion = -1;
  /** Ids handed out or reserved explicitly (survive re-scans). */
  readonly #reserved = new IdSets();
  #held = 0;
  readonly #rand = prng(0x7475_7262); // 'turb'

  private constructor(doc: WordDocument) {
    this.#doc = doc;
  }

  /** The registry of `doc` (one per document package, cached). */
  static of(doc: WordDocument): DocumentIds {
    let r = registries.get(doc.package);
    if (r === undefined) {
      r = new DocumentIds(doc);
      registries.set(doc.package, r);
    }
    return r;
  }

  /** A fresh unused numeric id of `kind` (and reserves it). */
  next(kind: NumericIdKind): number {
    const scans = this.#fresh();
    let max = this.#reserved.max[kind];
    for (const s of scans) if (s.ids.max[kind] > max) max = s.ids.max[kind];
    let n = Math.max(MIN[kind], max + 1);
    if (n > MAX_INT) {
      n = MIN[kind];
      while (n <= MAX_INT && this.#hasNum(scans, kind, n)) n++;
      if (n > MAX_INT) throw new RangeError(`No free ${kind} id left`);
    }
    this.#reserved.addNum(kind, n);
    return n;
  }

  /** A fresh unused 8-hex id of `kind` (and reserves it). */
  nextHex(kind: HexIdKind): string {
    const scans = this.#fresh();
    for (;;) {
      const v = this.#rand() & MAX_INT;
      if (v === 0) continue;
      const h = v.toString(16).toUpperCase().padStart(8, '0');
      if (h === '77777777' || this.#hasHex(scans, kind, h)) continue;
      this.#reserved.addHex(kind, h);
      return h;
    }
  }

  /** A fresh VML shape id (`_x0000_s1025`, ...). */
  nextVmlShapeId(): string {
    const scans = this.#fresh();
    let max = Math.max(VML_BASE, this.#reserved.vmlMax);
    for (const s of scans) if (s.ids.vmlMax > max) max = s.ids.vmlMax;
    const id = `_x0000_s${max + 1}`;
    this.#reserved.addVml(id);
    return id;
  }

  /** True when the id is already used. */
  has(kind: NumericIdKind | HexIdKind, id: number | string): boolean {
    const scans = this.#fresh();
    if (isHexKind(kind)) {
      const h = normHex(String(id));
      return h !== undefined && this.#hasHex(scans, kind, h);
    }
    const n = toInt(id);
    return n !== undefined && this.#hasNum(scans, kind, n);
  }

  /** Marks an id as used. */
  reserve(kind: NumericIdKind | HexIdKind, id: number | string): void {
    if (isHexKind(kind)) this.#reserved.addHex(kind, String(id));
    else this.#reserved.addNum(kind, id);
  }

  /** True when VML shape id `id` (or another `_x0000_?<n>` id with its number) is used. */
  hasVmlShapeId(id: string): boolean {
    const m = VML_ID.exec(id);
    if (m === null) return false;
    const n = Number(m[1]);
    if (this.#reserved.vml.has(n)) return true;
    return this.#fresh().some((s) => s.ids.vml.has(n));
  }

  /** Marks a VML shape id as used. */
  reserveVmlShapeId(id: string): void {
    this.#reserved.addVml(id);
  }

  /** Forgets cached scans (call after editing ids outside the registry). Reserved ids stay reserved. */
  invalidate(): void {
    this.#scans = undefined;
    this.#relsDoc = undefined;
    this.#relsVersion = -1;
  }

  /**
   * Suspends re-scanning until the returned function is called: for a batch of edits that only
   * write ids obtained from (or reserved in) this registry. Scans happen at most once while held.
   * Holds nest.
   */
  hold(): () => void {
    this.#fresh();
    this.#held++;
    let released = false;
    return () => {
      if (released) return;
      released = true;
      this.#held--;
    };
  }

  #hasNum(scans: readonly Scan[], kind: NumericIdKind, n: number): boolean {
    if (this.#reserved.num[kind].has(n)) return true;
    for (const s of scans) if (s.ids.num[kind].has(n)) return true;
    return false;
  }

  #hasHex(scans: readonly Scan[], kind: HexIdKind, h: string): boolean {
    if (this.#reserved.hex[kind].has(h)) return true;
    for (const s of scans) if (s.ids.hex[kind].has(h)) return true;
    return false;
  }

  /** Up-to-date scans (re-scanning changed parts; nothing while held, after the first scan). */
  #fresh(): Scan[] {
    const cached = this.#scans;
    if (cached !== undefined && this.#held > 0) return cached;
    const main = this.#doc.mainPart;
    const relsPart = this.#doc.package.part(relsPathFor(main.name));
    const relsDoc = relsPart === undefined ? undefined : safeXml(relsPart);
    const relsVersion = relsDoc?.version ?? -1;
    if (cached !== undefined && relsDoc === this.#relsDoc && relsVersion === this.#relsVersion) {
      let changed = false;
      const next = cached.map((s) => {
        const doc = safeXml(s.part);
        if (doc === s.doc && (doc?.version ?? -1) === s.version) return s;
        changed = true;
        return scanPart(s.part, doc);
      });
      if (changed) this.#scans = next;
      return this.#scans!;
    }
    const byPart = new Map<Part, Scan>();
    for (const s of cached ?? []) byPart.set(s.part, s);
    const out: Scan[] = [];
    for (const part of this.#storyParts()) {
      const doc = safeXml(part);
      if (doc === undefined) continue;
      const old = byPart.get(part);
      out.push(
        old !== undefined && old.doc === doc && old.version === doc.version
          ? old
          : scanPart(part, doc),
      );
    }
    this.#scans = out;
    this.#relsDoc = relsDoc;
    this.#relsVersion = relsVersion;
    return out;
  }

  /** The main part and every id-bearing part related from it, each once. */
  #storyParts(): Part[] {
    const main = this.#doc.mainPart;
    const out: Part[] = [main];
    const seen = new Set<Part>(out);
    for (const rel of main.rels.all()) {
      if (rel.external || !STORY_RELS.some((k) => relTypeIs(rel.type, k))) continue;
      const part = main.related(rel.id);
      if (part === undefined || seen.has(part)) continue;
      seen.add(part);
      out.push(part);
    }
    return out;
  }
}

function isHexKind(kind: NumericIdKind | HexIdKind): kind is HexIdKind {
  return (HEX as readonly string[]).includes(kind);
}

/** The part's DOM, or undefined when it does not parse (validatePackage reports that). */
function safeXml(part: Part): XDocument | undefined {
  try {
    return part.xml();
  } catch {
    return undefined;
  }
}

function scanPart(part: Part, doc: XDocument | undefined): Scan {
  let root: XElement | undefined;
  for (const k of doc?.children ?? []) {
    if (k.nodeType === 1) {
      root = k;
      break;
    }
  }
  return {
    part,
    doc,
    version: doc?.version ?? -1,
    ids: root === undefined ? new IdSets() : scanTree(root),
  };
}
