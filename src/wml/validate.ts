// Package invariants that make Word refuse a file or offer to "repair" it.
//
// This is not a schema validator. It checks the structural rules that editing and composing
// break in practice: content types, relationships, relationship ids used by markup, the main
// document, drawing ids, section-properties placement, table cells ending with a paragraph,
// well-formed XML and markup-compatibility declarations.
//
// Issues come out in a fixed order: checks run in the order of the codes below and parts are
// visited sorted by name (plain code-unit order, never locale-dependent), so the same package
// always yields the same list.
//
// Codes (all errors unless noted):
//   XML_MALFORMED        an XML part does not parse
//   CT_MISSING           no [Content_Types].xml
//   NO_CONTENT_TYPE      a part has neither a Default nor an Override content type
//   DANGLING_OVERRIDE    (warning) an Override names a part that does not exist
//   DUPLICATE_REL_ID     two relationships with the same Id in one rels part
//   REL_TARGET_MISSING   an internal relationship targets a missing part
//   NO_MAIN_DOCUMENT     no officeDocument relationship, or its target is missing
//   MAIN_CONTENT_TYPE    the main part's content type is not a WordprocessingML main type
//   REL_ID_UNRESOLVED    an r:* attribute (r:id, r:embed, r:dm, ...) is not an Id in the part's rels
//   UNDECLARED_IGNORABLE mc:Ignorable lists a prefix not declared in scope
//   SECTPR_PLACEMENT     body sectPr not last in w:body, or a sectPr outside w:body / w:pPr
//   CELL_NO_PARAGRAPH    a w:tc whose last block-level element is not a w:p
//   DUPLICATE_DOCPR_ID   wp:docPr/@id repeated across the main document and its stories

import type { ContentTypes } from '../opc/content-types.js';
import { OpcPackage } from '../opc/package.js';
import type { Part } from '../opc/part.js';
import { CT, relTypeIs } from '../opc/rel-types.js';
import { partDir, partKey, relsPathFor, resolveTarget, sourceOfRels } from '../opc/uri.js';
import { ELEMENT_NODE, type XDocument, type XElement } from '../xml/dom.js';
import { hasLocal, NS } from '../xml/ns.js';

/** One problem found by {@link validatePackage}. */
export interface ValidationIssue {
  /** 'error': Word refuses or repairs the file; 'warning': harmless but untidy. */
  severity: 'error' | 'warning';
  /** Stable machine-readable code, e.g. `'REL_TARGET_MISSING'`. */
  code: string;
  message: string;
  /** Name of the part the issue is about, when there is one. */
  part?: string;
}

const CT_PART = '[Content_Types].xml';
const MAIN_TYPES: ReadonlySet<string> = new Set([
  CT.document,
  CT.documentMacro,
  CT.template,
  CT.templateMacro,
]);
/** Relationship kinds (from the main part) whose targets are stories checked like the body. */
const STORY_KINDS = ['header', 'footer', 'footnotes', 'endnotes', 'comments'];
/** WordprocessingML elements that are block-level content of a cell. */
const BLOCKS: ReadonlySet<string> = new Set(['p', 'tbl', 'sdt', 'customXml', 'altChunk']);

interface Rel {
  readonly id: string;
  readonly type: string;
  readonly external: boolean;
  /** Resolved target part name (internal relationships only). */
  readonly partName: string | undefined;
}

/**
 * Checks package invariants (see the codes in this file's header) and returns the issues found,
 * in a deterministic order. Parses every XML part (media and other binary parts are skipped).
 */
export function validatePackage(
  target: OpcPackage | { readonly package: OpcPackage },
): ValidationIssue[] {
  const pkg = target instanceof OpcPackage ? target : target.package;
  return new Validator(pkg).run();
}

const key: (name: string) => string = partKey;
const XML_NAME = /\.(?:xml|rels)$/i;

/** Part.isXml, falling back to the name when content types are unreadable. */
function isXml(part: Part): boolean {
  try {
    return part.isXml;
  } catch {
    return XML_NAME.test(part.name);
  }
}
const byName = (a: Part, b: Part): number => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0);

class Validator {
  readonly #pkg: OpcPackage;
  readonly #issues: ValidationIssue[] = [];
  /** Parts sorted by name. */
  readonly #parts: Part[];
  /** Parsed XML parts by lower-case name; null = failed to parse. */
  readonly #docs = new Map<string, XDocument | null>();
  /** Relationships by lower-case source part name ('' = package), from parsed rels parts. */
  readonly #rels = new Map<string, Rel[]>();

  constructor(pkg: OpcPackage) {
    this.#pkg = pkg;
    this.#parts = pkg.parts().filter((p) => !p.name.endsWith('/'));
    this.#parts.sort(byName);
  }

  run(): ValidationIssue[] {
    this.#parse();
    const ct = this.#contentTypes();
    this.#relationships();
    const main = this.#mainDocument(ct);
    this.#markup();
    if (main !== undefined) this.#stories(main);
    return this.#issues;
  }

  #add(severity: ValidationIssue['severity'], code: string, message: string, part?: string): void {
    const issue: ValidationIssue = { severity, code, message };
    if (part !== undefined) issue.part = part;
    this.#issues.push(issue);
  }

  #root(name: string): XElement | undefined {
    const doc = this.#docs.get(key(name));
    if (doc == null) return undefined;
    for (const k of doc.children) if (k.nodeType === ELEMENT_NODE) return k;
    return undefined;
  }

  /** XML_MALFORMED: parse every XML part once. */
  #parse(): void {
    for (const part of this.#parts) {
      if (!isXml(part)) continue;
      try {
        this.#docs.set(key(part.name), part.xml());
      } catch (err) {
        this.#docs.set(key(part.name), null);
        const why = err instanceof Error ? err.message : String(err);
        this.#add(
          'error',
          'XML_MALFORMED',
          `${part.name} is not well-formed XML: ${why}`,
          part.name,
        );
      }
    }
  }

  /** CT_MISSING, NO_CONTENT_TYPE, DANGLING_OVERRIDE. */
  #contentTypes(): ContentTypes | undefined {
    const pkg = this.#pkg;
    if (pkg.part(CT_PART) === undefined) {
      this.#add('error', 'CT_MISSING', `The package has no ${CT_PART}`);
      return undefined;
    }
    if (this.#docs.get(key(CT_PART)) === null) return undefined; // already XML_MALFORMED
    let ct: ContentTypes;
    try {
      ct = pkg.contentTypes;
    } catch {
      return undefined;
    }
    for (const part of this.#parts) {
      if (key(part.name) === key(CT_PART)) continue;
      if (ct.get(part.name) === undefined) {
        this.#add(
          'error',
          'NO_CONTENT_TYPE',
          `${part.name} has no content type (no Default for its extension, no Override)`,
          part.name,
        );
      }
    }
    const overrides = [...ct.overrides().keys()].sort();
    for (const name of overrides) {
      if (pkg.part(name) === undefined) {
        this.#add('warning', 'DANGLING_OVERRIDE', `Override for missing part ${name}`, name);
      }
    }
    return ct;
  }

  /**
   * DUPLICATE_REL_ID, REL_TARGET_MISSING. Reads the rels DOMs directly (to see duplicates), as
   * leniently as the OPC layer: `Relationship` in any namespace, TargetMode case-insensitive.
   */
  #relationships(): void {
    const pkg = this.#pkg;
    for (const part of this.#parts) {
      const source = sourceOfRels(part.name);
      if (source === undefined) continue;
      if (source !== '' && pkg.part(source) === undefined) continue; // orphan rels part
      const root = this.#root(part.name);
      if (root === undefined) continue;
      const base = source === '' ? '' : partDir(source);
      const rels: Rel[] = [];
      const seen = new Set<string>();
      for (const r of root.elements(undefined, 'Relationship')) {
        const id = r.attr('Id') ?? '';
        const type = r.attr('Type') ?? '';
        if (seen.has(id)) {
          this.#add(
            'error',
            'DUPLICATE_REL_ID',
            `Relationship Id "${id}" is used more than once`,
            part.name,
          );
        }
        seen.add(id);
        const external = r.attr('TargetMode')?.toLowerCase() === 'external';
        let partName: string | undefined;
        if (!external) {
          const t = r.attr('Target') ?? '';
          try {
            partName = t === '' ? undefined : resolveTarget(base, t);
          } catch {
            partName = undefined;
          }
          if (partName === undefined || pkg.part(partName) === undefined) {
            this.#add(
              'error',
              'REL_TARGET_MISSING',
              `Relationship ${id} (${type}) targets missing part ${partName ?? `"${t}"`}`,
              part.name,
            );
          }
        }
        rels.push({ id, type, external, partName });
      }
      this.#rels.set(key(source), rels);
    }
  }

  /** NO_MAIN_DOCUMENT, MAIN_CONTENT_TYPE. Returns the main part when it exists. */
  #mainDocument(ct: ContentTypes | undefined): Part | undefined {
    const rel = (this.#rels.get('') ?? []).find(
      (r) => !r.external && relTypeIs(r.type, 'officeDocument'),
    );
    if (rel === undefined) {
      this.#add('error', 'NO_MAIN_DOCUMENT', 'The package has no officeDocument relationship');
      return undefined;
    }
    const main = rel.partName === undefined ? undefined : this.#pkg.part(rel.partName);
    if (main === undefined) {
      this.#add(
        'error',
        'NO_MAIN_DOCUMENT',
        `Main document part ${rel.partName ?? ''} is missing`,
        rel.partName,
      );
      return undefined;
    }
    const type = ct?.get(main.name);
    if (type !== undefined && !MAIN_TYPES.has(type)) {
      this.#add(
        'error',
        'MAIN_CONTENT_TYPE',
        `Main document ${main.name} has content type ${type}, not a WordprocessingML main type`,
        main.name,
      );
    }
    return main;
  }

  /** REL_ID_UNRESOLVED, UNDECLARED_IGNORABLE: one pass over every element of every XML part. */
  #markup(): void {
    for (const part of this.#parts) {
      if (key(part.name) === key(CT_PART) || sourceOfRels(part.name) !== undefined) continue;
      const root = this.#root(part.name);
      if (root === undefined) continue;
      // Ids of the part's rels; a rels part that failed to parse was reported already.
      if (this.#docs.get(key(relsPathFor(part.name))) === null) continue;
      const ids = new Set((this.#rels.get(key(part.name)) ?? []).map((r) => r.id));
      const reported = new Set<string>();
      const elements = root.descendants();
      elements.unshift(root);
      for (const el of elements) {
        const a = el._attrs;
        for (let i = 0; i < a.length; i += 3) {
          const uri = a[i + 1];
          if (uri === NS.r) {
            const v = a[i + 2]!;
            // Word writes empty references (`r:blip=""`, `r:id=""`) routinely; they point nowhere.
            if (v === '' || ids.has(v) || reported.has(v)) continue;
            reported.add(v);
            this.#add(
              'error',
              'REL_ID_UNRESOLVED',
              `${a[i]}="${v}" on <${el.name}> does not match any relationship of ${part.name}`,
              part.name,
            );
          } else if (uri === NS.mc && hasLocal(a[i]!, 'Ignorable')) {
            for (const prefix of a[i + 2]!.split(/\s+/)) {
              if (prefix === '' || el.lookupNamespace(prefix) !== undefined) continue;
              this.#add(
                'error',
                'UNDECLARED_IGNORABLE',
                `mc:Ignorable on <${el.name}> lists undeclared prefix "${prefix}"`,
                part.name,
              );
            }
          }
        }
      }
    }
  }

  /** SECTPR_PLACEMENT (main part), CELL_NO_PARAGRAPH and DUPLICATE_DOCPR_ID (all stories). */
  #stories(main: Part): void {
    const mainRoot = this.#root(main.name);
    if (mainRoot === undefined) return;
    this.#sectPr(main.name, mainRoot);

    const names = new Set<string>();
    const others: string[] = [];
    for (const r of this.#rels.get(key(main.name)) ?? []) {
      if (r.external || r.partName === undefined) continue;
      if (!STORY_KINDS.some((k) => relTypeIs(r.type, k))) continue;
      const part = this.#pkg.part(r.partName);
      if (part === undefined || names.has(key(part.name)) || key(part.name) === key(main.name)) {
        continue;
      }
      names.add(key(part.name));
      others.push(part.name);
    }
    others.sort();
    const stories: [string, XElement][] = [[main.name, mainRoot]];
    for (const name of others) {
      const root = this.#root(name);
      if (root !== undefined) stories.push([name, root]);
    }

    for (const [name, root] of stories) this.#cells(name, root);

    // docPr ids: Word only reads one branch of mc:AlternateContent, and the Fallback commonly
    // repeats the Choice's drawing, so docPr inside mc:Fallback is not counted.
    const uses = new Map<string, string[]>();
    for (const [name, root] of stories) {
      for (const d of root.descendants(NS.wp, 'docPr')) {
        const raw = d.attr('id');
        if (raw === undefined || d.closest(NS.mc, 'Fallback') !== undefined) continue;
        const id = /^\d+$/.test(raw) ? raw.replace(/^0+(?=\d)/, '') : raw;
        const list = uses.get(id);
        if (list === undefined) uses.set(id, [name]);
        else list.push(name);
      }
    }
    for (const [id, list] of uses) {
      if (list.length < 2) continue;
      const where = [...new Set(list)].join(', ');
      this.#add(
        'error',
        'DUPLICATE_DOCPR_ID',
        `wp:docPr id="${id}" is used ${list.length} times (${where})`,
        list[1],
      );
    }
  }

  #sectPr(name: string, root: XElement): void {
    for (const sp of root.descendants(NS.w, 'sectPr')) {
      const parent = sp.parent;
      if (parent === null || parent.nodeType !== ELEMENT_NODE) continue;
      if (parent.is(NS.w, 'body')) {
        if (parent.lastElement() !== sp) {
          this.#add(
            'error',
            'SECTPR_PLACEMENT',
            'The body-level w:sectPr is not the last child of w:body',
            name,
          );
        }
      } else if (!parent.is(NS.w, 'pPr') && !parent.is(NS.w, 'sectPrChange')) {
        this.#add(
          'error',
          'SECTPR_PLACEMENT',
          `w:sectPr inside <${parent.name}> (allowed only in w:body or w:pPr)`,
          name,
        );
      }
    }
  }

  #cells(name: string, root: XElement): void {
    let n = 0;
    for (const tc of root.descendants(NS.w, 'tc')) {
      n++;
      const last = lastBlock(tc);
      if (last?.is(NS.w, 'p')) continue;
      const what = last === undefined ? 'has no block content' : `ends with <${last.name}>`;
      this.#add(
        'error',
        'CELL_NO_PARAGRAPH',
        `Table cell #${n} ${what}; a cell must end with a w:p`,
        name,
      );
    }
  }
}

/**
 * Last block-level element of a cell, looking through block-level content controls and custom
 * XML (whose content is what Word lays out). Undefined when there is none.
 */
function lastBlock(container: XElement): XElement | undefined {
  let c: XElement = container;
  for (;;) {
    let found: XElement | undefined;
    const kids = c.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === ELEMENT_NODE && k.uri === NS.w && BLOCKS.has(k.local)) {
        found = k;
        break;
      }
    }
    if (found === undefined) return undefined;
    if (found.local === 'sdt') {
      const content = found.element(NS.w, 'sdtContent');
      if (content === undefined) return undefined;
      c = content;
    } else if (found.local === 'customXml') {
      c = found;
    } else {
      return found;
    }
  }
}
