// View over [Content_Types].xml: `<Default Extension ContentType>` maps an extension to a content
// type, `<Override PartName ContentType>` names one part explicitly and wins over the default.
//
// The view keeps no state of its own: every call reads the DOM, so edits made through `xml`
// directly are seen, and every edit made here goes through DOM methods (the part becomes dirty
// and is re-serialized on save). Part names compare case-insensitively and percent-decoded.

import type { XDocument, XElement } from '../xml/dom.js';
import { normalizePartName, partExtension, partKey } from './uri.js';
import { childElement, childrenNamed } from './xml-util.js';

const normalizeExt = (ext: string): string =>
  (ext.startsWith('.') ? ext.slice(1) : ext).toLowerCase();

export class ContentTypes {
  readonly #doc: XDocument;

  /** @internal */
  constructor(doc: XDocument) {
    this.#doc = doc;
  }

  /** The underlying DOM. */
  get xml(): XDocument {
    return this.#doc;
  }

  /** Content type of a part name (override first, case-insensitive; else default by extension). */
  get(partName: string): string | undefined {
    const override = this.getOverride(partName);
    if (override !== undefined) return override;
    const ext = partExtension(normalizePartName(partName));
    return ext === undefined ? undefined : this.getDefault(ext);
  }

  /** Content type of `<Default Extension="ext">` (case-insensitive). */
  getDefault(ext: string): string | undefined {
    return this.#findDefault(normalizeExt(ext))?.attr('ContentType');
  }

  /** Sets the default content type of an extension (a new `<Default>` goes after the last one). */
  setDefault(ext: string, contentType: string): void {
    const bare = ext.startsWith('.') ? ext.slice(1) : ext;
    const existing = this.#findDefault(bare.toLowerCase());
    if (existing !== undefined) {
      existing.setAttr('ContentType', contentType);
      return;
    }
    const root = this.#doc.root;
    const el = childElement(root, 'Default', [
      ['Extension', bare],
      ['ContentType', contentType],
    ]);
    const last = childrenNamed(root, 'Default').at(-1);
    if (last !== undefined) root.insertAfter(el, last);
    else root.prepend(el);
  }

  /** Content type of `<Override PartName="/partName">` (case-insensitive). */
  getOverride(partName: string): string | undefined {
    return this.#findOverrides(partName)[0]?.attr('ContentType');
  }

  /** Sets a part's override (updated in place, or a new `<Override>` appended at the end). */
  setOverride(partName: string, contentType: string): void {
    const existing = this.#findOverrides(partName)[0];
    if (existing !== undefined) {
      existing.setAttr('ContentType', contentType);
      return;
    }
    const root = this.#doc.root;
    root.append(
      childElement(root, 'Override', [
        ['PartName', `/${normalizePartName(partName)}`],
        ['ContentType', contentType],
      ]),
    );
  }

  /** Removes a part's override(s); returns whether there was one. */
  removeOverride(partName: string): boolean {
    const found = this.#findOverrides(partName);
    for (const el of found) el.remove();
    return found.length > 0;
  }

  /** Extension (lower-case) → content type. */
  defaults(): Map<string, string> {
    const out = new Map<string, string>();
    for (const el of childrenNamed(this.#doc.root, 'Default')) {
      const ext = el.attr('Extension');
      const ct = el.attr('ContentType');
      if (ext === undefined || ct === undefined) continue;
      const key = normalizeExt(ext);
      if (!out.has(key)) out.set(key, ct);
    }
    return out;
  }

  /** Part name (no leading slash) → content type. */
  overrides(): Map<string, string> {
    const out = new Map<string, string>();
    for (const el of childrenNamed(this.#doc.root, 'Override')) {
      const name = el.attr('PartName');
      const ct = el.attr('ContentType');
      if (name === undefined || ct === undefined) continue;
      const key = normalizePartName(name);
      if (!out.has(key)) out.set(key, ct);
    }
    return out;
  }

  #findDefault(lowerExt: string): XElement | undefined {
    for (const el of childrenNamed(this.#doc.root, 'Default')) {
      const ext = el.attr('Extension');
      if (ext !== undefined && normalizeExt(ext) === lowerExt) return el;
    }
    return undefined;
  }

  #findOverrides(partName: string): XElement[] {
    const key = partKey(partName);
    return childrenNamed(this.#doc.root, 'Override').filter((el) => {
      const name = el.attr('PartName');
      return name !== undefined && partKey(name) === key;
    });
  }
}
