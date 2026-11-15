// Content controls (`w:sdt`): discovery, typed reading and setValue.
//
// A ContentControl is a live view over one `w:sdt` element (like every turboword view it holds
// only the element), so it never goes stale and anything it does not model survives a save.
//
// - Discovery walks every story (body, headers/footers, notes, comments, and the text boxes in
//   them) in document order, outer controls before nested ones. Of an `mc:AlternateContent`
//   only the Choice is walked: Word reads the Choice, so that is the copy worth editing.
// - The type comes from the `w:sdtPr` child that names it (`w:text`, `w14:checkbox`,
//   `w15:repeatingSection`, ...); a control with none is rich text (ECMA-376 §17.5.2).
// - Values written to a control with a `w:dataBinding` are also written to the bound node of
//   its custom XML part when that node exists: Word refreshes bound controls from the XML when
//   it opens a file, so a value written only to the document would be lost.

import { UsageError } from '../core/errors.js';
import type { Part } from '../opc/part.js';
import type { WordDocument } from '../wml/document.js';
import { addImagePart, inlineDrawing } from '../wml/media.js';
import { ensureChild, ensureFirstChild, RPR_ORDER, setWAttr, wElement } from '../wml/order.js';
import type { Story } from '../wml/story.js';
import { blockParagraphs, paragraphText } from '../wml/text.js';
import { XElement } from '../xml/dom.js';
import { conventionalPrefix, NS } from '../xml/ns.js';
import {
  ensureContent,
  firstRun,
  type Level,
  levelOf,
  makeRun,
  replaceContent,
  unwrap,
} from './content.js';
import { customXmlItems, findCustomXml } from './customxml.js';
import { formatWordDate, isoDateTime, parseIsoDate } from './date-format.js';
import { type IdSource, idSource } from './ids.js';
import type { ContentControl, ContentControlType, ContentControlValue } from './types.js';
import { evaluateXPath, writeHit } from './xpath.js';

/** State shared by the controls of one operation. */
export class FormsContext {
  readonly doc: WordDocument;
  /** Write values through to the bound custom XML node (off while binding from XML). */
  writeThrough = true;
  #placeholder: Set<string> | undefined;
  #ids: IdSource | undefined;

  constructor(doc: WordDocument) {
    this.doc = doc;
  }

  /** Style ids of the "Placeholder Text" character style (localized ids included). */
  placeholderStyles(): Set<string> {
    if (this.#placeholder === undefined) {
      const out = new Set(['PlaceholderText']);
      const styles = this.doc.mainPart.relatedByType('styles')[0];
      if (styles !== undefined) {
        for (const s of styles.xml().root.elements(NS.w, 'style')) {
          const name = s.element(NS.w, 'name')?.attrNS(NS.w, 'val');
          const id = s.attrNS(NS.w, 'styleId');
          if (id !== undefined && name?.toLowerCase() === 'placeholder text') out.add(id);
        }
      }
      this.#placeholder = out;
    }
    return this.#placeholder;
  }

  /** Fresh-id source (created on first use). */
  ids(): IdSource {
    this.#ids ??= idSource(this.doc);
    return this.#ids;
  }
}

/** `w:sdtPr` child that names the type → type. */
const TYPE_OF: ReadonlyMap<string, ContentControlType> = /* @__PURE__ */ new Map([
  [`${NS.w} text`, 'plainText'],
  [`${NS.w} richText`, 'richText'],
  [`${NS.w} date`, 'date'],
  [`${NS.w} dropDownList`, 'dropDownList'],
  [`${NS.w} comboBox`, 'comboBox'],
  [`${NS.w} picture`, 'picture'],
  [`${NS.w} group`, 'group'],
  [`${NS.w} docPartObj`, 'buildingBlockGallery'],
  [`${NS.w} docPartList`, 'buildingBlockGallery'],
  [`${NS.w} citation`, 'citation'],
  [`${NS.w} bibliography`, 'bibliography'],
  [`${NS.w} equation`, 'equation'],
  [`${NS.w14} checkbox`, 'checkbox'],
  [`${NS.w14} entityPicker`, 'unknown'],
  [`${NS.w15} repeatingSection`, 'repeatingSection'],
  [`${NS.w15} repeatingSectionItem`, 'repeatingSectionItem'],
]);

/** ST_OnOff value (absent attribute = `dflt`). */
function onOff(v: string | undefined, dflt: boolean): boolean {
  return v === undefined ? dflt : !(v === '0' || v === 'false' || v === 'off');
}

/** `w:<local>/@w:val` of `parent`. */
function val(parent: XElement | undefined, local: string): string | undefined {
  return parent?.element(NS.w, local)?.attrNS(NS.w, 'val');
}

/** New element `{uri}local` named with the prefix in scope at `scope` (else the conventional one). */
export function nsElement(scope: XElement, uri: string, local: string): XElement {
  const p = scope.lookupPrefix(uri) ?? conventionalPrefix(uri) ?? '';
  return new XElement(p === '' ? local : `${p}:${local}`, uri);
}

/** Base64 of bytes (web-standard `btoa`, chunked). */
function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(s);
}

/** Removes the relationships in `ids` that no `r:` attribute of `part` references any more. */
function dropUnusedRels(part: Part, ids: Iterable<string>): void {
  const left = new Set(ids);
  if (left.size === 0) return;
  for (const el of part.xml().root.descendants()) {
    const a = el._attrs;
    for (let i = 0; i < a.length; i += 3) if (a[i + 1] === NS.r) left.delete(a[i + 2]!);
    if (left.size === 0) return;
  }
  for (const id of left) part.rels.remove(id);
}

/** A content control (`w:sdt`); see {@link ContentControl}. */
export class Sdt implements ContentControl {
  readonly element: XElement;
  readonly story: Story;
  readonly #ctx: FormsContext;

  constructor(element: XElement, story: Story, ctx: FormsContext) {
    this.element = element;
    this.story = story;
    this.#ctx = ctx;
  }

  /** `w:sdtPr`. */
  get pr(): XElement | undefined {
    return this.element.element(NS.w, 'sdtPr');
  }

  /** `w:sdtContent`. */
  get content(): XElement | undefined {
    return this.element.element(NS.w, 'sdtContent');
  }

  get type(): ContentControlType {
    for (const k of this.pr?.elements() ?? []) {
      const t = TYPE_OF.get(`${k.uri} ${k.local}`);
      if (t !== undefined) return t;
    }
    return 'richText';
  }

  get tag(): string | undefined {
    return val(this.pr, 'tag');
  }

  get alias(): string | undefined {
    return val(this.pr, 'alias');
  }

  get id(): number | undefined {
    const v = val(this.pr, 'id');
    const n = v === undefined ? Number.NaN : Number(v);
    return Number.isInteger(n) ? n : undefined;
  }

  get level(): Level {
    return levelOf(this.element);
  }

  get text(): string {
    const c = this.content;
    if (c === undefined) return '';
    if (this.level === 'inline') return paragraphText(c);
    return blockParagraphs(c).map(paragraphText).join('\n');
  }

  get value(): ContentControlValue {
    const pr = this.pr;
    switch (this.type) {
      case 'checkbox':
        return onOff(
          pr?.element(NS.w14, 'checkbox')?.element(NS.w14, 'checked')?.attrNS(NS.w14, 'val'),
          false,
        );
      case 'date': {
        const full = pr?.element(NS.w, 'date')?.attrNS(NS.w, 'fullDate');
        return (full === undefined ? undefined : parseIsoDate(full)) ?? null;
      }
      case 'dropDownList':
      case 'comboBox': {
        if (this.placeholderShown) return null;
        const text = this.text;
        const last = this.#list()?.attrNS(NS.w, 'lastValue');
        const items = this.items;
        const hit =
          items.find((i) => i.value === last && i.display === text) ??
          items.find((i) => i.display === text);
        return hit?.value ?? text;
      }
      case 'picture': {
        const blip = this.content?.descendants(NS.a, 'blip')[0];
        const id = blip?.attrNS(NS.r, 'embed');
        const part = id === undefined ? undefined : this.story.part.related(id);
        return part?.bytes() ?? null;
      }
      default:
        return this.placeholderShown ? null : this.text;
    }
  }

  get binding(): { xpath: string; storeItemId?: string; prefixMappings?: string } | undefined {
    const b = this.#bindingElement();
    const xpath = b?.attrNS(NS.w, 'xpath');
    if (b === undefined || xpath === undefined) return undefined;
    const out: { xpath: string; storeItemId?: string; prefixMappings?: string } = { xpath };
    const store = b.attrNS(NS.w, 'storeItemID');
    const maps = b.attrNS(NS.w, 'prefixMappings');
    if (store !== undefined && store !== '') out.storeItemId = store;
    if (maps !== undefined && maps !== '') out.prefixMappings = maps;
    return out;
  }

  get placeholderShown(): boolean {
    const el = this.pr?.element(NS.w, 'showingPlcHdr');
    return el !== undefined && onOff(el.attrNS(NS.w, 'val'), true);
  }

  get locked(): { content: boolean; control: boolean } {
    const v = val(this.pr, 'lock');
    return {
      content: v === 'contentLocked' || v === 'sdtContentLocked',
      control: v === 'sdtLocked' || v === 'sdtContentLocked',
    };
  }

  get items(): readonly { display: string; value: string }[] {
    return (this.#list()?.elements(NS.w, 'listItem') ?? []).map((li) => {
      const v = li.attrNS(NS.w, 'value');
      const d = li.attrNS(NS.w, 'displayText');
      return { display: d ?? v ?? '', value: v ?? d ?? '' };
    });
  }

  setValue(v: ContentControlValue): void {
    if (this.locked.content) {
      throw new UsageError(`Content control ${this.#label()} is locked (contentLocked)`);
    }
    this.apply(v);
  }

  /** Sets the value without the lock check (data binding updates locked controls too). */
  apply(v: ContentControlValue): void {
    const stored = this.#write(v);
    if (this.#ctx.writeThrough && stored !== undefined) this.#writeBound(stored);
    this.pr?.element(NS.w, 'showingPlcHdr')?.remove();
    const temp = this.pr?.element(NS.w, 'temporary');
    if (temp !== undefined && onOff(temp.attrNS(NS.w, 'val'), true)) unwrap(this.element);
  }

  /**
   * Shows `text` as the content without interpreting it (a bound date that does not parse, a
   * bound dropdown value with no item); the stored date or selection is cleared.
   */
  showText(text: string): void {
    this.pr?.element(NS.w, 'date')?.removeAttrNS(NS.w, 'fullDate');
    this.#list()?.removeAttrNS(NS.w, 'lastValue');
    this.#setText(text, false);
    this.pr?.element(NS.w, 'showingPlcHdr')?.remove();
  }

  /** Writes `v` into the content; returns its form for the bound XML node. */
  #write(v: ContentControlValue): string | undefined {
    const type = this.type;
    switch (type) {
      case 'checkbox': {
        const b = toBool(v, this.#label());
        this.#setCheckbox(b);
        return b ? 'true' : 'false';
      }
      case 'date':
        return this.#setDate(v);
      case 'dropDownList':
      case 'comboBox':
        return this.#setChoice(toText(v, this.#label()));
      case 'picture':
        if (!(v instanceof Uint8Array)) {
          throw new UsageError(`Picture content control ${this.#label()} takes image bytes`);
        }
        this.#setPicture(v);
        return toBase64(v);
      case 'richText':
      case 'plainText':
      case 'unknown': {
        const s = toText(v, this.#label());
        this.#setText(s, type !== 'plainText');
        return s;
      }
      default:
        throw new UsageError(
          `Content control ${this.#label()} is a ${type} control: set its nested controls instead`,
        );
    }
  }

  #label(): string {
    const name = this.tag ?? this.alias;
    return name === undefined ? `#${this.id ?? '?'}` : JSON.stringify(name);
  }

  #list(): XElement | undefined {
    const pr = this.pr;
    return pr?.element(NS.w, 'dropDownList') ?? pr?.element(NS.w, 'comboBox');
  }

  #bindingElement(): XElement | undefined {
    const pr = this.pr;
    return pr?.element(NS.w, 'dataBinding') ?? pr?.element(NS.w15, 'dataBinding');
  }

  /** Run properties for new text: the first run's (or the control's while the placeholder shows). */
  #runProps(): XElement | undefined {
    const own = this.pr?.element(NS.w, 'rPr');
    const run = firstRun(this.content)?.element(NS.w, 'rPr');
    const src = this.placeholderShown ? (own ?? run) : (run ?? own);
    if (src === undefined) return undefined;
    const r = src.clone();
    const style = r.element(NS.w, 'rStyle');
    if (style !== undefined && this.#ctx.placeholderStyles().has(style.attrNS(NS.w, 'val') ?? '')) {
      style.remove();
    }
    return r.firstElement() !== undefined ? r : undefined;
  }

  #setText(text: string, rich: boolean): void {
    const level = this.level;
    const rPr = this.#runProps();
    const lines = rich && level !== 'inline' ? text.split(/\r\n|\r|\n/) : [text];
    replaceContent(
      this.element,
      level,
      lines.map((l) => (l === '' ? [] : [makeRun(this.element, rPr, l)])),
    );
  }

  #setCheckbox(checked: boolean): void {
    const cb = this.pr!.element(NS.w14, 'checkbox')!;
    let ce = cb.element(NS.w14, 'checked');
    if (ce === undefined) {
      ce = nsElement(cb, NS.w14, 'checked');
      cb.prepend(ce);
    }
    ce.setAttrNS(NS.w14, 'val', checked ? '1' : '0');
    const state = cb.element(NS.w14, checked ? 'checkedState' : 'uncheckedState');
    const hex = state?.attrNS(NS.w14, 'val') ?? (checked ? '2612' : '2610');
    const font = state?.attrNS(NS.w14, 'font') ?? 'MS Gothic';
    const cp = Number.parseInt(hex, 16);
    const glyph = String.fromCodePoint(cp > 0 && cp <= 0x10ffff ? cp : checked ? 0x2612 : 0x2610);
    let run = firstRun(this.content);
    if (run === undefined) {
      run = makeRun(this.element, this.#runProps(), '');
      replaceContent(this.element, this.level, [[run]]);
    }
    const sym = run.element(NS.w, 'sym');
    if (sym !== undefined) {
      setWAttr(sym, 'font', font);
      setWAttr(sym, 'char', hex);
    } else {
      for (const k of run.children.slice()) {
        if (!(k.nodeType === 1 && k.is(NS.w, 'rPr'))) k.remove();
      }
      const t = wElement(run, 't');
      t.text = glyph;
      run.append(t);
    }
    const fonts = ensureChild(ensureFirstChild(run, 'rPr'), 'rFonts', RPR_ORDER);
    for (const a of ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme']) {
      fonts.removeAttrNS(NS.w, a);
    }
    for (const a of ['ascii', 'hAnsi', 'eastAsia', 'cs']) setWAttr(fonts, a, font);
  }

  #setDate(v: ContentControlValue): string {
    const date = this.pr!.element(NS.w, 'date')!;
    if (v === null || v === '') {
      date.removeAttrNS(NS.w, 'fullDate');
      this.#setText('', false);
      return '';
    }
    let d: Date | undefined;
    if (v instanceof Date) d = Number.isNaN(v.getTime()) ? undefined : v;
    else if (typeof v === 'string') d = parseIsoDate(v);
    if (d === undefined) {
      throw new UsageError(
        `Date content control ${this.#label()} takes a Date or an ISO date string (yyyy-mm-dd)`,
      );
    }
    const iso = isoDateTime(d);
    setWAttr(date, 'fullDate', iso);
    // No w:lid: Word formats in the language of the control's text.
    const lang =
      val(date, 'lid') ||
      val(this.pr?.element(NS.w, 'rPr'), 'lang') ||
      val(firstRun(this.content)?.element(NS.w, 'rPr'), 'lang') ||
      'en-US';
    const text = formatWordDate(d, val(date, 'dateFormat') || 'M/d/yyyy', lang);
    this.#setText(text, false);
    const mode = val(date, 'storeMappedDataAs');
    return mode === 'date' ? iso.slice(0, 10) : mode === 'dateTime' ? iso.slice(0, 19) : text;
  }

  #setChoice(s: string): string {
    const items = this.items;
    const hit = items.find((i) => i.value === s) ?? items.find((i) => i.display === s);
    if (hit === undefined && this.type === 'dropDownList' && s !== '') {
      const list = items.map((i) => JSON.stringify(i.value)).join(', ');
      throw new UsageError(
        `${JSON.stringify(s)} is not an item of dropdown ${this.#label()} (items: ${list})`,
      );
    }
    const list = this.#list()!;
    if (hit !== undefined) setWAttr(list, 'lastValue', hit.value);
    else list.removeAttrNS(NS.w, 'lastValue');
    this.#setText(hit?.display ?? s, false);
    return hit?.value ?? s;
  }

  #setPicture(bytes: Uint8Array): void {
    const doc = this.#ctx.doc;
    const part = this.story.part;
    const added = addImagePart(doc, part, bytes);
    const content = ensureContent(this.element);
    const blips = content.descendants(NS.a, 'blip');
    const vml = content.descendants(NS.v, 'imagedata');
    const old = new Set<string>();
    if (blips.length === 0 && vml.length === 0) {
      // No picture yet: an inline drawing at the image's natural size.
      const { width, height, dpiX, dpiY } = added.info;
      const emu = (px: number, dpi: number): number => Math.round((px * 914400) / (dpi || 96));
      const run = wElement(this.element, 'r');
      run.append(inlineDrawing(doc, part, added.relId, emu(width, dpiX), emu(height, dpiY)));
      replaceContent(this.element, this.level, [[run]]);
    }
    for (const b of blips) {
      const id = b.attrNS(NS.r, 'embed');
      if (id !== undefined && id !== added.relId) old.add(id);
      b.setAttrNS(NS.r, 'embed', added.relId);
    }
    for (const im of vml) {
      const id = im.attrNS(NS.r, 'id');
      if (id !== undefined && id !== added.relId) old.add(id);
      im.setAttrNS(NS.r, 'id', added.relId);
    }
    dropUnusedRels(part, old);
  }

  /** Writes `value` to the bound custom XML node, when the store and the node exist. */
  #writeBound(value: string): void {
    const b = this.binding;
    if (b === undefined) return;
    const doc = this.#ctx.doc;
    const stores =
      b.storeItemId === undefined
        ? customXmlItems(doc)
        : [findCustomXml(doc, b.storeItemId)].filter((x) => x !== undefined);
    for (const store of stores) {
      let xml: ReturnType<Part['xml']>;
      try {
        xml = store.part.xml();
      } catch {
        continue;
      }
      const hit = evaluateXPath(xml, b.xpath, b.prefixMappings);
      if (hit !== undefined) {
        writeHit(hit, value);
        return;
      }
    }
  }
}

/** String form of a value for text-like controls. */
function toText(v: ContentControlValue, label: string): string {
  if (typeof v === 'string') return v;
  if (v === null) return '';
  if (typeof v === 'boolean') return String(v);
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v.toISOString().slice(0, 10);
  throw new UsageError(`Content control ${label} takes text, not ${describe(v)}`);
}

function toBool(v: ContentControlValue, label: string): boolean {
  if (typeof v === 'boolean') return v;
  if (v === null) return false;
  if (typeof v === 'string') {
    if (/^\s*(1|true|on|yes|x)\s*$/i.test(v)) return true;
    if (/^\s*(0|false|off|no|)\s*$/i.test(v)) return false;
  }
  throw new UsageError(`Checkbox content control ${label} takes a boolean, not ${describe(v)}`);
}

function describe(v: ContentControlValue): string {
  if (v instanceof Uint8Array) return 'bytes';
  if (v instanceof Date) return 'a Date';
  return JSON.stringify(v);
}

/** Every `w:sdt` under `root` (itself included when `self`), document order, Choice branches only. */
export function controlsUnder(
  ctx: FormsContext,
  story: Story,
  root: XElement,
  self: boolean,
): Sdt[] {
  const out: Sdt[] = [];
  const stack: XElement[] = [root];
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el.uri === NS.w) {
      if (el.local === 'sdt') {
        if (self || el !== root) out.push(new Sdt(el, story, ctx));
      } else if (el.local === 'sdtPr' || el.local === 'rPr' || el.local === 'pPr') {
        continue;
      }
    } else if (el.uri === NS.mc && el.local === 'Fallback') {
      continue;
    }
    const kids = el.children;
    for (let i = kids.length - 1; i >= 0; i--) {
      const k = kids[i]!;
      if (k.nodeType === 1) stack.push(k);
    }
  }
  return out;
}

/** Every content control of the document (all stories, document order). */
export function allControls(ctx: FormsContext): Sdt[] {
  const out: Sdt[] = [];
  for (const story of ctx.doc.stories())
    out.push(...controlsUnder(ctx, story, story.element, false));
  return out;
}
