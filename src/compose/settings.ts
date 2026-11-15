// mergeSettings: the document-wide parts after an import.
//
// - settings.xml: the target's settings win (compatibility mode, compat options, tab stops,
//   protection, ...): imported content is laid out the target's way, as Word does when you
//   insert a file. Only additive data is merged: document variables (`w:docVars`, missing names
//   added) and revision-save ids (`w:rsids`, extended when the target keeps a list, so the rsid
//   attributes of imported content resolve). `w:evenAndOddHeaders` is decided by the sections
//   step, which needs the section-by-section context.
// - fontTable.xml: the source's fonts the target does not declare are added. Embedded fonts come
//   with their font parts, copied byte for byte together with their `w:fontKey` (the GUID that
//   de-obfuscates them), and `w:embedTrueTypeFonts` is switched on so Word uses them.
// - theme: the target's (default) or, with `theme: 'source'`, the source's theme replaces it.
// - webSettings.xml stays the target's.

import type { Part } from '../opc/part.js';
import { CT, REL } from '../opc/rel-types.js';
import { partKey, relativeTarget } from '../opc/uri.js';
import type { XElement } from '../xml/dom.js';
import { NS } from '../xml/ns.js';
import {
  addMainPart,
  ensureSettingsChild,
  mainRelated,
  newRoot,
  onOff,
  partRoot,
  relationshipTo,
  relTypeFor,
} from './parts-util.js';
import type { ImportContext } from './types.js';

/** Embedded font references of a `w:font`. */
const EMBEDS: ReadonlySet<string> = new Set([
  'embedRegular',
  'embedBold',
  'embedItalic',
  'embedBoldItalic',
]);

/** Merges settings, font table, theme and web settings as the options say. */
export function mergeSettings(ctx: ImportContext): void {
  mergeSettingsPart(ctx);
  const embedded = mergeFonts(ctx);
  if (embedded) enableEmbeddedFonts(ctx);
  if (ctx.options.theme === 'source') takeSourceTheme(ctx);
}

function rootOf(
  ctx: ImportContext,
  which: 'source' | 'target',
  kind: string,
): XElement | undefined {
  const part = mainRelated(which === 'source' ? ctx.source : ctx.target, kind);
  return part === undefined ? undefined : partRoot(part);
}

// ---- settings.xml ---------------------------------------------------------------------------

function mergeSettingsPart(ctx: ImportContext): void {
  const src = rootOf(ctx, 'source', 'settings');
  const tgt = rootOf(ctx, 'target', 'settings');
  if (src === undefined || tgt === undefined) return;

  const srcVars = src.element(NS.w, 'docVars')?.elements(NS.w, 'docVar') ?? [];
  if (srcVars.length > 0) {
    const vars = ensureSettingsChild(tgt, 'docVars');
    const names = new Set(vars.elements(NS.w, 'docVar').map((v) => v.attrNS(NS.w, 'name')));
    for (const v of srcVars) {
      const name = v.attrNS(NS.w, 'name');
      if (name === undefined || names.has(name)) continue;
      names.add(name);
      vars.append(v.clone());
    }
  }

  const tgtRsids = tgt.element(NS.w, 'rsids');
  const srcRsids = src.element(NS.w, 'rsids');
  if (tgtRsids !== undefined && srcRsids !== undefined) {
    const known = new Set(
      tgtRsids.elements(NS.w, 'rsid').map((r) => (r.attrNS(NS.w, 'val') ?? '').toUpperCase()),
    );
    for (const r of srcRsids.elements(NS.w, 'rsid')) {
      const val = (r.attrNS(NS.w, 'val') ?? '').toUpperCase();
      if (val === '' || known.has(val)) continue;
      known.add(val);
      tgtRsids.append(r.clone());
    }
  }
}

// ---- fontTable.xml --------------------------------------------------------------------------

/** Adds the source's missing fonts; returns true when an embedded font was copied. */
function mergeFonts(ctx: ImportContext): boolean {
  const srcPart = mainRelated(ctx.source, 'fontTable');
  const src = srcPart === undefined ? undefined : partRoot(srcPart);
  if (srcPart === undefined || src === undefined) return false;
  let tgtPart = mainRelated(ctx.target, 'fontTable');
  let tgt = tgtPart === undefined ? undefined : partRoot(tgtPart);
  if (tgtPart === undefined || tgt === undefined) {
    tgt = newRoot(ctx.target, NS.w, 'fonts');
    tgtPart = addMainPart(
      ctx,
      'word/fontTable.xml',
      'word/fontTable{n}.xml',
      CT.fontTable,
      REL.fontTable,
      tgt,
    );
  }
  const known = new Set(
    tgt.elements(NS.w, 'font').map((f) => (f.attrNS(NS.w, 'name') ?? '').toLowerCase()),
  );
  let embedded = false;
  for (const font of src.elements(NS.w, 'font')) {
    const name = font.attrNS(NS.w, 'name');
    if (name === undefined || known.has(name.toLowerCase())) continue;
    known.add(name.toLowerCase());
    const copy = font.clone();
    for (const e of copy.elements(NS.w)) {
      if (!EMBEDS.has(e.local)) continue;
      const rid = e.attrNS(NS.r, 'id');
      const rel = rid === undefined ? undefined : srcPart.rels.get(rid);
      if (rid === undefined || rel === undefined) {
        e.remove();
        continue;
      }
      e.setAttrNS(NS.r, 'id', ctx.copier.copyRelationship(srcPart, rid, tgtPart));
      embedded = true;
    }
    tgt.append(copy);
  }
  return embedded;
}

/** Turns on `w:embedTrueTypeFonts` (and the source's subsetting flag) in the target settings. */
function enableEmbeddedFonts(ctx: ImportContext): void {
  const tgt = rootOf(ctx, 'target', 'settings');
  if (tgt === undefined) return;
  const embed = tgt.element(NS.w, 'embedTrueTypeFonts');
  if (!onOff(embed)) ensureSettingsChild(tgt, 'embedTrueTypeFonts').removeAttrNS(NS.w, 'val');
  const src = rootOf(ctx, 'source', 'settings');
  if (
    onOff(src?.element(NS.w, 'saveSubsetFonts')) &&
    tgt.element(NS.w, 'saveSubsetFonts') === undefined
  ) {
    ensureSettingsChild(tgt, 'saveSubsetFonts');
  }
}

// ---- theme ----------------------------------------------------------------------------------

/**
 * Replaces the target's theme with a copy of the source's (copied with its own relationships,
 * e.g. theme images); the target's theme relationship keeps its id and the old part is removed.
 */
function takeSourceTheme(ctx: ImportContext): void {
  const srcTheme = mainRelated(ctx.source, 'theme');
  if (srcTheme === undefined) return;
  const main = ctx.target.mainPart;
  const copy: Part = ctx.copier.copyPart(srcTheme);
  const old = main.rels.byType('theme').filter((r) => !r.external);
  if (old.length === 0) {
    relationshipTo(main, copy, relTypeFor(ctx.target, REL.theme));
    return;
  }
  const copyKey = partKey(copy.name);
  for (const rel of old) {
    if (rel.partName !== undefined && partKey(rel.partName) === copyKey) continue;
    main.rels.remove(rel.id);
    main.rels.add(rel.type, relativeTarget(main.name, copy.name), { id: rel.id });
    const stillUsed =
      rel.partName !== undefined && main.rels.all().some((r) => r.partName === rel.partName);
    if (rel.partName !== undefined && !stillUsed) ctx.target.package.removePart(rel.partName);
  }
}
