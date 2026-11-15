// The font manager of one layout: collects the faces runs need, resolves them, hands them out.
//
// Resolution of a (family, bold, italic) request, first hit wins (plan decision D7):
//   A. the exact style from the "specific" families — the family itself, its font-table altName,
//      the caller's substitution, the metric-compatible substitutes — each looked up in the
//      fonts embedded in the docx, then in the caller's FontResolvers (sync or async);
//   B. a face a resolver returned for the exact request even though its style flags differ;
//   C. lesser styles of the specific families (bold italic → bold → italic → regular), to be
//      synthesized (synthetic());
//   D. the generic families of the family's class (serif / sans-serif / monospace; East Asian
//      families first for the eastAsia slot), any style;
//   E. a standard-14 PDF face of the class in the exact style (metrics only, never embedded).
// fontFor() then picks the rFonts slot of each character (slots.ts), and when the slot's face has
// no glyph for it, the run's other faces and then every resolved face (same style first) are
// tried before a standard-14 face. Everything is cached and deterministic: faces are resolved
// in want() order, resolvers are called one at a time, and fonts() lists faces in first-use order.

import { warning } from '../../core/warnings.js';
import type { ResolvedRunProps } from '../../resolve/types.js';
import type { WordDocument } from '../../wml/document.js';
import type { FontManager } from '../engine.js';
import type { Font, FontRequest, FontResolver, LayoutOptions } from '../types.js';
import { type EmbeddedFace, embeddedFaces, type FontTableEntry, fontTable } from './embedded.js';
import { loadFont } from './sfnt.js';
import { type FontSlot, type SlotProps, slotFor } from './slots.js';
import { standard14For, standardFont } from './standard14.js';
import { baseFont, SubstituteFont } from './substitute.js';
import {
  FONT_SUBSTITUTIONS,
  type GenericFamily,
  genericFamily,
  isSymbolFamily,
  STANDARD14_FALLBACKS,
  substituteMetricsOf,
  substitutesOf,
  type VerticalMetrics,
} from './substitutions.js';

/** The FontManager contract plus synthetic-style reporting. */
export interface LayoutFontManager extends FontManager {
  /**
   * Whether text set in `font` for a run with `props` must be emboldened / slanted by the
   * renderer (TextItem.fakeBold / fakeItalic): the run asks for a style the face lacks.
   * `codePoint` selects the slot (cs uses boldCs/italicCs); without it the run's cs/rtl decides.
   */
  synthetic(props: ResolvedRunProps, font: Font, codePoint?: number): SyntheticStyle;
}

export interface SyntheticStyle {
  bold: boolean;
  italic: boolean;
}

/** East Asian families asked for when an eastAsia-slot family is unavailable. */
export const EAST_ASIAN_FALLBACKS: readonly string[] = [
  'Noto Sans CJK SC',
  'Noto Serif CJK SC',
  'Source Han Sans',
  'Source Han Serif',
  'Noto Sans CJK JP',
  'Noto Sans CJK KR',
  'WenQuanYi Zen Hei',
  'Droid Sans Fallback',
];

/** Word's default when a run names no font at all. */
const DEFAULT_FAMILY = 'Times New Roman';
const SLOTS: readonly FontSlot[] = ['ascii', 'hAnsi', 'eastAsia', 'cs'];

interface Ask {
  family: string;
  bold: boolean;
  italic: boolean;
}

interface RunContext {
  faces: Record<FontSlot, Font>;
  fallback: Map<number, Font>;
  generation: number;
}

type Script = FontRequest['script'];

const styleKey = (bold: boolean, italic: boolean): number => (bold ? 1 : 0) | (italic ? 2 : 0);

/** Characters that never need a visible glyph (no fallback search for them). */
function invisible(cp: number): boolean {
  return (
    cp < 0x20 ||
    (cp >= 0x7f && cp < 0xa0) ||
    cp === 0xad ||
    (cp >= 0x200b && cp <= 0x200f) ||
    (cp >= 0x2028 && cp <= 0x202e) ||
    (cp >= 0x2060 && cp <= 0x206f) ||
    (cp >= 0xfe00 && cp <= 0xfe0f) ||
    cp === 0xfeff ||
    (cp >= 0xe0000 && cp <= 0xe0fff)
  );
}

function slotFamily(props: ResolvedRunProps, slot: FontSlot): string {
  const f = props.fonts;
  return (f[slot] || f.ascii || f.hAnsi || f.eastAsia || f.cs || DEFAULT_FAMILY).trim();
}

function slotStyle(props: ResolvedRunProps, slot: FontSlot): [boolean, boolean] {
  return slot === 'cs' ? [props.boldCs, props.italicCs] : [props.bold, props.italic];
}

function scriptOf(slot: FontSlot, family: string): Script {
  if (isSymbolFamily(family)) return 'symbol';
  return slot === 'eastAsia' ? 'eastAsia' : slot === 'cs' ? 'cs' : 'latin';
}

/** True when `font` already has `original`'s line metrics (within a font unit): no wrapper. */
function sameLineMetrics(font: Font, original: VerticalMetrics): boolean {
  const m = font.metrics;
  const k = m.unitsPerEm / original.unitsPerEm;
  const near = (a: number, b: number): boolean => Math.abs(a - b * k) <= 1;
  return (
    near(m.ascender, original.ascender) &&
    near(m.descender, original.descender) &&
    near(m.lineGap, original.lineGap) &&
    near(m.winAscent, original.winAscent) &&
    near(m.winDescent, original.winDescent)
  );
}

/** The font manager for one layout of `doc`. */
export function createFontManager(doc: WordDocument, o: LayoutOptions = {}): LayoutFontManager {
  const resolvers: readonly FontResolver[] =
    o.fonts === undefined ? [] : typeof o.fonts === 'function' ? [o.fonts] : o.fonts;
  const userSubs = new Map<string, string>();
  for (const [k, v] of Object.entries(o.substitutions ?? {}))
    userSubs.set(k.trim().toLowerCase(), v);

  let table: Map<string, FontTableEntry> | undefined;
  const tableEntry = (family: string): FontTableEntry | undefined => {
    table ??= fontTable(doc);
    return table.get(family.toLowerCase());
  };
  let embedded: Map<string, EmbeddedFace[]> | undefined;
  const embeddedFor = (family: string): EmbeddedFace[] => {
    if (embedded === undefined) {
      embedded = new Map();
      if (o.useEmbeddedFonts !== false) {
        for (const f of embeddedFaces(doc, o.onWarning)) {
          for (const name of new Set([f.name.toLowerCase(), f.font.family.toLowerCase()])) {
            const list = embedded.get(name) ?? [];
            list.push(f);
            embedded.set(name, list);
          }
        }
      }
    }
    return embedded.get(family.toLowerCase()) ?? [];
  };

  const warned = new Set<string>();
  const warn = (code: string, message: string): void => {
    const key = `${code}\0${message}`;
    if (warned.has(key)) return;
    warned.add(key);
    o.onWarning?.(warning(code, message));
  };

  // Resolver results per (family, style, script): a Font, or null = not available.
  const asked = new Map<string, Font | null>();
  const fromBytes = new WeakMap<Uint8Array, Font | null>();
  // Final faces per (family, style).
  const resolved = new Map<string, Font>();
  // Distinct faces in resolution order (per-glyph fallback candidates).
  const pool: Font[] = [];
  const pending = new Map<
    string,
    { family: string; bold: boolean; italic: boolean; slot: FontSlot }
  >();
  let generation = 0;
  const handed: Font[] = [];
  const handedIndex = new Map<Font, number>();
  const contexts = new WeakMap<ResolvedRunProps, RunContext>();
  const standIns = new Map<Font, Map<string, SubstituteFont>>();

  const faceKey = (family: string, bold: boolean, italic: boolean, script: Script): string =>
    `${family.toLowerCase()}|${styleKey(bold, italic)}|${script}`;

  function addToPool(font: Font): void {
    if (!pool.includes(font)) {
      pool.push(font);
      generation++;
    }
  }

  function embeddedFace(ask: Ask): Font | undefined {
    const want = ask.bold
      ? ask.italic
        ? 'boldItalic'
        : 'bold'
      : ask.italic
        ? 'italic'
        : 'regular';
    return embeddedFor(ask.family).find((f) => f.style === want)?.font;
  }

  function accept(result: Font | Uint8Array | undefined, family: string): Font | undefined {
    if (result === undefined || result === null) return undefined;
    if (result instanceof Uint8Array) {
      const cached = fromBytes.get(result);
      if (cached !== undefined) return cached ?? undefined;
      try {
        const font = loadFont(result);
        fromBytes.set(result, font);
        return font;
      } catch (err) {
        fromBytes.set(result, null);
        const why = err instanceof Error ? err.message : String(err);
        warn('FONT_INVALID', `Font data supplied for ${family} is not usable: ${why}`);
        return undefined;
      }
    }
    if (typeof result === 'object' && typeof result.glyphId === 'function') return result;
    return undefined;
  }

  function resolverError(family: string, err: unknown): void {
    const why = err instanceof Error ? err.message : String(err);
    warn('FONT_RESOLVER_ERROR', `Font resolver failed for ${family}: ${why}`);
  }

  async function lookupAsync(ask: Ask, script: Script): Promise<Font | undefined> {
    const local = embeddedFace(ask);
    if (local !== undefined) return local;
    if (resolvers.length === 0) return undefined;
    const key = faceKey(ask.family, ask.bold, ask.italic, script);
    const hit = asked.get(key);
    if (hit !== undefined) return hit ?? undefined;
    const req: FontRequest = { family: ask.family, bold: ask.bold, italic: ask.italic, script };
    let found: Font | undefined;
    for (const resolve of resolvers) {
      try {
        found = accept(await resolve({ ...req }), ask.family);
      } catch (err) {
        resolverError(ask.family, err);
      }
      if (found !== undefined) break;
    }
    asked.set(key, found ?? null);
    return found;
  }

  function lookupSync(ask: Ask, script: Script): Font | undefined {
    const local = embeddedFace(ask);
    if (local !== undefined) return local;
    if (resolvers.length === 0) return undefined;
    const key = faceKey(ask.family, ask.bold, ask.italic, script);
    const hit = asked.get(key);
    if (hit !== undefined) return hit ?? undefined;
    const req: FontRequest = { family: ask.family, bold: ask.bold, italic: ask.italic, script };
    let found: Font | undefined;
    let skipped = false;
    for (const resolve of resolvers) {
      try {
        const r = resolve({ ...req });
        if (r instanceof Promise) {
          r.catch(() => undefined);
          skipped = true;
          continue;
        }
        found = accept(r, ask.family);
      } catch (err) {
        resolverError(ask.family, err);
      }
      if (found !== undefined) break;
    }
    if (skipped) {
      // An async resolver earlier in the list might have answered differently: do not cache.
      warn(
        'FONT_NOT_LOADED',
        `Font ${ask.family} was needed before load() resolved it; async resolvers were skipped`,
      );
    } else {
      asked.set(key, found ?? null);
    }
    return found;
  }

  function matches(font: Font, bold: boolean, italic: boolean): boolean {
    return font.bold === bold && font.italic === italic;
  }

  /** The search order of the file header, as a generator of lookups. */
  function* search(
    family: string,
    bold: boolean,
    italic: boolean,
    slot: FontSlot,
  ): Generator<Ask, Font, Font | undefined> {
    const entry = tableEntry(family);
    const specific: string[] = [family];
    const add = (f: string | undefined): void => {
      if (
        f !== undefined &&
        f !== '' &&
        !specific.some((s) => s.toLowerCase() === f.toLowerCase())
      ) {
        specific.push(f);
      }
    };
    add(entry?.altName);
    add(userSubs.get(family.toLowerCase()));
    for (const s of substitutesOf(family)) add(s);
    if (entry?.altName !== undefined) for (const s of substitutesOf(entry.altName)) add(s);

    // A + B (a face found under another name stands in for `family`: standIn)
    let inexact: Font | undefined;
    for (const f of specific) {
      const found = yield { family: f, bold, italic };
      if (found === undefined) continue;
      const font = f === family ? found : standIn(found, family);
      if (matches(font, bold, italic)) return font;
      inexact ??= font;
    }
    if (inexact !== undefined) return inexact;
    // C
    const lesser: [boolean, boolean][] = [];
    if (bold && italic) lesser.push([true, false], [false, true]);
    if (bold || italic) lesser.push([false, false]);
    for (const [b, i] of lesser) {
      for (const f of specific) {
        const font = yield { family: f, bold: b, italic: i };
        if (font !== undefined) return f === family ? font : standIn(font, family);
      }
    }
    // D
    const generic: GenericFamily = genericFamily(family, entry);
    const generics = [
      ...(slot === 'eastAsia' ? EAST_ASIAN_FALLBACKS : []),
      ...(generic === 'symbol' ? [] : (FONT_SUBSTITUTIONS[generic] ?? [])),
    ].filter((g) => !specific.some((s) => s.toLowerCase() === g.toLowerCase()));
    for (const [b, i] of [[bold, italic] as [boolean, boolean], ...lesser]) {
      for (const f of generics) {
        const font = yield { family: f, bold: b, italic: i };
        if (font !== undefined) return standIn(font, family);
      }
    }
    // E
    const std =
      family.replace(/\s+/g, '').toLowerCase() === 'zapfdingbats'
        ? 'ZapfDingbats'
        : STANDARD14_FALLBACKS[generic];
    const face = standardFont(standard14For(std, bold, italic));
    warn(
      'FONT_FALLBACK',
      `Font ${family} is not available; using the standard PDF font ${face.postscriptName}`,
    );
    return standIn(face, family);
  }

  /** `font` reporting the vertical metrics of `family` when they are known (SUBSTITUTE_METRICS). */
  function standIn(font: Font, family: string): Font {
    const original = substituteMetricsOf(family);
    const base = baseFont(font);
    if (original === undefined || sameLineMetrics(base, original)) return base;
    let byFamily = standIns.get(base);
    if (byFamily === undefined) {
      byFamily = new Map();
      standIns.set(base, byFamily);
    }
    const key = family.toLowerCase();
    let wrapped = byFamily.get(key);
    if (wrapped === undefined) {
      wrapped = new SubstituteFont(base, family, original);
      byFamily.set(key, wrapped);
    }
    return wrapped;
  }

  function finish(key: string, font: Font): Font {
    resolved.set(key, font);
    addToPool(font);
    return font;
  }

  async function resolveAsync(
    family: string,
    bold: boolean,
    italic: boolean,
    slot: FontSlot,
  ): Promise<Font> {
    const script = scriptOf(slot, family);
    const key = faceKey(family, bold, italic, script);
    const done = resolved.get(key);
    if (done !== undefined) return done;
    const gen = search(family, bold, italic, slot);
    let step = gen.next();
    while (!step.done) step = gen.next(await lookupAsync(step.value, script));
    return finish(key, step.value);
  }

  function resolveSync(family: string, bold: boolean, italic: boolean, slot: FontSlot): Font {
    const script = scriptOf(slot, family);
    const key = faceKey(family, bold, italic, script);
    const done = resolved.get(key);
    if (done !== undefined) return done;
    const gen = search(family, bold, italic, slot);
    let step = gen.next();
    while (!step.done) step = gen.next(lookupSync(step.value, script));
    return finish(key, step.value);
  }

  function context(props: ResolvedRunProps): RunContext {
    let ctx = contexts.get(props);
    if (ctx === undefined) {
      const faces = {} as Record<FontSlot, Font>;
      for (const slot of SLOTS) {
        const [b, i] = slotStyle(props, slot);
        faces[slot] = resolveSync(slotFamily(props, slot), b, i, slot);
      }
      ctx = { faces, fallback: new Map(), generation };
      contexts.set(props, ctx);
    }
    if (ctx.generation !== generation) {
      ctx.fallback.clear();
      ctx.generation = generation;
    }
    return ctx;
  }

  function handOut(font: Font): Font {
    if (!handedIndex.has(font)) {
      handedIndex.set(font, handed.length);
      handed.push(font);
    }
    return font;
  }

  function fallbackFor(props: ResolvedRunProps, slot: FontSlot, primary: Font, cp: number): Font {
    const [bold, italic] = slotStyle(props, slot);
    const ctx = context(props);
    const candidates: Font[] = [];
    for (const s of ['eastAsia', 'hAnsi', 'ascii', 'cs'] as const) {
      const f = ctx.faces[s];
      if (f !== primary && !candidates.includes(f)) candidates.push(f);
    }
    for (const f of pool) if (f !== primary && !candidates.includes(f)) candidates.push(f);
    const real = candidates.filter((f) => f.outlines !== 'standard14');
    const sameStyle = real.filter((f) => matches(f, bold, italic));
    for (const list of [sameStyle, real]) {
      for (const f of list) if (f.glyphId(cp) !== 0) return f;
    }
    // Last resort: a standard-14 face of the primary's class that can encode the character.
    const generic = genericFamily(primary.family, tableEntry(primary.family));
    const stdFamilies =
      generic === 'symbol'
        ? ['Symbol', 'ZapfDingbats', 'Helvetica']
        : [STANDARD14_FALLBACKS[generic], 'Symbol', 'ZapfDingbats'];
    for (const fam of stdFamilies) {
      const f = standardFont(standard14For(fam, bold, italic));
      if (f !== primary && f.glyphId(cp) !== 0) return f;
    }
    return primary;
  }

  const manager: LayoutFontManager = {
    want(props: ResolvedRunProps): void {
      for (const slot of SLOTS) {
        const family = slotFamily(props, slot);
        const [bold, italic] = slotStyle(props, slot);
        const script = scriptOf(slot, family);
        const key = faceKey(family, bold, italic, script);
        if (resolved.has(key) || pending.has(key)) continue;
        pending.set(key, { family, bold, italic, slot });
      }
    },

    async load(): Promise<void> {
      while (pending.size > 0) {
        const [key, req] = pending.entries().next().value!;
        pending.delete(key);
        await resolveAsync(req.family, req.bold, req.italic, req.slot);
      }
    },

    fontFor(props: ResolvedRunProps, codePoint: number): Font {
      const ctx = context(props);
      const slot = slotFor(props as SlotProps, codePoint);
      const primary = ctx.faces[slot];
      if (invisible(codePoint) || primary.glyphId(codePoint) !== 0) return handOut(primary);
      let f = ctx.fallback.get(codePoint);
      if (f === undefined) {
        f = fallbackFor(props, slot, primary, codePoint);
        ctx.fallback.set(codePoint, f);
      }
      return handOut(f);
    },

    fonts(): Font[] {
      return handed.slice();
    },

    indexOf(font: Font): number {
      const i = handedIndex.get(font);
      if (i !== undefined) return i;
      handOut(font);
      return handed.length - 1;
    },

    synthetic(props: ResolvedRunProps, font: Font, codePoint?: number): SyntheticStyle {
      const slot: FontSlot =
        codePoint !== undefined
          ? slotFor(props as SlotProps, codePoint)
          : props.cs || props.rtl
            ? 'cs'
            : 'ascii';
      const [bold, italic] = slotStyle(props, slot);
      return { bold: bold && !font.bold, italic: italic && !font.italic };
    },
  };
  return manager;
}
