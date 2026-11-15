// Markup Compatibility (ECMA-376 Part 3) helpers: `mc:Ignorable` lists and `mc:AlternateContent`.
//
// Word marks extension namespaces (w14, w15, ...) as ignorable and wraps newer constructs in
// `mc:AlternateContent` with a `mc:Choice` per requirement set and a `mc:Fallback`. turboword keeps
// all branches on round-trip; these helpers are for code that must *read* the content (pick the
// branch it understands) or keep `mc:Ignorable` consistent after adding namespaces.

import { UsageError } from '../core/errors.js';
import type { XElement } from './dom.js';
import { canonicalNs, conventionalUri, NS } from './ns.js';

/** One `mc:Choice` of an `mc:AlternateContent`. */
export interface McChoice {
  /** Prefixes listed in its `Requires` attribute. */
  requires: string[];
  /** The `mc:Choice` element. */
  el: XElement;
}

/** The branches of an `mc:AlternateContent` element, in document order. */
export interface AlternateContent {
  choices: McChoice[];
  /** The `mc:Fallback` element, when present. */
  fallback?: XElement;
}

function splitList(v: string | undefined): string[] {
  return v === undefined ? [] : v.split(/[ \t\r\n]+/).filter((t) => t !== '');
}

/** Prefixes listed in `el`'s `mc:Ignorable` attribute (empty when absent). */
export function ignorablePrefixes(el: XElement): string[] {
  return splitList(el.attrNS(NS.mc, 'Ignorable'));
}

/**
 * Sets `el`'s `mc:Ignorable` to `prefixes` (deduplicated, in order); removes the attribute when
 * the list is empty. The `mc` prefix is declared on save if needed.
 */
export function setIgnorable(el: XElement, prefixes: readonly string[]): void {
  const list = [...new Set(prefixes.filter((p) => p !== ''))];
  if (list.length === 0) el.removeAttrNS(NS.mc, 'Ignorable');
  else el.setAttrNS(NS.mc, 'Ignorable', list.join(' '));
}

/** The `mc:Choice` branches and `mc:Fallback` of an `mc:AlternateContent` element. */
export function alternateContent(el: XElement): AlternateContent {
  if (!el.is(NS.mc, 'AlternateContent')) {
    throw new UsageError(`Expected mc:AlternateContent, got ${el.name}`);
  }
  const choices = el
    .elements(NS.mc, 'Choice')
    .map((c) => ({ requires: splitList(c.attrNS('', 'Requires')), el: c }));
  const fallback = el.element(NS.mc, 'Fallback');
  return fallback === undefined ? { choices } : { choices, fallback };
}

/**
 * The branch a consumer should read: the first `mc:Choice` whose every required prefix resolves
 * (in scope, else leniently to its conventional URI) to a namespace `understood` accepts, else
 * the `mc:Fallback`, else undefined. `understood` receives canonical URIs.
 */
export function selectBranch(
  ac: AlternateContent,
  understood: (uri: string) => boolean,
): XElement | undefined {
  for (const c of ac.choices) {
    const ok = c.requires.every((p) => {
      const uri = c.el.lookupNamespace(p) ?? conventionalUri(p);
      return uri !== undefined && understood(canonicalNs(uri));
    });
    if (ok) return c.el;
  }
  return ac.fallback;
}
