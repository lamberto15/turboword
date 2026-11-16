// Repo conventions enforced as tests, so a mistake fails fast on every OS instead of surfacing as a
// bundler error, a non-deterministic output or a Windows-only test failure much later:
// - layers: imports only go downward (core <- zip, xml <- opc <- wml <- ...), see ARCHITECTURE.md;
// - multi-runtime: no static `node:` imports or `require(` in src; only src/zip/node.ts may look up
//   Node built-ins (via process.getBuiltinModule);
// - determinism: no Math.random / Date.now / argument-less `new Date()` in src;
// - ESM: every relative import in src ends with `.js`;
// - tests derive paths with fileURLToPath (URL.pathname breaks on Windows: "/D:/...");
// - licensing: the dual licence is declared and shipped.
// Sources are scanned with a small comment/string-aware lexer rather than raw regexes, so a comment
// that *mentions* Date.now() does not fail the build and a commented-out import does not count.

import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { join, posix, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const SRC = join(ROOT, 'src');

/**
 * Layer rank of each top-level directory of src (from the plan's layer table). A file may import
 * its own directory or a directory of strictly lower rank, so same-rank siblings (zip/xml) stay
 * independent. Files directly in src/ (the entry points) may import anything; nothing may import
 * them. A new directory must be added here before it can be imported.
 */
const RANK: Readonly<Record<string, number>> = {
  core: 0,
  zip: 1,
  xml: 1,
  opc: 2,
  wml: 3,
  ids: 3,
  resolve: 4,
  text: 4,
  validate: 4,
  compose: 5,
  fields: 5,
  forms: 5,
  protect: 5,
  crypto: 5,
  layout: 6,
  mailmerge: 6,
  render: 7,
  pdf: 7,
  svg: 7,
};
/** The one file allowed to reach Node built-ins (through process.getBuiltinModule). */
const NODE_GATEWAY = 'zip/node.ts';

function walk(dir: string, accept: (name: string) => boolean, skip: string[] = []): string[] {
  const out: string[] = [];
  if (!existsSync(dir)) return out;
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) {
      if (!skip.includes(name)) out.push(...walk(p, accept, skip));
    } else if (accept(name)) out.push(p);
  }
  return out;
}

/** Path relative to `base` with forward slashes (stable across OSes in failure messages). */
const rel = (base: string, file: string): string => relative(base, file).split(sep).join('/');

/**
 * `src` with comments blanked out (newlines kept, so offsets and line numbers survive). With
 * `blankStrings`, the contents of string, template and regex literals are blanked too, which is
 * what identifier checks want; import parsing keeps them (specifiers are strings).
 */
function strip(src: string, blankStrings: boolean): string {
  const blank = (s: string): string => s.replace(/[^\n]/g, ' ');
  let out = '';
  let prev = ''; // last significant code character: decides regex literal vs division
  let i = 0;
  const n = src.length;
  while (i < n) {
    const c = src[i]!;
    const d = src[i + 1];
    if (c === '/' && d === '/') {
      const e = src.indexOf('\n', i);
      const end = e < 0 ? n : e;
      out += blank(src.slice(i, end));
      i = end;
      continue;
    }
    if (c === '/' && d === '*') {
      const e = src.indexOf('*/', i + 2);
      const end = e < 0 ? n : e + 2;
      out += blank(src.slice(i, end));
      i = end;
      continue;
    }
    if (c === '"' || c === "'" || c === '`') {
      let j = i + 1;
      while (j < n && src[j] !== c) {
        if (src[j] === '\\') j++;
        else if (c !== '`' && src[j] === '\n') break;
        j++;
      }
      const end = Math.min(j + 1, n);
      const body = src.slice(i + 1, end - 1);
      out += c + (blankStrings ? blank(body) : body) + (end - 1 > i ? src[end - 1] : '');
      prev = c;
      i = end;
      continue;
    }
    const regexStart =
      c === '/' &&
      (prev === '' ||
        '(,=:[!&|?{};+-*%<>~^'.includes(prev) ||
        /\b(?:return|typeof|case|do|else|in|of|void|yield|await|throw)$/.test(out.trimEnd()));
    if (regexStart) {
      let j = i + 1;
      let inClass = false;
      while (j < n && src[j] !== '\n') {
        const ch = src[j];
        if (ch === '\\') j++;
        else if (ch === '[') inClass = true;
        else if (ch === ']') inClass = false;
        else if (ch === '/' && !inClass) break;
        j++;
      }
      j++;
      while (j < n && /[a-z]/i.test(src[j]!)) j++;
      const body = src.slice(i, j);
      out += blankStrings ? blank(body) : body;
      prev = '/';
      i = j;
      continue;
    }
    out += c;
    if (!/\s/.test(c)) prev = c;
    i++;
  }
  return out;
}

const STATIC_IMPORT =
  /^[ \t]*(?:import|export)\s+(?:type\s+)?(?:[\w$*{},\s]+?\s+from\s+)?(['"])([^'"\n]+)\1/gm;
const DYNAMIC_IMPORT = /\bimport\s*\(\s*(['"])([^'"\n]+)\1/g;

interface SourceFile {
  /** Path relative to src/, forward slashes. */
  readonly path: string;
  /** Comments and literal contents blanked: for identifier checks. */
  readonly bare: string;
  /** Import/export specifiers (static and dynamic). */
  readonly specifiers: readonly string[];
}

function scan(path: string, text: string): SourceFile {
  const code = strip(text, false);
  const specifiers = [...code.matchAll(STATIC_IMPORT), ...code.matchAll(DYNAMIC_IMPORT)].map(
    (m) => m[2]!,
  );
  return { path, bare: strip(text, true), specifiers };
}

const layerOf = (path: string): string => (path.includes('/') ? path.split('/')[0]! : '<entry>');
const imports = (
  sources: readonly SourceFile[],
  offence: (s: SourceFile, spec: string) => string | false,
): string[] =>
  sources.flatMap((s) =>
    s.specifiers.flatMap((spec) => {
      const why = offence(s, spec);
      return why === false ? [] : [`${s.path}: '${spec}'${why ? ` (${why})` : ''}`];
    }),
  );
const files = (sources: readonly SourceFile[], bad: (s: SourceFile) => boolean): string[] =>
  sources.filter(bad).map((s) => s.path);

/** Each check returns offenders as readable `path: detail` strings (empty = clean). */
const CHECKS = {
  unknownLayers: (sources: readonly SourceFile[]): string[] => [
    ...new Set(
      sources.map((s) => layerOf(s.path)).filter((l) => l !== '<entry>' && RANK[l] === undefined),
    ),
  ],

  layering: (sources: readonly SourceFile[]): string[] =>
    imports(sources, (s, spec) => {
      if (!spec.startsWith('.')) return false;
      const target = posix.normalize(posix.join(posix.dirname(s.path), spec));
      if (target.startsWith('../')) return 'leaves src/';
      const from = layerOf(s.path);
      const to = layerOf(target);
      if (to === from || from === '<entry>') return false;
      const down = to !== '<entry>' && (RANK[to] ?? Infinity) < (RANK[from] ?? -Infinity);
      return down ? false : `${from} -> ${to}`;
    }),

  opcImportsWml: (sources: readonly SourceFile[]): string[] =>
    imports(sources, (s, spec) =>
      layerOf(s.path) === 'opc' &&
      spec.startsWith('.') &&
      layerOf(posix.normalize(posix.join(posix.dirname(s.path), spec))) === 'wml'
        ? ''
        : false,
    ),

  packages: (sources: readonly SourceFile[], deps: readonly string[]): string[] =>
    imports(sources, (s, spec) => {
      if (spec.startsWith('.') || spec.startsWith('node:')) return false;
      const parts = spec.split('/');
      const pkg = spec.startsWith('@') ? parts.slice(0, 2).join('/') : parts[0]!;
      if (layerOf(s.path) === 'core') return 'core imports nothing outside core';
      return deps.includes(pkg) ? false : 'not a runtime dependency';
    }),

  nodeImports: (sources: readonly SourceFile[]): string[] =>
    imports(sources, (_s, spec) => (spec.startsWith('node:') ? '' : false)),

  requireCalls: (sources: readonly SourceFile[]): string[] =>
    files(sources, (s) => /\brequire\s*\(/.test(s.bare)),

  builtinLookups: (sources: readonly SourceFile[]): string[] =>
    files(sources, (s) => s.path !== NODE_GATEWAY && /\bgetBuiltinModule\b/.test(s.bare)),

  extensionless: (sources: readonly SourceFile[]): string[] =>
    imports(sources, (_s, spec) => (spec.startsWith('.') && !spec.endsWith('.js') ? '' : false)),

  nondeterminism: (sources: readonly SourceFile[]): string[] => {
    const rules: [RegExp, string][] = [
      [/\bMath\s*\.\s*random\b/, 'Math.random'],
      [/\bDate\s*\.\s*now\b/, 'Date.now'],
      [/\bnew\s+Date\s*\(\s*\)/, 'new Date()'],
    ];
    return sources.flatMap((s) =>
      s.bare
        .split('\n')
        .flatMap((line, i) =>
          rules.filter(([re]) => re.test(line)).map(([, what]) => `${s.path}:${i + 1}: ${what}`),
        ),
    );
  },
};

const SOURCES: readonly SourceFile[] = walk(SRC, (n) => n.endsWith('.ts')).map((f) =>
  scan(rel(SRC, f), readFileSync(f, 'utf8')),
);
const PKG = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')) as {
  license?: string;
  files?: string[];
  dependencies?: Record<string, string>;
};
const RUNTIME_DEPS = Object.keys(PKG.dependencies ?? {});

describe('the checks themselves', () => {
  it('the scanner blanks comments but keeps code, strings and regex literals', () => {
    const src = [
      "import { a } from './a.js'; // Date.now()",
      '/* Math.random() */ const re = /\\/\\*not a comment/g;',
      "const s = '// not a comment'; const t = `new Date()`;",
      'const q = x / 2; // trailing',
    ].join('\n');
    const code = strip(src, false);
    expect(code).toContain("from './a.js'");
    expect(code).not.toMatch(/Date\.now|Math\.random|trailing/);
    expect(code).toContain('/\\/\\*not a comment/g');
    expect(code).toContain("'// not a comment'");
    expect(code).toContain('x / 2');
    expect(code.split('\n')).toHaveLength(4);
    const bare = strip(src, true);
    expect(bare).not.toMatch(/new Date|not a comment|a\.js/);
    expect(bare).toContain('x / 2');
  });

  it('flag synthetic violations', () => {
    const bad = [
      scan('core/a.ts', "import { x } from '../xml/dom.js';\nimport { y } from 'fflate';"),
      scan('zip/a.ts', "export { z } from '../xml/ns.js';"),
      scan('xml/a.ts', "import type {\n  Part,\n} from '../opc/part.js';\nimport('./lazy');"),
      scan('opc/a.ts', "import { readDocx } from '../wml/document.js';"),
      scan('wml/a.ts', "import '../index.js';\nimport { readFileSync } from 'node:fs';"),
      scan('wml/b.ts', "import x from 'left-pad';\nconst m = require('m');"),
      scan('xml/b.ts', 'const k = process.getBuiltinModule("node:zlib");'),
      scan('mystery/a.ts', 'const t = Date.now() + Math.random();\nconst d = new Date( );'),
    ];
    expect(CHECKS.unknownLayers(bad)).toEqual(['mystery']);
    expect(CHECKS.layering(bad)).toEqual([
      "core/a.ts: '../xml/dom.js' (core -> xml)",
      "zip/a.ts: '../xml/ns.js' (zip -> xml)",
      "xml/a.ts: '../opc/part.js' (xml -> opc)",
      "opc/a.ts: '../wml/document.js' (opc -> wml)",
      "wml/a.ts: '../index.js' (wml -> <entry>)",
    ]);
    expect(CHECKS.opcImportsWml(bad)).toEqual(["opc/a.ts: '../wml/document.js'"]);
    expect(CHECKS.packages(bad, ['fflate'])).toEqual([
      "core/a.ts: 'fflate' (core imports nothing outside core)",
      "wml/b.ts: 'left-pad' (not a runtime dependency)",
    ]);
    expect(CHECKS.nodeImports(bad)).toEqual(["wml/a.ts: 'node:fs'"]);
    expect(CHECKS.requireCalls(bad)).toEqual(['wml/b.ts']);
    expect(CHECKS.builtinLookups(bad)).toEqual(['xml/b.ts']);
    expect(CHECKS.extensionless(bad)).toEqual(["xml/a.ts: './lazy'"]);
    expect(CHECKS.nondeterminism(bad)).toEqual([
      'mystery/a.ts:1: Math.random',
      'mystery/a.ts:1: Date.now',
      'mystery/a.ts:2: new Date()',
    ]);
  });

  it('find the real source files and their imports', () => {
    expect(SOURCES.length).toBeGreaterThan(0);
    expect(SOURCES.find((s) => s.path === 'xml/index.ts')?.specifiers).toContain('./dom.js');
  });
});

describe('layers (imports only go downward)', () => {
  it('every top-level src directory has a layer rank', () => {
    expect(CHECKS.unknownLayers(SOURCES), 'add it to RANK in test/conventions.test.ts').toEqual([]);
  });

  it('relative imports stay inside src and only point to lower layers', () => {
    expect(CHECKS.layering(SOURCES), 'core <- zip, xml <- opc <- wml <- ...').toEqual([]);
  });

  it('src/opc never imports wml', () => {
    expect(CHECKS.opcImportsWml(SOURCES)).toEqual([]);
  });

  it('packages: none in core, elsewhere only runtime dependencies', () => {
    expect(CHECKS.packages(SOURCES, RUNTIME_DEPS)).toEqual([]);
  });
});

describe('multi-runtime', () => {
  it('no node: imports in src (static or dynamic)', () => {
    expect(CHECKS.nodeImports(SOURCES), `use process.getBuiltinModule in ${NODE_GATEWAY}`).toEqual(
      [],
    );
  });

  it('no require( in src', () => {
    expect(CHECKS.requireCalls(SOURCES)).toEqual([]);
  });

  it(`getBuiltinModule appears only in src/${NODE_GATEWAY}`, () => {
    expect(CHECKS.builtinLookups(SOURCES)).toEqual([]);
  });

  it('every relative import in src ends with .js', () => {
    expect(CHECKS.extensionless(SOURCES)).toEqual([]);
  });
});

describe('determinism', () => {
  it('no Math.random, Date.now or argument-less new Date() in src', () => {
    expect(
      CHECKS.nondeterminism(SOURCES),
      'use counters, the seeded PRNG or a date option',
    ).toEqual([]);
  });
});

describe('test conventions', () => {
  it('never derives file paths from URL.pathname (breaks on Windows: "/D:/...")', () => {
    const scripts = ['test', 'bench', 'scripts'].flatMap((d) =>
      walk(join(ROOT, d), (n) => /\.(ts|mts|mjs|js)$/.test(n), ['fixtures', 'node_modules']),
    );
    const offenders = scripts
      .filter((f) => /import\.meta\.url\)\s*\.pathname/.test(readFileSync(f, 'utf8')))
      .map((f) => rel(ROOT, f));
    expect(offenders, 'use fileURLToPath(new URL(..., import.meta.url)) instead').toEqual([]);
  });
});

describe('licensing', () => {
  const LICENCE_FILES = ['LICENSE-MIT', 'LICENSE-APACHE', 'NOTICE'];

  it('package.json declares MIT OR Apache-2.0', () => {
    expect(PKG.license).toBe('MIT OR Apache-2.0');
  });

  it('both licence texts and NOTICE are published and exist', () => {
    expect(PKG.files ?? []).toEqual(expect.arrayContaining(LICENCE_FILES));
    expect(LICENCE_FILES.filter((f) => !existsSync(join(ROOT, f)))).toEqual([]);
  });
});
