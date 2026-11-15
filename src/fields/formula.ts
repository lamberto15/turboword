// The `=` (formula) field: arithmetic, comparisons, functions, bookmark and table references
// (ECMA-376 Part 1 §17.16.3).
//
// - operators: + - * / ^, postfix % (percent), = <> < <= > >= (1 or 0); unary minus binds
//   tighter than ^, as in Excel;
// - functions: ABS AND AVERAGE COUNT DEFINED FALSE IF INT MAX MIN MOD NOT OR PRODUCT ROUND SIGN
//   SUM TRUE (case-insensitive);
// - references: inside a table cell, ABOVE/BELOW/LEFT/RIGHT (contiguous numeric cells in that
//   direction, stopping at the first empty or text cell), A1-style cells and ranges (A1:B3) and
//   R1C1 cells of the containing table; anywhere, bookmark names (their text as a number);
// - literals use the document's decimal symbol; function arguments are separated by the list
//   separator (and `;`); currency symbols are ignored, so nested results like "$1,234.50" work.
// Errors (syntax, unknown bookmark, division by zero) throw FormulaError; the caller leaves the
// field unchanged, where Word would show "!Syntax Error" and the like.

export class FormulaError extends Error {}

/** One cell of the containing table, as formulas see it. */
export interface CellValue {
  /** Identity of the cell (merged cells cover several grid slots). */
  id: object;
  /** The cell's number, or null when it is empty or holds text. */
  value: number | null;
}

/** The table around a formula field. */
export interface TableEnv {
  row: number;
  col: number;
  rows: number;
  cols: number;
  cell(r: number, c: number): CellValue | undefined;
}

export interface FormulaEnv {
  decimal: string;
  group: string;
  listSep: string;
  /** Number held by a bookmark; undefined when there is no such bookmark. */
  bookmark(name: string): number | undefined;
  table?: TableEnv | undefined;
}

type Tok =
  | { t: 'num'; v: number }
  | { t: 'id'; v: string }
  | { t: 'op'; v: string }
  | { t: '(' | ')' | 'sep' | ':' };

type Node =
  | { k: 'num'; v: number }
  | { k: 'ref'; name: string }
  | { k: 'range'; a: string; b: string }
  | { k: 'un'; op: string; x: Node }
  | { k: 'bin'; op: string; a: Node; b: Node }
  | { k: 'pct'; x: Node }
  | { k: 'call'; f: string; args: Node[] };

type Value = number | number[];

const MAX_DEPTH = 64;
/** Keeps left-deep operator chains well within the call stack when evaluated. */
const MAX_TOKENS = 4000;

function lex(src: string, env: FormulaEnv): Tok[] {
  const out: Tok[] = [];
  let depth = 0;
  let i = 0;
  const isDigit = (c: string | undefined): boolean => c !== undefined && c >= '0' && c <= '9';
  while (i < src.length) {
    const c = src[i]!;
    if (/[\s\u00a0"$€£¥]/.test(c)) {
      i++;
      continue;
    }
    if (isDigit(c) || (c === env.decimal && isDigit(src[i + 1]))) {
      let s = '';
      while (i < src.length) {
        const d = src[i]!;
        if (isDigit(d)) s += d;
        else if (d === env.decimal && !s.includes('.')) s += '.';
        else if (
          d === env.group &&
          !s.includes('.') &&
          (depth === 0 || env.group !== env.listSep) &&
          /^\d{3}(?!\d)/.test(src.slice(i + 1))
        ) {
          // grouping symbol inside a number
        } else break;
        i++;
      }
      if (/^[eE][+-]?\d/.test(src.slice(i))) {
        const m = /^[eE][+-]?\d+/.exec(src.slice(i))!;
        s += m[0];
        i += m[0].length;
      }
      out.push({ t: 'num', v: Number(s) });
      continue;
    }
    if (/[\p{L}_]/u.test(c)) {
      let s = '';
      while (i < src.length && /[\p{L}\p{N}_]/u.test(src[i]!)) s += src[i++];
      out.push({ t: 'id', v: s });
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      out.push({ t: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^%=<>'.includes(c)) out.push({ t: 'op', v: c });
    else if (c === '(') {
      depth++;
      out.push({ t: '(' });
    } else if (c === ')') {
      depth--;
      out.push({ t: ')' });
    } else if (c === env.listSep || c === ';' || c === ',') out.push({ t: 'sep' });
    else if (c === ':') out.push({ t: ':' });
    else throw new FormulaError(`Unexpected character ${c}`);
    i++;
  }
  return out;
}

class Parser {
  readonly toks: Tok[];
  pos = 0;
  depth = 0;

  constructor(toks: Tok[]) {
    this.toks = toks;
  }

  peek(): Tok | undefined {
    return this.toks[this.pos];
  }

  isOp(v: string): boolean {
    const t = this.peek();
    return t?.t === 'op' && t.v === v;
  }

  expect(t: Tok['t']): void {
    if (this.peek()?.t !== t) throw new FormulaError('Syntax error');
    this.pos++;
  }

  parse(): Node {
    const n = this.cmp();
    if (this.pos < this.toks.length) throw new FormulaError('Syntax error');
    return n;
  }

  cmp(): Node {
    let a = this.add();
    for (;;) {
      const t = this.peek();
      if (t?.t === 'op' && ['=', '<>', '<', '<=', '>', '>='].includes(t.v)) {
        this.pos++;
        a = { k: 'bin', op: t.v, a, b: this.add() };
      } else return a;
    }
  }

  add(): Node {
    let a = this.mul();
    while (this.isOp('+') || this.isOp('-')) {
      const op = (this.toks[this.pos++] as { v: string }).v;
      a = { k: 'bin', op, a, b: this.mul() };
    }
    return a;
  }

  mul(): Node {
    let a = this.pow();
    while (this.isOp('*') || this.isOp('/')) {
      const op = (this.toks[this.pos++] as { v: string }).v;
      a = { k: 'bin', op, a, b: this.pow() };
    }
    return a;
  }

  pow(): Node {
    let a = this.unary();
    while (this.isOp('^')) {
      this.pos++;
      a = { k: 'bin', op: '^', a, b: this.unary() };
    }
    return a;
  }

  unary(): Node {
    if (this.isOp('-') || this.isOp('+')) {
      const op = (this.toks[this.pos++] as { v: string }).v;
      if (++this.depth > MAX_DEPTH) throw new FormulaError('Formula too deep');
      const x = this.unary();
      this.depth--;
      return { k: 'un', op, x };
    }
    let x = this.primary();
    while (this.isOp('%')) {
      this.pos++;
      x = { k: 'pct', x };
    }
    return x;
  }

  primary(): Node {
    const t = this.peek();
    if (t === undefined) throw new FormulaError('Syntax error');
    if (t.t === 'num') {
      this.pos++;
      return { k: 'num', v: t.v };
    }
    if (t.t === '(') {
      this.pos++;
      if (++this.depth > MAX_DEPTH) throw new FormulaError('Formula too deep');
      const n = this.cmp();
      this.depth--;
      this.expect(')');
      return n;
    }
    if (t.t === 'id') {
      this.pos++;
      if (this.peek()?.t === '(') {
        this.pos++;
        if (++this.depth > MAX_DEPTH) throw new FormulaError('Formula too deep');
        const args: Node[] = [];
        if (this.peek()?.t !== ')') {
          args.push(this.cmp());
          while (this.peek()?.t === 'sep') {
            this.pos++;
            args.push(this.cmp());
          }
        }
        this.depth--;
        this.expect(')');
        return { k: 'call', f: t.v.toUpperCase(), args };
      }
      if (this.peek()?.t === ':') {
        this.pos++;
        const b = this.peek();
        if (b?.t !== 'id') throw new FormulaError('Syntax error');
        this.pos++;
        return { k: 'range', a: t.v, b: b.v };
      }
      return { k: 'ref', name: t.v };
    }
    throw new FormulaError('Syntax error');
  }
}

/** Grid position of an A1 or R1C1 reference, or undefined when `name` is neither. */
function cellRef(name: string): { r: number; c: number } | undefined {
  const a1 = /^([A-Za-z]{1,2})(\d{1,4})$/.exec(name);
  if (a1) {
    const letters = a1[1]!.toUpperCase();
    let c = 0;
    for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64);
    return { r: Number(a1[2]) - 1, c: c - 1 };
  }
  const rc = /^R(\d{1,4})C(\d{1,4})$/i.exec(name);
  if (rc) return { r: Number(rc[1]) - 1, c: Number(rc[2]) - 1 };
  return undefined;
}

const DIRECTIONS: Readonly<Record<string, readonly [number, number]>> = {
  ABOVE: [-1, 0],
  BELOW: [1, 0],
  LEFT: [0, -1],
  RIGHT: [0, 1],
};

function scalar(v: Value): number {
  if (typeof v === 'number') return v;
  if (v.length === 1) return v[0]!;
  throw new FormulaError('A list is not a number');
}

function flatten(vals: Value[]): number[] {
  const out: number[] = [];
  for (const v of vals) {
    if (typeof v === 'number') out.push(v);
    else out.push(...v);
  }
  return out;
}

/** Rounds half away from zero to `places` decimals. */
export function roundHalfAway(x: number, places: number): number {
  const p = Math.trunc(places);
  const r = Math.round(Number(`${Math.abs(x)}e${p}`));
  return Math.sign(x) * Number(`${r}e${-p}`);
}

class Evaluator {
  readonly env: FormulaEnv;

  constructor(env: FormulaEnv) {
    this.env = env;
  }

  ref(name: string): Value {
    const upper = name.toUpperCase();
    if (upper === 'TRUE') return 1;
    if (upper === 'FALSE') return 0;
    const table = this.env.table;
    const dir = DIRECTIONS[upper];
    if (dir) return this.direction(dir);
    const pos = table ? cellRef(name) : undefined;
    if (table && pos) {
      const cell = table.cell(pos.r, pos.c);
      if (cell === undefined) throw new FormulaError(`No cell ${name}`);
      return cell.value ?? 0;
    }
    const b = this.env.bookmark(name);
    if (b === undefined) throw new FormulaError(`Undefined bookmark ${name}`);
    return b;
  }

  direction([dr, dc]: readonly [number, number]): number[] {
    const t = this.env.table;
    if (t === undefined) throw new FormulaError('Not in a table');
    const out: number[] = [];
    const seen = new Set<object>();
    const own = t.cell(t.row, t.col)?.id;
    if (own) seen.add(own);
    for (let r = t.row + dr, c = t.col + dc; r >= 0 && c >= 0 && r < t.rows && c < t.cols; ) {
      const cell = t.cell(r, c);
      if (cell !== undefined && !seen.has(cell.id)) {
        seen.add(cell.id);
        if (cell.value === null) break;
        out.push(cell.value);
      }
      r += dr;
      c += dc;
    }
    return out;
  }

  range(a: string, b: string): number[] {
    const t = this.env.table;
    const p = cellRef(a);
    const q = cellRef(b);
    if (t === undefined || p === undefined || q === undefined) {
      throw new FormulaError('Invalid range');
    }
    const out: number[] = [];
    const seen = new Set<object>();
    for (let r = Math.min(p.r, q.r); r <= Math.max(p.r, q.r); r++) {
      for (let c = Math.min(p.c, q.c); c <= Math.max(p.c, q.c); c++) {
        const cell = t.cell(r, c);
        if (cell === undefined || seen.has(cell.id)) continue;
        seen.add(cell.id);
        if (cell.value !== null) out.push(cell.value);
      }
    }
    return out;
  }

  eval(n: Node): Value {
    switch (n.k) {
      case 'num':
        return n.v;
      case 'ref':
        return this.ref(n.name);
      case 'range':
        return this.range(n.a, n.b);
      case 'pct':
        return scalar(this.eval(n.x)) / 100;
      case 'un': {
        const x = scalar(this.eval(n.x));
        return n.op === '-' ? -x : x;
      }
      case 'bin':
        return this.binary(n.op, scalar(this.eval(n.a)), scalar(this.eval(n.b)));
      case 'call':
        return this.call(n.f, n.args);
    }
  }

  binary(op: string, a: number, b: number): number {
    switch (op) {
      case '+':
        return a + b;
      case '-':
        return a - b;
      case '*':
        return a * b;
      case '/':
        if (b === 0) throw new FormulaError('Zero divide');
        return a / b;
      case '^':
        return a ** b;
      case '=':
        return a === b ? 1 : 0;
      case '<>':
        return a !== b ? 1 : 0;
      case '<':
        return a < b ? 1 : 0;
      case '<=':
        return a <= b ? 1 : 0;
      case '>':
        return a > b ? 1 : 0;
      default:
        return a >= b ? 1 : 0;
    }
  }

  call(f: string, args: Node[]): Value {
    if (f === 'DEFINED') {
      try {
        for (const a of args) this.eval(a);
        return args.length > 0 ? 1 : 0;
      } catch {
        return 0;
      }
    }
    if (f === 'IF') {
      if (args.length !== 3) throw new FormulaError('IF takes 3 arguments');
      return scalar(this.eval(args[0]!)) !== 0 ? this.eval(args[1]!) : this.eval(args[2]!);
    }
    const vals = args.map((a) => this.eval(a));
    const one = (): number => {
      if (vals.length !== 1) throw new FormulaError(`${f} takes 1 argument`);
      return scalar(vals[0]!);
    };
    const two = (): [number, number] => {
      if (vals.length !== 2) throw new FormulaError(`${f} takes 2 arguments`);
      return [scalar(vals[0]!), scalar(vals[1]!)];
    };
    const list = flatten(vals);
    switch (f) {
      case 'SUM':
        return list.reduce((s, x) => s + x, 0);
      case 'PRODUCT':
        return list.reduce((s, x) => s * x, 1);
      case 'AVERAGE':
        if (list.length === 0) throw new FormulaError('AVERAGE of nothing');
        return list.reduce((s, x) => s + x, 0) / list.length;
      case 'COUNT':
        return list.length;
      case 'MIN':
        if (list.length === 0) throw new FormulaError('MIN of nothing');
        return Math.min(...list);
      case 'MAX':
        if (list.length === 0) throw new FormulaError('MAX of nothing');
        return Math.max(...list);
      case 'ABS':
        return Math.abs(one());
      case 'INT':
        return Math.trunc(one());
      case 'SIGN':
        return Math.sign(one());
      case 'NOT':
        return one() === 0 ? 1 : 0;
      case 'AND':
        if (list.length === 0) throw new FormulaError('AND of nothing');
        return list.every((x) => x !== 0) ? 1 : 0;
      case 'OR':
        if (list.length === 0) throw new FormulaError('OR of nothing');
        return list.some((x) => x !== 0) ? 1 : 0;
      case 'ROUND': {
        const [x, p] = two();
        return roundHalfAway(x, p);
      }
      case 'MOD': {
        const [x, y] = two();
        if (y === 0) throw new FormulaError('Zero divide');
        return x - y * Math.floor(x / y);
      }
      case 'TRUE':
        return 1;
      case 'FALSE':
        return 0;
      default:
        throw new FormulaError(`Unknown function ${f}`);
    }
  }
}

/** Evaluates a formula expression. Throws FormulaError on any error. */
export function evaluateFormula(expr: string, env: FormulaEnv): number {
  const toks = lex(expr, env);
  if (toks.length === 0) throw new FormulaError('Empty formula');
  if (toks.length > MAX_TOKENS) throw new FormulaError('Formula too long');
  const ast = new Parser(toks).parse();
  const v = scalar(new Evaluator(env).eval(ast));
  if (!Number.isFinite(v)) throw new FormulaError('Overflow');
  return v;
}
