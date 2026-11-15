// A small deterministic PDF 1.7 object writer. Objects are numbered in allocation order, bodies are
// plain strings (dictionaries are written by hand: there are only a dozen object shapes) plus raw
// stream bytes, and the file ends with a classic cross-reference table. The trailer /ID is the MD5
// of the file body, so identical input gives identical bytes and no clock is ever read.

import { zlibSync } from 'fflate';
import { md5 } from './md5.js';

/** Formats a number for PDF: at most `digits` decimals (default 3), no exponent, no `-0`. */
export function num(n: number, digits = 3): string {
  if (!Number.isFinite(n)) return '0';
  const m = 10 ** digits;
  const r = Math.round(n * m) / m;
  if (r === 0) return '0';
  if (Number.isInteger(r)) return String(r);
  return r.toFixed(digits).replace(/0+$/, '');
}

/** Hex digits of `bytes`, upper-case. */
export function hex(bytes: Uint8Array): string {
  let s = '';
  for (const b of bytes) s += (b < 16 ? '0' : '') + b.toString(16).toUpperCase();
  return s;
}

/**
 * A PDF text string: a literal `( )` string when `s` is printable ASCII, otherwise UTF-16BE hex
 * with a byte-order mark. Equal inputs always give equal outputs (name-tree keys rely on it).
 */
export function str(s: string): string {
  if (/^[\x20-\x7e]*$/.test(s)) return `(${s.replace(/[\\()]/g, '\\$&')})`;
  let h = '<FEFF';
  for (let i = 0; i < s.length; i++)
    h += s.charCodeAt(i).toString(16).toUpperCase().padStart(4, '0');
  return `${h}>`;
}

/** A PDF name object (`/Name`), escaping delimiters and non-regular bytes as `#xx`. */
export function name(s: string): string {
  let out = '/';
  for (const b of new TextEncoder().encode(s)) {
    const c = String.fromCharCode(b);
    out +=
      b < 0x21 || b > 0x7e || '()<>[]{}/%#'.includes(c)
        ? `#${b.toString(16).toUpperCase().padStart(2, '0')}`
        : c;
  }
  return out;
}

/** Latin-1 bytes of a string whose characters are all < 256 (content streams, dictionaries). */
export function latin1(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
}

/** `n 0 R`. */
export const ref = (id: number): string => `${id} 0 R`;

/** Collects numbered objects and serializes them into a complete PDF file. */
export class PdfWriter {
  private readonly objects: (Uint8Array[] | undefined)[] = [];

  /** Reserves the next object number (so it can be referenced before it is written). */
  alloc(): number {
    this.objects.push(undefined);
    return this.objects.length;
  }

  /** Sets the body of a non-stream object (a dictionary, array, ...). */
  set(id: number, body: string): void {
    this.objects[id - 1] = [latin1(`${id} 0 obj\n${body}\nendobj\n`)];
  }

  /**
   * Sets a stream object. `entries` are the dictionary entries without `<< >>` and without
   * /Length; with `deflate` the data is compressed and /Filter /FlateDecode is added.
   */
  stream(id: number, entries: string, data: Uint8Array, deflate: boolean): void {
    let bytes = data;
    let dict = entries;
    if (deflate) {
      bytes = zlibSync(data, { level: 6 });
      dict = `/Filter /FlateDecode${dict ? ` ${dict}` : ''}`;
    }
    const head = latin1(
      `${id} 0 obj\n<<${dict ? ` ${dict}` : ''} /Length ${bytes.length} >>\nstream\n`,
    );
    this.objects[id - 1] = [head, bytes, latin1('\nendstream\nendobj\n')];
  }

  /** alloc + set. */
  add(body: string): number {
    const id = this.alloc();
    this.set(id, body);
    return id;
  }

  /** alloc + stream. */
  addStream(entries: string, data: Uint8Array, deflate: boolean): number {
    const id = this.alloc();
    this.stream(id, entries, data, deflate);
    return id;
  }

  /** Serializes every object, the xref table and the trailer. */
  finish(root: number, info: number): Uint8Array {
    const chunks: Uint8Array[] = [latin1('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n')];
    const offsets: number[] = [];
    let pos = chunks[0]!.length;
    this.objects.forEach((parts, i) => {
      offsets.push(pos);
      for (const p of parts ?? [latin1(`${i + 1} 0 obj\nnull\nendobj\n`)]) {
        chunks.push(p);
        pos += p.length;
      }
    });
    const body = concat(chunks, pos);
    const id = hex(md5(body));
    let xref = `xref\n0 ${offsets.length + 1}\n0000000000 65535 f \n`;
    for (const o of offsets) xref += `${String(o).padStart(10, '0')} 00000 n \n`;
    xref +=
      `trailer\n<< /Size ${offsets.length + 1} /Root ${ref(root)} /Info ${ref(info)} ` +
      `/ID [<${id}> <${id}>] >>\nstartxref\n${pos}\n%%EOF\n`;
    const tail = latin1(xref);
    const out = new Uint8Array(pos + tail.length);
    out.set(body);
    out.set(tail, pos);
    return out;
  }
}

/** Concatenates byte chunks (`total` = their summed length when known). */
export function concat(chunks: readonly Uint8Array[], total?: number): Uint8Array {
  const out = new Uint8Array(total ?? chunks.reduce((n, c) => n + c.length, 0));
  let p = 0;
  for (const c of chunks) {
    out.set(c, p);
    p += c.length;
  }
  return out;
}
