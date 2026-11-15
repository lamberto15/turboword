// Bounds-checked big-endian reader over font bytes.
//
// Font files come from documents and users, so every read is checked: a truncated or hostile
// file must end in a TurbowordError('FONT_FORMAT'), never in a RangeError from DataView or in a
// silent read past the end. Offsets are absolute within the underlying byte array; table parsers
// receive the table's [start, end) range and check against that, so one table cannot read
// another's bytes by accident.

import { TurbowordError } from '../../core/errors.js';

/** Throws the error every malformed-font path ends in. */
export function fontError(message: string): never {
  throw new TurbowordError('FONT_FORMAT', message);
}

/** A window [start, end) over the font bytes with checked big-endian reads. */
export class Reader {
  readonly bytes: Uint8Array;
  readonly view: DataView;
  readonly start: number;
  readonly end: number;
  readonly what: string;

  constructor(bytes: Uint8Array, start: number, end: number, what: string) {
    if (!(start >= 0 && end <= bytes.length && start <= end)) {
      fontError(`${what}: range ${start}..${end} outside the font (${bytes.length} bytes)`);
    }
    this.bytes = bytes;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.start = start;
    this.end = end;
    this.what = what;
  }

  get length(): number {
    return this.end - this.start;
  }

  /** A sub-window at `offset` (relative to this window's start) of `length` bytes (default: to the end). */
  sub(offset: number, length?: number, what?: string): Reader {
    const s = this.start + offset;
    const e = length === undefined ? this.end : s + length;
    if (!(offset >= 0 && s <= this.end && e <= this.end && e >= s)) {
      fontError(`${what ?? this.what}: offset ${offset} outside ${this.what}`);
    }
    return new Reader(this.bytes, s, e, what ?? this.what);
  }

  /** True when `size` bytes at relative `offset` are inside the window. */
  has(offset: number, size: number): boolean {
    return offset >= 0 && size >= 0 && offset + size <= this.length;
  }

  check(offset: number, size: number): number {
    if (!(offset >= 0 && offset + size <= this.length)) {
      fontError(`${this.what}: read of ${size} bytes at ${offset} past the end (${this.length})`);
    }
    return this.start + offset;
  }

  u8(o: number): number {
    return this.view.getUint8(this.check(o, 1));
  }
  i8(o: number): number {
    return this.view.getInt8(this.check(o, 1));
  }
  u16(o: number): number {
    return this.view.getUint16(this.check(o, 2));
  }
  i16(o: number): number {
    return this.view.getInt16(this.check(o, 2));
  }
  u24(o: number): number {
    const a = this.check(o, 3);
    return (this.view.getUint8(a) << 16) | this.view.getUint16(a + 1);
  }
  u32(o: number): number {
    return this.view.getUint32(this.check(o, 4));
  }
  i32(o: number): number {
    return this.view.getInt32(this.check(o, 4));
  }
  /** 16.16 fixed. */
  fixed(o: number): number {
    return this.i32(o) / 65536;
  }
  tag(o: number): string {
    const a = this.check(o, 4);
    const b = this.bytes;
    return String.fromCharCode(b[a]!, b[a + 1]!, b[a + 2]!, b[a + 3]!);
  }

  /** Fails unless `count` records of `size` bytes fit at relative `offset`. */
  array(offset: number, count: number, size: number): void {
    if (!(count >= 0 && this.has(offset, count * size))) {
      fontError(`${this.what}: ${count} records of ${size} bytes at ${offset} exceed the table`);
    }
  }
}
