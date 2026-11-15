// The sfnt `name` table: family, subfamily, full and PostScript names.
//
// Word matches fonts by the legacy (RIBBI) family name, name ID 1 ("Calibri Light", "Arial
// Narrow"), not the typographic family (ID 16), so `family` is ID 1. Among the records of one ID
// the Windows-platform US-English one wins, then any Windows English, any Windows, Mac Roman
// English, then Unicode-platform records.

import { MAC_ROMAN_HIGH } from './cmap.js';
import type { Reader } from './reader.js';

export interface FontNames {
  family: string;
  subfamily: string;
  postscriptName: string;
  typographicFamily: string;
  fullName: string;
}

const IDS = [1, 2, 4, 6, 16] as const;

export function decodeNames(r: Reader): FontNames {
  const count = r.u16(2);
  const storage = r.u16(4);
  r.array(6, count, 12);
  const best = new Map<number, { score: number; value: string }>();
  for (let i = 0; i < count; i++) {
    const p = 6 + i * 12;
    const nameId = r.u16(p + 6);
    if (!(IDS as readonly number[]).includes(nameId)) continue;
    const platform = r.u16(p);
    const encoding = r.u16(p + 2);
    const language = r.u16(p + 4);
    const score = scoreOf(platform, encoding, language);
    if (score <= 0 || (best.get(nameId)?.score ?? 0) >= score) continue;
    const length = r.u16(p + 8);
    const offset = storage + r.u16(p + 10);
    if (!r.has(offset, length)) continue;
    const value = platform === 1 ? macRoman(r, offset, length) : utf16be(r, offset, length);
    const clean = value.replace(/\0/g, '').trim();
    if (clean !== '') best.set(nameId, { score, value: clean });
  }
  const get = (id: number): string => best.get(id)?.value ?? '';
  return {
    family: get(1),
    subfamily: get(2),
    fullName: get(4),
    postscriptName: get(6),
    typographicFamily: get(16),
  };
}

function scoreOf(platform: number, encoding: number, language: number): number {
  if (platform === 3 && (encoding === 1 || encoding === 10 || encoding === 0)) {
    if (language === 0x409) return 50;
    if ((language & 0x3ff) === 0x09) return 40;
    return 30;
  }
  if (platform === 1 && encoding === 0) return language === 0 ? 20 : 10;
  if (platform === 0) return 15;
  return 0;
}

function utf16be(r: Reader, offset: number, length: number): string {
  let s = '';
  for (let i = 0; i + 1 < length; i += 2) s += String.fromCharCode(r.u16(offset + i));
  return s;
}

function macRoman(r: Reader, offset: number, length: number): string {
  let s = '';
  for (let i = 0; i < length; i++) {
    const c = r.u8(offset + i);
    s += String.fromCodePoint(c < 0x80 ? c : MAC_ROMAN_HIGH[c - 0x80]!);
  }
  return s;
}
