// Document protection (w:documentProtection): Word's password hash against known vectors and
// Word-written files, the settings XML we write (attribute form, schema position), round trips,
// and documents without a settings part or in Strict OOXML.
//
// Vectors:
// - Word-written: Apache POI test-data bug56076.docx (Word 2010: trackedChanges, SHA-1, 100,000
//   spins, password "Example"; POI's TestDocumentProtection validates it), in
//   test/fixtures/crypto/apache-poi/ with its licence.
// - Legacy key: "Example" → 0x64CEED7E (ECMA-376 Part 4 §14.8.2.4 example; POI, Aspose),
//   "34579" → 0x0005CB00 (POI, leading-zero case), "test" → 0x1FC6CBEB and a long password →
//   0x46FFEEFF (Aspose.Words FOSS TestProtection.cs, MIT).
// - Full hash: Aspose's "test" with salt hIqOl99kIr/+N0hyovEq7Q==, SHA-1, 50,000 spins →
//   wZko1pqoSiQdGEuyf0evzPwm9AE= (TestProtection.UnifiedTestProtectSetPassword).
// - SHA-256/384/512 against an independent node:crypto implementation of the same algorithm.

import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { sha1, sha512 } from '../src/core/sha512.js';
import {
  createDocx,
  NS,
  readDocx,
  serializeXml,
  UsageError,
  validatePackage,
  type WordDocument,
  writeDocx,
} from '../src/index.js';
import { legacyKeyString, legacyPasswordKey, protectionHash } from '../src/protect/hash.js';
import {
  protectDocument,
  protectionOf,
  unprotectDocument,
  verifyProtectionPassword,
} from '../src/protect/index.js';
import { fixture as corpusFixture } from './helpers/fixtures.js';

const FIX = fileURLToPath(new URL('./fixtures/crypto/apache-poi/', import.meta.url));
const poi = (name: string): WordDocument => readDocx(new Uint8Array(readFileSync(join(FIX, name))));
const OUT = fileURLToPath(new URL('../out/protect-crypto/', import.meta.url));

const b64 = (s: string): Uint8Array => new Uint8Array(Buffer.from(s, 'base64'));
const SALT = b64('hIqOl99kIr/+N0hyovEq7Q==');
const PW = 'Example';

const settingsRoot = (doc: WordDocument) => doc.mainPart.relatedByType('settings')[0]!.xml().root;
const settingsXml = (doc: WordDocument): string =>
  serializeXml(doc.mainPart.relatedByType('settings')[0]!.xml());

/** Attributes of w:documentProtection, by qualified name. */
function protection(doc: WordDocument): Record<string, string> {
  const el = settingsRoot(doc).element(NS.w, 'documentProtection')!;
  return Object.fromEntries(el.attributes().map((a) => [a.name, a.value]));
}

/** Qualified names of the w:settings children, in order. */
const settingsOrder = (doc: WordDocument): string[] =>
  settingsRoot(doc)
    .elements()
    .map((e) => e.name);

/** Iterated salted hash on node:crypto: H0 = H(salt + UTF-16LE(input)), Hn = H(Hn-1 + LE32). */
function spun(alg: string, salt: Uint8Array, input: string, spin: number): string {
  let h = createHash(alg).update(salt).update(Buffer.from(input, 'utf16le')).digest();
  const it = Buffer.alloc(4);
  for (let i = 0; i < spin; i++) {
    it.writeUInt32LE(i);
    h = createHash(alg).update(h).update(it).digest();
  }
  return h.toString('base64');
}

/**
 * Independent reference of Word's document-protection hash on node:crypto. The legacy key comes
 * from the implementation (checked against the published vectors below); the byte reversal, hex
 * form and iteration are re-implemented here.
 */
function refHash(alg: string, salt: Uint8Array, password: string, spin: number): string {
  const key = legacyPasswordKey(password);
  const hex = [0, 8, 16, 24]
    .map((s) => ((key >>> s) & 0xff).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
  return spun(alg, salt, hex, spin);
}

describe('legacy password key', () => {
  it('matches the published vectors', () => {
    expect(legacyPasswordKey('Example')).toBe(0x64ceed7e);
    expect(legacyKeyString('Example')).toBe('7EEDCE64');
    expect(legacyPasswordKey('34579')).toBe(0x0005cb00);
    expect(legacyKeyString('34579')).toBe('00CB0500');
    expect(legacyPasswordKey('test')).toBe(0x1fc6cbeb);
    expect(legacyPasswordKey('This is a very long password.')).toBe(0x46ffeeff);
    expect(
      legacyPasswordKey('This is a very long password, but even longer than the previous one.'),
    ).toBe(0x46ffeeff);
    expect(legacyPasswordKey('')).toBe(0);
  });

  it('uses the high byte of characters whose low byte is zero', () => {
    // U+0100 has low byte 0 -> high byte 0x01, same as U+0001.
    expect(legacyPasswordKey('aĀb')).toBe(legacyPasswordKey('a\u0001b'));
    // U+0141 keeps its low byte 0x41, same as 'A'.
    expect(legacyPasswordKey('Ł')).toBe(legacyPasswordKey('A'));
  });
});

describe('protection hash', () => {
  it('matches Aspose.Words (SHA-1, 50,000 spins)', () => {
    const h = protectionHash(sha1, SALT, 'test', 50000);
    expect(Buffer.from(h).toString('base64')).toBe('wZko1pqoSiQdGEuyf0evzPwm9AE=');
  });

  it('matches an independent node:crypto implementation for every SHA', () => {
    const salt = b64('2Z+i7o/0EZyUNakVeWzU/w==');
    expect(Buffer.from(protectionHash(sha512, salt, 'pässwörd', 1000)).toString('base64')).toBe(
      refHash('sha512', salt, 'pässwörd', 1000),
    );
    const doc = createDocx();
    for (const [algorithm, node] of [
      ['SHA-1', 'sha1'],
      ['SHA-256', 'sha256'],
      ['SHA-384', 'sha384'],
      ['SHA-512', 'sha512'],
    ] as const) {
      protectDocument(doc, {
        mode: 'readOnly',
        password: 'Secret',
        algorithm,
        salt,
        spinCount: 77,
      });
      expect(protection(doc)['w:hash'], algorithm).toBe(refHash(node, salt, 'Secret', 77));
      expect(verifyProtectionPassword(doc, 'Secret')).toBe(true);
      expect(verifyProtectionPassword(doc, 'secret')).toBe(false);
    }
  });
});

describe('Word-written protection', () => {
  it('verifies bug56076.docx (Word 2010, SHA-1, 100,000 spins, "Example")', () => {
    const doc = poi('bug56076.docx');
    expect(protectionOf(doc)).toEqual({ mode: 'trackedChanges', hasPassword: true });
    expect(verifyProtectionPassword(doc, 'Example')).toBe(true);
    expect(verifyProtectionPassword(doc, 'example')).toBe(false);
    expect(verifyProtectionPassword(doc, '')).toBe(false);
  });

  it('reads protection without password, unenforced and formatting-only restrictions', () => {
    const ro = poi('documentProtection_readonly_no_password.docx');
    expect(protectionOf(ro)).toEqual({ mode: 'readOnly', hasPassword: false });
    expect(verifyProtectionPassword(ro, '')).toBe(true);
    expect(verifyProtectionPassword(ro, 'x')).toBe(false);
    // w:enforcement="0"
    expect(protectionOf(poi('documentProtection_no_protection_tag_existing.docx'))).toBe(undefined);
    // w:formatting="1" only, no w:edit
    expect(protectionOf(poi('EnforcedWith.docx'))).toBe(undefined);
    // Word 2007 file (SHA-1, 50,000 spins), password unknown
    const ps = poi('protected_sample.docx');
    expect(protectionOf(ps)).toEqual({ mode: 'readOnly', hasPassword: true });
    expect(verifyProtectionPassword(ps, 'Example')).toBe(false);
    expect(protectionOf(createDocx())).toBe(undefined);
    expect(verifyProtectionPassword(createDocx(), 'x')).toBe(false);
  });

  it('unprotects a Word file and keeps the rest of settings', () => {
    const doc = poi('bug56076.docx');
    const before = settingsOrder(doc);
    unprotectDocument(doc);
    expect(protectionOf(doc)).toBe(undefined);
    expect(settingsOrder(doc)).toEqual(before.filter((n) => n !== 'w:documentProtection'));
    const back = readDocx(writeDocx(doc));
    expect(protectionOf(back)).toBe(undefined);
    expect(settingsXml(back)).not.toContain('documentProtection');
  });
});

describe('protectDocument', () => {
  it('writes Word 2013+ attributes (rsaAES, sid 14) at the schema position', () => {
    const doc = createDocx();
    protectDocument(doc, { mode: 'readOnly', password: PW, salt: SALT, spinCount: 1000 });
    const hash = refHash('sha512', SALT, PW, 1000);
    expect(settingsXml(doc)).toContain(
      `<w:zoom w:percent="100"/><w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="1000" w:hash="${hash}" w:salt="hIqOl99kIr/+N0hyovEq7Q=="/><w:defaultTabStop`,
    );
    expect(protectionOf(doc)).toEqual({ mode: 'readOnly', hasPassword: true });
    expect(verifyProtectionPassword(doc, PW)).toBe(true);
    expect(verifyProtectionPassword(doc, 'Exampl')).toBe(false);
  });

  it('reproduces the Aspose vector through the public API (SHA-1 → rsaFull, sid 4)', () => {
    const doc = createDocx();
    protectDocument(doc, {
      mode: 'forms',
      password: 'test',
      algorithm: 'SHA-1',
      salt: SALT,
      spinCount: 50000,
    });
    expect(protection(doc)).toEqual({
      'w:edit': 'forms',
      'w:enforcement': '1',
      'w:cryptProviderType': 'rsaFull',
      'w:cryptAlgorithmClass': 'hash',
      'w:cryptAlgorithmType': 'typeAny',
      'w:cryptAlgorithmSid': '4',
      'w:cryptSpinCount': '50000',
      'w:hash': 'wZko1pqoSiQdGEuyf0evzPwm9AE=',
      'w:salt': 'hIqOl99kIr/+N0hyovEq7Q==',
    });
  });

  it('defaults: SHA-512, 100,000 spins, random 16-byte salt', () => {
    const doc = createDocx();
    protectDocument(doc, { mode: 'comments', password: 'pw' });
    const p = protection(doc);
    expect(p['w:cryptAlgorithmSid']).toBe('14');
    expect(p['w:cryptSpinCount']).toBe('100000');
    expect(b64(p['w:salt']!).length).toBe(16);
    expect(b64(p['w:hash']!).length).toBe(64);
    expect(verifyProtectionPassword(doc, 'pw')).toBe(true);
    const other = createDocx();
    protectDocument(other, { mode: 'comments', password: 'pw' });
    expect(protection(other)['w:salt']).not.toBe(p['w:salt']);
  });

  it('is deterministic with a fixed salt', () => {
    const run = () => {
      const doc = createDocx();
      protectDocument(doc, { mode: 'readOnly', password: 'x', salt: SALT, spinCount: 10 });
      return writeDocx(doc);
    };
    expect(run()).toEqual(run());
  });

  it('writes every mode; no password means no hash attributes', () => {
    for (const mode of ['readOnly', 'comments', 'trackedChanges', 'forms'] as const) {
      const doc = createDocx();
      protectDocument(doc, { mode });
      expect(protection(doc)).toEqual({ 'w:edit': mode, 'w:enforcement': '1' });
      expect(protectionOf(doc)).toEqual({ mode, hasPassword: false });
      protectDocument(doc, { mode, password: '' });
      expect(protectionOf(doc)).toEqual({ mode, hasPassword: false });
      expect(verifyProtectionPassword(doc, '')).toBe(true);
    }
  });

  it('trackedChanges turns w:trackRevisions on, in schema order', () => {
    const doc = createDocx();
    protectDocument(doc, { mode: 'trackedChanges' });
    expect(settingsOrder(doc).slice(0, 3)).toEqual([
      'w:zoom',
      'w:trackRevisions',
      'w:documentProtection',
    ]);
    // An existing w:trackRevisions w:val="0" is switched on, not duplicated.
    const root = settingsRoot(doc);
    root.element(NS.w, 'trackRevisions')!.setAttrNS(NS.w, 'val', '0');
    protectDocument(doc, { mode: 'trackedChanges', password: 'p', spinCount: 1 });
    expect(root.elements(NS.w, 'trackRevisions')).toHaveLength(1);
    expect(root.element(NS.w, 'trackRevisions')!.attrNS(NS.w, 'val')).toBe(undefined);
    // Unprotecting leaves tracking on (as Word does).
    unprotectDocument(doc);
    expect(root.element(NS.w, 'trackRevisions')).toBeDefined();
    expect(root.element(NS.w, 'documentProtection')).toBe(undefined);
  });

  it('replaces an existing protection in place and keeps w:formatting', () => {
    const doc = poi('EnforcedWith.docx');
    const before = settingsOrder(doc);
    protectDocument(doc, { mode: 'comments', password: 'p', spinCount: 1 });
    expect(settingsOrder(doc)).toEqual(before);
    expect(protection(doc)).toMatchObject({
      'w:edit': 'comments',
      'w:formatting': '1',
      'w:enforcement': '1',
    });
    const doc2 = poi('bug56076.docx');
    protectDocument(doc2, { mode: 'readOnly' });
    expect(protection(doc2)).toEqual({ 'w:edit': 'readOnly', 'w:enforcement': '1' });
    expect(verifyProtectionPassword(doc2, 'Example')).toBe(false);
  });

  it('inserts at the schema position in real Word settings (before w14/w15 extensions)', () => {
    for (const name of ['docxcompose/numberings.docx', 'python-docx/comments-rich-para.docx']) {
      const doc = readDocx(corpusFixture(name).bytes);
      protectDocument(doc, { mode: 'readOnly' });
      const order = settingsOrder(doc);
      const i = order.indexOf('w:documentProtection');
      expect(i, name).toBeGreaterThanOrEqual(0);
      const ext = order.findIndex((n) => !n.startsWith('w:'));
      if (ext >= 0) expect(i, name).toBeLessThan(ext);
      for (const after of ['w:defaultTabStop', 'w:compat', 'w:rsids', 'w:decimalSymbol']) {
        const j = order.indexOf(after);
        if (j >= 0) expect(i, `${name} ${after}`).toBeLessThan(j);
      }
      for (const before of ['w:zoom', 'w:proofState', 'w:trackRevisions']) {
        const j = order.indexOf(before);
        if (j >= 0) expect(i, `${name} ${before}`).toBeGreaterThan(j);
      }
    }
  });

  it('creates the settings part when the document has none', () => {
    const doc = createDocx();
    doc.mainPart.rels.remove(doc.mainPart.rels.byType('settings')[0]!.id);
    doc.package.removePart('word/settings.xml');
    expect(protectionOf(doc)).toBe(undefined);
    unprotectDocument(doc); // no-op: creates nothing
    expect(doc.mainPart.relatedByType('settings')).toHaveLength(0);

    protectDocument(doc, { mode: 'forms', password: 'p', spinCount: 1 });
    const part = doc.mainPart.relatedByType('settings')[0]!;
    expect(part.name).toBe('word/settings.xml');
    expect(part.contentType).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    );
    const back = readDocx(writeDocx(doc));
    expect(protectionOf(back)).toEqual({ mode: 'forms', hasPassword: true });
    expect(verifyProtectionPassword(back, 'p')).toBe(true);
    expect(validatePackage(back).filter((i) => i.severity === 'error')).toEqual([]);
  });

  it('writes the ISO attributes in Strict documents', () => {
    const doc = readDocx(corpusFixture('open-xml-sdk/strict-CommentRangeStart.docx').bytes);
    protectDocument(doc, { mode: 'readOnly', password: PW, salt: SALT, spinCount: 1000 });
    const p = protection(doc);
    expect(p).toMatchObject({
      'w:edit': 'readOnly',
      'w:enforcement': '1',
      'w:algorithmName': 'SHA-512',
      'w:hashValue': refHash('sha512', SALT, PW, 1000),
      'w:saltValue': 'hIqOl99kIr/+N0hyovEq7Q==',
      'w:spinCount': '1000',
    });
    expect(Object.keys(p).some((k) => k.startsWith('w:crypt'))).toBe(false);
    const back = readDocx(writeDocx(doc));
    expect(settingsXml(back)).toContain('http://purl.oclc.org/ooxml/wordprocessingml/main');
    expect(verifyProtectionPassword(back, PW)).toBe(true);
    expect(verifyProtectionPassword(back, 'example')).toBe(false);
  });

  it('rejects bad options with UsageError', () => {
    const doc = createDocx();
    expect(() => protectDocument(doc, { mode: 'none' as 'forms' })).toThrow(UsageError);
    expect(() => protectDocument(doc, { mode: 'forms', spinCount: -1 })).toThrow(UsageError);
    expect(() => protectDocument(doc, { mode: 'forms', spinCount: 1.5 })).toThrow(UsageError);
    expect(() => protectDocument(doc, { mode: 'forms', spinCount: 10_000_001 })).toThrow(
      UsageError,
    );
    expect(() =>
      protectDocument(doc, { mode: 'forms', password: 'x', algorithm: 'MD5' as 'SHA-1' }),
    ).toThrow(UsageError);
    expect(() =>
      protectDocument(doc, { mode: 'forms', password: 'x', salt: new Uint8Array(0) }),
    ).toThrow(UsageError);
    expect(protectionOf(doc)).toBe(undefined);
  });
});

describe('verifyProtectionPassword (ISO attribute form)', () => {
  /** A document whose w:documentProtection carries exactly `attrs`. */
  function withProtection(attrs: Record<string, string>): WordDocument {
    const doc = createDocx();
    const settings = settingsRoot(doc);
    const el = settings.ownerDocument()!.createElement(NS.w, 'documentProtection');
    for (const [k, v] of Object.entries(attrs)) el.setAttrNS(NS.w, k, v);
    settings.prepend(el);
    return doc;
  }
  const iso = (hashValue: string, algorithmName = 'SHA-512') => ({
    edit: 'readOnly',
    enforcement: '1',
    algorithmName,
    hashValue,
    saltValue: 'hIqOl99kIr/+N0hyovEq7Q==',
    spinCount: '100',
  });

  it('accepts Word pre-processing (legacy key) and raw-password hashes', () => {
    const legacy = withProtection(iso(spun('sha512', SALT, '7EEDCE64', 100)));
    expect(verifyProtectionPassword(legacy, 'Example')).toBe(true);
    expect(verifyProtectionPassword(legacy, 'Example2')).toBe(false);
    const raw = withProtection(iso(spun('sha256', SALT, 'Example', 100), 'SHA-256'));
    expect(verifyProtectionPassword(raw, 'Example')).toBe(true);
    expect(verifyProtectionPassword(raw, 'example')).toBe(false);
  });

  it('prefers the transitional attributes when both forms are present (as Word)', () => {
    const doc = withProtection({
      ...iso(spun('sha512', SALT, '7EEDCE64', 100)),
      cryptAlgorithmSid: '14',
      cryptSpinCount: '100',
      hash: spun('sha512', SALT, legacyKeyString('Other'), 100),
      salt: 'hIqOl99kIr/+N0hyovEq7Q==',
    });
    expect(verifyProtectionPassword(doc, 'Other')).toBe(true);
    expect(verifyProtectionPassword(doc, 'Example')).toBe(false);
  });

  it('never matches unsupported algorithms, absurd spin counts or broken base64', () => {
    const md5 = withProtection({
      ...iso(spun('md5', SALT, '7EEDCE64', 100)),
      algorithmName: 'MD5',
    });
    expect(verifyProtectionPassword(md5, 'Example')).toBe(false);
    const spin = withProtection({ ...iso('AAAA'), spinCount: '20000000' });
    expect(verifyProtectionPassword(spin, 'Example')).toBe(false);
    const broken = withProtection({ ...iso('%%%'), saltValue: '***' });
    expect(verifyProtectionPassword(broken, 'Example')).toBe(false);
    const sid1 = withProtection({ cryptAlgorithmSid: '1', hash: 'AAAA', salt: 'AAAA' });
    expect(verifyProtectionPassword(sid1, 'Example')).toBe(false);
  });

  it('treats a missing w:enforcement as enforced (as Word)', () => {
    expect(protectionOf(withProtection({ edit: 'forms' }))).toEqual({
      mode: 'forms',
      hasPassword: false,
    });
    expect(protectionOf(withProtection({ edit: 'none', enforcement: '1' }))).toBe(undefined);
    expect(protectionOf(withProtection({ edit: 'readOnly', enforcement: 'false' }))).toBe(
      undefined,
    );
  });
});

describe('outputs', () => {
  it('stay valid packages (written to out/protect-crypto/ for the LibreOffice check)', () => {
    mkdirSync(OUT, { recursive: true });
    const cases: [string, WordDocument][] = [];
    for (const mode of ['readOnly', 'comments', 'trackedChanges', 'forms'] as const) {
      const doc = createDocx();
      protectDocument(doc, { mode, password: PW, salt: SALT });
      cases.push([`created-${mode}.docx`, doc]);
    }
    const word = poi('bug56076.docx');
    protectDocument(word, { mode: 'readOnly', password: PW, salt: SALT });
    cases.push(['bug56076-readonly.docx', word]);
    const strict = readDocx(corpusFixture('open-xml-sdk/strict-footer-bookmarkStart.docx').bytes);
    protectDocument(strict, { mode: 'forms', password: PW, salt: SALT });
    cases.push(['strict-forms.docx', strict]);
    const noSettings = createDocx();
    noSettings.mainPart.rels.remove(noSettings.mainPart.rels.byType('settings')[0]!.id);
    noSettings.package.removePart('word/settings.xml');
    protectDocument(noSettings, { mode: 'comments' });
    cases.push(['no-settings-comments.docx', noSettings]);
    for (const [name, doc] of cases) {
      const bytes = writeDocx(doc);
      writeFileSync(join(OUT, name), bytes);
      const back = readDocx(bytes);
      expect(
        validatePackage(back).filter((i) => i.severity === 'error'),
        name,
      ).toEqual([]);
      expect(protectionOf(back), name).toBeDefined();
    }
  });
});
