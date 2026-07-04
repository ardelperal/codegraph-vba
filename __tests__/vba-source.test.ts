/**
 * Issue #53: encoding-robust source reading for VBA-family files.
 *
 * Unit tests for `src/extraction/vba-source.ts`. Buffers are injected via the
 * `readFile` DI so no temp files are needed — each atom drives the decode path
 * directly:
 *   - happy UTF-8 (no BOM)         -> byte-identical, no fallback
 *   - UTF-8 with BOM (EF BB BF)    -> BOM stripped, no U+FEFF leak
 *   - CP1252 (0xF3 = 'ó')          -> fatal UTF-8 throws, windows-1252 fallback
 *   - empty buffer                 -> '' / bomStripped false
 *   - BOM + `Attribute VB_Name`    -> module-name line survives BOM strip
 *   - onFallback callback          -> fires only on the CP1252 path
 *   - isVbaFamilyFile truth table  -> extension routing, case-insensitive
 */
import { describe, it, expect } from 'vitest';
import { readVbaSource, isVbaFamilyFile } from '../src/extraction/vba-source';

/** Build a `readFile` DI that always returns the given buffer. */
function inject(buf: Buffer): { readFile: (p: string) => Buffer } {
  return { readFile: () => buf };
}

describe('readVbaSource - encoding-robust decode (Issue #53)', () => {
  it('decodes UTF-8 without a BOM byte-identically and fires no fallback', () => {
    const src = 'Attribute VB_Name = "modUtils"\nSub Foo()\nEnd Sub\n';
    const buf = Buffer.from(src, 'utf8');

    let fallbackCalls = 0;
    const result = readVbaSource('fake.bas', {
      readFile: () => buf,
      onFallback: () => { fallbackCalls++; },
    });

    expect(result.text).toBe(src);
    expect(result.bomStripped).toBe(false);
    expect(fallbackCalls).toBe(0);
  });

  it('strips a leading UTF-8 BOM (EF BB BF) and reports bomStripped=true', () => {
    const body = 'Attribute VB_Name = "modUtils"';
    const buf = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(body, 'utf8')]);

    const result = readVbaSource('fake.cls', inject(buf));

    expect(result.text.charCodeAt(0)).not.toBe(0xfeff);
    expect(result.text.startsWith('\uFEFF')).toBe(false);
    expect(result.text).toBe(body);
    expect(result.bomStripped).toBe(true);
  });

  it('falls back to windows-1252 for a CP1252 buffer with 0xF3 ("ó")', () => {
    // 0xF3 alone is an invalid UTF-8 lead byte, so the fatal decode throws and
    // the helper decodes the bytes as windows-1252 → U+00F3 ('ó').
    const buf = Buffer.concat([Buffer.from('valor = ', 'latin1'), Buffer.from([0xf3])]);

    const result = readVbaSource('fake.bas', inject(buf));

    expect(result.text.endsWith(String.fromCharCode(0xf3))).toBe(true);
    expect(result.text).toBe('valor = ó');
    // Round-trip sanity: the decoded string re-encodes to the original bytes.
    expect(Buffer.from(result.text, 'latin1').equals(buf)).toBe(true);
    expect(result.bomStripped).toBe(false);
  });

  it('returns empty text for an empty buffer', () => {
    const result = readVbaSource('fake.bas', inject(Buffer.alloc(0)));

    expect(result).toEqual({ text: '', bomStripped: false });
  });

  it('recovers the module name from a BOM + UTF-8 Attribute VB_Name line', () => {
    // Spanish-accented Access module: BOM-carrying UTF-8 whose first line is the
    // VB_Name attribute. The BOM must not corrupt the first char or the name.
    const body = 'Attribute VB_Name = "MiClase"\n\' Módulo de configuración\n';
    const buf = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(body, 'utf8')]);

    const result = readVbaSource('MiClase.cls', inject(buf));

    expect(result.text.charCodeAt(0)).not.toBe(0xfeff);
    expect(result.text.includes('"MiClase"')).toBe(true);
    expect(result.text.includes('Módulo de configuración')).toBe(true);
    expect(result.bomStripped).toBe(true);
  });

  it('invokes onFallback (with path + reason) only on the CP1252 path, never on happy UTF-8', () => {
    // CP1252 path → callback fires with the file path and a non-empty reason.
    const cp1252 = Buffer.concat([Buffer.from('Sub Situaci', 'latin1'), Buffer.from([0xf3, 0x6e])]);
    const cp1252Calls: Array<{ filePath: string; reason: string }> = [];
    readVbaSource('accented.bas', {
      readFile: () => cp1252,
      onFallback: (filePath, reason) => cp1252Calls.push({ filePath, reason }),
    });

    expect(cp1252Calls).toHaveLength(1);
    expect(cp1252Calls[0].filePath).toBe('accented.bas');
    expect(typeof cp1252Calls[0].reason).toBe('string');
    expect(cp1252Calls[0].reason.length).toBeGreaterThan(0);

    // Happy UTF-8 path → callback must NOT fire.
    const utf8 = Buffer.from('Sub Ok()\nEnd Sub\n', 'utf8');
    let happyCalls = 0;
    readVbaSource('ok.bas', {
      readFile: () => utf8,
      onFallback: () => { happyCalls++; },
    });

    expect(happyCalls).toBe(0);
  });
});

describe('isVbaFamilyFile - extension routing (Issue #53)', () => {
  it('returns true for VBA-family extensions (case-insensitive)', () => {
    for (const p of [
      'modUtils.bas',
      'MiClase.cls',
      'Form_Login.form.txt',
      'Report_Ventas.report.txt',
      'qryGetRiesgos.sql',
      'MODUTILS.BAS',
      'MiClase.CLS',
      'Form_Login.FORM.TXT',
      'qryGetRiesgos.SQL',
    ]) {
      expect(isVbaFamilyFile(p)).toBe(true);
    }
  });

  it('returns false for non-VBA extensions and near-miss names', () => {
    for (const p of [
      'index.ts',
      'app.js',
      'package.json',
      'README.md',
      'Weird.clsss',
      'MyModule.txt', // NOT `.form.txt` / `.report.txt`
      'notes.form.md',
    ]) {
      expect(isVbaFamilyFile(p)).toBe(false);
    }
  });
});
