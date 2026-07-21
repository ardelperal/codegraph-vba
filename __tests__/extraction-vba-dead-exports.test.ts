import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const readSource = (relativePath: string): string =>
  readFileSync(new URL(`../src/${relativePath}`, import.meta.url), 'utf8');

describe('VBA extraction public surface', () => {
  it('keeps the split-once classifiers free of legacy source wrappers', () => {
    const wrappers = [
      ['extraction/vba/procedures.ts', 'sweepProcedures'],
      ['extraction/vba/dims.ts', 'sweepDimsAndWithEvents'],
      ['extraction/vba/declarations.ts', 'sweepEventsTypesAndDeclares'],
      ['extraction/vba/implements.ts', 'sweepImplements'],
      ['extraction/vba/enums-consts.ts', 'sweepEnumsAndConsts'],
      ['extraction/vba/call-sweep.ts', 'sweepCallsAndSql'],
    ] as const;

    for (const [path, wrapper] of wrappers) {
      const source = readSource(path);
      expect(source, path).not.toContain(`function ${wrapper}`);
      expect(source, path).not.toContain("src.split('\\n')");
    }
  });

  it('removes or narrows internal-only exports while preserving RUNTIME_OBJECTS', () => {
    expect(readSource('extraction/vba-source.ts')).not.toContain('function stripUtf8Bom');
    expect(readSource('extraction/vba/text-utils.ts')).not.toContain(
      'export function splitOutsideVbaStrings',
    );
    expect(readSource('extraction/vba-test-manifest-extractor.ts')).not.toContain(
      'export function isVbaTestManifestShape',
    );

    const runtimeObjects = readSource('resolution/vba-runtime-objects.ts');
    expect(runtimeObjects).not.toContain('export const VBA_STDLIB_FUNCTIONS');
    expect(runtimeObjects).toContain('export const RUNTIME_OBJECTS');
  });
});
