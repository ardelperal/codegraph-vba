/**
 * VbaTestManifestExtractor — SUB-1: file detection + extractor skeleton.
 *
 * Dysflow exports VBA test manifests as `tests(.<slice>)*.json` (e.g.
 * `tests/tests.vba.smoke.json`), each listing `{ procedure, name?, tags? }`
 * entries that name a `Test_*` VBA atom. This slice adds:
 *   - `isVbaTestManifestFile(path)` detection (basename regex),
 *   - the `VbaTestManifestExtractor` skeleton: guarded JSON.parse, a
 *     content-shape gate (top-level `tests` array whose items carry a string
 *     `procedure`), and — when the shape matches — a `file` node.
 *
 * Reference emission (per-entry `vba-test-manifest` UnresolvedReferences) and
 * resolution are later slices (SUB-2 / SUB-3), so this file does not assert on
 * references — only detection, the file node, and graceful failure.
 */
import { describe, it, expect } from 'vitest';
import { VbaTestManifestExtractor } from '../src/extraction/vba-test-manifest-extractor';
import { isVbaTestManifestFile } from '../src/extraction/grammars';

function extract(filePath: string, source: string) {
  return new VbaTestManifestExtractor(filePath, source).extract();
}

describe('isVbaTestManifestFile — detection', () => {
  it('matches tests.json and dotted tests.<slice>.json (case-insensitive)', () => {
    expect(isVbaTestManifestFile('tests/tests.json')).toBe(true);
    expect(isVbaTestManifestFile('tests/tests.vba.smoke.json')).toBe(true);
    expect(isVbaTestManifestFile('a/b/tests.vba.b2-punto-15.json')).toBe(true);
    expect(isVbaTestManifestFile('TESTS.VBA.JSON')).toBe(true);
    expect(isVbaTestManifestFile('C:\\proj\\tests\\tests.vba.smoke.json')).toBe(true);
  });

  it('rejects non-manifest JSON by basename', () => {
    expect(isVbaTestManifestFile('package.json')).toBe(false);
    expect(isVbaTestManifestFile('tsconfig.json')).toBe(false);
    expect(isVbaTestManifestFile('src/mytests.json')).toBe(false);
    expect(isVbaTestManifestFile('tests.txt')).toBe(false);
    expect(isVbaTestManifestFile('queries.json')).toBe(false);
  });
});

describe('VbaTestManifestExtractor — file node', () => {
  it('emits a file node for a well-formed manifest', () => {
    const r = extract(
      'tests/tests.vba.smoke.json',
      JSON.stringify({ tests: [{ procedure: 'Test_X_RunAll' }] }),
    );
    const file = r.nodes.find((n) => n.kind === 'file');
    expect(file).toBeDefined();
    expect(file?.name).toBe('tests.vba.smoke.json');
    expect(r.errors).toHaveLength(0);
  });
});

describe('VbaTestManifestExtractor — content-shape gate', () => {
  it('emits nothing for JSON without a tests array of procedures', () => {
    // A tests.*.json that is NOT a VBA manifest (sequences shape) — out of scope.
    const r = extract(
      'tests/tests.config.json',
      JSON.stringify({ runnerPolicy: {}, procedures: ['A'] }),
    );
    expect(r.nodes).toHaveLength(0);
    expect(r.errors).toHaveLength(0);
  });

  it('emits nothing for a tests array whose items lack a string procedure', () => {
    const r = extract(
      'tests/tests.weird.json',
      JSON.stringify({ tests: [{ name: 'no procedure here' }] }),
    );
    expect(r.nodes).toHaveLength(0);
  });
});

describe('VbaTestManifestExtractor — malformed JSON', () => {
  it('records a low-severity error and never throws', () => {
    const r = extract('tests/tests.vba.broken.json', '{ "tests": [ { "procedure": ');
    expect(r.nodes).toHaveLength(0);
    expect(r.errors).toHaveLength(1);
    expect(r.errors[0]?.severity).toBe('warning');
  });
});

describe('VbaTestManifestExtractor — manifest references (SUB-2)', () => {
  it('emits one vba-test-manifest reference per entry with metadata', () => {
    const manifestFile = 'tests/tests.vba.smoke.json';
    const r = extract(
      manifestFile,
      JSON.stringify({
        tests: [
          { procedure: 'Test_A', name: 'A happy', tags: ['presenter', 'happy'] },
          { procedure: 'Test_B', name: 'B path', tags: ['b2'] },
          { procedure: 'Test_C', name: 'C path', tags: [] },
        ],
      }),
    );
    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-test-manifest',
    );
    expect(refs).toHaveLength(3);

    const fileNode = r.nodes.find((n) => n.kind === 'file');
    for (const ref of refs) {
      expect(ref.referenceKind).toBe('references');
      expect(ref.fromNodeId).toBe(fileNode?.id);
      expect(ref.metadata?.manifestFile).toBe(manifestFile);
    }

    const a = refs.find((u) => u.referenceName === 'Test_A');
    expect(a?.metadata?.testName).toBe('A happy');
    expect(a?.metadata?.tags).toEqual(['presenter', 'happy']);
  });

  it('defaults testName to the procedure and tags to [] when absent', () => {
    const r = extract(
      'tests/tests.vba.min.json',
      JSON.stringify({ tests: [{ procedure: 'Test_Y' }] }),
    );
    const ref = r.unresolvedReferences.find((u) => u.referenceName === 'Test_Y');
    expect(ref).toBeDefined();
    expect(ref?.metadata?.testName).toBe('Test_Y');
    expect(ref?.metadata?.tags).toEqual([]);
  });

  it('skips entries lacking a string procedure but keeps the valid ones', () => {
    const r = extract(
      'tests/tests.vba.mixed.json',
      JSON.stringify({
        tests: [
          { procedure: 'Test_Ok' },
          { name: 'no procedure' },
          { procedure: 123 },
        ],
      }),
    );
    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-test-manifest',
    );
    expect(refs.map((u) => u.referenceName)).toEqual(['Test_Ok']);
  });
});
