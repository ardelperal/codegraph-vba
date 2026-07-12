/**
 * VbaTestSequenceExtractor — SUB-6 of epic #91 (issue #97):
 * parse Dysflow `tests/sequences/*.json` files carrying the
 * `runnerPolicy + procedures[]` shape.
 *
 * The Dysflow orchestrator groups manifest-bound test atoms into named
 * sequences under `tests/sequences/` (e.g. `cache-riesgo.json`); each
 * sequence carries:
 *
 *   { "runnerPolicy": { tool, mode, sequential, ... },
 *     "procedures":  ["Test_Cache_Riesgo_RunSlice", ...] }
 *
 * This extractor emits ONE `file` node + ONE `UnresolvedReference` per
 * `procedures[]` entry with metadata `{ synthesizedBy: 'vba-test-sequence',
 * runnerPolicy, sequenceFile, procedureIndex }`. The reference is the same
 * `references` shape SUB-3 already binds, so SUB-6 just extends the resolver
 * to also pick up `synthesizedBy: 'vba-test-sequence'` and carry the new
 * metadata through.
 *
 * SCOPE (deliberate):
 *   - ONLY the `runnerPolicy + procedures[]` shape. Other Dysflow sequence
 *     shapes — strict-sequence (`executionUnits`) and slices (`slices[]`
 *     with submanifests) — are documented as future work and must NOT be
 *     implemented here.
 *
 * Path detection (separate gate):
 *   - `isVbaTestSequenceFile(p)` is true iff `p` lives under a `sequences/`
 *     directory AND ends in `.json`. The basename does not have to start
 *     with `tests` (a `tests/sequences/cache-riesgo.json` is in scope; a
 *     `package.json` is NOT).
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import {
  VbaTestSequenceExtractor,
  isVbaTestSequenceShape,
} from '../src/extraction/vba-test-sequence-extractor';
import { isVbaTestSequenceFile } from '../src/extraction/grammars';
import { clearProjectConfigCache } from '../src/project-config';
import CodeGraph from '../src/index';

function extract(filePath: string, source: string) {
  return new VbaTestSequenceExtractor(filePath, source).extract();
}

describe('isVbaTestSequenceFile — path detection', () => {
  it('matches any .json file under a sequences/ directory', () => {
    expect(isVbaTestSequenceFile('tests/sequences/cache-riesgo.json')).toBe(true);
    expect(isVbaTestSequenceFile('tests/sequences/nested/deeper/x.json')).toBe(true);
    expect(isVbaTestSequenceFile('C:\\proj\\tests\\sequences\\cache-riesgo.json')).toBe(true);
    expect(isVbaTestSequenceFile('tests/sequences/cache-riesgo.JSON')).toBe(true);
    expect(isVbaTestSequenceFile('a/b/sequences/foo.json')).toBe(true);
  });

  it('rejects files that are not under a sequences/ directory', () => {
    expect(isVbaTestSequenceFile('tests/cache-riesgo.json')).toBe(false);
    expect(isVbaTestSequenceFile('tests/tests.vba.smoke.json')).toBe(false);
    expect(isVbaTestSequenceFile('package.json')).toBe(false);
    expect(isVbaTestSequenceFile('tsconfig.json')).toBe(false);
    expect(isVbaTestSequenceFile('sequences/foo.txt')).toBe(false);
    expect(isVbaTestSequenceFile('')).toBe(false);
  });
});

describe('isVbaTestSequenceShape — content-shape gate', () => {
  it('accepts { runnerPolicy: object, procedures: string[] }', () => {
    expect(
      isVbaTestSequenceShape({
        runnerPolicy: { tool: 'test_vba' },
        procedures: ['Test_X'],
      }),
    ).toBe(true);
  });

  it('accepts an empty description / arbitrary other keys alongside the gate', () => {
    expect(
      isVbaTestSequenceShape({
        description: 'anything',
        runnerPolicy: { tool: 'test_vba', mode: 'procedureName' },
        procedures: ['Test_A', 'Test_B'],
      }),
    ).toBe(true);
  });

  it('rejects missing procedures', () => {
    expect(isVbaTestSequenceShape({ runnerPolicy: { tool: 'test_vba' } })).toBe(false);
  });

  it('accepts an empty procedures array (extractor still emits the file node)', () => {
    // An empty sequence still passes the shape gate — the extractor emits
    // a file node and zero references, which keeps empty plans visible in
    // the graph without leaving orphan `references` edges.
    expect(isVbaTestSequenceShape({ runnerPolicy: {}, procedures: [] })).toBe(true);
  });

  it('rejects missing runnerPolicy', () => {
    expect(isVbaTestSequenceShape({ procedures: ['Test_X'] })).toBe(false);
  });

  it('rejects wrong types (runnerPolicy not object, procedures not array of strings)', () => {
    expect(isVbaTestSequenceShape({ runnerPolicy: 'oops', procedures: ['Test_X'] })).toBe(false);
    expect(isVbaTestSequenceShape({ runnerPolicy: null, procedures: ['Test_X'] })).toBe(false);
    expect(isVbaTestSequenceShape({ runnerPolicy: {}, procedures: 'Test_X' })).toBe(false);
    expect(
      isVbaTestSequenceShape({ runnerPolicy: {}, procedures: ['Test_X', 123] }),
    ).toBe(false);
    expect(isVbaTestSequenceShape({ runnerPolicy: {}, procedures: [null] })).toBe(false);
    expect(isVbaTestSequenceShape(null)).toBe(false);
    expect(isVbaTestSequenceShape(undefined)).toBe(false);
    expect(isVbaTestSequenceShape('a string')).toBe(false);
  });
});

describe('VbaTestSequenceExtractor — file node + procedure references', () => {
  it('emits one file node and one reference per procedures[] entry with metadata', () => {
    const r = extract(
      'tests/sequences/cache-riesgo.json',
      JSON.stringify({
        description: 'plan ejecutable por orquestador MCP',
        runnerPolicy: {
          tool: 'test_vba',
          mode: 'procedureName',
          sequential: true,
          stopOnFirstFunctionalFailure: true,
        },
        procedures: ['Test_Cache_Riesgo_RunSlice', 'Test_Cache_Riesgo_ResetSlice'],
      }),
    );

    expect(r.errors).toHaveLength(0);
    const file = r.nodes.find((n) => n.kind === 'file');
    expect(file).toBeDefined();
    expect(file?.name).toBe('cache-riesgo.json');

    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-test-sequence',
    );
    expect(refs).toHaveLength(2);

    for (const ref of refs) {
      expect(ref.referenceKind).toBe('references');
      expect(ref.fromNodeId).toBe(file?.id);
      expect(ref.language).toBe('vba');
      expect(ref.metadata?.sequenceFile).toBe('tests/sequences/cache-riesgo.json');
      expect(ref.metadata?.runnerPolicy).toEqual({
        tool: 'test_vba',
        mode: 'procedureName',
        sequential: true,
        stopOnFirstFunctionalFailure: true,
      });
    }

    expect(refs[0]?.referenceName).toBe('Test_Cache_Riesgo_RunSlice');
    expect(refs[0]?.metadata?.procedureIndex).toBe(0);
    expect(refs[1]?.referenceName).toBe('Test_Cache_Riesgo_ResetSlice');
    expect(refs[1]?.metadata?.procedureIndex).toBe(1);
  });

  it('emits only a file node when procedures is empty', () => {
    const r = extract(
      'tests/sequences/empty.json',
      JSON.stringify({ runnerPolicy: {}, procedures: [] }),
    );
    const file = r.nodes.find((n) => n.kind === 'file');
    expect(file).toBeDefined();
    const refs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-test-sequence',
    );
    expect(refs).toHaveLength(0);
    expect(r.errors).toHaveLength(0);
  });
});

describe('VbaTestSequenceExtractor — malformed JSON', () => {
  it('records a low-severity error and emits nothing (no throw)', () => {
    const r = extract('tests/sequences/broken.json', '{ "runnerPolicy": ');
    expect(r.nodes).toHaveLength(0);
    expect(r.unresolvedReferences).toHaveLength(0);
    expect(r.errors).toHaveLength(1);
    expect(r.errors[0]?.severity).toBe('warning');
  });
});

describe('VbaTestSequenceExtractor — deferred shapes emit nothing, no error', () => {
  it('rejects strict-sequence (has executionUnits) — emits nothing', () => {
    const r = extract(
      'tests/sequences/strict.json',
      JSON.stringify({
        runnerPolicy: { sequential: true },
        executionUnits: [
          'tests/sequences/cache-riesgo.json',
          'tests/tests.vba.cache-edicion.json',
        ],
      }),
    );
    expect(r.nodes).toHaveLength(0);
    expect(r.unresolvedReferences).toHaveLength(0);
    expect(r.errors).toHaveLength(0);
  });

  it('rejects slices (has slices[]) — emits nothing', () => {
    const r = extract(
      'tests/sequences/slices.json',
      JSON.stringify({
        version: 1,
        sourceManifest: 'tests/tests.vba.json',
        slices: [{ name: 'cache', filter: 'cache', purpose: 'baseline cache' }],
      }),
    );
    expect(r.nodes).toHaveLength(0);
    expect(r.unresolvedReferences).toHaveLength(0);
    expect(r.errors).toHaveLength(0);
  });
});

describe('VbaTestSequenceExtractor — resolver integration', () => {
  let dir: string;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-seq-'));
    clearProjectConfigCache();
  });

  afterEach(async () => {
    clearProjectConfigCache();
    await new Promise((resolve) => setTimeout(resolve, 100));
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // ignore cleanup EPERM on Windows
    }
  });

  it('binds a sequence reference to its Test_ procedure as a references edge with metadata', async () => {
    fs.writeFileSync(
      path.join(dir, 'TestCacheRiesgo.bas'),
      [
        'Public Sub Test_Cache_Riesgo_RunSlice()',
        'End Sub',
        'Public Sub Test_Cache_Riesgo_ResetSlice()',
        'End Sub',
      ].join('\n'),
    );
    fs.mkdirSync(path.join(dir, 'tests', 'sequences'), { recursive: true });
    const sequenceFile = path.join(dir, 'tests', 'sequences', 'cache-riesgo.json');
    fs.writeFileSync(
      sequenceFile,
      JSON.stringify({
        runnerPolicy: { tool: 'test_vba', mode: 'procedureName', sequential: true },
        procedures: ['Test_Cache_Riesgo_RunSlice', 'Test_Cache_Riesgo_ResetSlice'],
      }),
    );

    const cg = CodeGraph.initSync(dir);
    await cg.indexAll();

    const target = cg
      .getNodesByName('Test_Cache_Riesgo_RunSlice')
      .find((n) => n.kind === 'function');
    expect(target).toBeDefined();

    const callers = cg.getCallers(target!.id);
    const seqEdge = callers.find(
      (c) => c.edge.metadata?.synthesizedBy === 'vba-test-sequence',
    );
    expect(seqEdge).toBeDefined();
    expect(seqEdge?.edge.kind).toBe('references');
    expect(seqEdge?.node.kind).toBe('file');
    expect(seqEdge?.node.name).toBe('cache-riesgo.json');
    expect(seqEdge?.edge.metadata?.procedureIndex).toBe(0);
    expect((seqEdge?.edge.metadata?.runnerPolicy as { tool?: string } | undefined)?.tool).toBe(
      'test_vba',
    );

    // Manifest edges from SUB-3 must NOT collide with the sequence edge:
    // a manifest carries `synthesizedBy: 'vba-test-manifest'`, this one
    // carries `'vba-test-sequence'`.
    expect(
      callers.some((c) => c.edge.metadata?.synthesizedBy === 'vba-test-manifest'),
    ).toBe(false);

    await cg.destroy();
  });

  it('a sequence naming a missing procedure leaves the reference unresolved (no phantom edge)', async () => {
    fs.writeFileSync(
      path.join(dir, 'TestCacheRiesgo.bas'),
      ['Public Sub Test_Cache_Riesgo_RunSlice()', 'End Sub'].join('\n'),
    );
    fs.mkdirSync(path.join(dir, 'tests', 'sequences'), { recursive: true });
    fs.writeFileSync(
      path.join(dir, 'tests', 'sequences', 'cache-riesgo.json'),
      JSON.stringify({
        runnerPolicy: { tool: 'test_vba' },
        procedures: ['Test_DoesNotExist'],
      }),
    );

    const cg = CodeGraph.initSync(dir);
    await cg.indexAll();

    expect(cg.getNodesByName('Test_DoesNotExist')).toHaveLength(0);
    const present = cg
      .getNodesByName('Test_Cache_Riesgo_RunSlice')
      .find((n) => n.kind === 'function');
    const callers = present ? cg.getCallers(present.id) : [];
    expect(
      callers.some((c) => c.edge.metadata?.synthesizedBy === 'vba-test-sequence'),
    ).toBe(false);

    await cg.destroy();
  });
});