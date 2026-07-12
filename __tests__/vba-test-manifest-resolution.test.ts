/**
 * VbaTestManifestExtractor — SUB-3: resolve manifest references to Test_* nodes.
 *
 * End-to-end: index a VBA module defining a `Test_*` procedure alongside a
 * `tests.vba.*.json` manifest naming it, then assert the `ReferenceResolver`
 * bound the `vba-test-manifest` reference into a `references` edge from the
 * manifest `file` node to the existing `function` node — surfaced here via
 * `getCallers(testNode)`. A manifest naming a missing procedure must produce no
 * edge (drift stays unresolved, never a phantom node).
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import { clearProjectConfigCache } from '../src/project-config';
import CodeGraph from '../src/index';

describe('VBA test-manifest resolution (integration)', () => {
  let dir: string;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-manifest-'));
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

  it('binds a manifest reference to its Test_ procedure as a references edge', async () => {
    fs.writeFileSync(
      path.join(dir, 'TestSuite.bas'),
      ['Public Sub Test_X_RunAll()', 'End Sub'].join('\n'),
    );
    fs.mkdirSync(path.join(dir, 'tests'), { recursive: true });
    fs.writeFileSync(
      path.join(dir, 'tests', 'tests.vba.smoke.json'),
      JSON.stringify({
        tests: [{ procedure: 'Test_X_RunAll', name: 'X all', tags: ['smoke'] }],
      }),
    );

    const cg = CodeGraph.initSync(dir);
    await cg.indexAll();

    const testNode = cg
      .getNodesByName('Test_X_RunAll')
      .find((n) => n.kind === 'function');
    expect(testNode).toBeDefined();

    const callers = cg.getCallers(testNode!.id);
    const manifestEdge = callers.find(
      (c) => c.edge.metadata?.synthesizedBy === 'vba-test-manifest',
    );
    expect(manifestEdge).toBeDefined();
    expect(manifestEdge?.edge.kind).toBe('references');
    expect(manifestEdge?.node.kind).toBe('file');
    expect(manifestEdge?.edge.metadata?.tags).toEqual(['smoke']);

    await cg.destroy();
  });

  it('leaves a manifest reference to a missing procedure unresolved (no edge)', async () => {
    fs.writeFileSync(
      path.join(dir, 'TestSuite.bas'),
      ['Public Sub Test_Present()', 'End Sub'].join('\n'),
    );
    fs.mkdirSync(path.join(dir, 'tests'), { recursive: true });
    fs.writeFileSync(
      path.join(dir, 'tests', 'tests.vba.smoke.json'),
      JSON.stringify({ tests: [{ procedure: 'Test_Removed' }] }),
    );

    const cg = CodeGraph.initSync(dir);
    await cg.indexAll();

    // No function node named Test_Removed, so nothing to bind to.
    expect(cg.getNodesByName('Test_Removed')).toHaveLength(0);
    // The present procedure has no manifest edge either (manifest names the other one).
    const present = cg
      .getNodesByName('Test_Present')
      .find((n) => n.kind === 'function');
    const callers = present ? cg.getCallers(present.id) : [];
    expect(
      callers.some((c) => c.edge.metadata?.synthesizedBy === 'vba-test-manifest'),
    ).toBe(false);

    await cg.destroy();
  });
});
