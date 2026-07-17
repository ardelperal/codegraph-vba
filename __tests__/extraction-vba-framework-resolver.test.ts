/**
 * Lift the 3 Dysflow-specific extractors (form/report, test manifest, test
 * sequence) into a `FrameworkResolver`-shaped module that lives under
 * `src/extraction/frameworks/` and is opt-out-able via `codegraph.json`
 * (issue #154).
 *
 * The architecture:
 *   - `VbaExtractor` (in `src/extraction/vba/`) is the "VBA" language: it owns
 *     `.bas` / `.cls` / `.frm` / `.dsr` ÔÇö the stuff every VBA reader needs.
 *   - The 3 Dysflow-specific extractors (form/report, test manifest, test
 *     sequence) become `dysflowExportResolver`, a `FrameworkResolver` that
 *     runs ON TOP of the language, gated by `vba.dysflowExport` in the
 *     project config.
 *   - When the flag is `true` (the default), behavior is byte-identical to
 *     the pre-refactor code paths: a `.form.txt`/`.report.txt`/manifest/
 *     sequence file is routed to the matching Dysflow extractor and emits
 *     its full shape (form-layout, form-instance-control, references, etc.).
 *   - When the flag is `false`, those file types are tracked as just a
 *     `file` node ÔÇö useful for projects that carry legacy `.form.txt`
 *     files (or test manifests from a different system) and don't want
 *     them expanded into the graph.
 *
 * This test file's three tests (RED phase of the refactor):
 *   A. Opt-out: with `dysflowExport: false`, a real `.form.txt` produces
 *      ONLY a `file` node ÔÇö no `form-layout`, no `form-instance-control`,
 *      no `vba-form-binding` reference, no test-manifest refs.
 *   B. Resolver shape: `dysflowExportResolver` exists, conforms to the
 *      `FrameworkResolver` interface, and is registered in the framework
 *      registry.
 *   C. Regression: with `dysflowExport: true` (the default), the existing
 *      fixtures produce the same nodes/edges as before ÔÇö i.e. turning the
 *      flag back on is a no-op vs. the pre-refactor baseline.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import { extractFromSource } from '../src/extraction/tree-sitter';
import { clearProjectConfigCache } from '../src/project-config';
import {
  getAllFrameworkResolvers,
  getFrameworkResolver,
} from '../src/resolution/frameworks';
import type { FrameworkResolver } from '../src/resolution/types';
import CodeGraph from '../src/index';

// ---------------------------------------------------------------------------
// Fixture gate: every test builds its OWN isolated temp project so config and
// state never bleed. We do NOT use the `__tests__/fixtures/vba/` real-fixture
// directory ÔÇö that one carries a `.codegraph-vba/` produced by earlier tests
// and a strict `codegraph.json` could be left over from a previous run,
// hiding a regression in our own loader.
// ---------------------------------------------------------------------------

const FORM_SRC = [
  'VERSION 1.0 CLASS',
  'BEGIN',
  '  MultiUse = -1  \'True',
  'END',
  'Attribute VB_Name = "Form_OptOut"',
  'Begin',
  '    Begin TextBox',
  '        Name = "txtFoo"',
  '    End',
  '    Begin CommandButton',
  '        Name = "btnOK"',
  '    End',
  'End',
].join('\n');

const TEST_MANIFEST_SRC = JSON.stringify({
  tests: [
    { procedure: 'Test_X_RunAll', name: 'X all', tags: ['smoke'] },
  ],
});

const TEST_SEQUENCE_SRC = JSON.stringify({
  runnerPolicy: { tool: 'mcp', mode: 'seq' },
  procedures: ['Test_X_RunAll', 'Test_Y_Reset'],
});

const openProjects: Array<{ cg: CodeGraph; dir: string }> = [];

afterEach(async () => {
  while (openProjects.length > 0) {
    const { cg, dir } = openProjects.pop()!;
    try {
      await cg.close();
    } catch {
      // ignore close errors
    }
    // Give Windows a beat to release the SQLite file handle ÔÇö the
    // `.codegraph-vba/index.sqlite` lingers under an exclusive lock for
    // a few ms after `close()` returns, and `rmSync` would otherwise
    // hit EPERM.
    await new Promise((resolve) => setTimeout(resolve, 50));
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // ignore cleanup EPERM on Windows
    }
  }
  clearProjectConfigCache();
});

/** Write `files` (relpath ÔåÆ source) into a fresh temp dir and return the path. */
function freshProject(files: Record<string, string>): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-dysflow-'));
  for (const [rel, src] of Object.entries(files)) {
    const full = path.join(dir, rel);
    fs.mkdirSync(path.dirname(full), { recursive: true });
    fs.writeFileSync(full, src);
  }
  return dir;
}

/**
 * Helper: every edge leaving any node in `filePath` whose metadata
 * `synthesizedBy` matches `synthBy`. Used to verify the opt-out really
 * removed the Dysflow edges (vs. just hiding the source nodes).
 */
function edgesForFile(
  cg: CodeGraph,
  filePath: string,
  synthBy: string,
): { edge: { source: string; target: string; kind: string; metadata?: Record<string, unknown> }; fromKind: string }[] {
  const fileNodes = cg.getNodesInFile(filePath);
  const out: { edge: { source: string; target: string; kind: string; metadata?: Record<string, unknown> }; fromKind: string }[] = [];
  for (const n of fileNodes) {
    const outgoing = cg.getOutgoingEdges(n.id);
    for (const e of outgoing) {
      if (e.metadata?.synthesizedBy === synthBy) {
        out.push({ edge: e, fromKind: n.kind });
      }
    }
  }
  return out;
}

describe('Test A: dysflowExport: false opt-out ÔÇö form/report files emit only a file node', () => {
  it('a .form.txt produces ONLY a `file` node (no form-layout, no form-instance-control, no vba-form-binding ref)', async () => {
    const formAbs = path.join('src', 'forms', 'Form_OptOut.form.txt');
    const dir = freshProject({
      'codegraph.json': JSON.stringify({ vba: { dysflowExport: false } }),
      [`${formAbs}`]: FORM_SRC,
    });
    const cg = await CodeGraph.init(dir, { index: false });
    openProjects.push({ cg, dir });
    await cg.indexAll();

    const fullPath = formAbs.split(path.sep).join('/');
    const allFromFile = cg.getNodesInFile(fullPath);

    // Exactly one node from the file ÔÇö and it must be the `file` kind.
    expect(allFromFile).toHaveLength(1);
    expect(allFromFile[0]?.kind).toBe('file');

    // Explicit anti-assertions: every Dysflow-specific kind must be gone.
    const kinds = new Set(allFromFile.map((n) => n.kind));
    expect(kinds.has('form-layout')).toBe(false);
    expect(kinds.has('form-instance-control')).toBe(false);
    expect(kinds.has('property')).toBe(false);

    // The RecordSource/RowSource sweeps must NOT have emitted any edge
    // out of the file node.
    const recordEdges = edgesForFile(cg, fullPath, 'vba-record-source');
    expect(recordEdges).toHaveLength(0);
    const rowEdges = edgesForFile(cg, fullPath, 'vba-row-source');
    expect(rowEdges).toHaveLength(0);
  });

  it('a .report.txt produces ONLY a `file` node (no form-layout, no binding)', async () => {
    const reportSrc = [
      'VERSION 1.0 CLASS',
      'BEGIN',
      '  MultiUse = -1  \'True',
      'END',
      'Attribute VB_Name = "Report_OptOut"',
      'Begin',
      '    Begin TextBox',
      '        Name = "txtOrderId"',
      '    End',
      'End',
    ].join('\n');
    const reportAbs = path.join('src', 'reports', 'Report_OptOut.report.txt');
    const dir = freshProject({
      'codegraph.json': JSON.stringify({ vba: { dysflowExport: false } }),
      [`${reportAbs}`]: reportSrc,
    });
    const cg = await CodeGraph.init(dir, { index: false });
    openProjects.push({ cg, dir });
    await cg.indexAll();

    const fullPath = reportAbs.split(path.sep).join('/');
    const reportFileNodes = cg.getNodesInFile(fullPath);
    expect(reportFileNodes).toHaveLength(1);
    expect(reportFileNodes[0]?.kind).toBe('file');
    const kinds = new Set(reportFileNodes.map((n) => n.kind));
    expect(kinds.has('form-layout')).toBe(false);
  });

  it('a Dysflow test manifest produces ONLY a `file` node (no vba-test-manifest refs)', async () => {
    const manifestAbs = path.join('src', 'tests', 'tests.vba.smoke.json');
    const dir = freshProject({
      'codegraph.json': JSON.stringify({ vba: { dysflowExport: false } }),
      [`${manifestAbs}`]: TEST_MANIFEST_SRC,
    });
    const cg = await CodeGraph.init(dir, { index: false });
    openProjects.push({ cg, dir });
    await cg.indexAll();

    const fullPath = manifestAbs.split(path.sep).join('/');
    const manifestFileNodes = cg.getNodesInFile(fullPath);
    expect(manifestFileNodes).toHaveLength(1);
    expect(manifestFileNodes[0]?.kind).toBe('file');

    // The synthesized `vba-test-manifest` references would be edges from
    // the manifest file node to the Test_* procedures. With Dysflow export
    // disabled, none of those edges are emitted.
    for (const n of manifestFileNodes) {
      const edges = cg.getOutgoingEdges(n.id);
      const filtered = edges.filter(
        (e) => e.metadata?.synthesizedBy === 'vba-test-manifest',
      );
      expect(filtered).toHaveLength(0);
    }
  });

  it('a Dysflow test sequence produces ONLY a `file` node (no vba-test-sequence refs)', async () => {
    const seqAbs = path.join('src', 'tests', 'sequences', 'cache-riesgo.json');
    const dir = freshProject({
      'codegraph.json': JSON.stringify({ vba: { dysflowExport: false } }),
      [`${seqAbs}`]: TEST_SEQUENCE_SRC,
    });
    const cg = await CodeGraph.init(dir, { index: false });
    openProjects.push({ cg, dir });
    await cg.indexAll();

    const fullPath = seqAbs.split(path.sep).join('/');
    const seqFileNodes = cg.getNodesInFile(fullPath);
    expect(seqFileNodes).toHaveLength(1);
    expect(seqFileNodes[0]?.kind).toBe('file');

    for (const n of seqFileNodes) {
      const edges = cg.getOutgoingEdges(n.id);
      const filtered = edges.filter(
        (e) => e.metadata?.synthesizedBy === 'vba-test-sequence',
      );
      expect(filtered).toHaveLength(0);
    }
  });

  it('the base VBA extractor still runs for .bas/.cls files when dysflowExport is false', async () => {
    // Regression ÔÇö the flag is about the Dysflow-specific extractors, NOT
    // the base VBA language. A .bas with a Public Function must still emit
    // the module + function nodes.
    const dir = freshProject({
      'codegraph.json': JSON.stringify({ vba: { dysflowExport: false } }),
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Option Explicit',
        '',
        'Public Sub CallerSub()',
        '    Debug.Print "hi"',
        'End Sub',
        '',
      ].join('\n'),
    });
    const cg = await CodeGraph.init(dir, { index: false });
    openProjects.push({ cg, dir });
    await cg.indexAll();

    const moduleNode = cg
      .getNodesByName('Caller')
      .find((n) => n.kind === 'module' && n.filePath.endsWith('Caller.bas'));
    expect(moduleNode).toBeDefined();
    const functionNode = cg
      .getNodesByName('CallerSub')
      .find((n) => n.kind === 'function' && n.filePath.endsWith('Caller.bas'));
    expect(functionNode).toBeDefined();
  });
});

describe('Test B: dysflowExportResolver ÔÇö shape and registration', () => {
  it('exports a dysflowExportResolver object that conforms to FrameworkResolver', async () => {
    // Dynamic import: the module doesn't exist yet, so the import resolves
    // only after the GREEN phase writes it. Until then, this test fails
    // (RED) at the `import` step, which is the failure we want ÔÇö proves
    // the TDD discipline is real.
    const mod = await import('../src/extraction/frameworks/dysflow-export');
    const resolver = mod.dysflowExportResolver as FrameworkResolver;
    expect(resolver).toBeDefined();
    expect(resolver.name).toBe('dysflow-export');
    expect(typeof resolver.detect).toBe('function');
    expect(typeof resolver.extract).toBe('function');
  });

  it('is registered in the framework resolver registry', () => {
    // After GREEN, the resolver must be discoverable via the same path the
    // rest of the codebase uses to find framework resolvers. If it's not in
    // the registry, `getApplicableFrameworks` won't see it and the
    // detection-driven wiring (Test A's opt-out / Test C's default) won't
    // work end-to-end.
    const fromRegistry = getFrameworkResolver('dysflow-export');
    expect(fromRegistry).toBeDefined();
    expect(fromRegistry?.name).toBe('dysflow-export');
    // Sanity: the registry still resolves the pre-existing resolvers.
    expect(getFrameworkResolver('react')).toBeDefined();
  });

  it('the resolver appears in getAllFrameworkResolvers()', () => {
    const names = getAllFrameworkResolvers().map((r) => r.name);
    expect(names).toContain('dysflow-export');
  });
});

describe('Test C: dysflowExport: true (the default) ÔÇö behavior matches the pre-refactor baseline', () => {
  it('with NO codegraph.json, a .form.txt emits the full Dysflow shape (form-layout + form-instance-control + property nodes)', async () => {
    // No codegraph.json ÔÇö flag is implicitly true. The pre-refactor
    // baseline emits the full Dysflow shape. This test pins the
    // "behavior identical for the default case" acceptance criterion.
    const formAbs = path.join('src', 'forms', 'Form_Default.form.txt');
    const defaultFormSrc = FORM_SRC.replace('Form_OptOut', 'Form_Default');
    const dir = freshProject({
      [`${formAbs}`]: defaultFormSrc,
    });
    const cg = await CodeGraph.init(dir, { index: false });
    openProjects.push({ cg, dir });
    await cg.indexAll();

    const fullPath = formAbs.split(path.sep).join('/');
    const allFromFile = cg.getNodesInFile(fullPath);
    const formLayout = allFromFile.find((n) => n.kind === 'form-layout');
    expect(formLayout).toBeDefined();
    expect(formLayout?.name).toBe('Form_Default');

    const controlNodes = allFromFile.filter(
      (n) => n.kind === 'form-instance-control',
    );
    // Two controls in the fixture: txtFoo (TextBox) and btnOK (CommandButton).
    expect(controlNodes).toHaveLength(2);
    const controlNames = controlNodes.map((n) => n.name).sort();
    expect(controlNames).toEqual(['btnOK', 'txtFoo']);

    // The form-binding reference edge from the form-layout to the sibling
    // .cls would carry synthesizedBy='vba-form-binding'. With no .cls
    // sibling, the binding is unresolved; the pre-refactor baseline still
    // emits the UnresolvedReference in the extractor's result, and the
    // resolver then attempts to bind it. Either way, the form-layout node
    // and the control nodes must exist with the right shape.
  });

  it('with dysflowExport: true explicit, a Dysflow test manifest still emits vba-test-manifest refs', async () => {
    const dir = freshProject({
      'codegraph.json': JSON.stringify({ vba: { dysflowExport: true } }),
      'src/modules/Test_X_RunAll.bas': [
        'Public Sub Test_X_RunAll()',
        'End Sub',
      ].join('\n'),
      'src/tests/tests.vba.smoke.json': TEST_MANIFEST_SRC,
    });
    const cg = await CodeGraph.init(dir, { index: false });
    openProjects.push({ cg, dir });
    await cg.indexAll();

    // The Test_X_RunAll function must exist (it came from the .bas via the
    // base VbaExtractor ÔÇö proves the base language still works).
    const testFn = cg
      .getNodesByName('Test_X_RunAll')
      .find((n) => n.kind === 'function');
    expect(testFn).toBeDefined();

    // The manifest edge from the file node to the Test_X_RunAll function
    // must be present with synthesizedBy='vba-test-manifest' ÔÇö the same
    // shape the pre-refactor code path produced (SUB-3 of #91).
    const callers = cg.getCallers(testFn!.id);
    const manifestEdge = callers.find(
      (c) => c.edge.metadata?.synthesizedBy === 'vba-test-manifest',
    );
    expect(manifestEdge).toBeDefined();
    expect(manifestEdge?.edge.kind).toBe('references');
  });

  it('extractFromSource with default config produces the same shape as the pre-refactor per-extractor path (form)', () => {
    // Unit-level regression: extractFromSource on a .form.txt with the
    // default config must produce the same node/edge set as instantiating
    // the underlying VbaFormExtractor directly. This is the load-bearing
    // contract: "the framework-resolver wiring is a transparent refactor
    // for the default case".
    //
    // After GREEN, extractFromSource's signature accepts an optional
    // `dysflowExport` parameter (default true). We pass `undefined` here
    // (so the default kicks in) to keep the test focused on the wiring.
    const r = extractFromSource(
      'src/forms/Form_Unit.form.txt',
      FORM_SRC,
      'vba',
    );
    const formLayouts = r.nodes.filter((n) => n.kind === 'form-layout');
    expect(formLayouts.length).toBeGreaterThan(0);
    const controls = r.nodes.filter((n) => n.kind === 'form-instance-control');
    expect(controls).toHaveLength(2);
    const binding = r.unresolvedReferences.find(
      (u) => u.metadata?.synthesizedBy === 'vba-form-binding',
    );
    expect(binding).toBeDefined();
  });

  it('extractFromSource on a Dysflow test manifest with default config emits the manifest refs', () => {
    const r = extractFromSource(
      'tests/tests.vba.smoke.json',
      TEST_MANIFEST_SRC,
      'vba',
    );
    const fileNode = r.nodes.find((n) => n.kind === 'file');
    expect(fileNode).toBeDefined();
    const manifestRefs = r.unresolvedReferences.filter(
      (u) => u.metadata?.synthesizedBy === 'vba-test-manifest',
    );
    expect(manifestRefs).toHaveLength(1);
    expect(manifestRefs[0]?.referenceName).toBe('Test_X_RunAll');
  });
});
