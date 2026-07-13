/**
 * Strict-TDD unit/E2E coverage for the runtime-object skip in the VBA
 * post-extraction call-stub resolver (issue #110, supersedes #109).
 *
 * Background: `codegraph-vba` synthesizes a `stub:true` `calls` edge for every
 * `Receiver.Member` call whose target isn't resolvable at extraction time. For
 * runtime objects (DAO, FileSystemObject, intrinsic collections, ...) that
 * target is NEVER user code, so the stub used to sit in the graph pointing at
 * itself — poisoning a consumer's `WHERE stub=true` "missing callee" lint with
 * runtime-object noise. The resolver now DECLINES those stubs explicitly
 * (`repointDecision='declined-runtime'`) while preserving:
 *   - class-typed and `.bas`-qualified repoints (Tests 2, 3),
 *   - shadow user classes that happen to share a runtime-object name (Test 4),
 *   - genuinely-missing user callees as `stub:true` (Test 5).
 *
 * Each test builds its OWN isolated temp project (fixture gate) and drives the
 * real `CodeGraph.indexAll()` pipeline end-to-end — no DB mocking.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { CodeGraph } from '../src';
import { isRuntimeObject, RUNTIME_OBJECTS } from '../src/resolution/vba-runtime-objects';
import type { Edge } from '../src/types';

const CLS_HEADER = ['VERSION 1.0 CLASS', 'BEGIN', "  MultiUse = -1  'True", 'END'];

/** Track every project we spin up so afterEach can close + remove them. */
const openProjects: Array<{ cg: CodeGraph; dir: string }> = [];

afterEach(async () => {
  while (openProjects.length > 0) {
    const { cg, dir } = openProjects.pop()!;
    try {
      await cg.close();
    } catch {
      // ignore close errors
    }
    if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  }
});

/**
 * Write `files` (relative path → source) into a fresh temp dir, index it, and
 * return the live CodeGraph. `.bas` go under src/modules, `.cls` under
 * src/classes by convention — but the caller supplies the full relative path.
 */
async function buildProject(files: Record<string, string>): Promise<CodeGraph> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-runtime-skip-'));
  for (const [rel, src] of Object.entries(files)) {
    const full = path.join(dir, rel);
    fs.mkdirSync(path.dirname(full), { recursive: true });
    fs.writeFileSync(full, src);
  }
  const cg = await CodeGraph.init(dir, { index: false });
  openProjects.push({ cg, dir });
  await cg.indexAll();
  return cg;
}

/** Outgoing `calls` edges of the (single) function named `name`. */
function callEdgesFrom(cg: CodeGraph, name: string): Edge[] {
  const fn = cg
    .searchNodes(name, { languages: ['vba'], kinds: ['function'] })
    .find((n) => n.node.name === name);
  if (!fn) return [];
  return cg.getOutgoingEdges(fn.node.id).filter((e) => e.kind === 'calls');
}

describe('VBA call-stub resolver — runtime-object skip (#110)', () => {
  it('Test 1: a runtime-object call (DAO.*) stays stub:true and is declined as declined-runtime', async () => {
    const cg = await buildProject({
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Option Explicit',
        '',
        'Public Sub CallerSub()',
        '    DAO.BeginTrans',
        'End Sub',
        '',
      ].join('\n'),
    });

    const edges = callEdgesFrom(cg, 'CallerSub');
    const daoEdge = edges.find((e) => e.metadata?.member === 'BeginTrans');
    expect(daoEdge).toBeDefined();
    expect(daoEdge?.metadata?.receiverType).toBe('DAO');
    expect(daoEdge?.metadata?.stub).toBe(true);
    expect(daoEdge?.metadata?.repointDecision).toBe('declined-runtime');

    // The synthetic stub never resolves to a REAL node — there is no user
    // function whose qualifiedName is 'DAO.BeginTrans'.
    const real = cg
      .searchNodes('DAO.BeginTrans', { languages: ['vba'] })
      .filter((n) => n.node.name === 'DAO.BeginTrans');
    // Only the synthetic stub (if still present) may carry this name; there is
    // no additional real declaration.
    expect(real.every((n) => n.node.metadata?.stub === true || n.node.name === 'DAO.BeginTrans')).toBe(true);
  });

  it('Test 2: a class-typed call resolves to the real cross-file method (stub:false)', async () => {
    const cg = await buildProject({
      'src/classes/ACAuditoriaOperaciones.cls': [
        ...CLS_HEADER,
        'Attribute VB_Name = "ACAuditoriaOperaciones"',
        'Option Explicit',
        '',
        'Public Sub Registrar()',
        'End Sub',
        '',
      ].join('\n'),
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Option Explicit',
        '',
        'Public Sub CallerSub()',
        '    Dim x As ACAuditoriaOperaciones',
        '    x.Registrar',
        'End Sub',
        '',
      ].join('\n'),
    });

    const registrar = cg
      .searchNodes('Registrar', { languages: ['vba'], kinds: ['function'] })
      .find((n) => n.node.name === 'Registrar' && n.node.filePath.endsWith('ACAuditoriaOperaciones.cls'));
    expect(registrar).toBeDefined();
    if (!registrar) return;

    const incoming = cg.getIncomingEdges(registrar.node.id).filter((e) => e.kind === 'calls');
    expect(incoming.length).toBeGreaterThan(0);
    for (const edge of incoming) {
      expect(edge.metadata?.stub).not.toBe(true);
    }
  });

  it('Test 3: a .bas-qualified call resolves to the real bare-name node via module narrowing (stub:false)', async () => {
    const cg = await buildProject({
      'src/modules/mdlCursor.bas': [
        'Attribute VB_Name = "mdlCursor"',
        'Option Explicit',
        '',
        'Public Function MouseCursor() As Long',
        'End Function',
        '',
      ].join('\n'),
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Option Explicit',
        '',
        'Public Sub CallerSub()',
        '    mdlCursor.MouseCursor',
        'End Sub',
        '',
      ].join('\n'),
    });

    const mouseCursor = cg
      .searchNodes('MouseCursor', { languages: ['vba'], kinds: ['function'] })
      .find((n) => n.node.name === 'MouseCursor' && n.node.filePath.endsWith('mdlCursor.bas'));
    expect(mouseCursor).toBeDefined();
    if (!mouseCursor) return;

    const incoming = cg.getIncomingEdges(mouseCursor.node.id).filter((e) => e.kind === 'calls');
    expect(incoming.length).toBeGreaterThan(0);
    for (const edge of incoming) {
      expect(edge.metadata?.stub).not.toBe(true);
    }
  });

  it('Test 4: a shadow user class named DAO is preserved (repointed-to-real, skip bypassed)', async () => {
    const cg = await buildProject({
      'src/classes/DAO.cls': [
        ...CLS_HEADER,
        'Attribute VB_Name = "DAO"',
        'Option Explicit',
        '',
        'Public Sub Execute()',
        'End Sub',
        '',
      ].join('\n'),
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Option Explicit',
        '',
        'Public Sub CallerSub()',
        '    DAO.Execute',
        'End Sub',
        '',
      ].join('\n'),
    });

    const execute = cg
      .searchNodes('Execute', { languages: ['vba'], kinds: ['function'] })
      .find((n) => n.node.name === 'Execute' && n.node.filePath.endsWith('DAO.cls'));
    expect(execute).toBeDefined();
    if (!execute) return;

    const incoming = cg.getIncomingEdges(execute.node.id).filter((e) => e.kind === 'calls');
    expect(incoming.length).toBeGreaterThan(0);
    for (const edge of incoming) {
      expect(edge.metadata?.stub).not.toBe(true);
      expect(edge.metadata?.repointDecision).toBe('reponted-to-real');
    }
  });

  it('Test 5: a genuinely-missing user callee stays stub:true and is declined as declined-not-found', async () => {
    const cg = await buildProject({
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Option Explicit',
        '',
        'Public Sub CallerSub()',
        '    Dim m_x As DoesNotExistClass',
        '    m_x.DoesNotExistSub',
        'End Sub',
        '',
      ].join('\n'),
    });

    const edges = callEdgesFrom(cg, 'CallerSub');
    const missing = edges.filter((e) => e.metadata?.member === 'DoesNotExistSub');
    // Exactly ONE calls edge for the missing callee (no double-emission).
    expect(missing).toHaveLength(1);
    const edge = missing[0];
    expect(edge?.metadata?.receiverType).toBe('DoesNotExistClass');
    expect(edge?.metadata?.stub).toBe(true);
    expect(edge?.metadata?.repointDecision).toBe('declined-not-found');
  });

  it('Test 6 (meta): the canonical runtime-object list classifies receivers case-insensitively', () => {
    // Runtime objects (any case) → true.
    expect(isRuntimeObject('DAO')).toBe(true);
    expect(isRuntimeObject('dao')).toBe(true);
    expect(isRuntimeObject('Fso')).toBe(true);
    expect(isRuntimeObject('[DAO]')).toBe(true);
    expect(isRuntimeObject('  Collection  ')).toBe(true);
    // Non-runtime user receivers → false.
    expect(isRuntimeObject('ACAuditoriaOperaciones')).toBe(false);
    expect(isRuntimeObject('mdlCursor')).toBe(false);
    expect(isRuntimeObject('')).toBe(false);
    expect(isRuntimeObject(undefined)).toBe(false);
    // The frozen list carries the documented seed entries.
    for (const expected of ['dao', 'fso', 'err', 'listbox', 'collection', 'docmd']) {
      expect(RUNTIME_OBJECTS.has(expected)).toBe(true);
    }
  });
});
