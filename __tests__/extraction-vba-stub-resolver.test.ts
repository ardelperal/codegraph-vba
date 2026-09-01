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
import { CodeGraph, DatabaseConnection, getDatabasePath } from '../src';
import {
  isRuntimeObject,
  isVbaStdlibFunction,
  RUNTIME_OBJECTS,
} from '../src/resolution/vba-runtime-objects';
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
  it('Test 1 (updated by #245): a runtime-object call (DAO.*) never becomes a node at all', async () => {
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

    // Before #245 this call synthesized a `DAO.BeginTrans` function node and
    // a `calls` edge to it, which the resolver then stamped
    // `repointDecision: 'declined-runtime'`. Edge consumers could filter on
    // that field; symbol search and node counts could not. The extractor now
    // refuses to create the node in the first place, so the whole
    // (node, edge, decision) triple is gone.
    const edges = callEdgesFrom(cg, 'CallerSub');
    expect(edges.find((e) => e.metadata?.member === 'BeginTrans')).toBeUndefined();

    const real = cg
      .searchNodes('DAO.BeginTrans', { languages: ['vba'] })
      .filter((n) => n.node.name === 'DAO.BeginTrans');
    expect(real).toHaveLength(0);
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

  it('Test 4 (updated by #245): a shadow user class named after a runtime object is no longer call-linked', async () => {
    // The resolver's shadow bypass (FR-2.1) can only fire on a stub that
    // reached it. VBA extraction is per-file: when `Caller.bas` is parsed
    // nothing knows a `DAO.cls` exists elsewhere in the project, so the
    // runtime-object node gate (#245) drops the stub before the resolver
    // ever sees it.
    //
    // This is the trade the extractor had ALREADY made for the fourteen
    // names in `RUNTIME_RECEIVER_BLACKLIST` — a user class named `DoCmd` or
    // `Application` has never been call-linked either. #245 reconciles the
    // two lists, so `DAO`, `VBA`, `Collection`, `fso` and the control types
    // now share that behavior. The class itself is still indexed; only the
    // heuristic call edge qualified by the runtime name is gone.
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

    // The real declaration is still a first-class node.
    const execute = cg
      .searchNodes('Execute', { languages: ['vba'], kinds: ['function'] })
      .find((n) => n.node.name === 'Execute' && n.node.filePath.endsWith('DAO.cls'));
    expect(execute).toBeDefined();
    if (!execute) return;

    // No stub survives to be repointed — and, crucially, no `DAO.Execute`
    // phantom symbol is left behind either.
    expect(cg.getIncomingEdges(execute.node.id).filter((e) => e.kind === 'calls')).toHaveLength(0);
    expect(
      cg.searchNodes('DAO.Execute', { languages: ['vba'] })
        .filter((n) => n.node.name === 'DAO.Execute'),
    ).toHaveLength(0);
  });

  it('Test 4b (#245): a user class whose MEMBER name collides with a runtime member still stubs', async () => {
    // `Add` is `Collection.Add`'s member name. The gate keys on the RESOLVED
    // RECEIVER TYPE, never on the member, so a user class keeps its stub and
    // the resolver repoints it exactly as before.
    const cg = await buildProject({
      'src/classes/MiClase.cls': [
        ...CLS_HEADER,
        'Attribute VB_Name = "MiClase"',
        'Option Explicit',
        '',
        'Public Sub Add()',
        'End Sub',
        '',
      ].join('\n'),
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Option Explicit',
        '',
        'Public Sub CallerSub()',
        '    Dim m As MiClase',
        '    Set m = New MiClase',
        '    m.Add',
        'End Sub',
        '',
      ].join('\n'),
    });

    const add = cg
      .searchNodes('Add', { languages: ['vba'], kinds: ['function'] })
      .find((n) => n.node.name === 'Add' && n.node.filePath.endsWith('MiClase.cls'));
    expect(add).toBeDefined();
    if (!add) return;

    const incoming = cg.getIncomingEdges(add.node.id).filter((e) => e.kind === 'calls');
    expect(incoming.length).toBeGreaterThan(0);
    for (const edge of incoming) {
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

  it('Test 7: VBA stdlib functions are classified case-insensitively', () => {
    for (const name of ['CStr', 'CLng', 'Nz', 'IsNull', 'Array', 'InStr', 'Len', 'Replace', 'TypeName', 'VarType']) {
      expect(isVbaStdlibFunction(name)).toBe(true);
      expect(isVbaStdlibFunction(name.toLowerCase())).toBe(true);
    }
    expect(isVbaStdlibFunction('ProjectHelper')).toBe(false);
    expect(isVbaStdlibFunction(undefined)).toBe(false);
  });

  it('Test 8: stdlib unresolved calls are parked as declined-runtime', async () => {
    const cg = await buildProject({
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Public Sub CallerSub()',
        '    value = CStr(42)',
        '    ProjectHelper(42)',
        'End Sub',
      ].join('\n'),
    });
    const projectRoot = (cg as unknown as { projectRoot: string }).projectRoot;
    const connection = DatabaseConnection.open(getDatabasePath(projectRoot));
    const rows = connection.getDb().prepare(
      "SELECT reference_name, status FROM unresolved_refs WHERE reference_name IN ('CStr', 'ProjectHelper')",
    ).all() as Array<{ reference_name: string; status: string }>;
    connection.close();
    expect(rows.find((row) => row.reference_name === 'CStr')?.status).toBe('declined-runtime');
    expect(rows.find((row) => row.reference_name === 'ProjectHelper')?.status).toBe('failed');
  });

  it('Test 9: stdlib statement calls are declined only when extraction marks statement-call shape', async () => {
    const cg = await buildProject({
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Option Explicit',
        'Public Const SomePublicConst As Long = 1',
        '',
        'Public Sub CallerSub()',
        '    MsgBox "hi"',
        '    DoEvents',
        '    Shell "calc.exe"',
        '    MsgBox("hi")',
        '    ProjectHelper "x"',
        '    SomePublicConst',
        'End Sub',
        '',
        'Public Sub ProjectHelper(s As String)',
        'End Sub',
      ].join('\n'),
    });
    const projectRoot = (cg as unknown as { projectRoot: string }).projectRoot;
    const connection = DatabaseConnection.open(getDatabasePath(projectRoot));
    const rows = connection.getDb().prepare(
      "SELECT reference_name, reference_kind, status, metadata FROM unresolved_refs WHERE reference_name IN ('MsgBox', 'DoEvents', 'Shell', 'ProjectHelper', 'SomePublicConst') ORDER BY line",
    ).all() as Array<{
      reference_name: string;
      reference_kind: string;
      status: string;
      metadata: string | null;
    }>;
    connection.close();

    const runtimeRows = rows
      .filter((row) => ['MsgBox', 'DoEvents', 'Shell'].includes(row.reference_name))
      .map((row) => ({
        name: row.reference_name,
        kind: row.reference_kind,
        status: row.status,
        synthesizedBy: row.metadata
          ? (JSON.parse(row.metadata) as { synthesizedBy?: string }).synthesizedBy
          : undefined,
      }));
    // Issue #265: `MsgBox "hi"` and `Shell "calc.exe"` carry an argument
    // list, so their syntax rules out a `Const` read and they now carry the
    // canonical `calls` kind. Bare `DoEvents` stays `unqualified-ident`.
    // The point of this test is unchanged and still holds: all three are
    // `declined-runtime`, because the stdlib gate accepts both the
    // statement-call stamp and the `calls` kind.
    expect(runtimeRows).toEqual([
      {
        name: 'MsgBox',
        kind: 'calls',
        status: 'declined-runtime',
        synthesizedBy: 'vba-statement-call-unresolved',
      },
      {
        name: 'DoEvents',
        kind: 'unqualified-ident',
        status: 'declined-runtime',
        synthesizedBy: 'vba-statement-call-unresolved',
      },
      {
        name: 'Shell',
        kind: 'calls',
        status: 'declined-runtime',
        synthesizedBy: 'vba-statement-call-unresolved',
      },
      {
        name: 'MsgBox',
        kind: 'calls',
        status: 'declined-runtime',
        synthesizedBy: 'vba-paren-call-unresolved',
      },
    ]);

    const caller = cg
      .searchNodes('CallerSub', { languages: ['vba'], kinds: ['function'] })
      .find((node) => node.node.name === 'CallerSub');
    const helper = cg
      .searchNodes('ProjectHelper', { languages: ['vba'], kinds: ['function'] })
      .find((node) => node.node.name === 'ProjectHelper');
    expect(caller).toBeDefined();
    expect(helper).toBeDefined();
    expect(rows.filter((row) => row.reference_name === 'ProjectHelper')).toEqual([]);
    expect(
      caller && helper
        ? cg.getOutgoingEdges(caller.node.id).filter((edge) => edge.kind === 'calls' && edge.target === helper.node.id)
        : [],
    ).toHaveLength(1);

    const constantRow = rows.find((row) => row.reference_name === 'SomePublicConst');
    if (constantRow) {
      expect(constantRow).toMatchObject({
        reference_kind: 'unqualified-ident',
        status: 'failed',
      });
    } else {
      const constant = cg
        .searchNodes('SomePublicConst', { languages: ['vba'], kinds: ['constant'] })
        .find((node) => node.node.name === 'SomePublicConst');
      expect(constant).toBeDefined();
      expect(
        caller && constant
          ? cg.getIncomingEdges(constant.node.id).filter((edge) => edge.source === caller.node.id)
          : [],
      ).toHaveLength(1);
    }
  });

  it('Test 10: a user-defined MsgBox shadow resolves before runtime classification', async () => {
    const cg = await buildProject({
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Public Sub CallerSub()',
        '    MsgBox "shadow"',
        'End Sub',
      ].join('\n'),
      'src/modules/Shadow.bas': [
        'Attribute VB_Name = "Shadow"',
        'Public Function MsgBox(prompt As String) As Long',
        '    MsgBox = 1',
        'End Function',
      ].join('\n'),
    });

    const caller = cg
      .searchNodes('CallerSub', { languages: ['vba'], kinds: ['function'] })
      .find((node) => node.node.name === 'CallerSub');
    const shadow = cg
      .searchNodes('MsgBox', { languages: ['vba'], kinds: ['function'] })
      .find((node) => node.node.name === 'MsgBox' && node.node.filePath.endsWith('Shadow.bas'));
    expect(caller).toBeDefined();
    expect(shadow).toBeDefined();
    const edges =
      caller && shadow
        ? cg
            .getOutgoingEdges(caller.node.id)
            .filter((edge) => (edge.kind === 'calls' || edge.kind === 'references') && edge.target === shadow.node.id)
        : [];
    expect(edges).toHaveLength(1);

    const projectRoot = (cg as unknown as { projectRoot: string }).projectRoot;
    const connection = DatabaseConnection.open(getDatabasePath(projectRoot));
    const rows = connection.getDb().prepare(
      "SELECT reference_name, status FROM unresolved_refs WHERE reference_name = 'MsgBox'",
    ).all();
    connection.close();
    expect(rows).toEqual([]);
  });
});
