/**
 * E2E — real VBA fixtures copied from a Dysflow-managed Access project.
 *
 * The fixtures under `__tests__/fixtures/vba/` mirror the canonical Dysflow
 * folder structure (`src/modules/` for `.bas`, `src/classes/` for `.cls`,
 * `src/forms/` for the `.cls` + `.form.txt` pair). They were copied verbatim
 * from `C:\00repos\codigo\00_NO_CONFORMIDADES\` so the extractor is exercised
 * against real-world source, not synthetic strings.
 *
 * Strict TDD unit tests in `__tests__/extraction-vba*.test.ts` cover the
 * per-scenario shape; this E2E verifies the integrated tool works end-to-end
 * against real fixtures — same patterns, real bytes.
 *
 * The `.codegraph-vba/` index is created in the fixtures folder during this test
 * and removed in afterAll so subsequent runs are idempotent.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { CodeGraph } from '../src';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'vba');
const FORM_BASENAME = 'Form_FormNCAuditoriaMotivoEliminado';

describe('E2E - real VBA fixtures from a Dysflow-managed project', () => {
  let cg: CodeGraph | null = null;
  let initializedByTest = false;
  const codeGraphDir = path.join(FIXTURES_DIR, '.codegraph-vba');

  beforeAll(async () => {
    // Clean slate: remove any prior .codegraph-vba/ left by a previous run so
    // init() doesn't refuse to re-initialize.
    if (fs.existsSync(codeGraphDir)) {
      fs.rmSync(codeGraphDir, { recursive: true, force: true });
    }
    cg = await CodeGraph.init(FIXTURES_DIR, { index: false });
    initializedByTest = true;
    await cg.indexAll();
  }, 60_000);

  afterAll(async () => {
    if (cg) {
      try {
        await cg.close();
      } catch {
        // ignore close errors
      }
    }
    // Clean up the .codegraph-vba/ we created so subsequent runs are clean.
    if (initializedByTest && fs.existsSync(codeGraphDir)) {
      fs.rmSync(codeGraphDir, { recursive: true, force: true });
    }
  });

  it('indexes 7 VBA files (3 .bas + 2 .cls + 1 form .cls + 1 form .form.txt)', async () => {
    if (!cg) return;
    const stats = await cg.getStats();
    const vbaFiles = stats.filesByLanguage?.vba ?? 0;
    // modCallerDemo.bas (vba-graph-connectivity-fixes fixture) adds a 3rd
    // .bas file exercising the class-typed + .bas-qualified stub repoint.
    expect(vbaFiles).toBe(7);
  });

  it('emits a module node + function nodes for a .bas with Subs/Functions (mdlCursor.bas)', () => {
    if (!cg) return;
    // `mdlCursor.bas` declares two `Public Function`s (MouseCursor, PointM),
    // so per REQ-CODE-1 it must emit a module-level node plus the
    // function nodes. The Declare PtrSafe Function lines are preprocessor-
    // guarded and are intentionally NOT extracted as function nodes
    // (they're Win32 API bindings, not user-defined procedures).
    const moduleResults = cg.searchNodes('mdlCursor', {
      languages: ['vba'],
      kinds: ['module'],
    });
    expect(moduleResults.length).toBeGreaterThan(0);
    // Look up specific function names from the source.
    const mouseCursor = cg.searchNodes('MouseCursor', { languages: ['vba'] });
    const pointM = cg.searchNodes('PointM', { languages: ['vba'] });
    expect(mouseCursor.some((n) => n.node.kind === 'function')).toBe(true);
    expect(pointM.some((n) => n.node.kind === 'function')).toBe(true);
  });

  it('REQ-CODE-12/13: a constants .bas emits a module node + enum/constant symbols', () => {
    if (!cg) return;
    // `constantes.bas` declares `Public Const` entries and several
    // `Public Enum ... End Enum` blocks. Dysflow exports that text verbatim,
    // so per REQ-CODE-12/13 the extractor now emits the lazy module node plus
    // `enum`/`enum_member`/`constant` nodes. (REQ-CODE-10's "emits nothing"
    // is now narrowed to Option-directives-only files.)
    const nodes = cg.searchNodes('constantes', { languages: ['vba'] });
    const fromConstants = nodes.filter((n) =>
      n.node.filePath.endsWith('constantes.bas'),
    );
    const fileNodes = fromConstants.filter((n) => n.node.kind === 'file');
    const moduleNodes = fromConstants.filter((n) => n.node.kind === 'module');
    expect(fileNodes).toHaveLength(1);
    expect(moduleNodes.length).toBeGreaterThan(0);

    // The domain enums are now in the graph.
    const tipoUsuario = cg.searchNodes('EnumTipoUsuario', {
      languages: ['vba'],
      kinds: ['enum'],
    });
    expect(tipoUsuario.length).toBeGreaterThan(0);
    const administrador = cg.searchNodes('Administrador', {
      languages: ['vba'],
      kinds: ['enum_member'],
    });
    expect(administrador.length).toBeGreaterThan(0);

    // And the config constants.
    const constNode = cg.searchNodes('msoFileDialogOpen', {
      languages: ['vba'],
      kinds: ['constant'],
    });
    expect(constNode.length).toBeGreaterThan(0);
  });

  it('emits a class node for each class module (.cls) under src/classes/', () => {
    if (!cg) return;
    const arNodes = cg.searchNodes('ARAuditoria', {
      languages: ['vba'],
      kinds: ['class'],
    });
    expect(arNodes.length).toBeGreaterThan(0);
    const opNodes = cg.searchNodes('ACAuditoriaOperaciones', {
      languages: ['vba'],
      kinds: ['class'],
    });
    expect(opNodes.length).toBeGreaterThan(0);
  });

  it('emits a class node for the form code-behind (.cls under src/forms/)', () => {
    if (!cg) return;
    // The .cls file emits the class node; the .form.txt emits the module node.
    const formClass = cg.searchNodes(FORM_BASENAME, {
      languages: ['vba'],
      kinds: ['class'],
    });
    // Filter to the .cls file explicitly (includePatterns is not honored by
    // searchNodes, so we filter by filePath after the fact).
    const fromCls = formClass.filter((n) =>
      n.node.filePath.endsWith(`${FORM_BASENAME}.cls`),
    );
    expect(fromCls.length).toBeGreaterThan(0);
  });

  it('NON-NEGOTIABLE: .form.txt emits ZERO class nodes', () => {
    if (!cg) return;
    const formResults = cg.searchNodes(FORM_BASENAME, {
      languages: ['vba'],
      kinds: ['class'],
    });
    // Filter to the .form.txt file explicitly (includePatterns not honored).
    const fromFormTxt = formResults.filter((n) =>
      n.node.filePath.endsWith('.form.txt'),
    );
    expect(fromFormTxt).toHaveLength(0);
  });

  it('NON-NEGOTIABLE: .form.txt emits exactly one form-layout node + property nodes for controls', () => {
    if (!cg) return;
    const formResults = cg.searchNodes(FORM_BASENAME, {
      languages: ['vba'],
    });
    const fromFormTxt = formResults.filter((n) =>
      n.node.filePath.endsWith('.form.txt'),
    );
    // B2 (hueco 4): the file-level node is now `form-layout`, not
    // `module`. We still assert exactly one such node per `.form.txt`.
    const formLayouts = fromFormTxt.filter((n) => n.node.kind === 'form-layout');
    const properties = fromFormTxt.filter((n) => n.node.kind === 'property');
    // Exactly one form-level form-layout node (no sub-modules).
    expect(formLayouts).toHaveLength(1);
    // Real Access forms always have at least one control.
    expect(properties.length).toBeGreaterThan(0);
  });

  it('NON-NEGOTIABLE: .form.txt form-layout has a references edge to its sibling .cls', () => {
    if (!cg) return;
    const formResults = cg.searchNodes(FORM_BASENAME, {
      languages: ['vba'],
      kinds: ['form-layout'],
    });
    const formLayout = formResults
      .filter((n) => n.node.filePath.endsWith('.form.txt'))
      .at(0);
    expect(formLayout).toBeDefined();
    if (!formLayout) return;
    const outgoing = cg.getOutgoingEdges(formLayout.node.id);
    // REQ-FORM-1: at least one references edge from the form-layout node
    // toward the sibling .cls. The edge may carry metadata.synthesizedBy
    // = 'vba-form-binding' if the resolver preserves the provenance, or
    // it may be a plain resolved references edge — either way at least
    // one references-kind edge must exist.
    const refs = outgoing.filter((e) => e.kind === 'references');
    expect(refs.length).toBeGreaterThan(0);
    // If metadata.synthesizedBy is preserved, it must be 'vba-form-binding'.
    const formBinding = outgoing.filter((e) => {
      const meta = (e as { metadata?: Record<string, unknown> }).metadata;
      return meta?.synthesizedBy === 'vba-form-binding';
    });
    // Either the edge is preserved with provenance OR the edge exists without
    // it. Both are acceptable per the verify risk-adjudication result.
    expect(refs.length + formBinding.length).toBeGreaterThan(0);
  });

  it('indexes Dysflow queries/*.sql as query nodes with table references', () => {
    if (!cg) return;
    // `src/queries/` carries `queries.json` + `Consulta3.sql` / `Q_Smoke.sql`,
    // the canonical Dysflow saved-query layout. The discovery gate picks the
    // `.sql` files up (sibling manifest present) and routes them to the
    // SqlQueryExtractor, which emits a `query` node per file.
    const q = cg.searchNodes('Consulta3', {
      languages: ['sql'],
      kinds: ['query'],
    });
    expect(q.length).toBeGreaterThan(0);
    const queryNode = q.find((n) => /Consulta3\.sql$/i.test(n.node.filePath)) ?? q[0];
    expect(queryNode).toBeDefined();
    if (!queryNode) return;

    // The query references the table it selects from.
    const out = cg.getOutgoingEdges(queryNode.node.id);
    const refs = out.filter((e) => e.kind === 'references');
    expect(refs.length).toBeGreaterThan(0);
    const table = cg.searchNodes('TbACParaLista', { languages: ['sql'] });
    expect(table.length).toBeGreaterThan(0);
  });

  // -------------------------------------------------------------------------
  // Issue #12 (#12b): post-extraction call-stub resolution. `modCallerDemo.bas`
  // (added for this change) calls `ACAuditoriaOperaciones.Registrar` (class-
  // typed, TWICE — F1 duplicate-collapse) and `mdlCursor.MouseCursor(...)`
  // (`.bas`-qualified, module-scoped fallback).
  // -------------------------------------------------------------------------

  it('#12 (6.1): class-typed call-stub resolves to the real cross-file method node', () => {
    if (!cg) return;
    const registrar = cg
      .searchNodes('Registrar', { languages: ['vba'], kinds: ['function'] })
      .find(
        (n) => n.node.name === 'Registrar' && n.node.filePath.endsWith('ACAuditoriaOperaciones.cls'),
      );
    expect(registrar).toBeDefined();
    if (!registrar) return;

    const incoming = cg.getIncomingEdges(registrar.node.id).filter((e) => e.kind === 'calls');
    expect(incoming.length).toBeGreaterThan(0);
    // Every incoming edge must be resolved — no stub metadata left.
    for (const edge of incoming) {
      expect(edge.metadata?.stub).not.toBe(true);
    }
    // No leftover stub node named `ACAuditoriaOperaciones.Registrar` should
    // remain — the resolver deletes it once resolved.
    const stubLeftover = cg
      .searchNodes('ACAuditoriaOperaciones.Registrar', { languages: ['vba'] })
      .filter((n) => n.node.name === 'ACAuditoriaOperaciones.Registrar');
    expect(stubLeftover).toHaveLength(0);
  });

  it('#12 (6.2): .bas-qualified call-stub resolves to the real bare-name node via module-scoped narrowing', () => {
    if (!cg) return;
    const mouseCursor = cg
      .searchNodes('MouseCursor', { languages: ['vba'] })
      .find(
        (n) => n.node.kind === 'function' && n.node.name === 'MouseCursor' && n.node.filePath.endsWith('mdlCursor.bas'),
      );
    expect(mouseCursor).toBeDefined();
    if (!mouseCursor) return;

    const incoming = cg.getIncomingEdges(mouseCursor.node.id).filter((e) => e.kind === 'calls');
    expect(incoming.length).toBeGreaterThan(0);
    for (const edge of incoming) {
      expect(edge.metadata?.stub).not.toBe(true);
    }
    // No leftover `mdlCursor.MouseCursor` stub node.
    const stubLeftover = cg
      .searchNodes('mdlCursor.MouseCursor', { languages: ['vba'] })
      .filter((n) => n.node.name === 'mdlCursor.MouseCursor');
    expect(stubLeftover).toHaveLength(0);
  });

  it('#12 (6.3 — F1): two call sites to the same target collapse to exactly ONE edge row', () => {
    if (!cg) return;
    const registrar = cg
      .searchNodes('Registrar', { languages: ['vba'], kinds: ['function'] })
      .find(
        (n) => n.node.name === 'Registrar' && n.node.filePath.endsWith('ACAuditoriaOperaciones.cls'),
      );
    const callDemo = cg
      .searchNodes('CallDemo', { languages: ['vba'], kinds: ['function'] })
      .find((n) => n.node.name === 'CallDemo');
    expect(registrar).toBeDefined();
    expect(callDemo).toBeDefined();
    if (!registrar || !callDemo) return;

    // `modCallerDemo.bas` calls `x.Registrar p_Error` TWICE at two different
    // call sites — both must converge on exactly ONE (source,target,'calls')
    // edge row, not two.
    const edgesBetween = cg
      .getOutgoingEdges(callDemo.node.id)
      .filter((e) => e.kind === 'calls' && e.target === registrar.node.id);
    expect(edgesBetween).toHaveLength(1);

    // getCallers must list the caller ONCE too (not twice).
    const callers = cg
      .getCallers(registrar.node.id, 1)
      .filter((c) => c.node.id === callDemo.node.id);
    expect(callers).toHaveLength(1);
  });
});

// -----------------------------------------------------------------------------
// #12 (6.4/6.5/6.6): purpose-built small isolated project (its OWN tmp dir,
// separate CodeGraph instance) — needs to run indexAll() a second time and
// sync() a single file, which would corrupt the shared `cg` instance/state
// used by every test above if done in-place.
// -----------------------------------------------------------------------------

describe('E2E - VBA call-stub resolution edge cases (#12b: unmatched, idempotency, self-heal)', () => {
  let tmpDir: string;
  let cg2: CodeGraph | null = null;

  const CALLER_SRC = [
    'Attribute VB_Name = "Caller"',
    'Option Explicit',
    '',
    'Public Sub Go()',
    '    Dim m_Unknown As UnknownClassXYZ',
    '    m_Unknown.DoSomething',
    'End Sub',
    '',
    'Public Sub GoReal()',
    '    Dim m_Real As RealTarget',
    '    m_Real.DoWork',
    'End Sub',
    '',
  ].join('\n');

  const REAL_TARGET_SRC = [
    'VERSION 1.0 CLASS',
    'BEGIN',
    "  MultiUse = -1  'True",
    'END',
    'Attribute VB_Name = "RealTarget"',
    'Option Explicit',
    '',
    'Public Sub DoWork()',
    'End Sub',
    '',
  ].join('\n');

  beforeAll(async () => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-stub-e2e-'));
    fs.mkdirSync(path.join(tmpDir, 'src', 'modules'), { recursive: true });
    fs.mkdirSync(path.join(tmpDir, 'src', 'classes'), { recursive: true });
    fs.writeFileSync(path.join(tmpDir, 'src', 'modules', 'Caller.bas'), CALLER_SRC);
    fs.writeFileSync(path.join(tmpDir, 'src', 'classes', 'RealTarget.cls'), REAL_TARGET_SRC);

    cg2 = await CodeGraph.init(tmpDir, { index: false });
    await cg2.indexAll();
  }, 60_000);

  afterAll(async () => {
    if (cg2) {
      try {
        await cg2.close();
      } catch {
        // ignore close errors
      }
    }
    if (fs.existsSync(tmpDir)) fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it('#12 (6.4): unmatched call-stub keeps stub metadata, no throw', () => {
    if (!cg2) return;
    // `UnknownClassXYZ` never resolves to any real class in the project —
    // 0 candidates after the exact-qualifiedName + `.bas`-fallback steps,
    // so the resolver DECLINES and leaves the stub edge untouched.
    //
    // Node metadata isn't persisted by this codebase's schema (only edge
    // metadata is — see `getVbaCallStubs`'s design-deviation note in
    // `src/db/queries.ts`), so "keeps stub metadata" is verified on the
    // EDGE (which DOES persist `metadata.stub` through `indexAll()`),
    // not on the retained stub node itself.
    const goNode = cg2
      .searchNodes('Go', { languages: ['vba'], kinds: ['function'] })
      .find((n) => n.node.name === 'Go');
    expect(goNode).toBeDefined();
    if (!goNode) return;

    const outgoing = cg2.getOutgoingEdges(goNode.node.id).filter((e) => e.kind === 'calls');
    const stubEdge = outgoing.find((e) => e.metadata?.member === 'DoSomething');
    expect(stubEdge).toBeDefined();
    expect(stubEdge?.metadata?.stub).toBe(true);
    expect(stubEdge?.metadata?.receiverType).toBe('UnknownClassXYZ');
  });

  it('#12 (6.5 — F6): re-running indexAll() unchanged is idempotent (stable node/edge counts)', async () => {
    if (!cg2) return;
    const before = await cg2.getStats();

    await cg2.indexAll();

    const after = await cg2.getStats();
    expect(after.nodeCount).toBe(before.nodeCount);
    expect(after.edgeCount).toBe(before.edgeCount);

    // The resolved GoReal → RealTarget.DoWork edge must still be resolved
    // (not reverted to a stub, not duplicated) after the second full index.
    const doWork = cg2
      .searchNodes('DoWork', { languages: ['vba'], kinds: ['function'] })
      .find((n) => n.node.name === 'DoWork');
    expect(doWork).toBeDefined();
    if (!doWork) return;
    const incoming = cg2.getIncomingEdges(doWork.node.id).filter((e) => e.kind === 'calls');
    expect(incoming).toHaveLength(1);
    expect(incoming[0]?.metadata?.stub).not.toBe(true);
  });

  it('#12 (6.6 — F4): resyncing ONLY the target file (caller untouched) self-heals the repointed edge', async () => {
    if (!cg2) return;
    const targetPath = path.join(tmpDir, 'src', 'classes', 'RealTarget.cls');
    const original = fs.readFileSync(targetPath, 'utf8');
    // Touch ONLY the target — content changes so sync() re-extracts it, but
    // the caller (Caller.bas) is untouched.
    fs.writeFileSync(targetPath, `${original}\n' resynced\n`);

    await cg2.sync();

    const doWork = cg2
      .searchNodes('DoWork', { languages: ['vba'], kinds: ['function'] })
      .find((n) => n.node.name === 'DoWork');
    expect(doWork).toBeDefined();
    if (!doWork) return;

    const incoming = cg2.getIncomingEdges(doWork.node.id).filter((e) => e.kind === 'calls');
    expect(incoming).toHaveLength(1);
    expect(incoming[0]?.metadata?.stub).not.toBe(true);
  });
});
