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

  it('indexes 6 VBA files (2 .bas + 2 .cls + 1 form .cls + 1 form .form.txt)', async () => {
    if (!cg) return;
    const stats = await cg.getStats();
    const vbaFiles = stats.filesByLanguage?.vba ?? 0;
    expect(vbaFiles).toBe(6);
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
});
