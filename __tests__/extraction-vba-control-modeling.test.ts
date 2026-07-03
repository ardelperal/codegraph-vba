/**
 * extraction-vba-control-modeling.test.ts
 *
 * RED tests for the six VBA form control-modeling huecos.
 *
 * Phase A — these tests must FAIL today. Each test is a sharp, isolated
 * reproduction of a known modeling gap. Phase B (a separate work unit)
 * will make them GREEN by changing the production extractors.
 *
 * Hueco map (symptom → test):
 *   1. Me.X reference resolution              → "Me.lblTitulo debe emitir referencia al control"
 *   2. .form.txt control NAME (not just TYPE) → "cada control declarado en .form.txt debe aparecer como nodo por su nombre"
 *   3. event-handler edge ctrl → handler Sub  → "ComandoAltaPM_Click debe tener arista al control"
 *   4. .form.txt kind=form-layout             → ".form.txt no debe emitir nodos kind=module"
 *   5. Form_Load qualifiedName carries form   → 'query "Form_Load" debe componer prefijo de form'
 *   6. DoCmd.OpenForm "FormTest" modeling     → 'DoCmd.OpenForm "FormTest" debe emitir arista opens-form'
 *
 * Fixture layout (under __tests__/fixtures/vba-control-modeling/):
 *   Form_TestForm.cls        — code-behind: Me.lblTitulo, ComandoAltaPM_Click, Form_Load, ...
 *   Form_TestForm.form.txt   — UI: ComandoAltaPM, ComandoBajaPM, lblTitulo, lblDescripcion,
 *                              txtDescripcion, txtCodigo, grpEstado, recMarco, lstRiesgos
 *   Form_OtherForm.cls       — second form with its own Form_Load (for hueco 5 disambiguation)
 *   Form_OtherForm.form.txt  — UI with lblTitulo (so the other form has a real label too)
 *   modTestHelper.bas        — standard module with DoCmd.OpenForm "FormTest"
 *
 * Hard rules (Phase A):
 *   - Tests are RED, no GREEN code is written here.
 *   - No src/extraction/* files are modified.
 *   - The CodeGraph index for hueco 5 is created under the fixture dir and
 *     cleaned up in afterAll so this test leaves no stale `.codegraph-vba/`
 *     behind in the repo.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';
import { generateNodeId } from '../src/extraction/tree-sitter-helpers';
import { CodeGraph } from '../src';

// =============================================================================
// Fixture paths — all RED tests resolve these from the repo root.
// =============================================================================
const FIXTURE_DIR = path.join(__dirname, 'fixtures', 'vba-control-modeling');
const TEST_FORM_CLS = path.join(FIXTURE_DIR, 'Form_TestForm.cls');
const TEST_FORM_TXT = path.join(FIXTURE_DIR, 'Form_TestForm.form.txt');
const MOD_TEST_HELPER_BAS = path.join(FIXTURE_DIR, 'modTestHelper.bas');
// Report_ fixtures (issue #41): the .cls code-behind must bind to its sibling
// .report.txt the same way Form_*.cls binds to its .form.txt. The sibling
// path is computed from the cls path by swapping `.cls` for `.report.txt`,
// keeping the same directory and basename.
const TEST_REPORT_CLS = path.join(FIXTURE_DIR, 'Report_PayrollSummary.cls');
const TEST_REPORT_TXT = path.join(FIXTURE_DIR, 'Report_PayrollSummary.report.txt');
const TEST_REPORT_NO_SIBLING_CLS = path.join(
  FIXTURE_DIR,
  'Report_NoSibling.cls',
);

function readFixture(relPath: string): string {
  return fs.readFileSync(relPath, 'utf8');
}

// =============================================================================
// HUECO 1 — `Me.<Control>` reference resolution
// =============================================================================
describe('hueco-1: Me.X reference resolution', () => {
  it('Me.lblTitulo debe emitir referencia al control lblTitulo', () => {
    const r = new VbaExtractor(TEST_FORM_CLS, readFixture(TEST_FORM_CLS)).extract();

    // Today the Me.* member-access chain is silently dropped: neither the
    // unresolvedReferences list nor the edges list carries `lblTitulo`.
    // Phase B must capture the name `lblTitulo` as either:
    //   (a) an UnresolvedReference with `referenceName === 'lblTitulo'`, OR
    //   (b) a `references` edge whose target node's name === 'lblTitulo'.
    //
    // We assert (a) here — that's the simplest and cheapest hook for the
    // resolver to pick up later.
    const unresolvedNames = r.unresolvedReferences.map((u) => u.referenceName);
    expect(unresolvedNames).toContain('lblTitulo');
  });
});

// =============================================================================
// HUECO 2 — .form.txt emits control NAME, not just TYPE
// =============================================================================
describe('hueco-2: .form.txt exposes control NAME (not just TYPE)', () => {
  it('cada control declarado en .form.txt debe aparecer como nodo por su nombre', () => {
    // Today the .form.txt extractor emits one `property` node per
    // `Begin <Type>` line whose `name` is the TYPE ('CommandButton',
    // 'Label', etc.). The control's `Name = "..."` attribute is discarded.
    // Phase B must emit one node per control whose `name` is the control
    // name (e.g. 'ComandoAltaPM').
    const r = new VbaFormExtractor(
      TEST_FORM_TXT,
      readFixture(TEST_FORM_TXT),
    ).extract();

    const expectedControlNames = [
      'ComandoAltaPM',
      'ComandoBajaPM',
      'lblTitulo',
      'lblDescripcion',
      'txtDescripcion',
      'txtCodigo',
      'grpEstado',
      'recMarco',
      'lstRiesgos',
    ];
    const missing = expectedControlNames.filter(
      (name) => !r.nodes.some((n) => n.name === name),
    );
    expect(
      missing,
      `expected .form.txt to emit one node per control name; missing: ${missing.join(', ')}`,
    ).toEqual([]);
  });
});

// =============================================================================
// HUECO 3 — event-handler edge from control → handler Sub
// =============================================================================
describe('hueco-3: event-handler edge from control to Sub', () => {
  it('ComandoAltaPM_Click debe tener arista al control ComandoAltaPM', () => {
    // Cross-extract: the .cls emits the Click handler function; the
    // .form.txt emits the control node. Today the bridge between them is
    // missing — the resolver has no way to walk from a control to its
    // event handlers.
    //
    // We look across the union of both extracts because that's what a real
    // index pass would do.

    const cls = new VbaExtractor(TEST_FORM_CLS, readFixture(TEST_FORM_CLS)).extract();
    const form = new VbaFormExtractor(
      TEST_FORM_TXT,
      readFixture(TEST_FORM_TXT),
    ).extract();

    const clickHandler = cls.nodes.find(
      (n) => n.kind === 'function' && n.name === 'ComandoAltaPM_Click',
    );
    const control = form.nodes.find((n) => n.name === 'ComandoAltaPM');

    expect(clickHandler, 'expected function node for ComandoAltaPM_Click').toBeDefined();
    expect(control, 'expected control node named ComandoAltaPM').toBeDefined();

    // Phase B will emit an edge with kind === 'event-handler' from the
    // control to the handler (or vice-versa). Today no such edge exists.
    //
    // NOTE: 'event-handler' is NOT in EdgeKind today. We compare as strings
    // so the test compiles cleanly and the assertion is a runtime RED
    // rather than a compile error.
    const EVT_KIND = 'event-handler';
    const allEdges = [...cls.edges, ...form.edges];
    const bridge = allEdges.find(
      (e) =>
        (e.kind as string) === EVT_KIND &&
        ((e.source === control?.id && e.target === clickHandler?.id) ||
          (e.source === clickHandler?.id && e.target === control?.id)),
    );
    expect(bridge, 'expected event-handler edge between control and Sub').toBeDefined();
  });
});

// =============================================================================
// HUECO 4 — .form.txt kind is `form-layout`, not `module`
// =============================================================================
describe('hueco-4: .form.txt emits kind=form-layout, NOT kind=module', () => {
  it('.form.txt no debe emitir nodos kind=module', () => {
    const r = new VbaFormExtractor(
      TEST_FORM_TXT,
      readFixture(TEST_FORM_TXT),
    ).extract();

    // Strict RED: NO node with kind `module` whose filePath ends in
    // .form.txt. Phase B will either rename the existing module node to
    // `form-layout`, or emit a parallel `form-layout` node — either way
    // the strict assertion (zero `module` nodes for .form.txt files) holds.
    const moduleNodes = r.nodes.filter(
      (n) =>
        n.kind === 'module' &&
        (n.filePath ?? '').endsWith('Form_TestForm.form.txt'),
    );
    expect(moduleNodes.length, 'no module nodes for .form.txt files').toBe(0);

    // AND there must be at least one node with kind `form-layout` for the
    // form, so future resolvers can dispatch on that kind specifically.
    // 'form-layout' is NOT in NodeKind today — compare with String() to
    // keep the test type-clean.
    const FORM_LAYOUT_KIND = 'form-layout';
    const formLayoutNodes = r.nodes.filter(
      (n) =>
        (n.kind as string) === FORM_LAYOUT_KIND &&
        (n.filePath ?? '').endsWith('Form_TestForm.form.txt'),
    );
    expect(
      formLayoutNodes.length,
      'at least one form-layout node per .form.txt',
    ).toBeGreaterThan(0);
  });
});

// =============================================================================
// HUECO 5 — `Form_Load` qualifiedName carries the form prefix
// =============================================================================
describe('huecos 3 & 5: VBA event-handler and Form_Load integration', () => {
  let cg: CodeGraph | null = null;
  const codeGraphDir = path.join(FIXTURE_DIR, '.codegraph-vba');
  let initialized = false;

  beforeAll(async () => {
    if (fs.existsSync(codeGraphDir)) {
      fs.rmSync(codeGraphDir, { recursive: true, force: true });
    }
    cg = await CodeGraph.init(FIXTURE_DIR, { index: false });
    initialized = true;
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
    if (initialized && fs.existsSync(codeGraphDir)) {
      fs.rmSync(codeGraphDir, { recursive: true, force: true });
    }
  });

  it('ComandoAltaPM_Click debe tener callees con el control ComandoAltaPM (hueco-3)', async () => {
    if (!cg) return;
    const matches = cg.searchNodes('ComandoAltaPM_Click', { languages: ['vba'] });
    const fnNode = matches.find((m) => m.node.name === 'ComandoAltaPM_Click')?.node;
    expect(fnNode).toBeDefined();
    if (!fnNode) return;

    const callees = cg.getCallees(fnNode.id);
    expect(callees).toContainEqual(
      expect.objectContaining({
        node: expect.objectContaining({
          kind: 'form-instance-control',
          name: 'ComandoAltaPM',
        }),
        edge: expect.objectContaining({
          kind: 'event-handler',
        }),
      })
    );
  });

  it('query "Form_Load" debe componer qualifiedName con prefijo de form', async () => {
    if (!cg) return;
    const hits = cg.searchNodes('Form_Load', { languages: ['vba'] });

    // Filter to function-kind hits (skip the file node and any incidental
    // text matches in identifier bodies).
    const fnHits = hits.filter((h) => h.node.kind === 'function');
    expect(fnHits.length).toBeGreaterThan(0);

    // Every Form_Load hit must be qualified with its owning form:
    //   Form_TestForm.Form_Load
    //   Form_OtherForm.Form_Load
    // Today both have qualifiedName === 'Form_Load' (no prefix), so this
    // assertion fails RED.
    for (const hit of fnHits) {
      expect(
        hit.node.qualifiedName,
        `qualifiedName ${hit.node.qualifiedName} (file ${hit.node.filePath}) must include form prefix`,
      ).toMatch(/^Form_[^.]+\.Form_Load$/);
    }
  });

  it('codegraph_search de Form_Load debe incluir el prefijo de form en la salida (hueco-5)', async () => {
    if (!cg) return;
    const { ToolHandler } = await import('../src/mcp/tools');
    const handler = new ToolHandler(cg);
    const res = await handler.execute('codegraph_search', { query: 'Form_Load' });
    const text = res.content?.[0]?.text ?? '';
    expect(text).toContain('Form_TestForm.Form_Load');
    expect(text).toContain('Form_OtherForm.Form_Load');
  });
});

// =============================================================================
// HUECO 6 — `DoCmd.OpenForm "FormTest"` modeling
// =============================================================================
describe('hueco-6: DoCmd.OpenForm built-in modeling', () => {
  it('DoCmd.OpenForm "FormTest" debe emitir arista opens-form', () => {
    const r = new VbaExtractor(
      MOD_TEST_HELPER_BAS,
      readFixture(MOD_TEST_HELPER_BAS),
    ).extract();

    // Today, `DoCmd.OpenForm` is absorbed by the runtime-receiver blacklist
    // and the string literal "FormTest" is silently discarded. Phase B must
    // emit an edge capturing the target form name.
    //
    // 'opens-form' is NOT in EdgeKind today — compare with String() to keep
    // the test type-clean.
    const OPENS_FORM_KIND = 'opens-form';
    const edges = r.edges;
    const targetCandidates = edges
      .filter((e) => (e.kind as string) === OPENS_FORM_KIND)
      .map((e) => {
        // The target form name may live on the edge itself (when the form
        // module is not yet indexed) or on the target node. We accept either.
        const meta = (e.metadata ?? {}) as Record<string, unknown>;
        const targetName =
          typeof meta.targetFormName === 'string'
            ? (meta.targetFormName as string)
            : undefined;
        return targetName;
      })
      .filter((n): n is string => typeof n === 'string');

    expect(targetCandidates).toContain('FormTest');
  });
});

// =============================================================================
// Issue #41 — Report_*.cls code-behind must bind to its .report.txt sibling
// (symmetric to Form_*.cls → .form.txt).
//
// Pre-fix the cls-side binding was Form_-only and hardcoded the sibling
// extension to `.form.txt`, so report code-behind Subs had no edges to
// their controls. The fix must be PREFIX-driven, NOT a copy of the Form_
// path; the same helper decides both. The stub `form-instance-control`
// id formula `generateNodeId(siblingPath, 'form-instance-control',
// controlName, 0)` must produce the SAME id whether the sibling is
// `.form.txt` or `.report.txt` so the per-file INSERT OR REPLACE works.
// =============================================================================
describe('issue-41: Report_*.cls code-behind binds to its .report.txt sibling', () => {
  // Sibling path the .cls side must compute: the .cls path with the .cls
  // extension swapped for .report.txt. The same string is what
  // VbaFormExtractor receives as `this.filePath` when it parses the
  // .report.txt, so the deterministic id derived from it must match.
  const reportSiblingPath = TEST_REPORT_CLS.replace(/\.cls$/i, '.report.txt');

  it('Report_PayrollSummary.cls con txtTotal_Click emite arista event-handler al control txtTotal', () => {
    // The cls body has `Private Sub txtTotal_Click()` and the sibling
    // .report.txt declares a TextBox named `txtTotal`. The bridge edge
    // (kind=event-handler) MUST exist between the handler Sub and the
    // synthesized stub form-instance-control node whose id matches the
    // real node VbaFormExtractor would emit for the same control.
    const cls = new VbaExtractor(TEST_REPORT_CLS, readFixture(TEST_REPORT_CLS)).extract();
    const form = new VbaFormExtractor(TEST_REPORT_TXT, readFixture(TEST_REPORT_TXT)).extract();

    const clickHandler = cls.nodes.find(
      (n) => n.kind === 'function' && n.name === 'txtTotal_Click',
    );
    expect(clickHandler, 'expected function node for txtTotal_Click').toBeDefined();
    if (!clickHandler) return;

    // Cross-check that VbaFormExtractor actually emits a form-instance-control
    // for the same control — this is the "real" node that the cls-side stub
    // must collide with via INSERT OR REPLACE.
    const expectedStubId = generateNodeId(
      reportSiblingPath,
      'form-instance-control',
      'txtTotal',
      0,
    );
    const realControlFromForm = form.nodes.find((n) => n.id === expectedStubId);
    expect(
      realControlFromForm,
      'VbaFormExtractor must emit a form-instance-control with id derived from the .report.txt path',
    ).toBeDefined();
    expect(realControlFromForm?.kind).toBe('form-instance-control');
    expect(realControlFromForm?.name).toBe('txtTotal');

    // The cls side must synthesize the SAME stub id so the per-file
    // edge filter accepts the edge and the eventual INSERT OR REPLACE
    // overwrites the stub with the real node.
    const clsStub = cls.nodes.find((n) => n.id === expectedStubId);
    expect(
      clsStub,
      `expected cls-side synthesized stub at id ${expectedStubId}; got ${clsStub?.id ?? 'none'}`,
    ).toBeDefined();
    expect(clsStub?.kind).toBe('form-instance-control');
    expect(clsStub?.filePath).toBe(reportSiblingPath);

    // And there must be an event-handler edge from the handler Sub to the stub.
    const EVT_KIND = 'event-handler';
    const evtEdge = cls.edges.find(
      (e) =>
        (e.kind as string) === EVT_KIND &&
        e.source === clickHandler.id &&
        e.target === expectedStubId,
    );
    expect(
      evtEdge,
      'expected event-handler edge from txtTotal_Click to the synthesized stub',
    ).toBeDefined();
  });

  it('Report_PayrollSummary.cls cubre todos los handlers cuyos controles existen en el sibling', () => {
    // Same happy-path coverage for txtCount (the second handler with a
    // matching control). Together with the txtTotal atom this proves the
    // prefix-driven helper handles multiple handlers per report, not just one.
    const cls = new VbaExtractor(TEST_REPORT_CLS, readFixture(TEST_REPORT_CLS)).extract();
    const expectedStubId = generateNodeId(
      reportSiblingPath,
      'form-instance-control',
      'txtCount',
      0,
    );

    const handler = cls.nodes.find(
      (n) => n.kind === 'function' && n.name === 'txtCount_Click',
    );
    expect(handler).toBeDefined();

    const stub = cls.nodes.find((n) => n.id === expectedStubId);
    expect(stub, 'cls-side synthesized stub for txtCount').toBeDefined();

    const EVT_KIND = 'event-handler';
    const edge = cls.edges.find(
      (e) =>
        (e.kind as string) === EVT_KIND &&
        e.source === handler?.id &&
        e.target === expectedStubId,
    );
    expect(edge, 'event-handler edge txtCount_Click → txtCount stub').toBeDefined();
  });

  it('regresión Form_: el stub id se deriva del path .form.txt, no del .report.txt', () => {
    // Critical regression guard: the fix is PREFIX-driven, NOT a copy of
    // the Form_ logic. After the refactor, the Form_ path must STILL
    // produce a stub id derived from the .form.txt sibling path. If a
    // future change accidentally unifies Report_ to also write to
    // .form.txt stubs, or vice-versa, this test fails.
    const cls = new VbaExtractor(TEST_FORM_CLS, readFixture(TEST_FORM_CLS)).extract();
    const formSiblingPath = TEST_FORM_CLS.replace(/\.cls$/i, '.form.txt');
    const expectedStubId = generateNodeId(
      formSiblingPath,
      'form-instance-control',
      'ComandoAltaPM',
      0,
    );

    const stub = cls.nodes.find((n) => n.id === expectedStubId);
    expect(
      stub,
      `Form_ stub id must match generateNodeId(.form.txt path); expected ${expectedStubId}`,
    ).toBeDefined();
    expect(stub?.filePath).toBe(formSiblingPath);
    // The stub path MUST end in .form.txt, never .report.txt.
    expect(stub?.filePath.endsWith('.form.txt')).toBe(true);
    expect(stub?.filePath.endsWith('.report.txt')).toBe(false);
  });

  it('Report_*.cls sin sibling .report.txt NO emite stub a .form.txt', () => {
    // Acceptance criterion 4: a Report_*.cls body without a matching
    // .report.txt sibling must NOT emit a synthesized stub with a
    // .form.txt extension. The bug being fixed is specifically the
    // silent `.form.txt` stub — the pre-fix code only matched `Form_*`
    // so Report_ produced no event-handler edges at all, but a future
    // regression that copy-pasted the Form_ logic would re-introduce
    // the wrong extension.
    //
    // The accepted invariant matches the existing Form_ behaviour: the
    // helper synthesizes a stub for the EXPECTED sibling path
    // (`.report.txt` here) without checking whether the sibling file
    // is actually on disk. The resolver and CodeGraph handle the
    // missing-sibling case downstream.
    const cls = new VbaExtractor(
      TEST_REPORT_NO_SIBLING_CLS,
      readFixture(TEST_REPORT_NO_SIBLING_CLS),
    ).extract();

    // Hard guard: no stub node with a .form.txt extension can exist
    // for a Report_*.cls body. This is the exact bug being fixed.
    const formTxtStubs = cls.nodes.filter(
      (n) =>
        n.kind === 'form-instance-control' &&
        typeof n.filePath === 'string' &&
        n.filePath.endsWith('.form.txt'),
    );
    expect(
      formTxtStubs,
      'no .form.txt stubs synthesized for Report_*.cls (the bug)',
    ).toEqual([]);

    // And every event-handler stub synthesized by the prefix-driven
    // helper for a Report_*.cls file must point to a .report.txt path,
    // never a .form.txt path. This holds whether or not the sibling
    // file is on disk.
    const EVT_KIND = 'event-handler';
    const reportStubs = cls.edges
      .filter((e) => (e.kind as string) === EVT_KIND)
      .map((e) => cls.nodes.find((n) => n.id === e.target))
      .filter((n): n is { filePath?: string } => n !== undefined);
    for (const stub of reportStubs) {
      const fp = stub.filePath ?? '';
      expect(
        fp.endsWith('.report.txt') && !fp.endsWith('.form.txt'),
        `Report_*.cls event-handler stub must use .report.txt, got: ${fp}`,
      ).toBe(true);
    }
  });
});