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
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';

// =============================================================================
// Fixture paths — all RED tests resolve these from the repo root.
// =============================================================================
const FIXTURE_DIR = path.join(__dirname, 'fixtures', 'vba-control-modeling');
const TEST_FORM_CLS = path.join(FIXTURE_DIR, 'Form_TestForm.cls');
const TEST_FORM_TXT = path.join(FIXTURE_DIR, 'Form_TestForm.form.txt');

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