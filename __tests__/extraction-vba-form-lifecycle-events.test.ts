/**
 * extraction-vba-form-lifecycle-events.test.ts
 *
 * Acceptance tests for issue #247 — a form's (or report's) OWN lifecycle
 * handlers must produce an `event-handler` edge to the sibling
 * `form-layout` / `report-layout` node.
 *
 * Before #247 those handlers produced nothing: `parseEventHandlerName`
 * refuses `Form_*` on purpose, because routing a form-level event through
 * the control path would synthesize a `form-instance-control` node
 * literally named `Form`. The fix keeps that refusal and gives the
 * form-level case its own target instead.
 *
 * The two invariants that make this safe are asserted directly:
 *   1. The stub the code-behind emits carries the SAME deterministic id
 *      `VbaFormExtractor` produces for the real layout node, so the
 *      `INSERT OR REPLACE` convergence works whichever file is indexed
 *      first. We do not hardcode the id — we run the real form extractor
 *      on the real sibling file and compare.
 *   2. NO extra `form-instance-control` node appears. A bogus `Form`
 *      control node is exactly the failure mode this change had to avoid.
 *
 * Fixture layout (under __tests__/fixtures/vba-form-lifecycle/):
 *   Form_Lifecycle.cls          — Form_Open / Form_Load / Form_Unload (events),
 *                                 Form_Helper (NOT an event), cmdSave_Click
 *                                 (a control handler, must be unchanged).
 *   Form_Lifecycle.form.txt     — UI with cmdSave + lblTitulo.
 *   Report_Lifecycle.cls        — Report_Open (event), Report_Helper (not an
 *                                 event), txtTotal_Click (control handler).
 *   Report_Lifecycle.report.txt — UI with txtTotal.
 *
 * Real files, real extractors, no mocking.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';
import type { Edge, ExtractionResult, Node } from '../src/types';

const FIXTURE_DIR = path.join(__dirname, 'fixtures', 'vba-form-lifecycle');
const FORM_CLS = path.join(FIXTURE_DIR, 'Form_Lifecycle.cls');
const FORM_TXT = path.join(FIXTURE_DIR, 'Form_Lifecycle.form.txt');
const REPORT_CLS = path.join(FIXTURE_DIR, 'Report_Lifecycle.cls');
const REPORT_TXT = path.join(FIXTURE_DIR, 'Report_Lifecycle.report.txt');

function readFixture(p: string): string {
  return fs.readFileSync(p, 'utf8');
}

function extractCode(p: string): ExtractionResult {
  return new VbaExtractor(p, readFixture(p)).extract();
}

function extractLayout(p: string): ExtractionResult {
  return new VbaFormExtractor(p, readFixture(p)).extract();
}

/** The `function` node for a Sub declared in the given extract. */
function fn(result: ExtractionResult, name: string): Node | undefined {
  return result.nodes.find((n) => n.kind === 'function' && n.name === name);
}

/** Every `event-handler` edge leaving the named Sub. */
function handlerEdges(result: ExtractionResult, subName: string): Edge[] {
  const sub = fn(result, subName);
  if (!sub) return [];
  return result.edges.filter(
    (e) => e.kind === 'event-handler' && e.source === sub.id,
  );
}

const formCode = extractCode(FORM_CLS);
const formLayout = extractLayout(FORM_TXT);
const reportCode = extractCode(REPORT_CLS);
const reportLayout = extractLayout(REPORT_TXT);

// =============================================================================
// AC #1 — `Private Sub Form_Load()` in `Form_X.cls` → `event-handler` edge to
// the `Form_X` layout node, `eventName: 'Load'`, `scope: 'form'`.
// =============================================================================
describe('issue-247 AC#1: form-level handlers bind to the sibling form-layout node', () => {
  const realLayout = formLayout.nodes.find((n) => n.kind === 'form-layout');

  it('the sibling .form.txt really does emit exactly one form-layout node', () => {
    // Guards the fixture itself — every id comparison below is against
    // this node, so an empty or duplicated layout extract would make the
    // rest of the suite vacuous.
    const layouts = formLayout.nodes.filter((n) => n.kind === 'form-layout');
    expect(layouts).toHaveLength(1);
  });

  it.each([
    ['Form_Load', 'Load'],
    ['Form_Open', 'Open'],
    ['Form_Unload', 'Unload'],
  ])(
    '%s emits one event-handler edge with eventName=%s and scope=form',
    (subName, eventName) => {
      const edges = handlerEdges(formCode, subName);
      expect(
        edges,
        `expected exactly one event-handler edge from ${subName}`,
      ).toHaveLength(1);
      const edge = edges[0]!;
      expect(edge.metadata?.eventName).toBe(eventName);
      expect(edge.metadata?.scope).toBe('form');
      expect(edge.provenance).toBe('heuristic');
      // The target id must be the one the REAL form extractor produces for
      // the sibling file — that identity is what makes INSERT OR REPLACE
      // converge on a single node instead of leaving an orphan stub.
      expect(edge.target).toBe(realLayout?.id);
    },
  );

  it('the local stub the code-behind emits matches the real layout node id and kind', () => {
    const stub = formCode.nodes.find((n) => n.kind === 'form-layout');
    expect(stub, 'expected a local form-layout stub in the .cls extract').toBeDefined();
    expect(stub?.id).toBe(realLayout?.id);
    expect(stub?.name).toBe(realLayout?.name);
    expect(stub?.filePath).toBe(FORM_TXT);
    // Exactly one distinct stub id, however many form-level handlers the
    // class declares — three handlers must not produce three different ids.
    const stubs = formCode.nodes.filter((n) => n.kind === 'form-layout');
    expect(new Set(stubs.map((n) => n.id)).size).toBe(1);
  });
});

// =============================================================================
// AC #2 — `Private Sub Report_Open()` in `Report_Y.cls` → the report layout
// node (kind `report-layout`, sibling `.report.txt`).
// =============================================================================
describe('issue-247 AC#2: report-level handlers bind to the sibling report-layout node', () => {
  it('Report_Open targets the real report-layout node', () => {
    const realLayout = reportLayout.nodes.find((n) => n.kind === 'report-layout');
    expect(realLayout, 'expected a report-layout node from the .report.txt').toBeDefined();

    const edges = handlerEdges(reportCode, 'Report_Open');
    expect(edges).toHaveLength(1);
    const edge = edges[0]!;
    expect(edge.target).toBe(realLayout?.id);
    expect(edge.metadata?.eventName).toBe('Open');
    expect(edge.metadata?.scope).toBe('form');
  });

  it('the report stub is a report-layout, never a form-layout', () => {
    expect(reportCode.nodes.filter((n) => n.kind === 'form-layout')).toHaveLength(0);
    expect(reportCode.nodes.filter((n) => n.kind === 'report-layout')).toHaveLength(1);
  });
});

// =============================================================================
// AC #3 — `Private Sub Form_Helper()` → NO edge. The gate is
// `isAccessEventName`, not the `Form_` prefix.
// =============================================================================
describe('issue-247 AC#3: a Form_-prefixed method that is not an Access event gets no edge', () => {
  it('Form_Helper emits no event-handler edge', () => {
    expect(fn(formCode, 'Form_Helper'), 'fixture must declare Form_Helper').toBeDefined();
    expect(handlerEdges(formCode, 'Form_Helper')).toHaveLength(0);
  });

  it('Report_Helper emits no event-handler edge', () => {
    expect(fn(reportCode, 'Report_Helper')).toBeDefined();
    expect(handlerEdges(reportCode, 'Report_Helper')).toHaveLength(0);
  });
});

// =============================================================================
// AC #4 — a control handler is completely unchanged: same node id, same edge.
// =============================================================================
describe('issue-247 AC#4: control handlers are untouched', () => {
  it('cmdSave_Click still targets the real form-instance-control node id', () => {
    const realControl = formLayout.nodes.find(
      (n) => n.kind === 'form-instance-control' && n.name === 'cmdSave',
    );
    expect(realControl, 'expected a cmdSave control node from the .form.txt').toBeDefined();

    const edges = handlerEdges(formCode, 'cmdSave_Click');
    expect(edges).toHaveLength(1);
    const edge = edges[0]!;
    expect(edge.target).toBe(realControl?.id);
    expect(edge.metadata?.eventName).toBe('Click');
    // A control handler carries NO scope — that field is what lets a
    // consumer separate the two populations.
    expect(edge.metadata?.scope).toBeUndefined();
  });

  it('txtTotal_Click on the report side is likewise unchanged', () => {
    const realControl = reportLayout.nodes.find(
      (n) => n.kind === 'form-instance-control' && n.name === 'txtTotal',
    );
    const edges = handlerEdges(reportCode, 'txtTotal_Click');
    expect(edges).toHaveLength(1);
    expect(edges[0]?.target).toBe(realControl?.id);
    expect(edges[0]?.metadata?.scope).toBeUndefined();
  });
});

// =============================================================================
// AC #5 — `form-instance-control` node count must not move. A bogus control
// node named `Form` (or `Report`) is the exact failure this change avoids.
// =============================================================================
describe('issue-247 AC#5: no bogus Form/Report control node is created', () => {
  it('the form code-behind emits one control stub — cmdSave, and nothing else', () => {
    const controls = formCode.nodes.filter((n) => n.kind === 'form-instance-control');
    expect(controls.map((n) => n.name)).toEqual(['cmdSave']);
  });

  it('the report code-behind emits one control stub — txtTotal, and nothing else', () => {
    const controls = reportCode.nodes.filter((n) => n.kind === 'form-instance-control');
    expect(controls.map((n) => n.name)).toEqual(['txtTotal']);
  });

  it('no control stub is named Form or Report in either code-behind', () => {
    const bogus = [...formCode.nodes, ...reportCode.nodes].filter(
      (n) =>
        n.kind === 'form-instance-control' &&
        ['form', 'report'].includes(n.name.toLowerCase()),
    );
    expect(bogus, 'a form-level event must never synthesize a control node').toHaveLength(0);
  });
});

// =============================================================================
// AC #6 — a non-code-behind `.cls` is still skipped entirely. A service class
// with an underscore method must not gain a layout stub just because the new
// branch exists.
// =============================================================================
describe('issue-247 AC#6: the Form_/Report_ basename guard still holds', () => {
  it('a plain service class with a Form_Load method emits no event-handler edge', () => {
    const servicePath = path.join(FIXTURE_DIR, 'ServicioInformes.cls');
    const source = [
      'Attribute VB_Name = "ServicioInformes"',
      'Option Explicit',
      '',
      'Public Sub Form_Load()',
      '    Debug.Print "not code-behind"',
      'End Sub',
      '',
      'Public Sub GenerarHTML_Principal()',
      '    Debug.Print "not an event"',
      'End Sub',
      '',
    ].join('\r\n');
    // In-memory source with a path that is NOT `Form_*` / `Report_*`. The
    // extractor never reads from disk, so no fixture file is needed.
    const result = new VbaExtractor(servicePath, source).extract();
    expect(fn(result, 'Form_Load')).toBeDefined();
    expect(result.edges.filter((e) => e.kind === 'event-handler')).toHaveLength(0);
    expect(result.nodes.filter((n) => n.kind === 'form-layout')).toHaveLength(0);
    expect(result.nodes.filter((n) => n.kind === 'form-instance-control')).toHaveLength(0);
  });
});

// =============================================================================
// AC #7 — a code-behind whose layout file was never exported still gets the
// node, and the id is still derived from the sibling PATH.
//
// This is deliberate, and it is where most of the corpus benefit comes from:
// a real Dysflow export can carry a `Form_X.cls` whose `Form_X.form.txt` was
// not exported, and those forms are the majority in some projects. The
// synthesized layout node is the only node the form has, so dropping it would
// drop the handler edge with it. It also normalizes to the same Access object
// identity a `DoCmd.OpenForm "X"` target does, so the navigation stub can
// resolve onto it.
//
// This mirrors the control branch's long-standing behaviour: a handler whose
// control is missing from the sibling still gets its stub.
// =============================================================================
describe('issue-247 AC#7: a code-behind with no exported layout still gets its node', () => {
  it('a Form_*.cls with no .form.txt beside it still emits the layout node and edge', () => {
    const orphanPath = path.join(FIXTURE_DIR, 'Form_NoLayoutSibling.cls');
    expect(
      fs.existsSync(path.join(FIXTURE_DIR, 'Form_NoLayoutSibling.form.txt')),
      'the fixture directory must NOT contain this sibling',
    ).toBe(false);

    const source = [
      'Attribute VB_Name = "Form_NoLayoutSibling"',
      'Option Explicit',
      '',
      'Private Sub Form_Load()',
      'End Sub',
      '',
    ].join('\r\n');
    const result = new VbaExtractor(orphanPath, source).extract();
    expect(fn(result, 'Form_Load')).toBeDefined();

    const layouts = result.nodes.filter((n) => n.kind === 'form-layout');
    expect(layouts).toHaveLength(1);
    // The id is keyed on the sibling `.form.txt` path, never on the `.cls`
    // path — that is what lets the real node overwrite it if the layout is
    // exported later.
    expect(layouts[0]?.filePath).toBe(
      path.join(FIXTURE_DIR, 'Form_NoLayoutSibling.form.txt'),
    );
    expect(handlerEdges(result, 'Form_Load')).toHaveLength(1);
  });
});
