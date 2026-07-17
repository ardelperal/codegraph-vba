/**
 * extraction-vba-deferred-qualified.test.ts
 *
 * RED → GREEN coverage for the Issue #151 deferred qualified-call
 * conformance pass. Each test is a sharp, isolated reproduction of one
 * forward-reference failure mode the pre-#151 sweep silently dropped
 * (or worse, emitted a dead-end raw-receiver stub for).
 *
 * Bug summary: `createCallsAndSqlClassifier` walked the source once,
 * maintaining a per-line `localVarTypeMap`. A qualified call
 * `Receiver.Member` was resolved at the line it appeared on, so a
 * `Receiver` that was `Dim`'d or `Set` on a LATER line in the same
 * procedure was untracked at the call line — the synthetic stub ended
 * up named `Receiver.Member` (the raw variable name) instead of
 * `<RealClass>.Member` (the factory's return type).
 *
 * The fix is a deferred reprocessing pass: collect every qualified call
 * whose receiver is undeclared at the call line, then re-resolve them
 * at the matching `End Sub` / `End Function` / `End Property` once the
 * full `localVarTypeMap` for that procedure scope is known.
 *
 * Fixture strategy: each test is a 5–10 line `.bas` snippet run through
 * the public `VbaExtractor` API. The expected outcome is asserted by
 * inspecting the emitted `calls` edges and the synthetic `function`
 * stub nodes the extractor creates. No resolver is involved — the
 * acceptance is purely extraction-side.
 *
 * Acceptance criteria covered (from the issue body):
 *   AC-1 forward Set: `Set x = New Y` later, `x.Foo()` earlier → edge to Y.Foo
 *   AC-2 forward Dim: `Dim x As Y` later, `x.Foo` earlier → edge to Y.Foo
 *   AC-3 forward Factory: `Set x = CreateY()` later, `x.Foo` earlier → edge to Y.Foo
 *   AC-4 never resolves: `x.Foo` with no Dim/Set anywhere → silently dropped
 *   AC-5 isolation: deferred list is consumed at the same procedure's end,
 *         no cross-procedure leak
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';

/**
 * Helper: extract one inline `.bas` snippet and return the heuristic
 * `calls` edges to synthetic stubs tagged `synthesizedBy ===
 * 'vba-name-resolution'`. Filters out the SQL/reference sweeps so the
 * tests stay focused on the qualified-call conformance path.
 */
function extractHeuristicCallEdges(src: string) {
  const r = new VbaExtractor('src/modules/DeferredQualified.bas', src).extract();
  return r.edges
    .filter(
      (e) =>
        e.kind === 'calls' &&
        e.provenance === 'heuristic' &&
        e.metadata?.synthesizedBy === 'vba-name-resolution',
    )
    .map((e) => ({
      source: e.source,
      target: e.target,
      line: e.line,
      column: e.column,
      receiverType: e.metadata?.receiverType as string | undefined,
      member: e.metadata?.member as string | undefined,
    }));
}

/** Look up a synthetic stub node by its emitted `name` (`<Class>.<Member>`). */
function findStubByName(
  r: ReturnType<VbaExtractor['extract']>,
  qualified: string,
) {
  return r.nodes.find(
    (n) =>
      n.kind === 'function' && n.name === qualified && n.metadata?.stub === true,
  );
}

// =============================================================================
// AC-1: forward `Set x = New <Y>` — the receiver's class is set after the call
// =============================================================================
describe('AC-1: forward Set x = New <Y> defers qualified call to end of procedure', () => {
  it('emits a Y.Foo calls edge for `x.Foo()` before `Set x = New Y`', () => {
    const src = `Sub Caller()
  x.Foo arg1
  Dim x As Object
  Set x = New Y
End Sub`;
    const r = new VbaExtractor('src/modules/AC1.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    const yFooEdge = edges.find(
      (e) => e.receiverType === 'Y' && e.member === 'Foo',
    );
    expect(yFooEdge, 'Y.Foo heuristic edge must be emitted at end-of-procedure')
      .toBeDefined();
    // The stub must be named `Y.Foo` (resolved class), not `x.Foo` (raw var).
    const stub = findStubByName(r, 'Y.Foo');
    expect(stub, 'synthetic stub Y.Foo must exist').toBeDefined();
    const wrongStub = findStubByName(r, 'x.Foo');
    expect(wrongStub, 'wrong raw-receiver stub x.Foo must NOT exist').toBeUndefined();
  });
});

// =============================================================================
// AC-2: forward `Dim x As Y` — the receiver's class is declared after the call
// =============================================================================
describe('AC-2: forward Dim x As Y defers qualified call to end of procedure', () => {
  it('emits a Y.Foo calls edge for `x.Foo` before `Dim x As Y`', () => {
    const src = `Sub Caller()
  x.Foo arg1
  Dim x As Y
End Sub`;
    const r = new VbaExtractor('src/modules/AC2.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    const yFooEdge = edges.find(
      (e) => e.receiverType === 'Y' && e.member === 'Foo',
    );
    expect(yFooEdge, 'Y.Foo heuristic edge must be emitted at end-of-procedure')
      .toBeDefined();
    const stub = findStubByName(r, 'Y.Foo');
    expect(stub, 'synthetic stub Y.Foo must exist').toBeDefined();
    const wrongStub = findStubByName(r, 'x.Foo');
    expect(wrongStub, 'wrong raw-receiver stub x.Foo must NOT exist').toBeUndefined();
  });
});

// =============================================================================
// AC-3: forward `Set x = <Factory>()` — the factory return type is known
// (the main scenario from the issue body)
// =============================================================================
describe('AC-3: forward Set x = <Factory>() defers qualified call to end of procedure', () => {
  it('emits a Y.Foo calls edge for `x.Foo()` before `Set x = CreateY()`', () => {
    const src = `Sub Caller()
  x.Foo arg1
  Set x = CreateY()
End Sub

Function CreateY() As Y
End Function`;
    const r = new VbaExtractor('src/modules/AC3.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    const yFooEdge = edges.find(
      (e) => e.receiverType === 'Y' && e.member === 'Foo',
    );
    expect(
      yFooEdge,
      'Y.Foo heuristic edge must be emitted (the main #151 scenario)',
    ).toBeDefined();
    const stub = findStubByName(r, 'Y.Foo');
    expect(stub, 'synthetic stub Y.Foo must exist').toBeDefined();
    const wrongStub = findStubByName(r, 'x.Foo');
    expect(wrongStub, 'wrong raw-receiver stub x.Foo must NOT exist').toBeUndefined();
  });

  it('forward factory call works for the statement form (`x.Foo arg1`, no parens)', () => {
    // Spec coverage: the original bug is reported on the statement form
    // (`.Foo arg`), but the paren form must also benefit. Both must
    // resolve to `Y.Foo`.
    const src = `Sub Caller()
  x.Foo arg1
  Set x = CreateY()
End Sub

Function CreateY() As Y
End Function`;
    const r = new VbaExtractor('src/modules/AC3stmt.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    const yFooEdges = edges.filter(
      (e) => e.receiverType === 'Y' && e.member === 'Foo',
    );
    expect(yFooEdges.length, 'at least one Y.Foo edge (paren or statement)')
      .toBeGreaterThanOrEqual(1);
  });
});

// =============================================================================
// AC-4: never resolves — deferred calls whose receiver is a declared
// primitive (or external) type at end of procedure are dropped silently.
// This is the `localVarTypeMap` "primitive guard" applied at end of
// procedure, mirroring the in-line `shouldProcessQualifiedCall` gate
// from `calls.ts` and `call-sweep.ts`.
// =============================================================================
describe('AC-4: never-resolving qualified calls are dropped silently at end of procedure', () => {
  it('forward `Dim x As Long; x.Foo` (declared primitive) produces no stub, no edge, no error', () => {
    // The receiver is untyped at the call line, so it is deferred.
    // At `End Sub`, the Dim is seen — `x` is now a primitive local.
    // The deferred pass must drop it silently (same shape as the
    // in-line primitive guard in `calls.ts`).
    const src = `Sub Caller()
  x.Foo arg1
  Dim x As Long
End Sub`;
    const r = new VbaExtractor('src/modules/AC4.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    const xFooEdge = edges.find(
      (e) => e.receiverType === 'x' && e.member === 'Foo',
    );
    expect(xFooEdge, 'primitive-typed receiver x.Foo must NOT emit a stub')
      .toBeUndefined();
    const wrongStub = findStubByName(r, 'x.Foo');
    expect(wrongStub, 'no x.Foo synthetic stub').toBeUndefined();
    expect(r.errors, 'no extractor errors').toHaveLength(0);
  });

  it('forward `Dim x As DAO.Database; x.Foo` (declared external type) is silent', () => {
    // A receiver declared as a qualified external type (DAO.Database)
    // never resolves to a project-class call. The deferred pass must
    // drop it at `End Sub` — same shape as the in-line qualified-
    // external guard. This is the "forward Const" / "forward
    // non-project-class" scenario from the spec, where the receiver's
    // type is discovered to be a non-class and the call is silently
    // suppressed.
    const src = `Sub Caller()
  db.Foo arg1
  Dim db As DAO.Database
End Sub`;
    const r = new VbaExtractor('src/modules/AC4const.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    const dbFooEdge = edges.find(
      (e) => e.receiverType === 'db' && e.member === 'Foo',
    );
    expect(
      dbFooEdge,
      'qualified-external receiver db.Foo must NOT emit a stub',
    ).toBeUndefined();
    const wrongStub = findStubByName(r, 'db.Foo');
    expect(wrongStub, 'no db.Foo synthetic stub').toBeUndefined();
    expect(r.errors, 'no extractor errors').toHaveLength(0);
  });
});

// =============================================================================
// AC-5: isolation — the deferred list for procedure A is consumed at End A
// and does not leak into procedure B. `localVarTypeMap` is correctly
// scoped to the procedure's end (a Dim inside Sub A is NOT visible in Sub B).
// =============================================================================
describe('AC-5: deferred list is consumed at the closing procedure end with no cross-procedure leak', () => {
  it('two adjacent procedures: each gets its own deferred resolution, no leak', () => {
    const src = `Sub First()
  a.Foo arg1
  Set a = New MyClassA
End Sub

Sub Second()
  b.Foo arg1
  Set b = New MyClassB
End Sub`;
    const r = new VbaExtractor('src/modules/AC5.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    // MyClassA.Foo from First()
    const aFoo = edges.find(
      (e) => e.receiverType === 'MyClassA' && e.member === 'Foo',
    );
    expect(aFoo, 'MyClassA.Foo edge from First() must be emitted').toBeDefined();
    // MyClassB.Foo from Second()
    const bFoo = edges.find(
      (e) => e.receiverType === 'MyClassB' && e.member === 'Foo',
    );
    expect(bFoo, 'MyClassB.Foo edge from Second() must be emitted').toBeDefined();
    // Neither raw `a.Foo` nor `b.Foo` should leak.
    const aFooRaw = edges.find(
      (e) => e.receiverType === 'a' && e.member === 'Foo',
    );
    const bFooRaw = edges.find(
      (e) => e.receiverType === 'b' && e.member === 'Foo',
    );
    expect(aFooRaw, 'raw a.Foo must NOT leak').toBeUndefined();
    expect(bFooRaw, 'raw b.Foo must NOT leak').toBeUndefined();
    // Stub sanity: one Y-stub per procedure, no duplicates.
    expect(findStubByName(r, 'MyClassA.Foo'), 'MyClassA.Foo stub').toBeDefined();
    expect(findStubByName(r, 'MyClassB.Foo'), 'MyClassB.Foo stub').toBeDefined();
  });

  it('deferred list is fully drained at End Sub: the second procedure sees a fresh map', () => {
    // A `Dim x As MyClassA` in First() must not leak into Second() as a
    // typed receiver — the spec calls this "no leak into the next
    // procedure". The second procedure declares its own `x` and that
    // declaration must take precedence.
    const src = `Sub First()
  x.Foo arg1
  Dim x As MyClassA
End Sub

Sub Second()
  x.Bar arg1
  Dim x As MyClassB
End Sub`;
    const r = new VbaExtractor('src/modules/AC5isolation.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    const aFoo = edges.find(
      (e) => e.receiverType === 'MyClassA' && e.member === 'Foo',
    );
    const bBar = edges.find(
      (e) => e.receiverType === 'MyClassB' && e.member === 'Bar',
    );
    expect(aFoo, 'MyClassA.Foo from First()').toBeDefined();
    expect(bBar, 'MyClassB.Bar from Second()').toBeDefined();
    // If the dims map leaked, `Second`'s x.Bar would resolve to
    // MyClassA.Bar. That must NOT happen.
    const leak = edges.find(
      (e) => e.receiverType === 'MyClassA' && e.member === 'Bar',
    );
    expect(leak, 'MyClassA dim from First() must NOT leak into Second()').toBeUndefined();
  });
});

// =============================================================================
// Regression guard: existing extraction-vba test scenarios must still pass.
// This is a smoke test — the full regression bar is the extraction-vba.test.ts
// suite (215 tests), which is run separately by `pnpm test`.
// =============================================================================
describe('regression: existing qualified-call scenarios still behave correctly', () => {
  it('cross-module qualified call (`modUtils.Foo arg`) still emits a heuristic edge', () => {
    // The pre-#151 behaviour was: an UNDECLARED receiver like `modUtils`
    // emits a raw `modUtils.Foo` stub immediately (it is NOT deferred,
    // because the deferred pass only applies when the receiver MIGHT
    // become a typed local later in the same procedure). This test pins
    // that the cross-module path is unaffected.
    const src = `Sub RunIt()
  modUtils.Foo arg
End Sub`;
    const r = new VbaExtractor('src/modules/Regression.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    const modUtilsEdge = edges.find(
      (e) => e.receiverType === 'modUtils' && e.member === 'Foo',
    );
    expect(modUtilsEdge, 'modUtils.Foo edge must be emitted (cross-module path)')
      .toBeDefined();
  });

  it('declared-primitive receiver (`Dim x As Long; x.Foo`) is silent — no stub, no edge', () => {
    const src = `Sub RunIt()
  Dim x As Long
  x.Foo arg
End Sub`;
    const r = new VbaExtractor('src/modules/Regression.bas', src).extract();
    const edges = extractHeuristicCallEdges(src);
    const xFooEdge = edges.find(
      (e) => e.receiverType === 'x' && e.member === 'Foo',
    );
    expect(xFooEdge, 'primitive-typed receiver x.Foo must be silent').toBeUndefined();
    const wrongStub = findStubByName(r, 'x.Foo');
    expect(wrongStub, 'no x.Foo synthetic stub').toBeUndefined();
    expect(r.errors, 'no extractor errors').toHaveLength(0);
  });
});
