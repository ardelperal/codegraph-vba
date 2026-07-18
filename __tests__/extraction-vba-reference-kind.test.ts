/**
 * extraction-vba-reference-kind.test.ts
 *
 * Round-3 (issue #108): classify `unresolved_refs.reference_kind` by
 * syntactic shape. Tests cover:
 *
 *   Test 1 (FR-2.1): Paren-form `Foo(args)` when unresolved → 'call'
 *   Test 2 (FR-1.1): `Me!Ctl` (bang read)                  → 'bang-get'
 *   Test 3 (FR-1.1): `Me.Ctl` (dot read)                   → 'property-get'
 *   Test 4 (DAO round-4): SKIPPED — TODO marker
 *   Test 5 (FR-1.2): `Me!Ctl = value` (bang assignment)    → 'bang-set'
 *   Test 6 (FR-1.4): `DoCmd.OpenQuery "Q"`                 → 'dao-query'
 *   Test 7 (FR-3.1): `If SomeConst Then` disambiguation    → 'unqualified-ident'
 *   Test 8 (FR-5):   Back-compat — at least one path still emits 'references'
 *
 * Pattern: instantiate `VbaExtractor` directly with an inline VBA source
 * string and inspect `result.unresolvedReferences`. Mirrors
 * `__tests__/extraction-vba-control-modeling.test.ts` fixture style.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { UnresolvedReference, ReferenceKind } from '../src/types';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

function refsFor(r: { unresolvedReferences: UnresolvedReference[] }): UnresolvedReference[] {
  return r.unresolvedReferences;
}

function findRefByName(
  r: { unresolvedReferences: UnresolvedReference[] },
  referenceName: string,
): UnresolvedReference | undefined {
  return r.unresolvedReferences.find((u) => u.referenceName === referenceName);
}

describe('reference-kind classification: paren-form call (FR-2.1)', () => {
  it('Foo(args) unresolved → referenceKind = "call"', () => {
    // `HelperFunction` is referenced paren-form (`Name(...)`) but NEVER
    // declared in this file, so `scanCallSites` falls through to silent-skip
    // today. Round-3 must instead emit a 'call' unresolved reference.
    const src = `Attribute VB_Name = "modCaller"
Public Sub Caller()
    HelperFunction(42)
End Sub
`;
    const r = extract('src/modCaller.bas', src);
    const ref = findRefByName(r, 'HelperFunction');
    expect(ref, 'expected unresolved ref for HelperFunction when unresolved').toBeDefined();
    expect(ref?.referenceKind).toBe('call');
  });
});

describe('reference-kind classification: Me-control reads (FR-1.1)', () => {
  it('Me!Ctl (bang read) → referenceKind = "bang-get"', () => {
    const src = `Attribute VB_Name = "Form_F"
Public Sub Foo()
    Dim v As Variant
    v = Me!SubCtl
End Sub
`;
    const r = extract('src/Form_F.cls', src);
    const ref = findRefByName(r, 'SubCtl');
    expect(ref).toBeDefined();
    expect(ref?.referenceKind).toBe('bang-get');
  });

  it('Me.Ctl (dot read) → referenceKind = "property-get"', () => {
    const src = `Attribute VB_Name = "Form_F"
Public Sub Foo()
    Dim x As String
    x = Me.Name
End Sub
`;
    const r = extract('src/Form_F.cls', src);
    const ref = findRefByName(r, 'Name');
    expect(ref).toBeDefined();
    expect(ref?.referenceKind).toBe('property-get');
  });
});

describe('reference-kind classification: DAO round-4 follow-up (FR-4)', () => {
  // FR-4 specifies that round-3 ships without a DAO field scanner; the
  // `dao-field-get`/`dao-field-set` literals in the classification table
  // are documented as deferred. This test stays skipped until round-4.
  it.skip('rs!Field DAO access → "dao-field-get" (TODO round-4 DAO scanner, issue #108 follow-up)', () => {
    const src = `Attribute VB_Name = "modDAO"
Public Sub Foo()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT IDContingencia FROM TbX")
    Dim x As Long
    x = rs!IDContingencia
End Sub
`;
    const r = extract('src/modDAO.bas', src);
    const ref = findRefByName(r, 'IDContingencia');
    expect(ref).toBeDefined();
    expect(ref?.referenceKind).toBe('dao-field-get');
  });
});

describe('reference-kind classification: Me-control assignments (FR-1.2)', () => {
  it('Me!Ctl = value (bang assignment) → referenceKind = "bang-set"', () => {
    const src = `Attribute VB_Name = "Form_F"
Public Sub Foo()
    Me!SubCtl = "X"
End Sub
`;
    const r = extract('src/Form_F.cls', src);
    const ref = findRefByName(r, 'SubCtl');
    expect(ref).toBeDefined();
    expect(ref?.referenceKind).toBe('bang-set');
  });
});

describe('reference-kind classification: DoCmd.OpenQuery (FR-1.4)', () => {
  it('DoCmd.OpenQuery "Q" → referenceKind = "dao-query"', () => {
    const src = `Attribute VB_Name = "modQ"
Public Sub Foo()
    DoCmd.OpenQuery "MyQuery"
End Sub
`;
    const r = extract('src/modQ.bas', src);
    const ref = findRefByName(r, 'MyQuery');
    expect(ref).toBeDefined();
    expect(ref?.referenceKind).toBe('dao-query');
  });
});

describe('reference-kind classification: bare ident disambiguation (FR-3.1)', () => {
  it('SomeUndeclaredSub (statement-form, not in locals) → referenceKind = "unqualified-ident"', () => {
    // FR-3.1 disambiguation: when `detectStatementCall` returns a non-null
    // bare identifier (no paren after) that is not a same-file function and
    // not a Const, classify as 'unqualified-ident' (NOT 'call') to avoid
    // the round-1/round-2 false-positive trap on statement-form Sub calls
    // whose target is undeclared.
    //
    // The fixture declares `Const HAY_ERROR = False` at module level so
    // the const-first branch of the rule can be exercised — the line
    //   `HayErrorEnRiesgo`
    // is NOT in `localConstants` and is NOT a same-file function, so it
    // must reach `unresolved_refs` as `unqualified-ident`. The const
    // itself (`HAY_ERROR`) should NOT appear as unresolved — it's resolved
    // by `resolveLocalConst`.
    const src = `Attribute VB_Name = "modRiesgo"
Private Const HAY_ERROR As Boolean = False

Public Sub Foo()
    HayErrorEnRiesgo
End Sub
`;
    const r = extract('src/modRiesgo.bas', src);
    // Confirm the const declaration does NOT bleed into unresolved_refs
    // (negative control on the disambiguation rule's const-first branch).
    expect(r.unresolvedReferences.find((u) => u.referenceName === 'HAY_ERROR')).toBeUndefined();
    // The bare ident on its own line must reach unresolved_refs as
    // 'unqualified-ident' (not 'call', not 'references').
    const ref = findRefByName(r, 'HayErrorEnRiesgo');
    expect(ref, 'expected unresolved ref for bare-ident HayErrorEnRiesgo').toBeDefined();
    expect(ref?.referenceKind).toBe('unqualified-ident');
  });

  it('VBA statement keywords never become unqualified-ident references', () => {
    const src = `Attribute VB_Name = "modKeywords"
Public Sub Foo()
    End
    On Error GoTo Handler
    Exit Sub
    Set value = Nothing
    Resume Next
    Kill "temp.txt"
    Close #1
    Open "temp.txt" For Output As #1
    Print #1, "value"
Handler:
End Sub
`;
    const r = extract('src/modKeywords.bas', src);
    const keywordRefs = r.unresolvedReferences.filter(
      (u) => u.referenceKind === 'unqualified-ident',
    );
    expect(keywordRefs).toEqual([]);
  });
});

describe('reference-kind classification: back-compat (FR-5)', () => {
  it('legacy "references" literal stays in the ReferenceKind union (FR-5 forward-compat)', () => {
    // FR-5 forward-compat guarantee: round-3 explicitly preserves the literal
    // `'references'` in the `ReferenceKind` union so any push to
    // `ctx.unresolvedReferences` from a path this round does NOT reclassify
    // (e.g. a future emitter, the index-time `resurrectRefFromDroppedEdge`
    // path in `src/extraction/index.ts`) continues to compile and produce a
    // row carrying `reference_kind = 'references'` in the SQL layer.
    //
    // After round-3 ships, every VBA-push-site the round enumerates emits a
    // shape-based literal (`call`, `bang-get`, …). The literal `'references'`
    // is reachable on rows from `resurrectRefFromDroppedEdge` (which uses
    // `e.metadata?.refKind ?? e.kind` and falls back to `e.kind`), but that
    // path lives in `src/extraction/index.ts` (an integration re-index flow),
    // not in `VbaExtractor.extract()` itself — exercising it requires a full
    // `CodeGraph.indexAll()` round-trip. The pure-compile-time assertion
    // below proves the literal is still accepted by the TS compiler.
    const legacyLiteral: ReferenceKind = 'references';
    expect(legacyLiteral).toBe('references');

    // AND the literal remains valid when constructing an UnresolvedReference
    // directly. This is the type-safety boundary FR-4.1 promises: any push
    // site I don't enumerate this round (or any future contributor writes)
    // can keep using `'references'` without a tsc error.
    const ref: UnresolvedReference = {
      fromNodeId: 'synthetic:test:fn:1',
      referenceName: 'SomeForwardCompatShape',
      referenceKind: 'references',
      line: 1,
      column: 0,
      filePath: 'src/test.bas',
      language: 'vba',
    };
    expect(ref.referenceKind).toBe('references');
  });
});
