/**
 * Issue #207: DIM_DECL_PREFIX_RE lookahead omits Declare/Event/Enum/Type.
 *
 * Before the fix the negative lookahead at `src/extraction/vba/dims.ts:27` was
 * `Function|Sub|Property|Const|WithEvents`. Any line starting with one of the
 * five `Dim|Private|Public|Global|Static` visibility keywords followed by
 * `Declare`, `Event`, `Enum`, or `Type` therefore passed the gate, and the
 * follow-up `DIM_ALL_VARS_RE` / `BARE_DIM_VAR_RE` sweep then:
 *
 *   - emitted phantom `vba-name-resolution` `references` edges attributed to
 *     the module, for the parameter types on a `Declare` / `Event` line; and
 *   - registered the header keyword (`Enum`, `Type`) or the first Declare /
 *     Event parameter name into `localVarTypeMap` file-globally.
 *
 * The `localVarTypeMap` pollution is the worse half because
 * `shouldProcessQualifiedCall` gates on it (#205 territory): a Declare
 * parameter name that happens to collide with a real local declared as `Dim`
 * elsewhere in the file silently changes that local's call-emission behavior.
 *
 * Acceptance criteria mirrored from the issue body, one regression test per
 * excluded keyword, plus a regression guard that an ordinary `Dim|Private|
 * Public|Global|Static` declaration still works after the fix.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaExtractorContext } from '../src/extraction/vba/context';
import { sweepDimsAndWithEvents } from '../src/extraction/vba/dims';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

/** Names emitted by `vba-name-resolution` edges in a given extraction result. */
function referencedTypeNames(r: ReturnType<typeof extract>): string[] {
  return r.edges
    .filter(
      (e) =>
        e.kind === 'references' &&
        e.metadata?.synthesizedBy === 'vba-name-resolution',
    )
    .map((e) => r.nodes.find((n) => n.id === e.target)?.name)
    .filter((n): n is string => Boolean(n));
}

/**
 * Run the Dim / WithEvents sweep on a single source line and return the
 * resulting `localVarTypeMap`. Used to assert that the Dim sweep does NOT
 * register `Enum`, `Type`, `Declare`, `Event`, or any of their parameter
 * names into the map — the pollution half of the issue. Direct unit
 * access avoids relying on downstream observable side effects (which are
 * hard to construct for `Enum` / `Type` because the bare-Dim path always
 * registers outer='variant', a primitive, so no edges are emitted).
 */
function dimMapFor(source: string): Map<
  string,
  { outer: string; qualified: boolean }
> {
  const ctx = new VbaExtractorContext('src/modules/m.bas');
  sweepDimsAndWithEvents(ctx, source);
  return ctx.localVarTypeMap;
}

describe('Issue #207 — DIM_DECL_PREFIX_RE lookahead must exclude Declare/Event/Enum/Type', () => {
  describe('Declare header (issue repro)', () => {
    it('emits no phantom vba-name-resolution references edge for Declare parameter types', () => {
      const src =
        'Private Declare PtrSafe Function Foo Lib "k32" (ByVal h As MyHandleClass, ByVal c As OtherClass) As Long';
      const r = extract('src/modules/modApi.bas', src);
      const refs = referencedTypeNames(r);
      // The exact case from the issue body. Before the fix the extractor
      // emitted a `references` edge for BOTH parameter types.
      expect(refs).not.toContain('MyHandleClass');
      expect(refs).not.toContain('OtherClass');
    });

    it('Declare parameter names do NOT enter localVarTypeMap (no MyHandleClass.Method stub from a later `h.Method(1)`)', () => {
      // Indirect proof of the map being clean: if `h` were polluted by
      // the Declare line, a later qualified call `h.Method(1)` inside a
      // Sub would synthesize a stub named `MyHandleClass.Method` (the
      // `receiverType` resolves via `localVarTypeMap`). With the fix, the
      // map is clean and `resolveReceiverType('h')` returns the raw `h`,
      // so the synthesized stub is `h.Method`, not `MyHandleClass.Method`.
      const src = [
        'Private Declare PtrSafe Function Foo Lib "k32" (ByVal h As MyHandleClass) As Long',
        'Sub Caller()',
        '  h.Method(1)',
        'End Sub',
      ].join('\n');
      const r = extract('src/modules/modApi.bas', src);
      const staleReceiverEdges = r.edges.filter(
        (e) =>
          e.kind === 'calls' &&
          e.provenance === 'heuristic' &&
          e.metadata?.synthesizedBy === 'vba-name-resolution' &&
          e.metadata?.receiverType === 'MyHandleClass',
      );
      expect(staleReceiverEdges).toHaveLength(0);
      const staleStubNodes = r.nodes.filter(
        (n) => n.kind === 'function' && n.name === 'MyHandleClass.Method',
      );
      expect(staleStubNodes).toHaveLength(0);
    });

    it('still emits the `declare` node via the events/types/declares classifier (regression: Declare sweep unaffected)', () => {
      const src =
        'Private Declare PtrSafe Function Foo Lib "k32" () As Long';
      const r = extract('src/modules/modApi.bas', src);
      const decl = r.nodes.find(
        (n) => n.kind === 'declare' && n.name === 'Foo',
      );
      expect(decl).toBeDefined();
      expect(decl?.metadata?.isDeclare).toBe(true);
      expect(decl?.metadata?.dll).toBe('k32');
    });
  });

  describe('Event header', () => {
    it('emits no phantom vba-name-resolution references edge for Event parameter types', () => {
      const src = 'Public Event SomethingHappened(ByVal payload As MyPayload)';
      const r = extract('src/modules/clsEvt.cls', src);
      const refs = referencedTypeNames(r);
      expect(refs).not.toContain('MyPayload');
    });

    it('Event parameter names do NOT enter localVarTypeMap', () => {
      // Direct unit-level check on the Dim sweep. Before the fix the
      // sweep captured `payload` from the Event header and registered
      // it with outer='MyPayload' (non-primitive) — a polluting entry
      // that survives any later `Dim payload` of the same name.
      const map = dimMapFor(
        'Public Event SomethingHappened(ByVal payload As MyPayload)',
      );
      expect(map.has('payload')).toBe(false);
      expect(map.has('event')).toBe(false);
    });

    it('still emits the `event` node via the events/types/declares classifier (regression: Event sweep unaffected)', () => {
      const src = 'Public Event SomethingHappened(ByVal payload As String)';
      const r = extract('src/modules/clsEvt.cls', src);
      const evt = r.nodes.find(
        (n) => n.kind === 'event' && n.name === 'SomethingHappened',
      );
      expect(evt).toBeDefined();
    });
  });

  describe('Enum header', () => {
    it('"Public Enum Foo" does NOT register "Enum" as a local variable in localVarTypeMap', () => {
      // Direct unit-level check on the Dim sweep. Before the fix the
      // bare-Dim path captured the keyword `Enum` as a variable name
      // and registered `localVarTypeMap['enum'] = { outer: 'variant' }`.
      // The pollution is hard to observe downstream because `variant` is
      // a primitive (gates qualified calls), so the bug is silent in
      // the edge graph. This test pins the regression at the sweep.
      const map = dimMapFor('Public Enum Foo');
      expect(map.has('enum')).toBe(false);
    });

    it('Enum members still emit as enum_member nodes (regression: Enum sweep unaffected)', () => {
      const src = ['Public Enum Foo', '  A = 1', 'End Enum'].join('\n');
      const r = extract('src/modules/modEnum.bas', src);
      const member = r.nodes.find(
        (n) => n.kind === 'enum_member' && n.name === 'A',
      );
      expect(member).toBeDefined();
    });
  });

  describe('Type header', () => {
    it('"Public Type Foo" does NOT register "Type" as a local variable in localVarTypeMap', () => {
      // Symmetric to the Enum test above. Before the fix the bare-Dim
      // sweep registered `localVarTypeMap['type'] = { outer: 'variant' }`.
      const map = dimMapFor('Public Type Foo');
      expect(map.has('type')).toBe(false);
    });

    it('Type members still emit as type_member nodes (regression: Type sweep unaffected)', () => {
      const src = ['Public Type Foo', '  X As Long', 'End Type'].join('\n');
      const r = extract('src/modules/modUdt.bas', src);
      const member = r.nodes.find(
        (n) => n.kind === 'type_member' && n.name === 'X',
      );
      expect(member).toBeDefined();
    });
  });

  describe('Regression guards — ordinary Dim/Private/Public still work', () => {
    it('Private m_Foo As ClsFoo still emits a vba-name-resolution references edge to ClsFoo', () => {
      const src = 'Private m_Foo As ClsFoo';
      const r = extract('src/modules/m.bas', src);
      expect(referencedTypeNames(r)).toContain('ClsFoo');
    });

    it('Public gsLong As Long (primitive) still emits no vba-name-resolution references edge', () => {
      const src = 'Public gsLong As Long';
      const r = extract('src/modules/m.bas', src);
      expect(referencedTypeNames(r)).toHaveLength(0);
    });

    it('Public Const MY_CONST = 1 stays OUT of the Dim sweep (Const has its own sweepEnumsAndConsts path)', () => {
      // Regression guard carried over from the #47 family: the negative
      // lookahead MUST keep `Const` routed to `sweepEnumsAndConsts` even
      // after extending the lookahead with `Declare|Event|Enum|Type`.
      const src = 'Public Const MY_CONST = 1';
      const r = extract('src/modules/m.bas', src);
      expect(referencedTypeNames(r)).toHaveLength(0);
      const c = r.nodes.find(
        (n) => n.kind === 'constant' && n.name === 'MY_CONST',
      );
      expect(c).toBeDefined();
    });

    it('Private Function / Sub / Property still bypass the Dim sweep', () => {
      // The original reason `Function|Sub|Property` was in the lookahead:
      // procedure headers must not be misinterpreted as variable
      // declarations. After extending the lookahead with Declare/Event/
      // Enum/Type, procedures must still be exempt.
      const src = [
        'Private Function MyFn() As Long',
        '    MyFn = 0',
        'End Function',
        'Private Sub MySub()',
        'End Sub',
        'Private Property Get MyProp() As Long',
        '    MyProp = 0',
        'End Property',
      ].join('\n');
      const r = extract('src/modules/m.bas', src);
      // None of the procedure headers should leak a `MyFn As Long` style
      // vba-name-resolution edge (the only emitted reference would be
      // for a real `Dim`, which is not present in this source).
      expect(referencedTypeNames(r)).toHaveLength(0);
    });
  });
});