/**
 * Issue #205 — `localVarTypeMap` is keyed on the bare variable name with no
 * procedure scope, so the last `Dim` in a file wins for the whole file and
 * qualified calls resolve to the wrong class.
 *
 * The fix mirrors the existing per-procedure `localConstants` bucket
 * (`context.ts:168` and `resolveLocalConst` at `context.ts:526-535`) and
 * the per-procedure `arrayParameters` rationale spelled out at
 * `calls.ts:86-90`: a same-named variable declared in two different
 * procedures must NOT silently make one procedure's qualified calls
 * resolve to the other procedure's type.
 *
 * The GREEN fix keys the map by `${procStartLine}\0${varName}` with a
 * `'module'` bucket fallback, parallel to `localConstants`, and adds
 * a `currentVarTypeProcKey` shared scope the dims classifier maintains
 * in parallel to the call-sweep's `currentProcKey`. See
 * `src/extraction/vba/context.ts:145-180` and the dims classifier at
 * `src/extraction/vba/dims.ts` for the implementation.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';

/**
 * Helper: extract a single .bas file's worth of nodes/edges. The fixture
 * uses `Attribute VB_Name = "mod205"` so the module node is created and
 * `pendingModuleOrClassSource` re-attribution runs.
 */
function extract(src: string): ReturnType<VbaExtractor['extract']> {
  return new VbaExtractor('src/modules/mod205.bas', src).extract();
}

/**
 * Helper: collect the `provenance: 'heuristic'` `calls` edges whose
 * `metadata.synthesizedBy === 'vba-name-resolution'` (the qualified-call
 * stub path) and group them by the `metadata.receiverType` so a test can
 * assert "in proc A, `item.Guardar` resolves to type X" without
 * depending on edge order.
 */
interface StubCall {
  line: number;
  qualified: string;
  member: string;
  receiverType: string;
}
function stubCallsByLine(r: ReturnType<VbaExtractor['extract']>): StubCall[] {
  return r.edges
    .filter(
      (e) =>
        e.kind === 'calls' &&
        e.provenance === 'heuristic' &&
        e.metadata?.synthesizedBy === 'vba-name-resolution',
    )
    .map((e) => ({
      line: e.line ?? 0,
      // The qualified stub node is `<receiverType>.<member>` and is the
      // edge's target — its `name` is the dotted string.
      qualified:
        r.nodes.find((n) => n.id === e.target)?.name ?? '',
      member: e.metadata?.member as string,
      receiverType: e.metadata?.receiverType as string,
    }));
}

// ============================================================================
// 1. Two procedures with same `Dim item As X` of different types
// ============================================================================

describe('Issue #205 — localVarTypeMap is procedure-scoped', () => {
  it('two procedures declaring the same variable name with different types each resolve to their own type', () => {
    const src = [
      'Attribute VB_Name = "mod205"',
      '',
      "' Each procedure declares `Dim item As <Type>` of a different type.",
      "' The qualified call `item.Guardar` MUST resolve to that procedure's own type.",
      'Public Sub AltaProducto()',
      '    Dim item As Producto',
      '    item.Guardar',
      'End Sub',
      '',
      'Public Sub AltaCliente()',
      '    Dim item As Cliente',
      '    item.Guardar',
      'End Sub',
      '',
    ].join('\n');

    const r = extract(src);
    const calls = stubCallsByLine(r);

    // Two `item.Guardar` calls — one per procedure. Each MUST resolve to
    // the type declared IN that procedure, not the type declared in the
    // other one (and not some last-Dim-wins value).
    const producto = calls.filter((c) => c.line === 7);
    const cliente = calls.filter((c) => c.line === 12);

    expect(producto).toHaveLength(1);
    expect(cliente).toHaveLength(1);
    expect(producto[0]!.qualified).toBe('Producto.Guardar');
    expect(producto[0]!.receiverType).toBe('Producto');
    expect(cliente[0]!.qualified).toBe('Cliente.Guardar');
    expect(cliente[0]!.receiverType).toBe('Cliente');
  });

  // --------------------------------------------------------------------
  // 2. Module-level Dim still resolves inside procedures
  // --------------------------------------------------------------------

  it('module-level `Dim x As SomeType` still resolves inside procedures that do not redeclare x', () => {
    const src = [
      'Attribute VB_Name = "mod205"',
      '',
      "' Module-level `Dim` — visible to every procedure that does NOT",
      "' shadow it with its own `Dim x` declaration.",
      'Dim globalItem As GlobalThing',
      '',
      'Public Sub Reader()',
      '    globalItem.Use',
      'End Sub',
      '',
    ].join('\n');

    const r = extract(src);
    const calls = stubCallsByLine(r);
    const useCall = calls.filter((c) => c.line === 8);
    expect(useCall).toHaveLength(1);
    expect(useCall[0]!.qualified).toBe('GlobalThing.Use');
    expect(useCall[0]!.receiverType).toBe('GlobalThing');
  });

  // --------------------------------------------------------------------
  // 3. Module-level fallback when the call site is in a procedure
  //    without its own local dim of that name
  // --------------------------------------------------------------------

  it('falls back to the module bucket when the call site has no proc-local declaration', () => {
    // `globalItem` is declared at module level. `Sub NoShadow` does NOT
    // redeclare it, so the call-site lookup must use the module bucket.
    const src = [
      'Attribute VB_Name = "mod205"',
      '',
      'Dim globalItem As GlobalThing',
      '',
      "' `RedundantProc` declares `localItem` — only that name is in its",
      "' own bucket. The module-level `globalItem` is the only type for `globalItem`.",
      'Public Sub RedundantProc()',
      '    Dim localItem As LocalThing',
      '    localItem.Use',
      '    globalItem.Use',
      'End Sub',
      '',
    ].join('\n');

    const r = extract(src);
    const calls = stubCallsByLine(r);
    const localCall = calls.filter((c) => c.line === 9);
    const globalCall = calls.filter((c) => c.line === 10);
    expect(localCall).toHaveLength(1);
    expect(localCall[0]!.qualified).toBe('LocalThing.Use');
    expect(localCall[0]!.receiverType).toBe('LocalThing');
    expect(globalCall).toHaveLength(1);
    expect(globalCall[0]!.qualified).toBe('GlobalThing.Use');
    expect(globalCall[0]!.receiverType).toBe('GlobalThing');
  });

  // --------------------------------------------------------------------
  // 4. Mixed: proc-local shadows module-level. The proc-local WINS
  //    inside the proc; the module-level is untouched.
  // --------------------------------------------------------------------

  it('a proc-local `Dim x` shadows the module-level declaration of the same name inside the proc', () => {
    const src = [
      'Attribute VB_Name = "mod205"',
      '',
      'Dim item As ModuleThing',
      '',
      'Public Sub WithShadow()',
      '    Dim item As ProcThing',
      '    item.Guardar',
      'End Sub',
      '',
      "' `OutsideProc` does NOT shadow — its `item.Guardar` uses the",
      "' module-level `ModuleThing`.",
      'Public Sub OutsideProc()',
      '    item.Guardar',
      'End Sub',
      '',
    ].join('\n');

    const r = extract(src);
    const calls = stubCallsByLine(r);
    const procCall = calls.filter((c) => c.line === 7);
    const outsideCall = calls.filter((c) => c.line === 13);
    expect(procCall).toHaveLength(1);
    expect(procCall[0]!.qualified).toBe('ProcThing.Guardar');
    expect(procCall[0]!.receiverType).toBe('ProcThing');
    expect(outsideCall).toHaveLength(1);
    expect(outsideCall[0]!.qualified).toBe('ModuleThing.Guardar');
    expect(outsideCall[0]!.receiverType).toBe('ModuleThing');
  });

  // --------------------------------------------------------------------
  // 5. The exact reproduction from issue #205's body. The bug is that
  //    BOTH call sites resolved to the last-declared `Cliente` (the
  //    second `Dim` won for the whole file). The fix gives each its
  //    own type.
  // --------------------------------------------------------------------

  it('reproduces issue #205 verbatim and resolves each call to its own procedure\'s type', () => {
    const src = [
      'Attribute VB_Name = "mod205"',
      '',
      'Public Sub AltaProducto()',
      '    Dim item As Producto',
      '    item.Guardar',
      'End Sub',
      '',
      'Public Sub AltaCliente()',
      '    Dim item As Cliente',
      '    item.Guardar',
      'End Sub',
      '',
    ].join('\n');

    const r = extract(src);
    const calls = stubCallsByLine(r);
    // Find the two `item.Guardar` calls.
    const guardar = calls.filter((c) => c.member === 'Guardar');
    expect(guardar).toHaveLength(2);
    // Each call site's `receiverType` MUST be the type its own procedure
    // declared — NOT a shared "last Dim wins" stub. The pre-fix code
    // would emit BOTH with `receiverType: 'Cliente'` (the last Dim).
    const types = guardar.map((c) => c.receiverType).sort();
    expect(types).toEqual(['Cliente', 'Producto']);
  });
});
