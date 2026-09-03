/**
 * Issue #263 (task E6 of `docs/vba-error-handling-plan.md`) — `label` nodes
 * and `handles-error` edges.
 *
 * #259 records *whether* a procedure has a handler; #260 marks *which* edges
 * come from inside one. Neither gives the handler an identity you can point
 * at, search for or traverse to. This does — and the risk it carries is the
 * reason most of the assertions below exist:
 *
 *   1. **Collision.** VBA scopes a line label to its procedure and this
 *      corpus writes the same label everywhere (`errores` is defined 3,735
 *      times). Without the procedure segment in `qualifiedName`, every
 *      handler in a project collapses into one symbol. `two procedures in
 *      one module` is the guard.
 *   2. **The precision gate.** A label nobody targets with `On Error GoTo`
 *      is control flow, not a handler. Calling one a handler is the worst
 *      answer available here — a confidently wrong "yes" to "does this
 *      handle errors".
 *   3. **No fabricated targets.** `On Error GoTo noExiste` must leave the
 *      graph honest: an unresolved reference and NO node. A graph that
 *      invents its own targets cannot be used to find the defect.
 *   4. **No re-parenting.** Calls inside a handler stay attributed to the
 *      enclosing procedure. `callers`/`callees` for every procedure with a
 *      handler must be untouched.
 *   5. **Flood control.** ~3,900 label nodes must stay out of the two
 *      default result surfaces, exactly as #257's parameters did.
 */
import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { RULES } from '../src/extraction/vba/errors';
import { parseQuery } from '../src/search/query-parser';
import { Edge, Node, UnresolvedReference } from '../src/types';

/**
 * Extract a `.bas` module. The two header lines are prepended here so a
 * fixture's own first line is source line 3 — the convention
 * `extraction-vba-error-policy.test.ts` and
 * `extraction-vba-error-handler-region.test.ts` both use.
 */
function extract(body: string[], filePath = 'src/modules/ModErrores.bas') {
  return new VbaExtractor(
    filePath,
    ['Attribute VB_Name = "ModErrores"', 'Option Explicit', ...body].join('\n'),
  ).extract();
}

function labels(nodes: Node[]): Node[] {
  return nodes.filter((n) => n.kind === 'label');
}

function label(nodes: Node[], qualifiedName: string): Node | undefined {
  return labels(nodes).find((n) => n.qualifiedName === qualifiedName);
}

/** Declared procedures only — call-target stubs carry `metadata.stub`. */
function procedure(nodes: Node[], name: string): Node | undefined {
  return nodes.find(
    (n) => n.kind === 'function' && n.metadata?.stub !== true && n.name === name,
  );
}

function edgesOfKind(edges: Edge[], kind: string, sourceId?: string): Edge[] {
  return edges.filter(
    (e) => e.kind === kind && (sourceId === undefined || e.source === sourceId),
  );
}

function refsTo(refs: UnresolvedReference[], name: string): UnresolvedReference[] {
  return refs.filter((r) => r.referenceName === name);
}

// ============================================================================
// 1. Node shape
// ============================================================================

describe('Issue #263 — label node shape', () => {
  it('emits one node per label definition, scoped to module and procedure', () => {
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    p_Error = "boom"',
      'End Sub',
    ]);

    const found = labels(nodes);
    expect(found).toHaveLength(1);
    const node = found[0]!;
    expect(node.name).toBe('errores');
    expect(node.qualifiedName).toBe('ModErrores.Guardar.errores');
    expect(node.filePath).toBe('src/modules/ModErrores.bas');
    expect(node.language).toBe('vba');
    // Header (2 lines) + `Public Sub` (3) + `On Error` (4) + `Call` (5) +
    // `Exit Sub` (6) → the label is line 7.
    expect(node.startLine).toBe(7);
  });

  it('spans a handler label to the procedure end and a control-flow label to its own line', () => {
    const { nodes } = extract([
      'Public Sub Recorrer()',
      '    On Error GoTo errores',
      '    Call Primero',
      'siguiente:',
      '    Call Segundo',
      '    Exit Sub',
      'errores:',
      '    p_Error = "boom"',
      'End Sub',
    ]);

    const siguiente = label(nodes, 'ModErrores.Recorrer.siguiente');
    const errores = label(nodes, 'ModErrores.Recorrer.errores');
    expect(siguiente).toBeDefined();
    expect(errores).toBeDefined();

    // `siguiente:` is line 6, `errores:` is line 9, `End Sub` is line 11.
    expect(siguiente!.startLine).toBe(6);
    expect(siguiente!.endLine).toBe(6);
    expect(errores!.startLine).toBe(9);
    expect(errores!.endLine).toBe(11);
  });

  it('copies handlerBehavior and the region lines from the procedure errorPolicy', () => {
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    p_Error = "boom"',
      '    MsgBox "boom"',
      'End Sub',
    ]);

    const policy = procedure(nodes, 'Guardar')!.metadata!.errorPolicy as {
      behavior: string | null;
      handlerStartLine: number | null;
      handlerEndLine: number | null;
    };
    const node = label(nodes, 'ModErrores.Guardar.errores')!;

    // Copied, never re-derived: the node must agree with the policy #260
    // already published, whatever that policy says.
    expect(policy.behavior).toBe('mixed');
    expect(node.metadata?.handlerBehavior).toBe(policy.behavior);
    expect(node.metadata?.regionStartLine).toBe(policy.handlerStartLine);
    expect(node.metadata?.regionEndLine).toBe(policy.handlerEndLine);
  });
});

// ============================================================================
// 2. The collision trap — the whole reason qualifiedName carries the procedure
// ============================================================================

describe('Issue #263 — procedure-scoped identity', () => {
  it('gives two procedures in one module that both define `errores` distinct nodes', () => {
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    p_Error = "guardar"',
      'End Sub',
      '',
      'Public Sub Borrar()',
      '    On Error GoTo errores',
      '    Call Eliminar',
      '    Exit Sub',
      'errores:',
      '    p_Error = "borrar"',
      'End Sub',
    ]);

    const found = labels(nodes);
    expect(found).toHaveLength(2);
    expect(found.map((n) => n.name)).toEqual(['errores', 'errores']);

    // Distinct qualified names…
    expect(found.map((n) => n.qualifiedName).sort()).toEqual([
      'ModErrores.Borrar.errores',
      'ModErrores.Guardar.errores',
    ]);
    // …and distinct ids. Without either, 3,735 handlers become one symbol.
    expect(new Set(found.map((n) => n.id)).size).toBe(2);

  });

  it('keeps the module prefix off a file with no VB_Name and still scopes to the procedure', () => {
    const { nodes } = new VbaExtractor(
      'src/modules/Sin Nombre.bas',
      ['Public Sub Guardar()', '    On Error GoTo errores', 'errores:', 'End Sub'].join(
        '\n',
      ),
    ).extract();

    // No `Attribute VB_Name` → `ctx.moduleName` falls back to the basename,
    // which is still a real scope. What must never happen is a bare `errores`.
    const node = labels(nodes)[0]!;
    expect(node.qualifiedName).toMatch(/\.Guardar\.errores$/);
    expect(node.qualifiedName).not.toBe('errores');
  });
});

// ============================================================================
// 3. isHandler — the precision gate
// ============================================================================

describe('Issue #263 — isHandler', () => {
  it('marks an On Error GoTo target as a handler', () => {
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      'errores:',
      '    p_Error = "boom"',
      'End Sub',
    ]);
    expect(label(nodes, 'ModErrores.Guardar.errores')!.metadata?.isHandler).toBe(
      true,
    );
  });

  it('marks a label nobody targets as control flow, with no region and no handles-error edge', () => {
    const { nodes, edges } = extract([
      'Public Sub Recorrer()',
      '    Dim i As Long',
      '    For i = 1 To 10',
      '        GoTo siguiente',
      'siguiente:',
      '    Next i',
      'End Sub',
    ]);

    const node = label(nodes, 'ModErrores.Recorrer.siguiente')!;
    expect(node.metadata?.isHandler).toBe(false);
    // No region at all — not a null region, no key.
    expect(node.metadata).not.toHaveProperty('handlerBehavior');
    expect(node.metadata).not.toHaveProperty('regionStartLine');
    expect(node.metadata).not.toHaveProperty('regionEndLine');
    // And nothing routes errors to it.
    expect(edgesOfKind(edges, 'handles-error')).toHaveLength(0);
  });

  it('does not treat a label mentioned inside a string literal as a handler', () => {
    const { nodes, edges } = extract([
      'Public Sub Guardar()',
      '    Dim s As String',
      '    s = "On Error GoTo errores"',
      'errores:',
      'End Sub',
    ]);

    // The label definition is real; the `On Error` inside the string is not,
    // so nothing targets it (#209 discipline).
    expect(label(nodes, 'ModErrores.Guardar.errores')!.metadata?.isHandler).toBe(
      false,
    );
    expect(edgesOfKind(edges, 'handles-error')).toHaveLength(0);
  });
});

// ============================================================================
// 4. Edges
// ============================================================================

describe('Issue #263 — edges onto the label', () => {
  it('emits a contains edge from the owning procedure for every label', () => {
    const { nodes, edges } = extract([
      'Public Sub Recorrer()',
      '    On Error GoTo errores',
      'siguiente:',
      '    Exit Sub',
      'errores:',
      'End Sub',
    ]);

    const proc = procedure(nodes, 'Recorrer')!;
    const labelIds = new Set(labels(nodes).map((n) => n.id));
    expect(labelIds.size).toBe(2);

    const contains = edgesOfKind(edges, 'contains', proc.id).filter((e) =>
      labelIds.has(e.target),
    );
    expect(contains).toHaveLength(2);
    expect(new Set(contains.map((e) => e.target))).toEqual(labelIds);
  });

  it('emits one handles-error edge per On Error GoTo statement, not per target', () => {
    const { nodes, edges } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    On Error GoTo errores',
      '    Call Confirmar',
      '    Exit Sub',
      'errores:',
      '    p_Error = "boom"',
      'End Sub',
    ]);

    const proc = procedure(nodes, 'Guardar')!;
    const target = label(nodes, 'ModErrores.Guardar.errores')!;
    const handles = edgesOfKind(edges, 'handles-error', proc.id);

    // TWO statements, TWO edges — deliberately not deduplicated. Each swap is
    // its own routing decision with its own line.
    expect(handles).toHaveLength(2);
    expect(handles.every((e) => e.target === target.id)).toBe(true);
    expect(handles.map((e) => e.line)).toEqual([4, 6]);
    expect(handles.every((e) => e.metadata?.synthesizedBy === 'vba-error-handler')).toBe(
      true,
    );
    expect(handles.every((e) => e.provenance === 'heuristic')).toBe(true);
  });

  it('emits one handles-error edge per label when a procedure swaps handlers', () => {
    const { nodes, edges } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    On Error GoTo errMem',
      '    Call Reservar',
      '    Exit Sub',
      'errores:',
      '    p_Error = "boom"',
      '    Exit Sub',
      'errMem:',
      '    p_Error = "sin memoria"',
      'End Sub',
    ]);

    const proc = procedure(nodes, 'Guardar')!;
    const handles = edgesOfKind(edges, 'handles-error', proc.id);
    expect(handles).toHaveLength(2);

    const byTarget = new Map(handles.map((e) => [e.target, e]));
    const errores = label(nodes, 'ModErrores.Guardar.errores')!;
    const errMem = label(nodes, 'ModErrores.Guardar.errMem')!;
    expect(byTarget.has(errores.id)).toBe(true);
    expect(byTarget.has(errMem.id)).toBe(true);

    // Both are handlers. Only the one whose region `errorPolicy` resolved —
    // the earliest targeted definition — carries the derived behaviour, since
    // that is the only value #260 computed. The second is NOT re-classified
    // here; that would be exactly the drift this design forbids.
    expect(errores.metadata?.isHandler).toBe(true);
    expect(errMem.metadata?.isHandler).toBe(true);
    expect(errores.metadata?.handlerBehavior).toBe('channel');
    expect(errMem.metadata).not.toHaveProperty('handlerBehavior');
  });

  it('emits a references edge tagged vba-goto for a plain GoTo', () => {
    const { nodes, edges } = extract([
      'Public Sub Recorrer()',
      '    Dim i As Long',
      '    If i = 0 Then GoTo salir',
      '    Call Trabajar',
      'salir:',
      'End Sub',
    ]);

    const proc = procedure(nodes, 'Recorrer')!;
    const salir = label(nodes, 'ModErrores.Recorrer.salir')!;
    const jumps = edgesOfKind(edges, 'references', proc.id).filter(
      (e) => e.target === salir.id,
    );

    expect(jumps).toHaveLength(1);
    expect(jumps[0]!.metadata?.synthesizedBy).toBe('vba-goto');
    expect(jumps[0]!.line).toBe(5);
    // A jump is not an error-handling fact; it must NOT borrow the new kind.
    expect(edgesOfKind(edges, 'handles-error')).toHaveLength(0);
  });

  it('does not mistake an On Error GoTo for a plain GoTo', () => {
    const { nodes, edges } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      'errores:',
      'End Sub',
    ]);

    const proc = procedure(nodes, 'Guardar')!;
    const errores = label(nodes, 'ModErrores.Guardar.errores')!;
    expect(edgesOfKind(edges, 'handles-error', proc.id)).toHaveLength(1);
    expect(
      edgesOfKind(edges, 'references', proc.id).filter(
        (e) => e.target === errores.id,
      ),
    ).toHaveLength(0);
  });

  it('ignores a numeric GoTo target — a VBA line number is not a label', () => {
    const { nodes, edges, unresolvedReferences } = extract([
      'Public Sub Antiguo()',
      '    GoTo 100',
      'End Sub',
    ]);

    expect(labels(nodes)).toHaveLength(0);
    expect(edgesOfKind(edges, 'references').filter((e) => e.metadata?.synthesizedBy === 'vba-goto')).toHaveLength(0);
    // And no fabricated dangling reference for legal code.
    expect(refsTo(unresolvedReferences, '100')).toHaveLength(0);
  });
});

// ============================================================================
// 5. Dangling targets — never fabricate a node
// ============================================================================

describe('Issue #263 — dangling targets', () => {
  it('emits an UnresolvedReference and no node for On Error GoTo noExiste', () => {
    const { nodes, edges, unresolvedReferences } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo noExiste',
      '    Call Escribir',
      'End Sub',
    ]);

    // No node. The defect is only findable while the graph stays honest.
    expect(labels(nodes)).toHaveLength(0);
    expect(edgesOfKind(edges, 'handles-error')).toHaveLength(0);

    const dangling = refsTo(unresolvedReferences, 'noExiste');
    expect(dangling).toHaveLength(1);
    expect(dangling[0]!.referenceKind).toBe('references');
    expect(dangling[0]!.metadata?.synthesizedBy).toBe('vba-goto-unresolved');
    expect(dangling[0]!.fromNodeId).toBe(procedure(nodes, 'Guardar')!.id);

    // …and `errorPolicy` says the same thing, from the other side.
    expect(
      (procedure(nodes, 'Guardar')!.metadata!.errorPolicy as { danglingTarget: string | null })
        .danglingTarget,
    ).toBe('noExiste');
  });

  it('treats a label defined in a SIBLING procedure as still dangling', () => {
    const { nodes, unresolvedReferences } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      'End Sub',
      '',
      'Public Sub Borrar()',
      'errores:',
      '    p_Error = "borrar"',
      'End Sub',
    ]);

    // VBA scopes labels to the procedure, so `Guardar` has no handler even
    // though the module contains an `errores:` somewhere else.
    expect(refsTo(unresolvedReferences, 'errores')).toHaveLength(1);
    expect(labels(nodes).map((n) => n.qualifiedName)).toEqual([
      'ModErrores.Borrar.errores',
    ]);
    // And that sibling label is control flow — nothing in ITS procedure
    // targets it.
    expect(label(nodes, 'ModErrores.Borrar.errores')!.metadata?.isHandler).toBe(
      false,
    );
  });
});

// ============================================================================
// 6. No re-parenting — the invariant that protects every existing consumer
// ============================================================================

describe('Issue #263 — calls inside a handler stay on the procedure', () => {
  it('leaves every call reference attributed to the enclosing procedure', () => {
    const { nodes, unresolvedReferences } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    Call Registrar',
      'End Sub',
    ]);

    const proc = procedure(nodes, 'Guardar')!;
    const labelIds = new Set(labels(nodes).map((n) => n.id));

    // The handler-body call is still the PROCEDURE's, flagged by #260 rather
    // than re-parented. Re-parenting would change `callers`/`callees` for
    // every procedure with a handler in the corpus.
    const registrar = refsTo(unresolvedReferences, 'Registrar');
    expect(registrar).toHaveLength(1);
    expect(registrar[0]!.fromNodeId).toBe(proc.id);
    expect(registrar[0]!.metadata?.inErrorHandler).toBe(true);

    // Nothing at all is sourced FROM a label node.
    expect(
      unresolvedReferences.some((r) => labelIds.has(r.fromNodeId)),
    ).toBe(false);
  });

  it('stamps inErrorHandler on a GoTo written inside the handler region', () => {
    const { nodes, edges } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    GoTo salir',
      'salir:',
      'End Sub',
    ]);

    const proc = procedure(nodes, 'Guardar')!;
    const salir = label(nodes, 'ModErrores.Guardar.salir')!;
    const jump = edgesOfKind(edges, 'references', proc.id).find(
      (e) => e.target === salir.id,
    );
    expect(jump).toBeDefined();
    // #260 owns the single stamping point; the new emitter must not opt out.
    expect(jump!.metadata?.inErrorHandler).toBe(true);
  });
});

// ============================================================================
// 7. The rule table
// ============================================================================

describe('Issue #263 — the goto-jump rule', () => {
  it('is registered on the errors table, masked and gated on an open procedure', () => {
    const rule = RULES.find((r) => r.id === 'goto-jump');
    expect(rule).toBeDefined();
    expect(rule!.scan).toBe('masked');
    expect(rule!.requires).toBe('inside-procedure');
  });

  it('records nothing for a GoTo outside any procedure body', () => {
    const { nodes, edges } = extract(['Public Const X As Long = 1']);
    expect(labels(nodes)).toHaveLength(0);
    expect(edgesOfKind(edges, 'handles-error')).toHaveLength(0);
  });
});

// ============================================================================
// 8. Search
// ============================================================================

describe('Issue #263 — kind:label is a search filter', () => {
  it('parses `kind:label` into a kind filter rather than free text', () => {
    const parsed = parseQuery('kind:label errores');
    expect(parsed.kinds).toContain('label');
    expect(parsed.text).toBe('errores');
  });
});

// ============================================================================
// 9. The two deliberate exclusions (the #257 precedent)
// ============================================================================

describe('Issue #263 — `label` stays out of the default result surfaces', () => {
  it('is absent from HIGH_VALUE_NODE_KINDS', () => {
    const source = fs.readFileSync(
      path.join(__dirname, '..', 'src', 'context', 'index.ts'),
      'utf-8',
    );
    const block = /const HIGH_VALUE_NODE_KINDS: NodeKind\[\] = \[([\s\S]*?)\];/.exec(
      source,
    );
    expect(block).not.toBeNull();
    expect(block![1]).not.toMatch(/'label'/);
  });

  it('is absent from CONTAINER_NODE_KINDS', () => {
    const source = fs.readFileSync(
      path.join(__dirname, '..', 'src', 'mcp', 'tools.ts'),
      'utf-8',
    );
    const block = /const CONTAINER_NODE_KINDS = new Set<NodeKind>\(\[([\s\S]*?)\]\);/.exec(
      source,
    );
    expect(block).not.toBeNull();
    expect(block![1]).not.toMatch(/'label'/);
  });
});
