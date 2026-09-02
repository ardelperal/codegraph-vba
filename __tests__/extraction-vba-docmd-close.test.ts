/**
 * Issue #246 (task T4 of `docs/vba-node-discovery-plan.md`) —
 * `DoCmd.Close acForm|acReport, "<Name>"`.
 *
 * `DoCmd.Close` is by far the most-called modellable `DoCmd` verb in the
 * Access corpus and produced nothing in the graph. It now closes the
 * lifecycle loop `opens-form` opens.
 *
 * The contract this suite pins:
 *   - only the STRING-LITERAL second argument emits, and only for
 *     `acForm` / `acReport`. A variable, `Me.Name`, a bare `DoCmd.Close` or
 *     any other object type is skipped SILENTLY — never a guessed target;
 *   - the edge kind is `references` (this wave adds no edge kinds), tagged
 *     `synthesizedBy: 'vba-closes-form'` + `targetFormName`;
 *   - the target is the SAME stub node `opens-form` / `opens-report` points
 *     at. One form opened and closed in one file yields ONE node and TWO
 *     edges — a duplicate stub with a different id would be worse than no
 *     edge at all.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import type { ExtractedEdge, ExtractedNode } from '../src/types';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

/** Wrap statements in a procedure so every call site has a caller. */
function inSub(...statements: string[]): string {
  return [
    'Public Sub Run()',
    ...statements.map((s) => `  ${s}`),
    'End Sub',
  ].join('\n');
}

function synthesizedByOf(edge: ExtractedEdge): string {
  return String(
    (edge.metadata as Record<string, unknown> | undefined)?.synthesizedBy ?? '',
  );
}

function closeEdges(result: { edges: ExtractedEdge[] }): ExtractedEdge[] {
  return result.edges.filter((e) => synthesizedByOf(e) === 'vba-closes-form');
}

function layoutNodes(result: { nodes: ExtractedNode[] }): ExtractedNode[] {
  return result.nodes.filter(
    (n) => n.kind === 'form-layout' || n.kind === 'report-layout',
  );
}

describe('DoCmd.Close — the emitting form', () => {
  it('emits one references edge to the Form_ stub for a string literal', () => {
    const result = extract(
      'modules/M.bas',
      inSub('DoCmd.Close acForm, "FormExpediente", acSaveNo'),
    );

    const edges = closeEdges(result);
    expect(edges).toHaveLength(1);
    const edge = edges[0]!;
    expect(edge.kind).toBe('references');
    expect(edge.provenance).toBe('heuristic');
    expect(edge.metadata).toMatchObject({
      synthesizedBy: 'vba-closes-form',
      targetFormName: 'FormExpediente',
    });

    const stubs = layoutNodes(result);
    expect(stubs).toHaveLength(1);
    const stub = stubs[0]!;
    expect(stub.kind).toBe('form-layout');
    expect(stub.name).toBe('FormExpediente');
    expect(stub.qualifiedName).toBe('Form_FormExpediente');
    expect(edge.target).toBe(stub.id);
  });

  it('uses the Report_ prefix convention for acReport', () => {
    const result = extract('modules/M.bas', inSub('DoCmd.Close acReport, "Rpt_X"'));

    const edges = closeEdges(result);
    expect(edges).toHaveLength(1);

    const stubs = layoutNodes(result);
    expect(stubs).toHaveLength(1);
    const stub = stubs[0]!;
    expect(stub.kind).toBe('report-layout');
    expect(stub.name).toBe('Rpt_X');
    expect(stub.qualifiedName).toBe('Report_Rpt_X');
    expect(edges[0]!.target).toBe(stub.id);
  });

  it('matches the parenthesised call form and is case-insensitive on the verb', () => {
    const result = extract(
      'modules/M.bas',
      inSub('docmd.close(acForm, "FormA")'),
    );

    const edges = closeEdges(result);
    expect(edges).toHaveLength(1);
    expect(edges[0]!.metadata).toMatchObject({ targetFormName: 'FormA' });
  });

  it('emits one edge per close call site on the same line', () => {
    const result = extract(
      'modules/M.bas',
      inSub('DoCmd.Close acForm, "FormA" : DoCmd.Close acForm, "FormB"'),
    );

    const edges = closeEdges(result);
    expect(edges).toHaveLength(2);
    expect(
      edges
        .map((e) => (e.metadata as Record<string, unknown>).targetFormName)
        .sort(),
    ).toEqual(['FormA', 'FormB']);
  });
});

describe('DoCmd.Close — the silent forms', () => {
  it.each([
    ['a Me.Name argument', 'DoCmd.Close acForm, Me.Name, acSaveNo'],
    ['a plain variable argument', 'DoCmd.Close acForm, strFormulario, acSaveNo'],
    ['a bare DoCmd.Close', 'DoCmd.Close'],
    ['a bare DoCmd.Close with a save option', 'DoCmd.Close , , acSaveNo'],
    ['an object type with no stub shape', 'DoCmd.Close acTable, "tblClientes"'],
    ['a concatenated name', 'DoCmd.Close acForm, "Form" & strSufijo'],
  ])('emits nothing for %s', (_label, statement) => {
    const result = extract('modules/M.bas', inSub(statement));

    expect(closeEdges(result)).toHaveLength(0);
    expect(layoutNodes(result)).toHaveLength(0);
  });

  it('ignores a DoCmd.Close that only appears inside a string literal', () => {
    const result = extract(
      'modules/M.bas',
      inSub('strLog = "DoCmd.Close acForm, ""FormA"""'),
    );

    expect(closeEdges(result)).toHaveLength(0);
    expect(layoutNodes(result)).toHaveLength(0);
  });
});

describe('DoCmd.Close — convergence with DoCmd.OpenForm', () => {
  it('an OpenForm and a Close of the same form share ONE stub node', () => {
    const result = extract(
      'modules/M.bas',
      [
        'Public Sub Abrir()',
        '  DoCmd.OpenForm "FormExpediente"',
        'End Sub',
        'Public Sub Cerrar()',
        '  DoCmd.Close acForm, "FormExpediente", acSaveNo',
        'End Sub',
      ].join('\n'),
    );

    const stubs = layoutNodes(result);
    expect(stubs).toHaveLength(1);
    const stubId = stubs[0]!.id;

    const opens = result.edges.filter((e) => e.kind === 'opens-form');
    const closes = closeEdges(result);
    expect(opens).toHaveLength(1);
    expect(closes).toHaveLength(1);

    // Same target node, two distinct edges from two distinct callers.
    expect(opens[0]!.target).toBe(stubId);
    expect(closes[0]!.target).toBe(stubId);
    expect(opens[0]!.source).not.toBe(closes[0]!.source);
    expect(opens[0]!.kind).not.toBe(closes[0]!.kind);
  });

  it('matches the stub case-insensitively, as OpenForm does', () => {
    const result = extract(
      'modules/M.bas',
      inSub('DoCmd.OpenForm "FormExpediente"', 'DoCmd.Close acForm, "formexpediente"'),
    );

    expect(layoutNodes(result)).toHaveLength(1);
    expect(closeEdges(result)).toHaveLength(1);
    expect(closeEdges(result)[0]!.target).toBe(layoutNodes(result)[0]!.id);
  });

  it('a form only ever closed still gets its stub, and the form/report buckets stay disjoint', () => {
    const result = extract(
      'modules/M.bas',
      inSub('DoCmd.Close acForm, "SoloCerrado"', 'DoCmd.Close acReport, "SoloCerrado"'),
    );

    const stubs = layoutNodes(result);
    expect(stubs).toHaveLength(2);
    expect(stubs.map((n) => n.kind).sort()).toEqual([
      'form-layout',
      'report-layout',
    ]);
    expect(new Set(stubs.map((n) => n.id)).size).toBe(2);
    expect(closeEdges(result)).toHaveLength(2);
  });
});
