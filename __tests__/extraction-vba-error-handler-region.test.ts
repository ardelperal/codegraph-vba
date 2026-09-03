/**
 * Issue #260 (task E3 of `docs/vba-error-handling-plan.md`) — everything a
 * procedure does ONLY when something has already gone wrong is marked as
 * such.
 *
 * Two things ship here, and they are judged by different yardsticks:
 *
 *   1. `metadata.inErrorHandler: true` on every edge and every
 *      `UnresolvedReference` emitted from inside the handler region E2
 *      resolved. The yardstick is the region itself: an edge is flagged iff
 *      its line falls in `[handlerStartLine, handlerEndLine]`. §3.4 of the
 *      plan is the motivation — today `Riesgo.Guardar -> MsgBox` and
 *      `Riesgo.Guardar -> Escribir` are indistinguishable in the graph, when
 *      one runs always and the other only on failure.
 *
 *   2. `errorPolicy.behavior`, which E2 deliberately left `null`. The
 *      yardstick here is NOT this file's opinion — it is
 *      `scripts/vba-coverage-probe.mjs`, the committed instrument the corpus
 *      census was measured with. `agrees with the probe's classifier` below
 *      runs BOTH classifiers over the same fixture and asserts they return
 *      the same string for every procedure, so a future edit to either one
 *      that silently forks the two fails here rather than in a corpus
 *      re-measurement nobody runs.
 *
 * The invariant that outranks both: this task adds **no node, no edge and no
 * unresolved reference**. It only stamps a field onto rows that already
 * existed. `adds no rows` pins that.
 */
import { describe, expect, it } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaExtractorContext } from '../src/extraction/vba/context';
import { newErrorPolicyState } from '../src/extraction/vba/errors';
import { Edge, Node, UnresolvedReference } from '../src/types';
// @ts-expect-error — plain ESM script, no type declarations by design.
import { analyzeVbaErrorHandling } from '../scripts/vba-coverage-probe.mjs';

interface ErrorPolicy {
  protection: 'handler' | 'resume-next' | 'none';
  handlerLabel: string | null;
  handlerStartLine: number | null;
  handlerEndLine: number | null;
  behavior: 'channel' | 'display' | 'reraise' | 'mixed' | 'unknown' | null;
  handlerCount: number;
  resumeNextOpen: boolean;
  danglingTarget: string | null;
}

/**
 * Extract a `.bas` module. The two header lines are prepended here (not in
 * each fixture) so a fixture's own first line is source line 3 — the same
 * convention `extraction-vba-error-policy.test.ts` uses.
 */
function extract(body: string[], filePath = 'src/modules/ModErrores.bas') {
  return new VbaExtractor(
    filePath,
    ['Attribute VB_Name = "ModErrores"', 'Option Explicit', ...body].join('\n'),
  ).extract();
}

/** Declared procedures only — call-target stubs carry `metadata.stub`. */
function procedure(nodes: Node[], name: string): Node | undefined {
  return nodes.find(
    (n) => n.kind === 'function' && n.metadata?.stub !== true && n.name === name,
  );
}

function policyOf(nodes: Node[], name: string): ErrorPolicy | undefined {
  return procedure(nodes, name)?.metadata?.errorPolicy as ErrorPolicy | undefined;
}

/** Every unresolved reference to `name`, in source order. */
function refsTo(refs: UnresolvedReference[], name: string) {
  return refs.filter((r) => r.referenceName === name);
}

function flagged(row: Edge | UnresolvedReference | undefined): unknown {
  return row?.metadata?.inErrorHandler;
}

describe('VBA error-handler region — inErrorHandler (issue #260)', () => {
  it('flags the call after the handler label and leaves the happy path alone', () => {
    const { unresolvedReferences } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    Call Registrar',
      'End Sub',
    ]);

    const escribir = refsTo(unresolvedReferences, 'Escribir');
    const registrar = refsTo(unresolvedReferences, 'Registrar');
    expect(escribir).toHaveLength(1);
    expect(registrar).toHaveLength(1);

    // The happy-path call carries NO flag at all — not `false`. Absence is
    // the encoding, so a consumer filtering on the key's presence is right.
    expect(escribir[0]!.metadata?.inErrorHandler).toBeUndefined();
    expect(flagged(registrar[0])).toBe(true);
  });

  it('flags nothing in a procedure with no handler', () => {
    const { edges, unresolvedReferences, nodes } = extract([
      'Public Sub Guardar()',
      '    Call Escribir',
      '    CurrentDb.Execute "INSERT INTO TbLog (x) VALUES (1)"',
      'End Sub',
    ]);

    expect(policyOf(nodes, 'Guardar')?.protection).toBe('none');
    expect(edges.some((e) => e.metadata?.inErrorHandler)).toBe(false);
    expect(
      unresolvedReferences.some((r) => r.metadata?.inErrorHandler),
    ).toBe(false);
  });

  it('flags nothing for a label nobody targets (control flow, not a handler)', () => {
    const { unresolvedReferences, nodes } = extract([
      'Public Sub Recorrer()',
      '    Call Escribir',
      'siguiente:',
      '    Call Registrar',
      'End Sub',
    ]);

    expect(policyOf(nodes, 'Recorrer')?.protection).toBe('none');
    expect(
      unresolvedReferences.some((r) => r.metadata?.inErrorHandler),
    ).toBe(false);
  });

  it('still emits the SQL table reference inside a handler, now flagged', () => {
    const { edges, nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    CurrentDb.Execute "UPDATE TbExpediente SET x = 1"',
      '    Exit Sub',
      'errores:',
      '    CurrentDb.Execute "INSERT INTO TbLog (msg) VALUES (1)"',
      'End Sub',
    ]);

    const tables = edges.filter(
      (e) => e.metadata?.synthesizedBy === 'vba-sql-table',
    );
    // Both tables are still there — the flag adds a field, it never filters.
    const names = tables.map((e) => e.target.replace(/^.*:/, ''));
    expect(tables).toHaveLength(2);
    expect(names.length).toBe(2);

    const happy = tables.find((e) => e.line === 5);
    const handler = tables.find((e) => e.line === 8);
    expect(happy?.metadata?.inErrorHandler).toBeUndefined();
    expect(flagged(handler)).toBe(true);
    // …and the access classification it already carried survives untouched.
    expect(handler?.metadata?.access).toBe('write');

    const policy = policyOf(nodes, 'Guardar')!;
    expect(policy.handlerStartLine).toBe(8);
    expect(policy.handlerEndLine).toBe(9);
  });

  it('ends the region at End Sub — a call in the NEXT procedure is not flagged', () => {
    const { unresolvedReferences } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Exit Sub',
      'errores:',
      '    Call Registrar',
      'End Sub',
      '',
      'Public Sub Siguiente()',
      '    Call Registrar',
      'End Sub',
    ]);

    const registrar = refsTo(unresolvedReferences, 'Registrar');
    expect(registrar.map((r) => r.line)).toEqual([7, 11]);
    expect(flagged(registrar[0])).toBe(true);
    // The off-by-one guard: the second procedure opens a fresh policy with no
    // handler, so nothing in it can inherit the first one's region.
    expect(registrar[1]!.metadata?.inErrorHandler).toBeUndefined();
  });

  it('flags a DoCmd.OpenForm edge opened from inside a handler', () => {
    // A different emitter and a different EDGE kind from the two above, which
    // is the point: the flag is stamped once, centrally, so `opens-form`
    // inherits it without `docmd.ts` knowing this feature exists.
    const { edges } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    DoCmd.OpenForm "FrmOk"',
      '    Exit Sub',
      'errores:',
      '    DoCmd.OpenForm "FrmError"',
      'End Sub',
    ]);

    const opens = edges.filter((e) => e.kind === 'opens-form');
    expect(opens.map((e) => e.metadata?.targetFormName)).toEqual([
      'FrmOk',
      'FrmError',
    ]);
    expect(opens[0]!.metadata?.inErrorHandler).toBeUndefined();
    expect(flagged(opens[1])).toBe(true);
  });

  it('adds no node, no edge and no unresolved reference', () => {
    const body = [
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    MsgBox "boom"',
      '    CurrentDb.Execute "INSERT INTO TbLog (msg) VALUES (1)"',
      'End Sub',
    ];
    const withHandler = extract(body);
    // The same module with the handler label renamed to a control-flow label
    // nobody targets: identical statements, no handler region.
    const withoutRegion = extract(
      body.map((l) =>
        l === '    On Error GoTo errores' ? '    Dim n As Long' : l,
      ),
    );

    expect(withHandler.nodes).toHaveLength(withoutRegion.nodes.length);
    expect(withHandler.edges).toHaveLength(withoutRegion.edges.length);
    expect(withHandler.unresolvedReferences).toHaveLength(
      withoutRegion.unresolvedReferences.length,
    );
    // …and the ONLY difference is the flag.
    expect(
      withHandler.unresolvedReferences.filter((r) => r.metadata?.inErrorHandler),
    ).toHaveLength(1);
    expect(
      withoutRegion.unresolvedReferences.filter(
        (r) => r.metadata?.inErrorHandler,
      ),
    ).toHaveLength(0);
  });
});

describe('VbaExtractorContext.inErrorHandler (issue #260)', () => {
  it('is false with no open procedure and false before the region opens', () => {
    const ctx = new VbaExtractorContext('src/modules/ModErrores.bas');
    expect(ctx.inErrorHandler(10)).toBe(false);

    ctx.vbaErrorPolicy = newErrorPolicyState(3, 0, 0);
    // A body is open but no targeted label has been seen yet.
    expect(ctx.inErrorHandler(10)).toBe(false);

    ctx.vbaErrorPolicy.handlerStartLine = 8;
    expect(ctx.inErrorHandler(7)).toBe(false);
    expect(ctx.inErrorHandler(8)).toBe(true);
    expect(ctx.inErrorHandler(99)).toBe(true);

    // The region's upper bound is the procedure body itself: closing it is
    // what ends the region, which is why the next procedure inherits nothing.
    ctx.vbaErrorPolicy = null;
    expect(ctx.inErrorHandler(99)).toBe(false);
  });
});

describe('VBA errorPolicy.behavior (issue #260)', () => {
  const cases: Array<[string, string[], ErrorPolicy['behavior']]> = [
    [
      'display',
      ['    MsgBox "Ha fallado"'],
      'display',
    ],
    [
      'display via Debug.Print',
      ['    Debug.Print Err.Description'],
      'display',
    ],
    [
      'channel',
      ['    If Err.Number <> 1000 Then p_Error = "Guardar: " & Err.Description'],
      'channel',
    ],
    [
      'reraise',
      ['    Err.Raise Err.Number, "Guardar", Err.Description'],
      'reraise',
    ],
    [
      'mixed',
      ['    Me.Error = Err.Description', '    MsgBox Err.Description'],
      'mixed',
    ],
    [
      'unknown',
      ['    Resume Next'],
      'unknown',
    ],
  ];

  for (const [name, handlerBody, expected] of cases) {
    it(`classifies a ${name} handler as ${expected}`, () => {
      const { nodes } = extract([
        'Public Sub Guardar()',
        '    On Error GoTo errores',
        '    Call Escribir',
        '    Exit Sub',
        'errores:',
        ...handlerBody,
        'End Sub',
      ]);
      expect(policyOf(nodes, 'Guardar')?.behavior).toBe(expected);
    });
  }

  it('classifies cleanup-then-record as channel, not as unknown', () => {
    // The fifth most common handler shape in the corpus: a `DoCmd` cleanup
    // call leads, and the channel write only happens under the sentinel
    // guard. A classifier that stops at the first statement — or that reads
    // the leading call as the handler's behaviour — gets this wrong.
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    DoCmd.Hourglass False',
      '    If Err.Number <> 1000 Then m_Error = "Guardar: " & Err.Description',
      'End Sub',
    ]);
    expect(policyOf(nodes, 'Guardar')?.behavior).toBe('channel');
  });

  it('reads only the handler body — a channel write on the happy path is not behaviour', () => {
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    m_Error = ""',
      '    If m_Error <> "" Then Err.Raise 1000',
      '    Exit Sub',
      'errores:',
      '    MsgBox Err.Description',
      'End Sub',
    ]);
    // `m_Error = ""` and `Err.Raise 1000` are both ABOVE the label, so the
    // only signal in the region is the MsgBox.
    expect(policyOf(nodes, 'Guardar')?.behavior).toBe('display');
  });

  it('ignores an error-channel LOOK-alike that is a read, not a write', () => {
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Exit Sub',
      'errores:',
      '    If m_Error <> "" Then Call Registrar',
      'End Sub',
    ]);
    expect(policyOf(nodes, 'Guardar')?.behavior).toBe('unknown');
  });

  it('does not match a name that merely contains a channel name', () => {
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Exit Sub',
      'errores:',
      '    ErrorCount = ErrorCount + 1',
      'End Sub',
    ]);
    expect(policyOf(nodes, 'Guardar')?.behavior).toBe('unknown');
  });

  it('leaves behavior null when the procedure has no handler', () => {
    const { nodes } = extract([
      'Public Sub Suprimir()',
      '    On Error Resume Next',
      '    MsgBox "no soy un manejador"',
      'End Sub',
      '',
      'Public Sub Sin()',
      '    MsgBox "tampoco"',
      'End Sub',
    ]);
    expect(policyOf(nodes, 'Suprimir')?.protection).toBe('resume-next');
    expect(policyOf(nodes, 'Suprimir')?.behavior).toBeNull();
    expect(policyOf(nodes, 'Sin')?.behavior).toBeNull();
  });

  it('is unknown for a handler whose target label is never defined', () => {
    const { nodes } = extract([
      'Public Sub Colgado()',
      '    On Error GoTo noExiste',
      '    MsgBox "esto no es el manejador"',
      'End Sub',
    ]);
    const policy = policyOf(nodes, 'Colgado')!;
    expect(policy.danglingTarget).toBe('noExiste');
    // No region, so no body to read — but the procedure IS `handler`-
    // protected, so `behavior` is `unknown` rather than `null`.
    expect(policy.behavior).toBe('unknown');
  });

  it('ignores a handler-shaped statement inside a string literal', () => {
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Exit Sub',
      'errores:',
      '    s = "MsgBox Err.Description"',
      'End Sub',
    ]);
    expect(policyOf(nodes, 'Guardar')?.behavior).toBe('unknown');
  });
});

/**
 * The cross-check that makes "agree with the probe" a test rather than a
 * promise. `scripts/vba-coverage-probe.mjs` is the instrument the corpus
 * figures in §2.3 of the plan were measured with; if the extractor's
 * `behavior` and the probe's diverge, the extractor's distribution can no
 * longer be compared against that census. Both classifiers read the SAME
 * fixture here and must return the same string for every procedure.
 */
describe('errorPolicy.behavior agrees with the probe classifier', () => {
  const SOURCE = [
    'Attribute VB_Name = "ModErrores"',
    'Option Explicit',
    '',
    'Public Sub Canal()',
    '    On Error GoTo errores',
    '    Call Escribir',
    '    Exit Sub',
    'errores:',
    '    If Err.Number <> 1000 Then p_Error = "Canal: " & Err.Description',
    'End Sub',
    '',
    'Public Sub Limpieza()',
    '    On Error GoTo errores',
    '    Call Escribir',
    '    Exit Sub',
    'errores:',
    '    DoCmd.Hourglass False',
    '    If Err.Number <> 1000 Then m_Error = "Limpieza: " & Err.Description',
    'End Sub',
    '',
    'Public Sub Pantalla()',
    '    On Error GoTo errores',
    '    Exit Sub',
    'errores:',
    '    MsgBox Err.Description',
    'End Sub',
    '',
    'Public Sub Relanza()',
    '    On Error GoTo errores',
    '    Exit Sub',
    'errores:',
    '    Err.Raise Err.Number, "Relanza", Err.Description',
    'End Sub',
    '',
    'Public Sub Mixto()',
    '    On Error GoTo errores',
    '    Exit Sub',
    'errores:',
    '    Me.Error = Err.Description',
    '    MsgBox Err.Description',
    'End Sub',
    '',
    'Public Sub Desconocido()',
    '    On Error GoTo errores',
    '    Exit Sub',
    'errores:',
    'End Sub',
    '',
    'Public Sub Suprime()',
    '    On Error Resume Next',
    '    MsgBox "no soy un manejador"',
    'End Sub',
    '',
    'Public Sub Desnudo()',
    '    Call Escribir',
    'End Sub',
    '',
    'Public Sub DosManejadores()',
    '    On Error GoTo errores',
    '    Call Escribir',
    '    On Error GoTo errores',
    '    Exit Sub',
    'errores:',
    '    g_Error = Err.Description',
    'End Sub',
  ].join('\n');

  it('returns the same behavior for every procedure in the fixture', () => {
    const filePath = 'src/modules/ModErrores.bas';
    const { nodes } = new VbaExtractor(filePath, SOURCE).extract();
    const declared = nodes.filter(
      (n) => n.kind === 'function' && n.metadata?.stub !== true,
    );
    expect(declared.length).toBe(9);

    const probeRecords = analyzeVbaErrorHandling(
      filePath,
      SOURCE,
      declared.map((n) => ({ startLine: n.startLine, name: n.name })),
    ) as Array<{ procedure: string; behavior: string | null }>;

    const fromProbe = new Map(
      probeRecords.map((r) => [r.procedure, r.behavior]),
    );
    const fromExtractor = new Map(
      declared.map((n) => [
        n.name,
        (n.metadata?.errorPolicy as ErrorPolicy | undefined)?.behavior ?? null,
      ]),
    );

    expect(Object.fromEntries(fromExtractor)).toEqual(
      Object.fromEntries(fromProbe),
    );
    // Guard against a vacuous pass: the fixture really does exercise every
    // bucket the plan names.
    expect(new Set(fromProbe.values())).toEqual(
      new Set(['channel', 'display', 'reraise', 'mixed', 'unknown', null]),
    );
  });
});
