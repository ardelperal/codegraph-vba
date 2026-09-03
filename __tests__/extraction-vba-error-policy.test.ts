/**
 * Issue #259 (task E2 of `docs/vba-error-handling-plan.md`) — every procedure
 * records HOW it handles errors, as an `errorPolicy` object on the `function`
 * node the procedures sweep already emits.
 *
 * Two questions motivate it, and neither is greppable:
 *
 *   - "Which procedures have no error handling at all?" — an ABSENCE scoped
 *     to a procedure body has no text to match.
 *   - "Which `On Error Resume Next` scopes are never closed?" — the answer
 *     depends on the whole body, not on any one line.
 *
 * The hard constraint of the whole error-handling wave is that this adds
 * **zero node kinds and zero edge kinds**: §4 of the plan rejects a `label`
 * node plus a `handles-error` edge because 96.5% of the corpus's line labels
 * are the same label (`errores`) doing the same job, i.e. one bit, i.e. a
 * field. `emits no node and no edge` below is that constraint as a test.
 *
 * The precision gate that carries the most weight: **a label defined but
 * never targeted by an `On Error GoTo` is control flow, not a handler**
 * (137 such labels in the corpus). Reporting one as a handler would give a
 * confidently wrong answer to "does this procedure handle errors", which is
 * the worst failure mode available here.
 */
import { describe, expect, it } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VBA_RULE_TABLES } from '../src/extraction/vba-extractor';
import {
  createErrorPolicyClassifier,
  RULES,
} from '../src/extraction/vba/errors';
import { VbaExtractorContext } from '../src/extraction/vba/context';
import { Edge, Node } from '../src/types';

interface ErrorPolicy {
  protection: 'handler' | 'resume-next' | 'none';
  handlerLabel: string | null;
  handlerStartLine: number | null;
  handlerEndLine: number | null;
  behavior: string | null;
  handlerCount: number;
  resumeNextOpen: boolean;
  danglingTarget: string | null;
  executableStatementCount: number;
}

/**
 * Extract a `.bas` module. The two header lines are prepended here (not in
 * each fixture) so every asserted line number below counts from the first
 * line the fixture itself writes — `line 1 of the fixture` is source line 3.
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

function policy(nodes: Node[], name: string): ErrorPolicy {
  const node = procedure(nodes, name);
  expect(node, `procedure ${name} must exist`).toBeDefined();
  const found = node!.metadata?.errorPolicy as ErrorPolicy | undefined;
  expect(found, `procedure ${name} must carry metadata.errorPolicy`).toBeDefined();
  return found!;
}

function countPhysicalProcedureLines(body: string[]): number | undefined {
  const lines = ['Public Sub Probe()', ...body];
  const ctx = new VbaExtractorContext('src/modules/Probe.bas');
  ctx.functionNodeByStartLine.set(1, {} as Node);
  const classifier = createErrorPolicyClassifier();
  lines.forEach((line, index) => classifier.classifyLine(line, index, ctx));
  return ctx.vbaErrorPolicy?.executableStatementCount;
}

describe('Issue #259: the rule table', () => {
  it('registers `errors` in VBA_RULE_TABLES with the four tabulated rules', () => {
    // The issue tabulates these four ids; they are also the handles
    // `codegraph stats vba-rules` reports. `goto-jump` was appended by issue
    // #263, which needs the plain-`GoTo` jumps this table had no reason to
    // look at while it emitted nothing.
    expect(VBA_RULE_TABLES.errors?.map((r) => r.id)).toEqual([
      'on-error-label',
      'on-error-resume-next',
      'on-error-reset',
      'line-label',
      'goto-jump',
    ]);
    expect(VBA_RULE_TABLES.errors).toBe(RULES);
  });

  it('scans every rule MASKED and gates every rule on an open procedure', () => {
    // `masked` is the #209 discipline: an `On Error` idiom quoted inside a
    // string literal must be invisible here. `inside-procedure` is what keeps
    // module-level noise out — there is no error policy without a procedure.
    for (const rule of RULES) {
      expect(rule.scan, `${rule.id}.scan`).toBe('masked');
      expect(rule.requires, `${rule.id}.requires`).toBe('inside-procedure');
    }
  });
});

describe('Issue #259: protection — the three classes', () => {
  it('classifies a handler, a resume-next and an unprotected procedure', () => {
    const { nodes } = extract([
      'Public Sub ConHandler()',           // 3
      '    On Error GoTo errores',         // 4
      '    Call Trabajo',                  // 5
      '    Exit Sub',                      // 6
      'errores:',                          // 7
      '    MsgBox Err.Description',        // 8
      'End Sub',                           // 9
      '',                                  // 10
      'Public Sub ConResumeNext()',        // 11
      '    On Error Resume Next',          // 12
      '    Call Trabajo',                  // 13
      'End Sub',                           // 14
      '',                                  // 15
      'Public Sub SinProteccion()',        // 16
      '    Call Trabajo',                  // 17
      'End Sub',                           // 18
    ]);

    expect(policy(nodes, 'ConHandler').protection).toBe('handler');
    expect(policy(nodes, 'ConResumeNext').protection).toBe('resume-next');
    expect(policy(nodes, 'SinProteccion').protection).toBe('none');
  });

  it('a handler outranks a blanket suppression in the same procedure', () => {
    // `On Error Resume Next` before an `On Error GoTo` must not downgrade the
    // procedure to `resume-next`: it really does have a handler. This is the
    // probe's precedence (`handlerTargets.length > 0` wins).
    const { nodes } = extract([
      'Public Sub Ambas()',
      '    On Error Resume Next',
      '    Call Trabajo',
      '    On Error GoTo errores',
      '    Call MasTrabajo',
      '    Exit Sub',
      'errores:',
      '    MsgBox Err.Description',
      'End Sub',
    ]);

    expect(policy(nodes, 'Ambas').protection).toBe('handler');
  });

  it('leaves an unprotected procedure with every handler field null', () => {
    const { nodes } = extract([
      'Public Sub SinNada()',
      '    Call Trabajo',
      'End Sub',
    ]);

    expect(policy(nodes, 'SinNada')).toEqual({
      protection: 'none',
      handlerLabel: null,
      handlerStartLine: null,
      handlerEndLine: null,
      behavior: null,
      handlerCount: 0,
      resumeNextOpen: false,
      danglingTarget: null,
      executableStatementCount: 1,
    });
  });

  it('counts executable statements rather than body lines, including the six-statement boundary', () => {
    const { nodes } = extract([
      'Public Sub ExactamenteSeis()',
      "    ' comment-only lines are not executable",
      '',
      '    value = 1',
      '    value = value + 1',
      '    value = value + 1',
      '    value = value + 1',
      '    value = value + 1',
      '    FileCopy "source.txt", "target.txt"',
      'End Sub',
      '',
      'Public Sub SoloCincoConRelleno()',
      '    value = 1',
      '',
      "    ' another non-executable line",
      '    value = value + 1',
      '    value = value + 1',
      '    value = value + 1',
      '    FileCopy "source.txt", "target.txt"',
      'End Sub',
    ]);

    expect(policy(nodes, 'ExactamenteSeis').executableStatementCount).toBe(6);
    expect(policy(nodes, 'SoloCincoConRelleno').executableStatementCount).toBe(5);
  });

  it('splits true separators without splitting named arguments, dates, or string colons', () => {
    const { nodes } = extract([
      'Public Sub TrueSeparators()',
      '    value = 1: value = 2: value = 3',
      'End Sub',
      '',
      'Public Sub NamedArgument()',
      '    Foo bar:=1',
      'End Sub',
      '',
      'Public Sub DateLiteral()',
      '    startedAt = #12:30:00#: finishedAt = #1/2/2026 13:45#',
      'End Sub',
      '',
      'Public Sub StringColon()',
      '    caption = "status: ready": value = 1',
      'End Sub',
    ]);

    expect(policy(nodes, 'TrueSeparators').executableStatementCount).toBe(3);
    expect(policy(nodes, 'NamedArgument').executableStatementCount).toBe(1);
        expect(policy(nodes, 'DateLiteral').executableStatementCount).toBe(2);
        expect(policy(nodes, 'StringColon').executableStatementCount).toBe(2);
      });

      it('treats the Double type suffix as a suffix, never as a date delimiter', () => {
        const { nodes } = extract([
          'Public Sub DoubleSuffixes()',
          '    value# = 1: other = 2',
          '    total = 1#: finalValue = 3',
          'End Sub',
        ]);

        expect(policy(nodes, 'DoubleSuffixes').executableStatementCount).toBe(4);
      });

      it('counts ReDim and ReDim Preserve as executable but excludes true declarations', () => {
        const { nodes } = extract([
          'Public Sub Arrays()',
          '    Dim values() As Double',
          '    Static cached As Boolean',
          '    Const limit As Long = 4',
          '    ReDim values(1 To limit)',
          '    ReDim Preserve values(1 To limit + 1)',
          'End Sub',
        ]);

        expect(policy(nodes, 'Arrays').executableStatementCount).toBe(2);
      });

      it('counts one executable statement continued across physical lines', () => {
        expect(countPhysicalProcedureLines([
          '    result = BuildValue( _',
          '        firstValue, _',
          '        secondValue)',
        ])).toBe(1);
      });

      it('counts a declaration continued across physical lines as zero', () => {
        expect(countPhysicalProcedureLines([
          '    Dim result As _',
          '        String',
        ])).toBe(0);
      });
});

describe('Issue #259: the handler region', () => {
  it('brackets exactly the handler body — first line after the label to the End', () => {
    const { nodes } = extract([
      'Public Sub Guardar()',              // 3
      '    On Error GoTo errores',         // 4
      '    Call Escribir',                 // 5
      '    Exit Sub',                      // 6
      'errores:',                          // 7  <- label definition
      '    Call Registrar',                // 8  <- handlerStartLine
      '    Resume Next',                   // 9
      'End Sub',                           // 10 <- handlerEndLine
    ]);

    const p = policy(nodes, 'Guardar');
    expect(p.handlerLabel).toBe('errores');
    expect(p.handlerStartLine).toBe(8);
    expect(p.handlerEndLine).toBe(10);
  });

  it('resolves a handler label defined BEFORE the `On Error GoTo` that targets it', () => {
    // Legal VBA, and the walk sees the definition first, so the region can
    // only be resolved once the whole body is known.
    const { nodes } = extract([
      'Public Sub Rara()',                 // 3
      '    GoTo arranque',                 // 4
      'reintento:',                        // 5
      '    Call Recuperar',                // 6
      '    Exit Sub',                      // 7
      'arranque:',                         // 8
      '    On Error GoTo reintento',       // 9
      '    Call Trabajo',                  // 10
      'End Sub',                           // 11
    ]);

    const p = policy(nodes, 'Rara');
    expect(p.protection).toBe('handler');
    expect(p.handlerLabel).toBe('reintento');
    expect(p.handlerStartLine).toBe(6);
    expect(p.handlerEndLine).toBe(11);
  });

  it('gives a `Property Get` with a handler the same treatment as a Sub', () => {
    const { nodes } = extract([
      'Public Property Get Titulo() As String',  // 3
      '    On Error GoTo errores',               // 4
      '    Titulo = Leer()',                     // 5
      '    Exit Property',                       // 6
      'errores:',                                // 7
      '    Titulo = ""',                         // 8
      'End Property',                            // 9
    ]);

    const p = policy(nodes, 'Titulo');
    expect(p.protection).toBe('handler');
    expect(p.handlerLabel).toBe('errores');
    expect(p.handlerStartLine).toBe(8);
    expect(p.handlerEndLine).toBe(9);
  });

  it('names `behavior` from the handler body (issue #260 filled in E2\'s null)', () => {
    // E2 shipped this field as a hard `null` and said E3 would derive it.
    // Issue #260 did; the full behaviour matrix lives in
    // `extraction-vba-error-handler-region.test.ts`. This assertion stays
    // here so the field's OWNER — the object E2 builds — keeps a test.
    const { nodes } = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    MsgBox Err.Description',
      'End Sub',
    ]);

    expect(policy(nodes, 'Guardar').behavior).toBe('display');
  });

  it('leaves `behavior` null when there is no handler to describe', () => {
    const { nodes } = extract([
      'Public Sub Sin()',
      '    Call Escribir',
      'End Sub',
    ]);

    expect(policy(nodes, 'Sin').protection).toBe('none');
    expect(policy(nodes, 'Sin').behavior).toBeNull();
  });
});

describe('Issue #259: the precision gate — an untargeted label is control flow', () => {
  it('a label nobody targets leaves the procedure unprotected and opens no region', () => {
    // 137 labels in the corpus are exactly this shape (`siguiente`, `salir`,
    // `fin`). Reading one as a handler is the worst failure mode here.
    const { nodes } = extract([
      'Public Sub Bucle()',
      '    Dim i As Long',
      '    For i = 1 To 10',
      '        If i = 3 Then GoTo siguiente',
      '        Call Trabajo',
      'siguiente:',
      '    Next i',
      'End Sub',
    ]);

    const p = policy(nodes, 'Bucle');
    expect(p.protection).toBe('none');
    expect(p.handlerLabel).toBeNull();
    expect(p.handlerStartLine).toBeNull();
    expect(p.handlerEndLine).toBeNull();
    expect(p.handlerCount).toBe(0);
  });

  it('a control-flow label sitting beside a real handler never becomes the handler', () => {
    const { nodes } = extract([
      'Public Sub Mixto()',                 // 3
      '    On Error GoTo errores',          // 4
      '    If Cancelar Then GoTo salir',    // 5
      '    Call Trabajo',                   // 6
      'salir:',                             // 7  <- control flow, not a handler
      '    Exit Sub',                       // 8
      'errores:',                           // 9
      '    MsgBox Err.Description',         // 10
      'End Sub',                            // 11
    ]);

    const p = policy(nodes, 'Mixto');
    expect(p.handlerLabel).toBe('errores');
    expect(p.handlerStartLine).toBe(10);
    expect(p.handlerEndLine).toBe(11);
  });
});

describe('Issue #259: handlerCount', () => {
  it('counts two `On Error GoTo` sites in one procedure', () => {
    const { nodes } = extract([
      'Public Sub Doble()',
      '    On Error GoTo primero',
      '    Call Uno',
      '    On Error GoTo segundo',
      '    Call Dos',
      '    Exit Sub',
      'primero:',
      '    Resume Next',
      'segundo:',
      '    MsgBox Err.Description',
      'End Sub',
    ]);

    expect(policy(nodes, 'Doble').handlerCount).toBe(2);
  });

  it('does not count `On Error GoTo 0` as a handler site', () => {
    const { nodes } = extract([
      'Public Sub Reset()',
      '    On Error Resume Next',
      '    Call Trabajo',
      '    On Error GoTo 0',
      'End Sub',
    ]);

    const p = policy(nodes, 'Reset');
    expect(p.handlerCount).toBe(0);
    expect(p.protection).toBe('resume-next');
    expect(p.handlerLabel).toBeNull();
  });
});

describe('Issue #259: resumeNextOpen', () => {
  it('is false when `On Error GoTo 0` closes the scope', () => {
    const { nodes } = extract([
      'Public Sub Cerrada()',
      '    On Error Resume Next',
      '    Call Trabajo',
      '    On Error GoTo 0',
      '    Call MasTrabajo',
      'End Sub',
    ]);

    expect(policy(nodes, 'Cerrada').resumeNextOpen).toBe(false);
  });

  it('is true when the scope runs to the end of the procedure', () => {
    const { nodes } = extract([
      'Public Sub Abierta()',
      '    On Error Resume Next',
      '    Call Trabajo',
      'End Sub',
    ]);

    expect(policy(nodes, 'Abierta').resumeNextOpen).toBe(true);
  });

  it('is true again when a second Resume Next reopens the scope after a reset', () => {
    const { nodes } = extract([
      'Public Sub Reabierta()',
      '    On Error Resume Next',
      '    Call Uno',
      '    On Error GoTo 0',
      '    On Error Resume Next',
      '    Call Dos',
      'End Sub',
    ]);

    expect(policy(nodes, 'Reabierta').resumeNextOpen).toBe(true);
  });

  it('treats `On Error GoTo -1` as a reset (valid VBA, 0 occurrences in the corpus)', () => {
    // Pinned deliberately: the plan asks for the zero-occurrence case to have
    // a fixture so it cannot rot untested.
    const { nodes } = extract([
      'Public Sub MenosUno()',
      '    On Error Resume Next',
      '    Call Trabajo',
      '    On Error GoTo -1',
      'End Sub',
    ]);

    const p = policy(nodes, 'MenosUno');
    expect(p.resumeNextOpen).toBe(false);
    expect(p.protection).toBe('resume-next');
    expect(p.handlerCount).toBe(0);
    expect(p.danglingTarget).toBeNull();
  });
});

describe('Issue #259: danglingTarget', () => {
  it('reports an `On Error GoTo` target this procedure never defines', () => {
    const { nodes } = extract([
      'Public Sub Rota()',
      '    On Error GoTo noExiste',
      '    Call Trabajo',
      'End Sub',
    ]);

    const p = policy(nodes, 'Rota');
    expect(p.danglingTarget).toBe('noExiste');
    expect(p.protection).toBe('handler');
    expect(p.handlerStartLine).toBeNull();
    expect(p.handlerEndLine).toBeNull();
  });

  it('is null when every target is defined in the same procedure', () => {
    const { nodes } = extract([
      'Public Sub Sana()',
      '    On Error GoTo errores',
      '    Call Trabajo',
      '    Exit Sub',
      'errores:',
      '    MsgBox Err.Description',
      'End Sub',
    ]);

    expect(policy(nodes, 'Sana').danglingTarget).toBeNull();
  });

  it('is scoped per procedure — a label defined in a SIBLING procedure is still dangling', () => {
    const { nodes } = extract([
      'Public Sub Primera()',
      '    On Error GoTo errores',
      '    Call Trabajo',
      'End Sub',
      '',
      'Public Sub Segunda()',
      '    Exit Sub',
      'errores:',
      '    MsgBox Err.Description',
      'End Sub',
    ]);

    expect(policy(nodes, 'Primera').danglingTarget).toBe('errores');
    expect(policy(nodes, 'Segunda').protection).toBe('none');
  });
});

describe('Issue #259: regression guards', () => {
  it('#208 — a colon-separated single-line procedure closes its own policy', () => {
    // `PROCEDURE_END_RE`'s `(?:^|:\s*)` prefix is what makes the single-line
    // form terminate. If it did not, `Corta`'s open resume-next scope would
    // leak into `Larga` and both procedures would be misreported.
    const { nodes } = extract([
      'Public Sub Corta(): On Error Resume Next: Call Trabajo: End Sub',  // 3
      '',                                                                 // 4
      'Public Sub Larga()',                                               // 5
      '    On Error GoTo errores',                                        // 6
      '    Call Trabajo',                                                 // 7
      '    Exit Sub',                                                     // 8
      'errores:',                                                         // 9
      '    MsgBox Err.Description',                                       // 10
      'End Sub',                                                          // 11
    ]);

    const corta = policy(nodes, 'Corta');
    expect(corta.protection).toBe('resume-next');
    expect(corta.resumeNextOpen).toBe(true);
    expect(corta.handlerCount).toBe(0);

    const larga = policy(nodes, 'Larga');
    expect(larga.protection).toBe('handler');
    expect(larga.handlerLabel).toBe('errores');
    expect(larga.handlerStartLine).toBe(10);
    expect(larga.handlerEndLine).toBe(11);
  });

  it('#209 — an `On Error GoTo` quoted inside a string literal sets nothing', () => {
    const { nodes } = extract([
      'Public Sub Ayuda()',
      '    Dim s As String',
      '    s = "On Error GoTo errores"',
      '    Debug.Print "errores: escriba On Error Resume Next"',
      'End Sub',
    ]);

    expect(policy(nodes, 'Ayuda')).toEqual({
      protection: 'none',
      handlerLabel: null,
      handlerStartLine: null,
      handlerEndLine: null,
      behavior: null,
      handlerCount: 0,
      resumeNextOpen: false,
      danglingTarget: null,
      executableStatementCount: 2,
    });
  });

  it('#209 — a real `On Error GoTo` on the same line as a quoted one still counts once', () => {
    const { nodes } = extract([
      'Public Sub Mezcla()',
      '    On Error GoTo errores: Debug.Print "On Error GoTo otro"',
      '    Call Trabajo',
      '    Exit Sub',
      'errores:',
      '    MsgBox Err.Description',
      'End Sub',
    ]);

    const p = policy(nodes, 'Mezcla');
    expect(p.handlerCount).toBe(1);
    expect(p.handlerLabel).toBe('errores');
    expect(p.danglingTarget).toBeNull();
  });
});

/**
 * Issue #263 (task E6) later added a `label` node and a `handles-error` edge
 * on top of this classifier — with maintainer sign-off, and against the
 * budget §4.3 of the plan sets out. It is the ONLY thing allowed to add rows
 * here, so the guard below still holds once its rows are set aside: the
 * error-POLICY classifier itself must remain a pure annotator.
 */
function withoutIssue263Rows(result: {
  nodes: Node[];
  edges: Edge[];
}): { nodes: Node[]; edges: Edge[] } {
  const labelIds = new Set(
    result.nodes.filter((n) => n.kind === 'label').map((n) => n.id),
  );
  return {
    nodes: result.nodes.filter((n) => n.kind !== 'label'),
    edges: result.edges.filter(
      (e) => e.kind !== 'handles-error' && !labelIds.has(e.target),
    ),
  };
}

describe('Issue #259: zero new node kinds, zero new edge kinds', () => {
  it('emits no node and no edge for the handler or the policy', () => {
    // The merge-blocking constraint of the whole error-handling wave. Setting
    // #263's label rows aside, the handler-bearing module must produce exactly
    // the nodes and edges the module WITHOUT any `On Error` produces, plus
    // nothing.
    const withHandler = extract([
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Call Escribir',
      '    Exit Sub',
      'errores:',
      '    Call Registrar',
      '    Resume Next',
      'End Sub',
    ]);
    const withoutHandler = extract([
      'Public Sub Guardar()',
      '    Call Escribir',
      '    Exit Sub',
      '    Call Registrar',
      'End Sub',
    ]);

    const withHandlerRows = withoutIssue263Rows(withHandler);
    const withoutHandlerRows = withoutIssue263Rows(withoutHandler);

    expect(withHandlerRows.nodes.map((n) => n.kind).sort()).toEqual(
      withoutHandlerRows.nodes.map((n) => n.kind).sort(),
    );
    expect(withHandlerRows.edges.map((e) => e.kind).sort()).toEqual(
      withoutHandlerRows.edges.map((e) => e.kind).sort(),
    );
    expect(withHandlerRows.nodes.length).toBe(withoutHandlerRows.nodes.length);
    expect(withHandlerRows.edges.length).toBe(withoutHandlerRows.edges.length);
    expect(withHandler.unresolvedReferences.length).toBe(
      withoutHandler.unresolvedReferences.length,
    );

    // Outside #263's own `label` kind, no node is named after the label.
    expect(withHandlerRows.nodes.some((n) => n.name === 'errores')).toBe(false);
  });

  it('a module-level `On Error` line alone still creates no module node', () => {
    // The classifier's `count` stays 0 for its whole life, so a file whose
    // only content is an `On Error` statement is still a file with no
    // symbols — it must not gain a `module` node. (The `file` node is
    // unconditional and is the baseline both sides share.)
    const withOnError = extract(['On Error Resume Next']);
    const empty = extract([]);

    expect(withOnError.nodes.map((n) => `${n.kind}:${n.name}`)).toEqual(
      empty.nodes.map((n) => `${n.kind}:${n.name}`),
    );
    expect(withOnError.nodes.some((n) => n.kind === 'module')).toBe(false);
  });
});
