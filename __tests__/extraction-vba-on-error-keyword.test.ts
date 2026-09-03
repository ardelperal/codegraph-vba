/**
 * Issue #292 — `On Error` is a statement keyword pair, not a variable read.
 *
 * `scanModuleVariableReferences` walks every identifier on a line and emits a
 * reference for any that names a module-level variable. It guarded against a
 * `.` / `!` prefix and against procedure-local shadowing (#205), but not
 * against VBA keyword context — so in a module declaring the house-convention
 * `Public Error As String`, every `On Error GoTo errores` reported a
 * `property-get` of a variable named `Error`.
 *
 * On its own that is a stray edge. Once #261 labels channel references with
 * `errorChannel: true` it becomes a confident, wrong claim that the line takes
 * part in error propagation — the failure mode `CLAUDE.md` and guardrail 1 of
 * `docs/vba-error-handling-plan.md` both name as the worst available here.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

/** Every module-variable reference the sweep emitted, in source order. */
function moduleVarRefs(result: ReturnType<typeof extract>) {
  return result.unresolvedReferences.filter(
    (u) =>
      (u.metadata as Record<string, unknown> | undefined)?.synthesizedBy ===
      'vba-module-var',
  );
}

function cls(...body: string[]): string {
  return ['Attribute VB_Name = "Riesgo"', 'Public Error As String', '', ...body, ''].join('\n');
}

describe('issue #292 — On Error is a keyword, not a read of a variable named Error', () => {
  it('emits only the genuine write when a handler assigns the channel', () => {
    const result = extract(
      'Riesgo.cls',
      cls(
        'Public Sub Guardar()',
        '  On Error GoTo errores',
        '  Exit Sub',
        'errores:',
        '  Error = "boom"',
        'End Sub',
      ),
    );
    const refs = moduleVarRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceKind).toBe('property-set');
    expect(refs[0]!.line).toBe(8);
  });

  it('emits nothing for On Error Resume Next', () => {
    const result = extract(
      'Riesgo.cls',
      cls('Public Sub Guardar()', '  On Error Resume Next', 'End Sub'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });

  it('emits nothing for On Error GoTo 0', () => {
    const result = extract(
      'Riesgo.cls',
      cls('Public Sub Guardar()', '  On Error GoTo 0', 'End Sub'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });

  it('emits nothing for On Error GoTo -1', () => {
    const result = extract(
      'Riesgo.cls',
      cls('Public Sub Guardar()', '  On Error GoTo -1', 'End Sub'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });

  it('tolerates extra whitespace between On and Error', () => {
    const result = extract(
      'Riesgo.cls',
      cls('Public Sub Guardar()', '  On   Error   GoTo errores', 'End Sub'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });

  it('still reports a genuine bare read', () => {
    const result = extract(
      'Riesgo.cls',
      cls('Public Sub Comprobar()', '  If Error <> "" Then Exit Sub', 'End Sub'),
    );
    const refs = moduleVarRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceKind).toBe('property-get');
  });

  it('still reports a genuine write', () => {
    const result = extract(
      'Riesgo.cls',
      cls('Public Sub Fallar()', '  Error = "x"', 'End Sub'),
    );
    const refs = moduleVarRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceKind).toBe('property-set');
  });

  it('keeps column offsets intact when a genuine reference follows On Error on one line', () => {
    // The mask replaces `On Error` with spaces of the SAME length, so the
    // column reported for the later reference must be its real offset. A
    // substitution that shortened the line would silently shift every column.
    const statement = '  On Error GoTo errores: Error = "x"';
    const result = extract(
      'Riesgo.cls',
      cls('Public Sub Mixto()', statement, 'End Sub'),
    );
    const refs = moduleVarRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.column).toBe(statement.indexOf('Error = "x"'));
    // The kind here is `property-get`, not `property-set`: after a colon the
    // prefix `isDirectAssignment` sees is `… GoTo errores: `, which is not
    // blank, so a colon-separated assignment reads as an access rather than a
    // write. That is a pre-existing limitation of the shared predicate, not
    // something this mask introduces — asserted so the behaviour is pinned
    // rather than silently depended on.
    expect(refs[0]!.referenceKind).toBe('property-get');
  });

  it('leaves the existing member-access guard alone', () => {
    const result = extract(
      'Riesgo.cls',
      cls('Public Sub Leer()', '  x = obj.Error', '  y = Me!Error', 'End Sub'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });

  it('does not mask a variable that merely starts with the word On', () => {
    const result = extract(
      'Riesgo.cls',
      ['Attribute VB_Name = "Riesgo"', 'Public OnError As String', '', 'Public Sub Leer()', '  x = OnError', 'End Sub', ''].join('\n'),
    );
    const refs = moduleVarRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceName).toBe('OnError');
  });
});
