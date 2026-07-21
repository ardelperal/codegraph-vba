/**
 * Issue #209 — `DoCmd.OpenForm` / `Forms!…!…` / `TempVars(…)` matched INSIDE
 * a string literal must NOT emit a real graph edge.
 *
 * Real Access code is full of log messages, docstrings, and TODO comments
 * that quote these idioms literally. The matchers in
 * `src/extraction/vba/{docmd,controls,tempvars}.ts` intentionally scan the
 * ORIGINAL (unmasked) line because the payload (a form name, a control
 * name, a TempVars key) lives inside a `"literal"` — masking would
 * obliterate the payload. The consequence is that the regex MATCHES inside
 * a string literal, and a `form-layout` stub node + an `opens-form` edge
 * end up in the graph even though no `DoCmd.OpenForm` was actually called.
 *
 * Acceptance criteria for this issue:
 *   1. `MsgBox "DoCmd.OpenForm ""frmX"""`  emits NO `opens-form` edge.
 *   2. `Debug.Print "Forms!frmY!ctl"`      emits NO control reference.
 *   3. A real `DoCmd.OpenForm "frmX"`      STILL emits its edge.
 *   4. A real `Forms!frmY!ctl`            STILL emits its reference.
 *
 * The fix must be local to the per-line scanners and must NOT touch the
 * masked-vs-original scan strategy that the regexes depend on (the
 * payload still has to come from a string literal). Per line, reject any
 * match whose `m.index` falls inside a string-literal span.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';

function extract(src: string) {
  return new VbaExtractor(
    'src/modules/ModIssue209.bas',
    [
      'Attribute VB_Name = "ModIssue209"',
      'Option Explicit',
      src,
    ].join('\n'),
  ).extract();
}

function heuristicEdges(r: ReturnType<typeof extract>) {
  return r.edges.filter((e) => e.provenance === 'heuristic');
}

describe('Issue #209 — string-literal matches emit no real graph edges', () => {
  it('a MsgBox string containing `DoCmd.OpenForm "frmX"` emits no opens-form edge', () => {
    const src = [
      'Public Sub Ayuda()',
      '    MsgBox "Para continuar use DoCmd.OpenForm ""frmClientes"" desde el menu"',
      'End Sub',
    ].join('\n');
    const r = extract(src);
    const opens = heuristicEdges(r).filter((e) => e.kind === 'opens-form');
    expect(opens, 'expected no opens-form edge from a MsgBox string').toHaveLength(0);
  });

  it('a Debug.Print string containing `Forms!frmY!ctl` emits no unresolved reference', () => {
    const src = [
      'Public Sub LogIt()',
      '    Debug.Print "pendiente: Forms!frmPedidos!txtTotal"',
      'End Sub',
    ].join('\n');
    const r = extract(src);
    const refs = r.unresolvedReferences.filter((u) => u.referenceKind === 'bang-get');
    expect(refs, 'expected no bang-get reference from a Debug.Print string').toHaveLength(0);
  });

  it('a Debug.Print string containing `TempVars("clave")` emits no tempvar reference', () => {
    const src = [
      'Public Sub LogIt()',
      '    Debug.Print "guarda TempVars(""clave"") para luego"',
      'End Sub',
    ].join('\n');
    const r = extract(src);
    const tempRefs = heuristicEdges(r).filter(
      (e) => e.metadata?.synthesizedBy === 'vba-tempvar',
    );
    expect(tempRefs, 'expected no vba-tempvar reference from a Debug.Print string').toHaveLength(0);
  });

  it('a real `DoCmd.OpenForm "frmX"` STILL emits its opens-form edge', () => {
    const src = [
      'Public Sub RealOpen()',
      '    DoCmd.OpenForm "frmClientes", acNormal',
      'End Sub',
    ].join('\n');
    const r = extract(src);
    const opens = heuristicEdges(r).filter((e) => e.kind === 'opens-form');
    expect(opens).toHaveLength(1);
    expect(opens[0]?.metadata?.targetFormName).toBe('frmClientes');
  });

  it('a real `Forms!frmY!ctl` STILL emits its bang-get reference', () => {
    const src = [
      'Public Sub RealRead()',
      '    Dim v As Variant',
      '    v = Forms!frmPedidos!txtTotal',
      'End Sub',
    ].join('\n');
    const r = extract(src);
    const refs = r.unresolvedReferences.filter((u) => u.referenceKind === 'bang-get');
    expect(refs.length).toBeGreaterThan(0);
    expect(refs.some((u) => u.referenceName === 'frmPedidos')).toBe(true);
  });

  it('a real `TempVars("k")` STILL emits a vba-tempvar reference', () => {
    const src = [
      'Public Sub RealTemp()',
      '    Debug.Print TempVars("k")',
      'End Sub',
    ].join('\n');
    const r = extract(src);
    const tempRefs = heuristicEdges(r).filter(
      (e) => e.metadata?.synthesizedBy === 'vba-tempvar',
    );
    expect(tempRefs.length).toBeGreaterThan(0);
  });

  it('a mixed line — real call OUTSIDE the literal + quoted call INSIDE the literal — emits exactly one edge for the real one', () => {
    const src = [
      'Public Sub Mix()',
      '    MsgBox "Tip: DoCmd.OpenForm ""frmMensaje"""',
      '    DoCmd.OpenForm "frmClientes"',
      'End Sub',
    ].join('\n');
    const r = extract(src);
    const opens = heuristicEdges(r).filter((e) => e.kind === 'opens-form');
    expect(opens).toHaveLength(1);
    expect(opens[0]?.metadata?.targetFormName).toBe('frmClientes');
  });
});
