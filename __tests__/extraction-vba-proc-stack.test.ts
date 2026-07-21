import { describe, expect, it } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaExtractorContext, ProcInfo } from '../src/extraction/vba/context';
import { createEnumsConstsClassifier } from '../src/extraction/vba/enums-consts';
import { createCallsAndSqlClassifier } from '../src/extraction/vba/call-sweep';

function extract(source: string) {
  return new VbaExtractor('src/modules/ProcStack.bas', source).extract();
}

function observedDepthWhenProcedureStartsAt(startLine: number): number {
  const lines = startLine === 1 ? ['Public Sub A()'] : ['Option Explicit', 'Public Sub A()'];
  const ctx = new VbaExtractorContext('src/modules/ProcStack.bas');
  const proc: ProcInfo = {
    name: 'A',
    qualifiedName: 'A',
    kind: 'sub',
    visibility: 'public',
    startLine,
  };
  ctx.localProcs.set('A', [proc]);
  const enums = createEnumsConstsClassifier();
  const calls = createCallsAndSqlClassifier(lines);
  for (let i = 0; i < lines.length; i++) {
    enums.classifyLine(lines[i]!, i, ctx);
    calls.classifyLine(lines[i]!, i, ctx);
  }
  return ctx.procStack.length;
}

describe('VbaExtractor — Issue #208 proc-stack ownership and single-line procedures', () => {
  it('scans the body and pops a colon-separated single-line procedure', () => {
    const result = extract([
      'Public Sub Helper()',
      'End Sub',
      'Public Sub A(): Helper: End Sub',
    ].join('\n'));

    const a = result.nodes.find((node) => node.kind === 'function' && node.name === 'A');
    const helper = result.nodes.find((node) => node.kind === 'function' && node.name === 'Helper');
    expect(result.edges).toContainEqual(expect.objectContaining({
      kind: 'calls',
      source: a?.id,
      target: helper?.id,
      line: 3,
    }));
    expect(a?.endLine).toBe(3);
  });

  it('returns to module scope after a single-line procedure', () => {
    const result = extract([
      'Public Sub A(): End Sub',
      'Public Const FORM_X As String = "FormAfterA"',
      'Public Sub OpenIt()',
      '  DoCmd.OpenForm FORM_X',
      'End Sub',
    ].join('\n'));

    expect(result.nodes).toContainEqual(expect.objectContaining({
      kind: 'constant',
      name: 'FORM_X',
    }));
    expect(result.edges).toContainEqual(expect.objectContaining({
      kind: 'opens-form',
      metadata: expect.objectContaining({ targetFormName: 'FormAfterA' }),
    }));
  });

  it('tracks one real frame regardless of whether the first procedure starts on line 1 or 2', () => {
    expect(observedDepthWhenProcedureStartsAt(1)).toBe(1);
    expect(observedDepthWhenProcedureStartsAt(2)).toBe(1);
  });
});
