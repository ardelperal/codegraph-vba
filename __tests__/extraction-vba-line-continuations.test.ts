import { describe, expect, it } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { joinLineContinuations } from '../src/extraction/vba-preprocess';

function extract(source: string) {
  return new VbaExtractor('src/modules/Continued.bas', source).extract();
}

describe('Issue #202: VBA line continuations are logical statements', () => {
  it('joins a continued Dim on its first physical line and preserves line count', () => {
    const source = ['Public Sub Go()', '  Dim m As _', '      MiClase', '  m.Hacer', 'End Sub'].join('\n');
    const joined = joinLineContinuations(source);

    expect(joined.split('\n')).toEqual([
      'Public Sub Go()',
      '  Dim m As       MiClase',
      '',
      '  m.Hacer',
      'End Sub',
    ]);

    const result = extract(source);
    const typeNode = result.nodes.find(
      (node) => node.kind === 'class' && node.name === 'MiClase',
    );
    const typeReference = result.edges.find(
      (edge) => edge.kind === 'references' && edge.target === typeNode?.id,
    );
    expect(typeReference?.line).toBe(2);
    expect(
      result.nodes.some(
        (node) => node.kind === 'function' && node.name === 'MiClase.Hacer',
      ),
    ).toBe(true);
  });

  it('parses a continued procedure header at its first physical line', () => {
    const result = extract([
      'Public Function Build( _',
      '    ByVal values() As String) As MiClase',
      'End Function',
    ].join('\n'));

    const procedure = result.nodes.find(
      (node) => node.kind === 'function' && node.name === 'Build',
    );
    expect(procedure?.startLine).toBe(1);
  });

  it('parses a continued Declare statement at its first physical line', () => {
    const result = extract([
      'Private Declare PtrSafe Function GetTickCount Lib "kernel32" _',
      '    Alias "GetTickCount" () As Long',
    ].join('\n'));

    const declaration = result.nodes.find(
      (node) => node.kind === 'declare' && node.name === 'GetTickCount',
    );
    expect(declaration?.startLine).toBe(1);
    expect(declaration?.metadata?.aliasName).toBe('GetTickCount');
  });
});
