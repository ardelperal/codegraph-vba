/**
 * Issue #250 (task T8 of the VBA node-discovery plan): a `Sub`, a `Function`,
 * a `Property Get` and a `Property Let` used to be four indistinguishable
 * `function` nodes with no `metadata` at all. These tests pin the procedure
 * shape the extractor now records on the node: `procKind`, `accessor`,
 * `isStatic`, `isFriend`, `returnType`, `params` and `arity`.
 */
import { describe, expect, it } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { Node } from '../src/types';

function extract(source: string, filePath = 'src/modules/Signatures.bas') {
  return new VbaExtractor(filePath, source).extract();
}

/** Declared procedures only — synthesized call-target stubs carry `metadata.stub`. */
function declaredProcedures(nodes: Node[]): Node[] {
  return nodes.filter(
    (node) => node.kind === 'function' && node.metadata?.stub !== true,
  );
}

function procedure(nodes: Node[], name: string): Node | undefined {
  return declaredProcedures(nodes).find((node) => node.name === name);
}

describe('Issue #250: procedure kind metadata', () => {
  it('records procKind for a Sub, a Function and a Property accessor', () => {
    const { nodes } = extract(
      [
        'Public Sub DoWork()',
        'End Sub',
        'Public Function Compute() As Long',
        'End Function',
        'Public Property Get Title() As String',
        'End Property',
      ].join('\n'),
    );

    expect(procedure(nodes, 'DoWork')?.metadata?.procKind).toBe('sub');
    expect(procedure(nodes, 'Compute')?.metadata?.procKind).toBe('function');
    expect(procedure(nodes, 'Title')?.metadata?.procKind).toBe('property');
  });

  it('records the accessor for each Property form and omits it elsewhere', () => {
    const { nodes } = extract(
      [
        'Public Property Get Alpha() As String',
        'End Property',
        'Public Property Let Beta(ByVal value As String)',
        'End Property',
        'Public Property Set Gamma(ByVal value As Object)',
        'End Property',
        'Public Sub Delta()',
        'End Sub',
      ].join('\n'),
    );

    expect(procedure(nodes, 'Alpha')?.metadata?.accessor).toBe('get');
    expect(procedure(nodes, 'Beta')?.metadata?.accessor).toBe('let');
    expect(procedure(nodes, 'Gamma')?.metadata?.accessor).toBe('set');
    expect(procedure(nodes, 'Delta')?.metadata?.accessor).toBeUndefined();
  });

  it('separates a Property Get from its matching Property Let in the same class', () => {
    const { nodes } = extract(
      [
        'Attribute VB_Name = "Cliente"',
        'Public Property Get Nombre() As String',
        'End Property',
        'Public Property Let Nombre(ByVal value As String)',
        'End Property',
      ].join('\n'),
      'src/classes/Cliente.cls',
    );

    const accessors = declaredProcedures(nodes).filter(
      (node) => node.name === 'Nombre',
    );
    expect(accessors).toHaveLength(2);
    expect(accessors.map((node) => node.metadata?.accessor).sort()).toEqual([
      'get',
      'let',
    ]);
    // Both still share the same qualifiedName — the accessor is what tells
    // them apart, which is the whole point of the issue.
    expect(new Set(accessors.map((node) => node.qualifiedName)).size).toBe(1);
  });

  it('never reports a Property accessor as a sub', () => {
    const { nodes } = extract(
      [
        'Public Property Get A() As Long',
        'End Property',
        'Private Property Let B(ByVal v As Long)',
        'End Property',
        'Friend Property Set C(ByVal v As Object)',
        'End Property',
        'Public Property Get D() As Long',
        'End Property',
      ].join('\n'),
    );

    for (const node of declaredProcedures(nodes)) {
      expect(node.metadata?.procKind).toBe('property');
    }
  });
});

describe('Issue #250: modifier metadata', () => {
  it('records isStatic for a bare Static procedure and a Public Static one', () => {
    const { nodes } = extract(
      [
        'Static Sub Counter()',
        'End Sub',
        'Public Static Function Tally() As Long',
        'End Function',
        'Public Sub Plain()',
        'End Sub',
      ].join('\n'),
    );

    expect(procedure(nodes, 'Counter')?.metadata?.isStatic).toBe(true);
    expect(procedure(nodes, 'Tally')?.metadata?.isStatic).toBe(true);
    expect(procedure(nodes, 'Plain')?.metadata?.isStatic).toBe(false);
  });

  it('keeps the Friend fact even though visibility folds to public', () => {
    const { nodes } = extract(
      [
        'Friend Sub Internal()',
        'End Sub',
        'Public Sub Exposed()',
        'End Sub',
        'Private Sub Hidden()',
        'End Sub',
      ].join('\n'),
    );

    const internal = procedure(nodes, 'Internal');
    expect(internal?.visibility).toBe('public');
    expect(internal?.metadata?.isFriend).toBe(true);
    expect(procedure(nodes, 'Exposed')?.metadata?.isFriend).toBe(false);
    expect(procedure(nodes, 'Hidden')?.visibility).toBe('private');
    expect(procedure(nodes, 'Hidden')?.metadata?.isFriend).toBe(false);
  });
});

describe('Issue #250: return type metadata', () => {
  it('records the return type of a Function and of a Property Get', () => {
    const { nodes } = extract(
      [
        'Public Function Build(ByVal name As String) As Collection',
        'End Function',
        'Public Property Get Total() As Currency',
        'End Property',
      ].join('\n'),
    );

    expect(procedure(nodes, 'Build')?.metadata?.returnType).toBe('Collection');
    expect(procedure(nodes, 'Total')?.metadata?.returnType).toBe('Currency');
  });

  it('omits the return type for a Sub, a Property Let and an untyped Function', () => {
    const { nodes } = extract(
      [
        'Public Sub Run(ByVal name As String)',
        'End Sub',
        'Public Property Let Total(ByVal value As Currency)',
        'End Property',
        'Public Function Loose(ByVal name As String)',
        'End Function',
      ].join('\n'),
    );

    expect(procedure(nodes, 'Run')?.metadata?.returnType).toBeUndefined();
    expect(procedure(nodes, 'Total')?.metadata?.returnType).toBeUndefined();
    expect(procedure(nodes, 'Loose')?.metadata?.returnType).toBeUndefined();
  });
});

describe('Issue #250: parameter metadata', () => {
  it('records name, type, byRef, optional, isArray and hasDefault per parameter', () => {
    const { nodes } = extract(
      [
        'Public Sub Notify(ByVal id As Long, Optional ByRef msg As String = "", logs() As String)',
        'End Sub',
      ].join('\n'),
    );

    expect(procedure(nodes, 'Notify')?.metadata?.params).toEqual([
      {
        name: 'id',
        type: 'Long',
        byRef: false,
        optional: false,
        isArray: false,
        hasDefault: false,
      },
      {
        name: 'msg',
        type: 'String',
        byRef: true,
        optional: true,
        isArray: false,
        hasDefault: true,
      },
      {
        name: 'logs',
        type: 'String',
        byRef: true,
        optional: false,
        isArray: true,
        hasDefault: false,
      },
    ]);
  });

  it('treats an unqualified parameter as ByRef, which is what VBA does', () => {
    const { nodes } = extract(['Public Sub Touch(target As Long)', 'End Sub'].join('\n'));

    const params = procedure(nodes, 'Touch')?.metadata?.params as Array<
      Record<string, unknown>
    >;
    expect(params[0].byRef).toBe(true);
    expect(params[0].type).toBe('Long');
  });

  it('records a null type for a parameter with no As clause', () => {
    const { nodes } = extract(['Public Sub Touch(anything)', 'End Sub'].join('\n'));

    expect(procedure(nodes, 'Touch')?.metadata?.params).toEqual([
      {
        name: 'anything',
        type: null,
        byRef: true,
        optional: false,
        isArray: false,
        hasDefault: false,
      },
    ]);
  });

  it('keeps a default value containing a comma as one parameter', () => {
    const { nodes } = extract(
      ['Public Sub Join(Optional ByVal sep As String = "a,b")', 'End Sub'].join('\n'),
    );

    const params = procedure(nodes, 'Join')?.metadata?.params as unknown[];
    expect(params).toHaveLength(1);
    expect(procedure(nodes, 'Join')?.metadata?.arity).toEqual({
      required: 0,
      total: 1,
    });
  });

  it('records an empty parameter list for a paren-less declaration', () => {
    const { nodes } = extract(['Public Sub Bare', 'End Sub'].join('\n'));

    expect(procedure(nodes, 'Bare')?.metadata?.params).toEqual([]);
    expect(procedure(nodes, 'Bare')?.metadata?.arity).toEqual({
      required: 0,
      total: 0,
    });
  });
});

describe('Issue #250: arity metadata', () => {
  it('counts required and total parameters', () => {
    const { nodes } = extract(
      [
        'Public Sub Send(ByVal to_ As String, ByVal body As String, Optional ByVal cc As String)',
        'End Sub',
      ].join('\n'),
    );

    expect(procedure(nodes, 'Send')?.metadata?.arity).toEqual({
      required: 2,
      total: 3,
    });
  });

  it('serialises an unbounded ParamArray total as null', () => {
    const { nodes } = extract(
      ['Public Sub Log(ByVal tag As String, ParamArray args() As Variant)', 'End Sub'].join(
        '\n',
      ),
    );

    const meta = procedure(nodes, 'Log')?.metadata;
    expect((meta?.arity as { total: unknown }).total).toBeNull();
    expect((meta?.arity as { required: number }).required).toBe(1);
    // The value must survive the worker boundary, so it must be JSON-safe.
    expect(JSON.parse(JSON.stringify(meta?.arity))).toEqual({
      required: 1,
      total: null,
    });
  });
});

describe('Issue #250: continued signatures', () => {
  it('parses a signature split with line continuations exactly like the one-line form', () => {
    const oneLine = extract(
      [
        'Public Function Build(ByVal id As Long, Optional ByRef msg As String = "") As Collection',
        'End Function',
      ].join('\n'),
    );
    const continued = extract(
      [
        'Public Function Build(ByVal id As Long, _',
        '                      Optional ByRef msg As String = "") _',
        '                      As Collection',
        'End Function',
      ].join('\n'),
    );

    const a = procedure(oneLine.nodes, 'Build');
    const b = procedure(continued.nodes, 'Build');
    expect(b?.metadata).toEqual(a?.metadata);
    expect(b?.metadata?.returnType).toBe('Collection');
    expect(b?.startLine).toBe(1);
  });
});

describe('Issue #250: coverage of the declared-procedure population', () => {
  it('gives every declared procedure a procKind and leaves call stubs untouched', () => {
    const { nodes } = extract(
      [
        'Public Sub Entry()',
        '  Call Helper',
        '  MissingProc 1, 2',
        'End Sub',
        'Private Function Helper() As Long',
        'End Function',
        'Public Property Get Value() As Long',
        'End Property',
      ].join('\n'),
    );

    const declared = declaredProcedures(nodes);
    expect(declared.map((node) => node.name).sort()).toEqual([
      'Entry',
      'Helper',
      'Value',
    ]);
    for (const node of declared) {
      expect(node.metadata?.procKind).toBeDefined();
      expect(node.metadata?.params).toBeDefined();
      expect(node.metadata?.arity).toBeDefined();
    }

    const stubs = nodes.filter(
      (node) => node.kind === 'function' && node.metadata?.stub === true,
    );
    for (const stub of stubs) {
      expect(stub.metadata?.procKind).toBeUndefined();
    }
  });
});
