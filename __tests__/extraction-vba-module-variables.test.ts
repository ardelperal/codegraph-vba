/**
 * Issue #251 — module-level variables are part of the graph.
 *
 * Before this change the dims sweep populated `localVarTypeMap` and emitted
 * type-reference edges for every `Public` / `Private` / `Global` declaration
 * but never a node, so "who reads `gblConn`?" had no answer at all. This
 * suite pins both halves of the fix:
 *
 *   1. a `variable` node + a module -> variable `contains` edge for every
 *      MODULE-LEVEL declaration (and for nothing else — procedure locals
 *      stay out of the graph);
 *   2. `property-get` / `property-set` unresolved references from the
 *      procedures that read and write those variables, gated strictly to
 *      names the same file declared at module level.
 *
 * The two behaviours that must NOT move are pinned here as well: the
 * `references` edge a qualified `As DAO.Database` declaration emits (still
 * exactly one), and the `subscribes-event` edge a `WithEvents` field emits.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { matchVbaModuleVariable } from '../src/resolution/name-matcher';
import type { ResolutionContext, UnresolvedRef } from '../src/resolution/types';
import type { Node } from '../src/types';

const MODULE_PATH = 'src/modules/modGlobals.bas';

function extract(src: string, filePath: string = MODULE_PATH) {
  return new VbaExtractor(filePath, src).extract();
}

function variableNodes(result: ReturnType<typeof extract>) {
  return result.nodes.filter((n) => n.kind === 'variable');
}

function moduleVarRefs(result: ReturnType<typeof extract>) {
  return result.unresolvedReferences.filter(
    (r) => r.metadata?.synthesizedBy === 'vba-module-var',
  );
}

/**
 * The `contains` edges whose source is the file's module/class node.
 *
 * `extract()` pushes that node LAST, after the walk — every earlier
 * `class` node is a synthetic type stub emitted by `emitReference` (`DAO`,
 * `Form_Pedido`, …), so the container is the final one.
 */
function containsTargets(result: ReturnType<typeof extract>): Set<string> {
  const container = result.nodes
    .filter((n) => n.kind === 'module' || n.kind === 'class')
    .at(-1);
  return new Set(
    result.edges
      .filter((e) => e.kind === 'contains' && e.source === container?.id)
      .map((e) => e.target),
  );
}

// ============================================================================
// 1. Nodes + contains, and the pre-existing type reference is untouched
// ============================================================================

describe('Issue #251 — module-level declarations become `variable` nodes', () => {
  it('`Public gblConn As DAO.Database` emits one variable node, one contains edge, and still exactly one DAO reference', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Option Explicit',
        'Public gblConn As DAO.Database',
        '',
        'Public Sub Ping()',
        '    Debug.Print 1',
        'End Sub',
      ].join('\n'),
    );

    const variables = variableNodes(result);
    expect(variables).toHaveLength(1);
    const gblConn = variables[0]!;
    expect(gblConn.name).toBe('gblConn');
    expect(gblConn.qualifiedName).toBe('modGlobals.gblConn');
    expect(gblConn.visibility).toBe('public');
    expect(gblConn.startLine).toBe(3);
    expect(gblConn.metadata).toMatchObject({
      declaredType: 'DAO.Database',
      isArray: false,
      isWithEvents: false,
      isConst: false,
    });

    expect(containsTargets(result).has(gblConn.id)).toBe(true);

    // The qualified-Dim `references` edge to the outer type is the
    // behaviour this issue must not double-count.
    const daoNode = result.nodes.find(
      (n) => n.kind === 'class' && n.name === 'DAO',
    );
    expect(daoNode).toBeDefined();
    const daoRefs = result.edges.filter(
      (e) => e.kind === 'references' && e.target === daoNode!.id,
    );
    expect(daoRefs).toHaveLength(1);
  });

  it('`Dim x As Long` inside a Sub emits no node', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public Sub Calcular()',
        '    Dim x As Long',
        '    x = 1',
        'End Sub',
      ].join('\n'),
    );
    expect(variableNodes(result)).toHaveLength(0);
    expect(moduleVarRefs(result)).toHaveLength(0);
  });

  it('a bare `Dim` with no `As` clause is Variant and module-private', () => {
    const result = extract(
      ['Attribute VB_Name = "modGlobals"', 'Dim gblFlag'].join('\n'),
    );
    const variables = variableNodes(result);
    expect(variables).toHaveLength(1);
    expect(variables[0]!.visibility).toBe('private');
    expect(variables[0]!.metadata?.declaredType).toBe('Variant');
  });

  it('an array declaration is flagged `isArray`', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Private m_items(1 To 10) As String',
      ].join('\n'),
    );
    expect(variableNodes(result)[0]!.metadata).toMatchObject({
      declaredType: 'String',
      isArray: true,
    });
  });

  it('a multi-variable declaration line emits one node per variable', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public gblAlpha As Long, gblBeta As Producto',
      ].join('\n'),
    );
    expect(variableNodes(result).map((n) => n.name).sort()).toEqual([
      'gblAlpha',
      'gblBeta',
    ]);
  });

  it('a re-declared name keeps a single node', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public gblFlag As Long',
        'Public gblFlag As Long',
      ].join('\n'),
    );
    expect(variableNodes(result)).toHaveLength(1);
  });

  it('a module whose only symbols are variables still gets its module node and attached contains edges', () => {
    const result = extract(
      ['Attribute VB_Name = "modGlobals"', 'Public gblConn As Object'].join('\n'),
    );
    const moduleNode = result.nodes.find((n) => n.kind === 'module');
    expect(moduleNode).toBeDefined();
    // A contains edge left with an empty source is dropped by `extract()`,
    // so its presence proves the pending re-attribution ran.
    expect(containsTargets(result).size).toBe(1);
  });
});

// ============================================================================
// 2. Visibility folding
// ============================================================================

describe('Issue #251 — declaration keyword decides visibility', () => {
  it.each([
    ['Public gblA As Long', 'public'],
    ['Global gblB As Long', 'public'],
    ['Private gblC As Long', 'private'],
    ['Dim gblD As Long', 'private'],
    ['Static gblE As Long', 'private'],
  ])('%s is %s', (declaration, expected) => {
    const result = extract(
      ['Attribute VB_Name = "modGlobals"', declaration].join('\n'),
    );
    expect(variableNodes(result)[0]!.visibility).toBe(expected);
  });
});

// ============================================================================
// 3. WithEvents
// ============================================================================

describe('Issue #251 — WithEvents fields are module state too', () => {
  it('emits a variable node flagged isWithEvents and leaves the subscribes-event edge alone', () => {
    const result = extract(
      [
        'VERSION 1.0 CLASS',
        'Attribute VB_Name = "Form_Cliente"',
        'Public WithEvents m_Form As Form_Pedido',
        '',
        'Private Sub m_Form_Load()',
        'End Sub',
      ].join('\n'),
      'src/forms/Form_Cliente.cls',
    );

    const variables = variableNodes(result);
    expect(variables).toHaveLength(1);
    expect(variables[0]!.name).toBe('m_Form');
    expect(variables[0]!.qualifiedName).toBe('Form_Cliente.m_Form');
    expect(variables[0]!.metadata).toMatchObject({
      declaredType: 'Form_Pedido',
      isWithEvents: true,
      isConst: false,
    });
    expect(containsTargets(result).has(variables[0]!.id)).toBe(true);

    const subscribes = result.edges.filter((e) => e.kind === 'subscribes-event');
    expect(subscribes).toHaveLength(1);
    expect(subscribes[0]!.metadata).toMatchObject({
      synthesizedBy: 'vba-withevents',
      variableName: 'm_Form',
    });
  });
});

// ============================================================================
// 4. Read / write references
// ============================================================================

describe('Issue #251 — reads and writes of module-level variables', () => {
  it('`x = gblConn` is a property-get and `gblConn = Nothing` is a property-set', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public gblConn As Object',
        '',
        'Public Sub Leer()',
        '    Dim x As Object',
        '    x = gblConn',
        'End Sub',
        '',
        'Public Sub Limpiar()',
        '    gblConn = Nothing',
        'End Sub',
      ].join('\n'),
    );

    const refs = moduleVarRefs(result);
    expect(refs).toHaveLength(2);

    const read = refs.find((r) => r.referenceKind === 'property-get');
    expect(read).toBeDefined();
    expect(read!.referenceName).toBe('gblConn');
    expect(read!.line).toBe(6);
    expect(read!.metadata).toMatchObject({
      synthesizedBy: 'vba-module-var',
      access: 'read',
    });

    const write = refs.find((r) => r.referenceKind === 'property-set');
    expect(write).toBeDefined();
    expect(write!.line).toBe(10);
    expect(write!.metadata).toMatchObject({
      synthesizedBy: 'vba-module-var',
      access: 'write',
    });

    // Each reference is attributed to the procedure it appears in.
    expect(read!.fromNodeId).not.toBe(write!.fromNodeId);
  });

  it('`Set gblConn = New Foo` is a write, not a read', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public gblConn As Conexion',
        '',
        'Public Sub Abrir()',
        '    Set gblConn = New Conexion',
        'End Sub',
      ].join('\n'),
    );
    const refs = moduleVarRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceKind).toBe('property-set');
  });

  it('a read inside a condition or an argument stays a property-get', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public gblFlag As Boolean',
        '',
        'Public Sub Comprobar()',
        '    If gblFlag = True Then Exit Sub',
        'End Sub',
      ].join('\n'),
    );
    const refs = moduleVarRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceKind).toBe('property-get');
  });

  it('repeated accesses of the same direction from one procedure collapse to one reference, reads and writes stay separate', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public gblContador As Long',
        '',
        'Public Sub Contar()',
        '    Debug.Print gblContador',
        '    Debug.Print gblContador',
        '    gblContador = 0',
        'End Sub',
      ].join('\n'),
    );
    const refs = moduleVarRefs(result);
    expect(refs.map((r) => r.referenceKind).sort()).toEqual([
      'property-get',
      'property-set',
    ]);
  });

  it('a name that only appears inside a string literal is not an access', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public gblConn As Object',
        '',
        'Public Sub Avisar()',
        '    MsgBox "gblConn no disponible"',
        'End Sub',
      ].join('\n'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });

  it('a member of another object spelled like the module variable is not an access', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public gblConn As Object',
        '',
        'Public Sub Copiar(ByVal origen As Object)',
        '    Debug.Print origen.gblConn',
        'End Sub',
      ].join('\n'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });

  it('a module with no module-level variables emits no module-variable references at all', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public Sub Trabajar()',
        '    Dim total As Long',
        '    total = 1',
        'End Sub',
      ].join('\n'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });
});

// ============================================================================
// 5. The issue #205 scoping rule must keep holding
// ============================================================================

describe('Issue #251 — procedure-local declarations shadow module state (#205)', () => {
  it('a procedure that declares its own `codigo` reports no access to the module-level `codigo`, while a procedure that does not still does', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public codigo As String',
        '',
        'Public Sub ConLocal()',
        '    Dim codigo As String',
        '    codigo = "X"',
        '    Debug.Print codigo',
        'End Sub',
        '',
        'Public Sub SinLocal()',
        '    Debug.Print codigo',
        'End Sub',
      ].join('\n'),
    );

    // The local declaration produces no node of its own.
    const variables = variableNodes(result);
    expect(variables).toHaveLength(1);
    expect(variables[0]!.startLine).toBe(2);

    const refs = moduleVarRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceKind).toBe('property-get');
    expect(refs[0]!.line).toBe(11);

    const sinLocal = result.nodes.find(
      (n) => n.kind === 'function' && n.name === 'SinLocal',
    );
    expect(refs[0]!.fromNodeId).toBe(sinLocal!.id);
  });

  it('a bare untyped local shadows the module variable too', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public codigo As String',
        '',
        'Public Sub ConLocal()',
        '    Dim codigo',
        '    codigo = "X"',
        'End Sub',
      ].join('\n'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });

  it('a parameter shadows the module variable for the whole procedure body', () => {
    const result = extract(
      [
        'Attribute VB_Name = "modGlobals"',
        'Public codigo As String',
        '',
        'Public Sub Guardar(ByVal codigo As String)',
        '    Debug.Print codigo',
        'End Sub',
      ].join('\n'),
    );
    expect(moduleVarRefs(result)).toHaveLength(0);
  });
});

// ============================================================================
// 6. Resolution is file-scoped
// ============================================================================

describe('Issue #251 — module-variable references resolve inside their own file only', () => {
  function variableNode(filePath: string, name: string, id: string): Node {
    return {
      id,
      kind: 'variable',
      name,
      qualifiedName: `mod.${name}`,
      filePath,
      language: 'vba',
      startLine: 1,
      endLine: 1,
      startColumn: 0,
      endColumn: 0,
      updatedAt: 0,
    };
  }

  function contextWith(nodesByFile: Record<string, Node[]>): ResolutionContext {
    return {
      getNodesInFile: (filePath: string) => nodesByFile[filePath] ?? [],
    } as unknown as ResolutionContext;
  }

  const ref: UnresolvedRef = {
    fromNodeId: 'function:caller',
    referenceName: 'gblConn',
    referenceKind: 'property-get',
    line: 5,
    column: 4,
    filePath: 'src/modules/modA.bas',
    language: 'vba',
    metadata: { synthesizedBy: 'vba-module-var', access: 'read' },
  };

  it('binds to the variable declared in the referencing file', () => {
    const resolved = matchVbaModuleVariable(
      ref,
      contextWith({
        'src/modules/modA.bas': [
          variableNode('src/modules/modA.bas', 'gblConn', 'variable:a'),
        ],
        'src/modules/modB.bas': [
          variableNode('src/modules/modB.bas', 'gblConn', 'variable:b'),
        ],
      }),
    );
    expect(resolved?.targetNodeId).toBe('variable:a');
  });

  it('declines rather than binding an identically-named variable in another file', () => {
    const resolved = matchVbaModuleVariable(
      ref,
      contextWith({
        'src/modules/modB.bas': [
          variableNode('src/modules/modB.bas', 'gblConn', 'variable:b'),
        ],
      }),
    );
    expect(resolved).toBeNull();
  });

  it('ignores references it does not own', () => {
    expect(
      matchVbaModuleVariable(
        { ...ref, metadata: { synthesizedBy: 'vba-me-control' } },
        contextWith({}),
      ),
    ).toBeNull();
  });
});
