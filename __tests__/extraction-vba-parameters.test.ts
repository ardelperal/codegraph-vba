/**
 * Issue #257 (task T16, half A) — `parameter` nodes for VBA procedure
 * signatures.
 *
 * Issue #250 already parses every signature and stamps `metadata.params` onto
 * the `function` node. This half turns that parsed shape into first-class
 * nodes so "which procedures take a `Cliente`?" has an answer, and pins:
 *
 *   1. the node shape — id scoping, `qualifiedName`, `startLine`, metadata;
 *   2. the `contains` edge from the owning procedure;
 *   3. `type_of` firing for a project class and NOT for a primitive;
 *   4. a `_`-continued signature producing exactly the one-line result;
 *   5. `ParamArray` and `Optional … = default` metadata;
 *   6. the two deliberate EXCLUSIONS — `parameter` must stay out of the
 *      context builder's default node filter and out of the MCP layer's
 *      container kinds, or ~4,700 parameters flood every response.
 */
import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph } from '../src';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { Node, Edge } from '../src/types';

function extract(source: string, filePath = 'src/modules/Signatures.bas') {
  return new VbaExtractor(filePath, source).extract();
}

function parameters(nodes: Node[]): Node[] {
  return nodes.filter((n) => n.kind === 'parameter');
}

function parameter(nodes: Node[], name: string): Node | undefined {
  return parameters(nodes).find((n) => n.name === name);
}

/** Declared procedures only — synthesized call targets carry `metadata.stub`. */
function procedure(nodes: Node[], name: string): Node | undefined {
  return nodes.find(
    (n) => n.kind === 'function' && n.name === name && n.metadata?.stub !== true,
  );
}

function containsTargets(edges: Edge[], sourceId: string): Set<string> {
  return new Set(
    edges
      .filter((e) => e.kind === 'contains' && e.source === sourceId)
      .map((e) => e.target),
  );
}

function typeOfEdges(edges: Edge[], sourceId: string): Edge[] {
  return edges.filter((e) => e.kind === 'type_of' && e.source === sourceId);
}

function synthTypeNode(nodes: Node[], name: string): Node | undefined {
  return nodes.find((n) => n.kind === 'class' && n.name === name);
}

// ============================================================================
// 1. Node shape
// ============================================================================

describe('Issue #257 — parameter node shape', () => {
  it('emits one node per parameter, qualifiedName scoped to module and procedure', () => {
    const { nodes } = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Option Explicit',
        '',
        'Public Sub Registrar(ByVal codigo As String, ByRef total As Currency)',
        'End Sub',
      ].join('\n'),
    );

    const params = parameters(nodes);
    expect(params.map((p) => p.name)).toEqual(['codigo', 'total']);
    expect(params.map((p) => p.qualifiedName)).toEqual([
      'modVentas.Registrar.codigo',
      'modVentas.Registrar.total',
    ]);
    for (const p of params) {
      expect(p.filePath).toBe('src/modules/Signatures.bas');
      expect(p.language).toBe('vba');
      // Parameters have no line of their own — they carry the procedure's.
      expect(p.startLine).toBe(4);
      expect(p.endLine).toBe(4);
    }
  });

  it('records position, byRef, optional, isArray, hasDefault and declaredType', () => {
    const { nodes } = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Public Sub Registrar(ByVal codigo As String, ByRef lineas() As Long, Optional ByVal nota As String = "")',
        'End Sub',
      ].join('\n'),
    );

    expect(parameter(nodes, 'codigo')?.metadata).toEqual({
      position: 0,
      byRef: false,
      optional: false,
      isArray: false,
      hasDefault: false,
      declaredType: 'String',
    });
    expect(parameter(nodes, 'lineas')?.metadata).toEqual({
      position: 1,
      byRef: true,
      optional: false,
      isArray: true,
      hasDefault: false,
      declaredType: 'Long',
    });
    expect(parameter(nodes, 'nota')?.metadata).toEqual({
      position: 2,
      byRef: false,
      optional: true,
      isArray: false,
      hasDefault: true,
      declaredType: 'String',
    });
  });

  it('stamps the VBA default type on an untyped parameter and emits no type edge', () => {
    const { nodes, edges } = extract(
      ['Attribute VB_Name = "modVentas"', 'Public Sub Ping(algo)', 'End Sub'].join('\n'),
    );

    const algo = parameter(nodes, 'algo');
    expect(algo?.metadata?.declaredType).toBe('Variant');
    expect(typeOfEdges(edges, algo!.id)).toHaveLength(0);
  });

  it('a parameterless procedure emits no parameter node at all', () => {
    const { nodes } = extract(
      ['Attribute VB_Name = "modVentas"', 'Public Sub Bare()', 'End Sub'].join('\n'),
    );
    expect(parameters(nodes)).toHaveLength(0);
  });

  it('gives each Property accessor its own parameter node', () => {
    const { nodes } = extract(
      [
        'Attribute VB_Name = "Cliente"',
        'Public Property Let Nombre(ByVal valor As String)',
        'End Property',
        'Public Property Set Origen(ByVal valor As Object)',
        'End Property',
      ].join('\n'),
      'src/classes/Cliente.cls',
    );

    const params = parameters(nodes);
    expect(params).toHaveLength(2);
    expect(new Set(params.map((p) => p.id)).size).toBe(2);
    expect(params.map((p) => p.qualifiedName).sort()).toEqual([
      'Cliente.Nombre.valor',
      'Cliente.Origen.valor',
    ]);
  });
});

// ============================================================================
// 2. contains
// ============================================================================

describe('Issue #257 — the owning procedure contains its parameters', () => {
  it('emits a contains edge from the function node to every parameter', () => {
    const { nodes, edges } = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Public Function Total(ByVal a As Long, ByVal b As Long) As Long',
        'End Function',
      ].join('\n'),
    );

    const fn = procedure(nodes, 'Total')!;
    const owned = containsTargets(edges, fn.id);
    expect(owned.has(parameter(nodes, 'a')!.id)).toBe(true);
    expect(owned.has(parameter(nodes, 'b')!.id)).toBe(true);
    expect(owned.size).toBe(2);
  });

  it('does not attach parameters to the module node', () => {
    const { nodes, edges } = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Public Sub Registrar(ByVal codigo As String)',
        'End Sub',
      ].join('\n'),
    );

    const moduleNode = nodes.find((n) => n.kind === 'module')!;
    const owned = containsTargets(edges, moduleNode.id);
    expect(owned.has(parameter(nodes, 'codigo')!.id)).toBe(false);
  });
});

// ============================================================================
// 3. type_of — project classes only
// ============================================================================

describe('Issue #257 — type_of fires only for non-primitive declared types', () => {
  it('emits type_of onto the type node for a project class parameter', () => {
    const { nodes, edges } = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Public Sub Guardar(ByVal cli As Cliente)',
        'End Sub',
      ].join('\n'),
    );

    const cli = parameter(nodes, 'cli')!;
    const typeEdges = typeOfEdges(edges, cli.id);
    expect(typeEdges).toHaveLength(1);
    expect(typeEdges[0]!.metadata?.synthesizedBy).toBe('vba-parameter-type');
    expect(typeEdges[0]!.target).toBe(synthTypeNode(nodes, 'Cliente')!.id);
  });

  it('emits nothing for primitive parameters', () => {
    const { nodes, edges } = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Public Sub Mezcla(ByVal n As Long, ByVal s As String, ByVal v As Variant, ByVal f As Boolean)',
        'End Sub',
      ].join('\n'),
    );

    for (const name of ['n', 's', 'v', 'f']) {
      expect(typeOfEdges(edges, parameter(nodes, name)!.id)).toHaveLength(0);
    }
    expect(edges.filter((e) => e.kind === 'type_of')).toHaveLength(0);
  });

  it('shares one type node between a parameter and a Dim of the same type', () => {
    const { nodes, edges } = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Public Sub Guardar(ByVal cli As Cliente)',
        '    Dim otro As Cliente',
        'End Sub',
      ].join('\n'),
    );

    const clienteNodes = nodes.filter(
      (n) => n.kind === 'class' && n.name === 'Cliente',
    );
    expect(clienteNodes).toHaveLength(1);
    expect(
      typeOfEdges(edges, parameter(nodes, 'cli')!.id)[0]!.target,
    ).toBe(clienteNodes[0]!.id);
  });

  it('points a qualified type at its outer segment, like the Dim sweep does', () => {
    const { nodes, edges } = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Public Sub Leer(ByVal rs As DAO.Recordset)',
        'End Sub',
      ].join('\n'),
    );

    const rs = parameter(nodes, 'rs')!;
    expect(rs.metadata?.declaredType).toBe('DAO.Recordset');
    const typeEdges = typeOfEdges(edges, rs.id);
    expect(typeEdges).toHaveLength(1);
    expect(typeEdges[0]!.target).toBe(synthTypeNode(nodes, 'DAO')!.id);
  });
});

// ============================================================================
// 4. Line continuations
// ============================================================================

describe('Issue #257 — a continued signature yields the one-line result', () => {
  it('produces identical parameter nodes for the joined and the split form', () => {
    const oneLine = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Public Sub Registrar(ByVal codigo As String, ByRef total As Currency, Optional ByVal nota As String = "")',
        'End Sub',
      ].join('\n'),
    );
    const continued = extract(
      [
        'Attribute VB_Name = "modVentas"',
        'Public Sub Registrar(ByVal codigo As String, _',
        '                     ByRef total As Currency, _',
        '                     Optional ByVal nota As String = "")',
        'End Sub',
      ].join('\n'),
    );

    const shape = (r: ReturnType<typeof extract>) =>
      parameters(r.nodes).map((p) => ({
        name: p.name,
        qualifiedName: p.qualifiedName,
        startLine: p.startLine,
        metadata: p.metadata,
      }));

    expect(shape(continued)).toEqual(shape(oneLine));
    expect(shape(continued)).toHaveLength(3);
    // Both forms declare the procedure on line 2, so the ids match too.
    expect(parameters(continued.nodes).map((p) => p.id)).toEqual(
      parameters(oneLine.nodes).map((p) => p.id),
    );
  });
});

// ============================================================================
// 5. ParamArray
// ============================================================================

describe('Issue #257 — ParamArray', () => {
  it('emits a node for the ParamArray tail flagged as an array, not as optional', () => {
    const { nodes } = extract(
      [
        'Attribute VB_Name = "modLog"',
        'Public Sub Escribir(ByVal nivel As Long, ParamArray partes() As Variant)',
        'End Sub',
      ].join('\n'),
    );

    expect(parameters(nodes).map((p) => p.name)).toEqual(['nivel', 'partes']);
    expect(parameter(nodes, 'partes')?.metadata).toEqual({
      position: 1,
      byRef: true,
      optional: false,
      isArray: true,
      hasDefault: false,
      declaredType: 'Variant',
    });
    // The unbounded arity stays where issue #250 put it — on the procedure.
    expect(procedure(nodes, 'Escribir')?.metadata?.arity).toEqual({
      required: 1,
      total: null,
    });
  });
});

// ============================================================================
// 6. The two deliberate exclusions
// ============================================================================

describe('Issue #257 — `parameter` stays out of the default result surfaces', () => {
  it('is absent from HIGH_VALUE_NODE_KINDS', () => {
    const source = fs.readFileSync(
      path.join(__dirname, '..', 'src', 'context', 'index.ts'),
      'utf-8',
    );
    const block = /const HIGH_VALUE_NODE_KINDS: NodeKind\[\] = \[([\s\S]*?)\];/.exec(
      source,
    );
    expect(block).not.toBeNull();
    expect(block![1]).not.toMatch(/'parameter'/);
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
    expect(block![1]).not.toMatch(/'parameter'/);
  });
});

// ============================================================================
// 7. End-to-end: the type_of edge survives stub resolution
// ============================================================================

/**
 * `resolveVbaReferenceStubs` DELETES a synthetic type stub once it repoints the
 * edges into it, and both edge endpoint columns are `ON DELETE CASCADE`. A
 * `type_of` edge left behind would therefore not merely stay unresolved — it
 * would be destroyed outright. This drives the real indexing pipeline to prove
 * the edge comes out the other side pointing at the real class.
 */
const openProjects: Array<{ cg: CodeGraph; dir: string }> = [];

afterEach(async () => {
  while (openProjects.length > 0) {
    const { cg, dir } = openProjects.pop()!;
    try {
      await cg.close();
    } catch {
      // ignore close errors
    }
    try {
      if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // Windows keeps a handle on the SQLite file for a moment after close.
    }
  }
});

describe('Issue #257 — a parameter type resolves to the real class', () => {
  it('repoints the parameter type_of edge onto the declared class instead of cascading it away', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-parameters-'));
    fs.mkdirSync(path.join(dir, 'src', 'modules'), { recursive: true });
    fs.mkdirSync(path.join(dir, 'src', 'classes'), { recursive: true });
    fs.writeFileSync(
      path.join(dir, 'src', 'modules', 'modVentas.bas'),
      [
        'Attribute VB_Name = "modVentas"',
        'Option Explicit',
        '',
        'Public Sub Guardar(ByVal cli As Cliente, ByVal n As Long)',
        'End Sub',
      ].join('\n'),
    );
    fs.writeFileSync(
      path.join(dir, 'src', 'classes', 'Cliente.cls'),
      [
        'VERSION 1.0 CLASS',
        'BEGIN',
        "  MultiUse = -1  'True",
        'END',
        'Attribute VB_Name = "Cliente"',
        'Option Explicit',
        '',
        'Public Sub Ping()',
        'End Sub',
      ].join('\n'),
    );

    const cg = await CodeGraph.init(dir, { index: false });
    openProjects.push({ cg, dir });
    await cg.indexAll();

    const cli = cg
      .searchNodes('cli', { languages: ['vba'], kinds: ['parameter'] })
      .map((r) => r.node)
      .find((n) => n.qualifiedName === 'modVentas.Guardar.cli');
    expect(cli).toBeDefined();

    const typeEdges = cg
      .getOutgoingEdges(cli!.id)
      .filter((e) => e.kind === 'type_of');
    expect(typeEdges).toHaveLength(1);

    const target = cg.getNode(typeEdges[0]!.target);
    expect(target?.kind).toBe('class');
    expect(target?.name).toBe('Cliente');
    // The real declaration, not the synthetic stub the extractor emitted in
    // modVentas.bas.
    expect(target?.filePath.replace(/\\/g, '/')).toContain(
      'src/classes/Cliente.cls',
    );

    // The primitive parameter still has no type edge after resolution.
    const n = cg
      .searchNodes('n', { languages: ['vba'], kinds: ['parameter'] })
      .map((r) => r.node)
      .find((node) => node.qualifiedName === 'modVentas.Guardar.n');
    expect(n).toBeDefined();
    expect(cg.getOutgoingEdges(n!.id).filter((e) => e.kind === 'type_of')).toHaveLength(0);
  });
});
