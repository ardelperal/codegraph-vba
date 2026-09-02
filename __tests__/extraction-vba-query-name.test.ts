/**
 * Issue #253 (task T11 of `docs/vba-node-discovery-plan.md`) — a saved query
 * named from code resolves to the `query` node built from `queries/<Name>.sql`.
 *
 * `SqlQueryExtractor` already emits one `query` node per saved query and links
 * it to its tables, but the only code-side path into that node was
 * `DoCmd.OpenQuery`, which this corpus never uses. `OpenRecordset("nombre")`,
 * `QueryDefs("nombre")` and the variable form all named a saved query and
 * produced nothing, leaving the data layer extracted and then orphaned.
 *
 * The contract:
 *   - a wrapper handed a literal that is NOT a statement but IS a bare Access
 *     object name emits ONE `dao-query` reference tagged `vba-query-name`,
 *     sourced from the CALLING PROCEDURE so `procedure -> query -> table` is
 *     traversable end to end;
 *   - verb detection wins — anything that reads as SQL keeps the table path;
 *   - a name matching no `query` node DECLINES: it stays a `failed` reference
 *     and never becomes a fabricated table placeholder.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { CodeGraph } from '../src';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

/** Wrap statements in a procedure so the call site has a caller. */
function inSub(...statements: string[]): string {
  return ['Public Sub Run()', ...statements.map((s) => `  ${s}`), 'End Sub'].join('\n');
}

function queryRefs(result: ReturnType<typeof extract>) {
  return result.unresolvedReferences.filter(
    (u) =>
      (u.metadata as Record<string, unknown> | undefined)?.synthesizedBy ===
      'vba-query-name',
  );
}

function tableEdges(result: ReturnType<typeof extract>) {
  return result.edges.filter(
    (e) =>
      (e.metadata as Record<string, unknown> | undefined)?.synthesizedBy ===
      'vba-sql-table',
  );
}

describe('issue #253 — a saved query name reaches its query node', () => {
  it('OpenRecordset with a bare name emits one dao-query reference', () => {
    const result = extract(
      'modDatos.bas',
      inSub('Set rs = getdb().OpenRecordset("qryPendientes")'),
    );
    const refs = queryRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceName).toBe('qryPendientes');
    expect(refs[0]!.referenceKind).toBe('dao-query');
  });

  it('sources the reference from the calling procedure, not the module', () => {
    const result = extract(
      'modDatos.bas',
      inSub('Set rs = getdb().OpenRecordset("qryPendientes")'),
    );
    const proc = result.nodes.find((n) => n.kind === 'function' && n.name === 'Run');
    expect(proc).toBeDefined();
    expect(queryRefs(result)[0]!.fromNodeId).toBe(proc!.id);
  });

  it('emits NO node for the query name', () => {
    const result = extract(
      'modDatos.bas',
      inSub('Set rs = getdb().OpenRecordset("qryPendientes")'),
    );
    expect(result.nodes.some((n) => n.name === 'qryPendientes')).toBe(false);
  });

  it('verb detection wins: a real statement still takes the table path', () => {
    const result = extract(
      'modDatos.bas',
      inSub('Set rs = getdb().OpenRecordset("SELECT * FROM TbUsuarios")'),
    );
    expect(queryRefs(result)).toHaveLength(0);
    expect(tableEdges(result)).toHaveLength(1);
    expect(result.nodes.some((n) => n.name === 'TbUsuarios')).toBe(true);
  });

  it('a concatenated statement is never mistaken for a query name', () => {
    const result = extract(
      'modDatos.bas',
      inSub('Set rs = getdb().OpenRecordset("SELECT * FROM " & tabla)'),
    );
    expect(queryRefs(result)).toHaveLength(0);
  });

  it('QueryDefs with a literal name emits one dao-query reference', () => {
    const result = extract('modDatos.bas', inSub('Set qdf = db.QueryDefs("qryPendientes")'));
    const refs = queryRefs(result);
    expect(refs).toHaveLength(1);
    expect(refs[0]!.referenceName).toBe('qryPendientes');
  });

  it('QueryDefs with a variable emits nothing', () => {
    const result = extract('modDatos.bas', inSub('Set qdf = db.QueryDefs(nombre)'));
    expect(queryRefs(result)).toHaveLength(0);
  });

  it('names a punctuated literal as neither a query nor a table', () => {
    const result = extract(
      'modDatos.bas',
      inSub('Set rs = getdb().OpenRecordset("no-es-un-nombre!")'),
    );
    expect(queryRefs(result)).toHaveLength(0);
    expect(tableEdges(result)).toHaveLength(0);
  });

  it('accepts an Access name containing spaces', () => {
    const result = extract(
      'modDatos.bas',
      inSub('Set rs = getdb().OpenRecordset("Consulta De Pendientes")'),
    );
    expect(queryRefs(result).map((r) => r.referenceName)).toEqual([
      'Consulta De Pendientes',
    ]);
  });
});

describe('issue #253 — end-to-end: procedure -> query -> table', () => {
  let tmpDir: string | undefined;

  afterEach(() => {
    if (tmpDir && fs.existsSync(tmpDir)) {
      // Windows holds the SQLite file open a moment after `close()`; a failed
      // temp-dir cleanup must not fail an otherwise-passing assertion.
      try {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      } catch {
        /* best effort — see the Windows notes in CLAUDE.md */
      }
    }
    tmpDir = undefined;
  });

  it('connects both hops in one indexed project', async () => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-query-name-'));
    fs.mkdirSync(path.join(tmpDir, 'src/queries'), { recursive: true });
    fs.mkdirSync(path.join(tmpDir, 'src/modules'), { recursive: true });
    fs.writeFileSync(
      path.join(tmpDir, 'src/queries/qryPendientes.sql'),
      'SELECT Id, Nombre FROM TbExpedientes WHERE Estado = 1;\n',
    );
    // A `.sql` only reaches extraction when its directory also holds a
    // `queries.json` — the Dysflow saved-query gate.
    fs.writeFileSync(
      path.join(tmpDir, 'src/queries/queries.json'),
      JSON.stringify({ queries: [{ name: 'qryPendientes' }] }, null, 2),
    );
    fs.writeFileSync(
      path.join(tmpDir, 'src/modules/modDatos.bas'),
      [
        'Attribute VB_Name = "modDatos"',
        'Public Sub CargarPendientes()',
        '  Set rs = getdb().OpenRecordset("qryPendientes")',
        'End Sub',
        '',
      ].join('\n'),
    );

    const cg = CodeGraph.initSync(tmpDir);
    await cg.indexAll();

    const proc = cg
      .getNodesByKind('function')
      .find((n) => n.name === 'CargarPendientes');
    expect(proc).toBeDefined();

    const query = cg.getNodesByKind('query').find((n) => n.name === 'qryPendientes');
    expect(query).toBeDefined();

    // Hop 1: the calling procedure reaches the query node.
    const hop1 = cg.getOutgoingEdges(proc!.id).find((e) => e.target === query!.id);
    expect(hop1).toBeDefined();
    expect(hop1!.kind).toBe('references');

    // Hop 2: the query reaches its table. Without this the flow is
    // half-bridged, which is worse than not bridging it at all.
    const hop2Targets = cg
      .getOutgoingEdges(query!.id)
      .map((e) => cg.getNode(e.target))
      .filter((n): n is NonNullable<typeof n> => Boolean(n));
    expect(hop2Targets.map((n) => n.name)).toContain('TbExpedientes');

    cg.close();
  });

  it('declines a name no query node backs, without inventing a table', async () => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-query-name-miss-'));
    fs.mkdirSync(path.join(tmpDir, 'src/modules'), { recursive: true });
    fs.writeFileSync(
      path.join(tmpDir, 'src/modules/modDatos.bas'),
      [
        'Attribute VB_Name = "modDatos"',
        'Public Sub CargarPendientes()',
        '  Set rs = getdb().OpenRecordset("qryNoExiste")',
        'End Sub',
        '',
      ].join('\n'),
    );

    const cg = CodeGraph.initSync(tmpDir);
    await cg.indexAll();

    // No fabricated target carries the missing name.
    expect(cg.getNodesByKind('query').filter((n) => n.name === 'qryNoExiste')).toEqual([]);
    expect(cg.getNodesByKind('class').filter((n) => n.name === 'qryNoExiste')).toEqual([]);

    const proc = cg
      .getNodesByKind('function')
      .find((n) => n.name === 'CargarPendientes');
    expect(proc).toBeDefined();
    const targets = cg
      .getOutgoingEdges(proc!.id)
      .map((e) => cg.getNode(e.target)?.name)
      .filter(Boolean);
    expect(targets).not.toContain('qryNoExiste');

    cg.close();
  });
});
