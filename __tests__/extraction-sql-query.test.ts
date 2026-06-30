/**
 * SqlQueryExtractor — tests for Dysflow-exported saved queries (`queries/*.sql`).
 *
 * Dysflow exports each saved Access QueryDef as `queries/<Name>.sql` (raw SQL)
 * alongside a `queries.json` manifest. Before this slice CodeGraph did not
 * recognize `.sql` at all, so the entire saved-query/data layer was invisible.
 *
 * This extractor models each query file as a `query` node and emits a
 * `references` edge to every table named after `FROM`/`JOIN`/`INTO`/`UPDATE`,
 * tagged `metadata.synthesizedBy = 'sql-query-table'`, so the data layer is
 * queryable and table usage is traceable.
 *
 * Detection (a `.sql` is only treated as a Dysflow query when a sibling
 * `queries.json` exists) is enforced at the directory-discovery layer and is
 * covered by the real-fixtures E2E, not here — this file unit-tests the
 * extractor's output shape given SQL text.
 */
import { describe, it, expect } from 'vitest';
import { SqlQueryExtractor } from '../src/extraction/sql-query-extractor';

function extract(filePath: string, source: string) {
  return new SqlQueryExtractor(filePath, source).extract();
}

describe('SqlQueryExtractor — query node', () => {
  it('emits a query node named after the file basename', () => {
    const r = extract('queries/Consulta3.sql', 'SELECT * FROM TbACParaLista;');
    const q = r.nodes.find((n) => n.kind === 'query');
    expect(q).toBeDefined();
    expect(q?.name).toBe('Consulta3');
    expect(q?.qualifiedName).toBe('Consulta3');
    expect(q?.language).toBe('sql');
  });

  it('emits a file node for the query file', () => {
    const r = extract('queries/Consulta3.sql', 'SELECT 1 AS One;');
    const file = r.nodes.find((n) => n.kind === 'file');
    expect(file).toBeDefined();
  });
});

describe('SqlQueryExtractor — table references', () => {
  it('emits a references edge to a FROM table', () => {
    const r = extract('queries/q.sql', 'SELECT * FROM TbACParaLista ORDER BY x;');
    const q = r.nodes.find((n) => n.kind === 'query');
    const table = r.nodes.find((n) => n.name === 'TbACParaLista');
    expect(table).toBeDefined();
    const edge = r.edges.find(
      (e) => e.kind === 'references' && e.source === q?.id && e.target === table?.id,
    );
    expect(edge).toBeDefined();
    expect(edge?.metadata?.synthesizedBy).toBe('sql-query-table');
  });

  it('captures a JOIN table', () => {
    const r = extract(
      'queries/q.sql',
      'SELECT a.* FROM TbA a INNER JOIN TbB b ON a.id = b.id;',
    );
    const names = r.nodes.filter((n) => n.name === 'TbA' || n.name === 'TbB');
    expect(names.map((n) => n.name).sort()).toEqual(['TbA', 'TbB']);
  });

  it('captures an INTO target table', () => {
    const r = extract('queries/q.sql', 'INSERT INTO TbAudit (Id) VALUES (1);');
    expect(r.nodes.some((n) => n.name === 'TbAudit')).toBe(true);
  });

  it('captures an UPDATE target table', () => {
    const r = extract('queries/q.sql', 'UPDATE TbOrders SET Status = 1;');
    expect(r.nodes.some((n) => n.name === 'TbOrders')).toBe(true);
  });

  it('strips brackets from a bracketed table name with spaces', () => {
    const r = extract('queries/q.sql', 'SELECT * FROM [Order Details];');
    expect(r.nodes.some((n) => n.name === 'Order Details')).toBe(true);
    // The bracketed literal must not leak into a node name.
    expect(r.nodes.some((n) => n.name === '[Order')).toBe(false);
  });

  it('deduplicates a table referenced twice in the same query', () => {
    const r = extract(
      'queries/q.sql',
      'SELECT * FROM TbX WHERE id IN (SELECT id FROM TbX);',
    );
    const tableNodes = r.nodes.filter((n) => n.name === 'TbX');
    expect(tableNodes).toHaveLength(1);
  });

  it('does not throw on a query with no tables', () => {
    const r = extract('queries/q.sql', 'SELECT 1 AS One;');
    expect(r.errors).toHaveLength(0);
    expect(r.nodes.some((n) => n.kind === 'query')).toBe(true);
  });
});
