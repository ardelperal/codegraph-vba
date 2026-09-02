/**
 * Issue #256 (plan task T15) — SQL clause coverage.
 *
 * `src/extraction/sql-table-scan.ts` is the single scanner every SQL path
 * in the project shares (`vba/sql-wrapper.ts`, `vba-form-extractor.ts`,
 * `sql-query-extractor.ts`). Before this change it captured table names
 * after `FROM` / `JOIN` / `INTO` / `UPDATE` only, so:
 *
 *   - `CREATE TABLE` / `ALTER TABLE` / `DROP TABLE` targets were invisible
 *     — a table a module creates or drops produced no reference at all;
 *   - the Access-specific `IN "<path>"` clause, which points a query at an
 *     external database file, had no representation in the graph, so a
 *     cross-backend read/write looked like a plain local table access.
 *
 * What this file pins:
 *
 *   1. The three DDL verbs are captured and carry `access: 'write'`.
 *   2. `IN "<path>"` yields an EXTERNAL BACKEND path, normalized so the
 *      same file named from two different queries converges on ONE node.
 *      The external node is `file`-kind (it IS a file, just not one this
 *      index parsed) with `metadata: { external: true, backendPath }`,
 *      and the edge to it is tagged `synthesizedBy: 'vba-external-backend'`.
 *   3. The reserved-word reject list — the project's main defence against
 *      emitting `WHERE` as a table name — is UNCHANGED and still rejects
 *      every one of its tokens, including after the new DDL verbs.
 *
 * Point 3 is the highest-risk part of this change and is therefore tested
 * explicitly (pinned list contents + per-token rejection sweep), not
 * incidentally.
 */
import { describe, it, expect } from 'vitest';
import {
  scanSqlTables,
  scanSqlExternalBackends,
  normalizeBackendPath,
  buildExternalBackendNode,
  EXTERNAL_BACKEND_SYNTHESIZED_BY,
  SQL_RESERVED_TABLE_TOKENS,
} from '../src/extraction/sql-table-scan';
import { SqlQueryExtractor } from '../src/extraction/sql-query-extractor';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';

// ---------------------------------------------------------------------------
// 1. DDL verbs
// ---------------------------------------------------------------------------

describe('scanSqlTables — DDL verbs (Issue #256)', () => {
  it('captures a CREATE TABLE target as access=write', () => {
    expect(scanSqlTables('CREATE TABLE TbNueva (Id COUNTER, Nombre TEXT(50))')).toEqual([
      { table: 'TbNueva', clause: 'CREATE TABLE', access: 'write' },
    ]);
  });

  it('captures a DROP TABLE target as access=write', () => {
    expect(scanSqlTables('DROP TABLE TbVieja')).toEqual([
      { table: 'TbVieja', clause: 'DROP TABLE', access: 'write' },
    ]);
  });

  it('captures an ALTER TABLE target as access=write', () => {
    expect(scanSqlTables('ALTER TABLE TbClientes ADD COLUMN Nif TEXT(9)')).toEqual([
      { table: 'TbClientes', clause: 'ALTER TABLE', access: 'write' },
    ]);
  });

  it('is case-insensitive and tolerates extra whitespace between the two keywords', () => {
    expect(scanSqlTables('create   table  TbNueva (Id LONG)')).toEqual([
      { table: 'TbNueva', clause: 'CREATE TABLE', access: 'write' },
    ]);
  });

  it('unwraps a bracketed DDL target', () => {
    expect(scanSqlTables('DROP TABLE [Order Details]')).toEqual([
      { table: 'Order Details', clause: 'DROP TABLE', access: 'write' },
    ]);
  });

  it('does not capture CREATE INDEX (only the three TABLE verbs are modelled)', () => {
    // `CREATE INDEX idxA ON TbClientes (Nif)` — `ON` is a reserved word, so
    // the trailing table is dropped rather than guessed at.
    expect(scanSqlTables('CREATE INDEX idxA ON TbClientes (Nif)')).toEqual([]);
  });

  it('still captures the FROM/JOIN/INTO/UPDATE clauses alongside a DDL verb', () => {
    const rows = scanSqlTables('CREATE TABLE TbCopia AS SELECT * FROM TbOrigen');
    expect(rows).toEqual([
      { table: 'TbCopia', clause: 'CREATE TABLE', access: 'write' },
      { table: 'TbOrigen', clause: 'FROM', access: 'read' },
    ]);
  });

  it('drops a DDL capture whose target was lost to a dropped concat operand', () => {
    // `"CREATE TABLE " & nombre & " (Id LONG)"` reaches the scanner as the
    // `?` sentinel shape `vba/sql-wrapper.ts` produces.
    expect(scanSqlTables('CREATE TABLE  ?  (Id LONG)')).toEqual([]);
    expect(scanSqlTables('DROP TABLE  ?')).toEqual([]);
    expect(scanSqlTables('ALTER TABLE  ?  ADD COLUMN X TEXT(5)')).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// 2. Reserved-word reject list — explicit regression (highest risk)
// ---------------------------------------------------------------------------

describe('SQL_RESERVED_TABLE_TOKENS — the reject list is unchanged (Issue #256 regression)', () => {
  // Pinned verbatim. Adding the DDL verbs must NOT remove a token: this
  // list is the only thing standing between a malformed concat and a
  // `WHERE` node in the graph.
  const EXPECTED = [
    'AS',
    'CROSS',
    'DISTINCT',
    'EXISTS',
    'FULL',
    'GROUP',
    'HAVING',
    'IN',
    'INNER',
    'JOIN',
    'LEFT',
    'ON',
    'ORDER',
    'OUTER',
    'RIGHT',
    'SELECT',
    'SET',
    'TOP',
    'UNION',
    'VALUES',
    'WHERE',
  ];

  it('contains exactly the 21 canonical tokens, no more and no fewer', () => {
    expect([...SQL_RESERVED_TABLE_TOKENS].sort()).toEqual(EXPECTED);
  });

  it.each(EXPECTED)('never emits %s as a table name after any capturing clause', (token) => {
    const clauses = [
      `SELECT * FROM   ${token} x`,
      `DELETE FROM   ${token} x`,
      `INSERT INTO   ${token} x`,
      `UPDATE   ${token} x`,
      `SELECT * FROM a INNER JOIN   ${token} x`,
      `CREATE TABLE   ${token} x`,
      `ALTER TABLE   ${token} x`,
      `DROP TABLE   ${token}`,
    ];
    for (const sql of clauses) {
      const captured = scanSqlTables(sql).map((r) => r.table.toUpperCase());
      expect(captured, `"${sql}" leaked ${token} as a table name`).not.toContain(token);
    }
  });

  it.each(EXPECTED)('never emits %s as a schema-qualified table name either', (token) => {
    for (const sql of [`SELECT * FROM ${token}.Id`, `DROP TABLE ${token}.Id`]) {
      const captured = scanSqlTables(sql).map((r) => r.table.toUpperCase());
      expect(captured).not.toContain(`${token}.ID`);
    }
  });

  it('keeps the original Issue #203 reproduction cases at zero captures', () => {
    expect(scanSqlTables(['DELETE FROM ', ' ', ' WHERE x'].join(' '))).toEqual([]);
    expect(scanSqlTables(['SELECT * FROM ', ' ', ' WHERE x=1'].join(' '))).toEqual([]);
    expect(scanSqlTables(['UPDATE ', ' ', ' SET a=1'].join(' '))).toEqual([]);
    expect(scanSqlTables('SELECT * FROM   GROUP BY x')).toEqual([]);
    expect(scanSqlTables('SELECT * FROM   HAVING count > 0')).toEqual([]);
    expect(scanSqlTables('INSERT INTO   SELECT 1')).toEqual([]);
    expect(scanSqlTables('SELECT * FROM [WHERE]')).toEqual([]);
    expect(scanSqlTables('select * from   where x=1')).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// 3. Access `IN "<path>"` external-backend clause
// ---------------------------------------------------------------------------

describe('normalizeBackendPath (Issue #256)', () => {
  it('lowercases and forward-slashes a Windows path', () => {
    expect(normalizeBackendPath('C:\\Datos\\Otra.accdb')).toBe('c:/datos/otra.accdb');
  });

  it('collapses repeated separators and drops a trailing one', () => {
    expect(normalizeBackendPath('C:\\\\datos\\\\\\otra.accdb\\')).toBe('c:/datos/otra.accdb');
  });

  it('preserves the leading double separator of a UNC share', () => {
    expect(normalizeBackendPath('\\\\SERVIDOR\\Datos\\otra.accdb')).toBe(
      '//servidor/datos/otra.accdb',
    );
  });

  it('returns an empty string for blank input', () => {
    expect(normalizeBackendPath('   ')).toBe('');
  });
});

describe('scanSqlExternalBackends (Issue #256)', () => {
  it('captures a double-quoted IN path', () => {
    expect(scanSqlExternalBackends('SELECT * FROM T IN "C:\\datos\\otra.accdb"')).toEqual([
      'c:/datos/otra.accdb',
    ]);
  });

  it('captures a single-quoted IN path', () => {
    expect(scanSqlExternalBackends("SELECT * FROM T IN 'C:\\datos\\otra.accdb'")).toEqual([
      'c:/datos/otra.accdb',
    ]);
  });

  it('captures a VBA doubled-quote IN path (the shape a VBA literal produces)', () => {
    expect(scanSqlExternalBackends('SELECT * FROM T IN ""C:\\datos\\otra.accdb""')).toEqual([
      'c:/datos/otra.accdb',
    ]);
  });

  it('leaves the local table capture untouched', () => {
    const sql = 'SELECT * FROM T IN "C:\\datos\\otra.accdb"';
    expect(scanSqlTables(sql)).toEqual([{ table: 'T', clause: 'FROM', access: 'read' }]);
  });

  it('deduplicates the same path named twice in one statement', () => {
    const sql =
      'SELECT * FROM A IN "C:\\datos\\otra.accdb" UNION SELECT * FROM B IN "C:\\Datos\\OTRA.accdb"';
    expect(scanSqlExternalBackends(sql)).toEqual(['c:/datos/otra.accdb']);
  });

  it('ignores the IN value-list operator', () => {
    expect(scanSqlExternalBackends('SELECT * FROM T WHERE Id IN (1, 2, 3)')).toEqual([]);
    expect(scanSqlExternalBackends('SELECT * FROM T WHERE N IN ("a", "b")')).toEqual([]);
  });

  it('ignores an empty IN path (the Access ODBC/dBASE connect-string form)', () => {
    expect(scanSqlExternalBackends('SELECT * FROM T IN "" [ODBC;DSN=x]')).toEqual([]);
  });

  it('returns [] for empty or backend-less SQL', () => {
    expect(scanSqlExternalBackends('')).toEqual([]);
    expect(scanSqlExternalBackends('SELECT * FROM T')).toEqual([]);
  });
});

describe('buildExternalBackendNode (Issue #256)', () => {
  it('builds a file-kind node carrying external metadata', () => {
    const node = buildExternalBackendNode('c:/datos/otra.accdb');
    expect(node.kind).toBe('file');
    expect(node.name).toBe('otra.accdb');
    expect(node.qualifiedName).toBe('c:/datos/otra.accdb');
    expect(node.metadata).toMatchObject({ external: true, backendPath: 'c:/datos/otra.accdb' });
  });

  it('is byte-identical for the same path, so two callers converge on ONE node', () => {
    const a = buildExternalBackendNode('c:/datos/otra.accdb');
    const b = buildExternalBackendNode('c:/datos/otra.accdb');
    expect(a.id).toBe(b.id);
    expect({ ...a, updatedAt: 0 }).toEqual({ ...b, updatedAt: 0 });
  });

  it('gives different paths different ids', () => {
    expect(buildExternalBackendNode('c:/datos/otra.accdb').id).not.toBe(
      buildExternalBackendNode('c:/datos/tercera.accdb').id,
    );
  });
});

// ---------------------------------------------------------------------------
// 4. The three consumers of the shared scanner
// ---------------------------------------------------------------------------

describe('SqlQueryExtractor — external backend (Issue #256)', () => {
  function extract(filePath: string, source: string) {
    return new SqlQueryExtractor(filePath, source).extract();
  }

  it('emits one external file node plus a vba-external-backend edge from the query', () => {
    const r = extract('queries/Externa.sql', 'SELECT * FROM T IN "C:\\datos\\otra.accdb";');
    const query = r.nodes.find((n) => n.kind === 'query');
    const backend = r.nodes.find((n) => n.metadata?.external === true);
    expect(backend).toBeDefined();
    expect(backend?.kind).toBe('file');
    expect(backend?.metadata?.backendPath).toBe('c:/datos/otra.accdb');

    const edges = r.edges.filter(
      (e) => e.metadata?.synthesizedBy === EXTERNAL_BACKEND_SYNTHESIZED_BY,
    );
    expect(edges).toHaveLength(1);
    expect(edges[0]?.kind).toBe('references');
    expect(edges[0]?.source).toBe(query?.id);
    expect(edges[0]?.target).toBe(backend?.id);
  });

  it('still emits the local table reference alongside the external backend', () => {
    const r = extract('queries/Externa.sql', 'SELECT * FROM TbLocal IN "C:\\datos\\otra.accdb";');
    expect(r.nodes.some((n) => n.name === 'TbLocal')).toBe(true);
  });

  it('the same path named from two different queries converges on ONE file node', () => {
    const a = extract('queries/Uno.sql', 'SELECT * FROM A IN "C:\\datos\\otra.accdb";');
    const b = extract('queries/Dos.sql', "SELECT * FROM B IN 'c:/DATOS/Otra.accdb';");
    const na = a.nodes.find((n) => n.metadata?.external === true);
    const nb = b.nodes.find((n) => n.metadata?.external === true);
    expect(na).toBeDefined();
    expect(nb).toBeDefined();
    expect(na?.id).toBe(nb?.id);
    expect({ ...na!, updatedAt: 0 }).toEqual({ ...nb!, updatedAt: 0 });
  });

  it('emits no external node for a query with no IN clause', () => {
    const r = extract('queries/Plana.sql', 'SELECT * FROM TbLocal;');
    expect(r.nodes.some((n) => n.metadata?.external === true)).toBe(false);
  });
});

describe('VbaExtractor — external backend from in-code SQL (Issue #256)', () => {
  function extract(filePath: string, source: string) {
    return new VbaExtractor(filePath, source).extract();
  }

  it('emits an external file node + vba-external-backend edge for an IN clause', () => {
    const src = `Public Sub LeeExterna()
    getdb().Execute "SELECT * FROM TbRemota IN ""C:\\datos\\otra.accdb"""
End Sub`;
    const r = extract('src/modules/Externa.bas', src);
    const backend = r.nodes.find((n) => n.metadata?.external === true);
    expect(backend).toBeDefined();
    expect(backend?.kind).toBe('file');
    expect(backend?.metadata?.backendPath).toBe('c:/datos/otra.accdb');

    const edges = r.edges.filter(
      (e) => e.metadata?.synthesizedBy === EXTERNAL_BACKEND_SYNTHESIZED_BY,
    );
    expect(edges).toHaveLength(1);
    expect(edges[0]?.target).toBe(backend?.id);
    expect(edges[0]?.source).not.toBe('');
  });

  it('still emits the local table reference for the same statement', () => {
    const src = `Public Sub LeeExterna()
    getdb().Execute "SELECT * FROM TbRemota IN ""C:\\datos\\otra.accdb"""
End Sub`;
    const r = extract('src/modules/Externa.bas', src);
    const sqlEdges = r.edges.filter((e) => e.metadata?.synthesizedBy === 'vba-sql-table');
    const names = sqlEdges.map((e) => r.nodes.find((n) => n.id === e.target)?.name);
    expect(names).toContain('TbRemota');
  });

  it('emits a write reference for a DDL statement executed in code', () => {
    const src = `Public Sub CreaTabla()
    getdb().Execute "CREATE TABLE TbNueva (Id LONG)"
    getdb().Execute "DROP TABLE TbVieja"
End Sub`;
    const r = extract('src/modules/Ddl.bas', src);
    const sqlEdges = r.edges.filter((e) => e.metadata?.synthesizedBy === 'vba-sql-table');
    const byName = new Map(
      sqlEdges.map((e) => [r.nodes.find((n) => n.id === e.target)?.name, e.metadata?.access]),
    );
    expect(byName.get('TbNueva')).toBe('write');
    expect(byName.get('TbVieja')).toBe('write');
  });

  it('emits no external node when the module has no IN clause', () => {
    const src = `Public Sub Normal()
    getdb().Execute "SELECT * FROM TbLocal"
End Sub`;
    const r = extract('src/modules/Normal.bas', src);
    expect(r.nodes.some((n) => n.metadata?.external === true)).toBe(false);
  });
});

describe('VbaFormExtractor — external backend from a RecordSource (Issue #256)', () => {
  function extract(filePath: string, source: string) {
    return new VbaFormExtractor(filePath, source).extract();
  }

  it('emits an external file node + vba-external-backend edge from the form-layout node', () => {
    const src = `Attribute VB_Name = "Form_Externa"
Begin Form
    RecordSource = "SELECT * FROM TbRemota IN ""C:\\datos\\otra.accdb"""
End`;
    const r = extract('src/forms/Form_Externa.form.txt', src);
    const formNode = r.nodes.find((n) => n.kind === 'form-layout');
    const backend = r.nodes.find((n) => n.metadata?.external === true);
    expect(backend).toBeDefined();
    expect(backend?.kind).toBe('file');
    expect(backend?.metadata?.backendPath).toBe('c:/datos/otra.accdb');

    const edges = r.edges.filter(
      (e) => e.metadata?.synthesizedBy === EXTERNAL_BACKEND_SYNTHESIZED_BY,
    );
    expect(edges).toHaveLength(1);
    expect(edges[0]?.source).toBe(formNode?.id);
    expect(edges[0]?.target).toBe(backend?.id);
  });

  it('still emits the bound table reference alongside the external backend', () => {
    const src = `Attribute VB_Name = "Form_Externa"
Begin Form
    RecordSource = "SELECT * FROM TbRemota IN ""C:\\datos\\otra.accdb"""
End`;
    const r = extract('src/forms/Form_Externa.form.txt', src);
    expect(r.nodes.some((n) => n.name === 'TbRemota' && n.kind === 'class')).toBe(true);
  });

  it('emits no external node for a plain bare-name RecordSource', () => {
    const src = `Attribute VB_Name = "Form_Local"
Begin Form
    RecordSource = "TbLocal"
End`;
    const r = extract('src/forms/Form_Local.form.txt', src);
    expect(r.nodes.some((n) => n.metadata?.external === true)).toBe(false);
  });
});
