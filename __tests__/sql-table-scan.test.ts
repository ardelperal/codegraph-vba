/**
 * Issue #203 — RED tests for the SQL-table scan consolidation.
 *
 * Bug:
 *   - `SQL_TABLE_RE` is duplicated in three places
 *     (`vba/sql-wrapper.ts:80`, `sql-query-extractor.ts:57`,
 *     `vba-form-extractor.ts:617`).
 *   - When SQL is built by `&`-concatenation, `collectSqlWrapperChain`
 *     silently drops non-literal operands and joins surviving fragments
 *     with a space, so `FROM <tabla> WHERE x` becomes `FROM   WHERE x`
 *     — `\s+` happily crosses the gap and `WHERE` is emitted as a
 *     synthetic table reference. Same for `ORDER`, `SET`, `GROUP`,
 *     `HAVING`, etc.
 *   - `classifySqlAccess` (`vba/sql-wrapper.ts:91-96`) tags
 *     `UPDATE->SET` as a write, so "who writes to table X" reports are
 *     poisoned.
 *
 * Acceptance criteria (issue #203):
 *   1. `"DELETE FROM " & tabla & " WHERE x"` emits no table reference.
 *   2. No SQL reserved word is ever emitted as a table name.
 *   3. The regex exists in exactly one module.
 *   4. All three call sites import the shared scanner.
 *   5. Existing SQL extraction tests still pass.
 *
 * These tests target the proposed leaf module
 * `src/extraction/sql-table-scan.ts` exporting
 * `scanSqlTables(sql): {table, clause, access}[]`. The bug-exposing
 * tests at the bottom (issue #203 reproduction) live alongside the
 * extractor-level tests in `extraction-vba.test.ts` /
 * `extraction-sql-query.test.ts`; this file owns the SHARED module's
 * API contract so the three call sites can adopt it without each
 * replicating the test.
 */
import { describe, it, expect } from 'vitest';
import { scanSqlTables, SQL_RESERVED_TABLE_TOKENS } from '../src/extraction/sql-table-scan';

describe('scanSqlTables — basic FROM/JOIN/INTO/UPDATE clauses', () => {
  it('captures a plain FROM table', () => {
    const rows = scanSqlTables('SELECT * FROM tblCustomers');
    expect(rows).toEqual([
      { table: 'tblCustomers', clause: 'FROM', access: 'read' },
    ]);
  });

  it('captures a JOIN table', () => {
    const rows = scanSqlTables('SELECT * FROM tblA INNER JOIN tblB ON tblA.Id = tblB.Id');
    const tables = rows.map((r) => r.table);
    expect(tables).toEqual(['tblA', 'tblB']);
    rows.forEach((r) => expect(r.access).toBe('read'));
  });

  it('captures an INSERT INTO target as access=write', () => {
    const rows = scanSqlTables('INSERT INTO tblAudit (Id) VALUES (1)');
    expect(rows).toEqual([
      { table: 'tblAudit', clause: 'INTO', access: 'write' },
    ]);
  });

  it('captures an UPDATE target as access=write', () => {
    const rows = scanSqlTables('UPDATE tblOrders SET Status = 1');
    expect(rows).toEqual([
      { table: 'tblOrders', clause: 'UPDATE', access: 'write' },
    ]);
  });

  it('captures a DELETE FROM target as access=write (FROM after DELETE is a write, not a read)', () => {
    const rows = scanSqlTables('DELETE FROM tblOld');
    expect(rows).toEqual([{ table: 'tblOld', clause: 'FROM', access: 'write' }]);
  });

  it('captures an INSERT INTO ... SELECT FROM distinctly (target=write, source=read)', () => {
    const rows = scanSqlTables('INSERT INTO tblArchive SELECT * FROM tblLive');
    expect(rows.find((r) => r.table === 'tblArchive')).toEqual({
      table: 'tblArchive',
      clause: 'INTO',
      access: 'write',
    });
    expect(rows.find((r) => r.table === 'tblLive')).toEqual({
      table: 'tblLive',
      clause: 'FROM',
      access: 'read',
    });
  });

  it('strips brackets from a bracketed table name with spaces', () => {
    const rows = scanSqlTables('SELECT * FROM [Order Details]');
    expect(rows).toEqual([{ table: 'Order Details', clause: 'FROM', access: 'read' }]);
  });
});

describe('scanSqlTables — schema-qualified table names', () => {
  it('captures a schema-qualified table (dbo.tblCustomers) as one composite reference', () => {
    const rows = scanSqlTables('SELECT * FROM dbo.tblCustomers');
    expect(rows).toEqual([
      { table: 'dbo.tblCustomers', clause: 'FROM', access: 'read' },
    ]);
  });

  it('captures a bracketed schema-qualified table ([My Schema].[My Table]) as one composite reference', () => {
    const rows = scanSqlTables('SELECT * FROM [My Schema].[My Table]');
    expect(rows).toEqual([
      { table: 'My Schema.My Table', clause: 'FROM', access: 'read' },
    ]);
  });

  it('captures a schema-qualified JOIN table', () => {
    const rows = scanSqlTables('SELECT a.* FROM dbo.TbA a INNER JOIN sales.TbB b ON a.Id = b.Id');
    const tables = rows.map((r) => r.table);
    expect(tables).toEqual(['dbo.TbA', 'sales.TbB']);
  });
});

describe('scanSqlTables — reserved-word rejection (Issue #203 AC #2)', () => {
  // Reserved SQL tokens that legitimately follow FROM/JOIN/INTO/UPDATE in
  // a real statement. Capturing one of these as a "table name" is the bug.
  // (SQL_RESERVED_TABLE_TOKENS is exported as the canonical reject list so
  // consumers can introspect it.)
  it('exports SQL_RESERVED_TABLE_TOKENS as the canonical reject list', () => {
    expect(SQL_RESERVED_TABLE_TOKENS).toBeInstanceOf(Set);
    // Spot-check the obvious offenders from issue #203:
    expect(SQL_RESERVED_TABLE_TOKENS.has('WHERE')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('ORDER')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('SET')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('GROUP')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('HAVING')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('INNER')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('LEFT')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('RIGHT')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('OUTER')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('FULL')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('CROSS')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('ON')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('VALUES')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('SELECT')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('UNION')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('AS')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('DISTINCT')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('TOP')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('IN')).toBe(true);
    expect(SQL_RESERVED_TABLE_TOKENS.has('EXISTS')).toBe(true);
  });

  it('rejects WHERE as a table name (issue #203 reproduction — FROM `  WHERE x=1`)', () => {
    const rows = scanSqlTables('SELECT * FROM   WHERE x=1');
    expect(rows).toEqual([]);
  });

  it('rejects ORDER as a table name (issue #203 reproduction — DELETE FROM ` ORDER BY a)', () => {
    const rows = scanSqlTables('DELETE FROM   ORDER BY a');
    expect(rows).toEqual([]);
  });

  it('rejects SET as a table name (issue #203 reproduction — UPDATE ` SET a=1)', () => {
    const rows = scanSqlTables('UPDATE   SET a=1');
    expect(rows).toEqual([]);
  });

  it('rejects GROUP as a table name (FROM ` GROUP BY x)', () => {
    const rows = scanSqlTables('SELECT * FROM   GROUP BY x');
    expect(rows).toEqual([]);
  });

  it('rejects HAVING as a table name (FROM ` HAVING count > 0)', () => {
    const rows = scanSqlTables('SELECT * FROM   HAVING count > 0');
    expect(rows).toEqual([]);
  });

  it('rejects INNER as a table name (FROM ` INNER JOIN ...)', () => {
    const rows = scanSqlTables('SELECT * FROM   INNER JOIN tblB ON 1=1');
    expect(rows).toEqual([{ table: 'tblB', clause: 'JOIN', access: 'read' }]);
  });

  it('rejects SELECT as a table name', () => {
    const rows = scanSqlTables('INSERT INTO   SELECT 1');
    expect(rows).toEqual([]);
  });

  it('rejects a reserved word even when it is bracketed (defensive — [WHERE] still WHERE)', () => {
    // The bracketed form is unusual but legal in some dialects. Even
    // there, the bare identifier is the SQL keyword and should not be
    // emitted as a table name.
    const rows = scanSqlTables('SELECT * FROM [WHERE]');
    expect(rows).toEqual([]);
  });

  it('case-insensitive reserved-word rejection (FROM where x=1)', () => {
    const rows = scanSqlTables('select * from   where x=1');
    expect(rows).toEqual([]);
  });
});

describe('scanSqlTables — interaction with chained literals (Issue #203 AC #1)', () => {
  // The `scanSqlTables` API receives the *joined* SQL string after
  // `collectSqlWrapperChain`/`collectStringLiteralText` have merged all
  // literal fragments with a single space. When the dropped operand
  // sits BETWEEN a FROM clause keyword and a SQL reserved word
  // (e.g. `FROM  WHERE`), the gap is collapsed to a single space
  // and the reserved-word reject list (above) catches the bad capture.
  // Here we exercise the actual reproduction snippet the issue lists.

  it('"DELETE FROM " & tabla & " WHERE x" emits no table reference', () => {
    // Simulate what `collectSqlWrapperChain` produces today:
    //   [literal1, dropped, literal2].join(' ')
    //   = "DELETE FROM " + " " + " WHERE x"
    //   = "DELETE FROM   WHERE x"
    // The reserved-word reject list must drop the WHERE capture.
    const joined = ['DELETE FROM ', ' ', ' WHERE x'].join(' ');
    const rows = scanSqlTables(joined);
    expect(rows).toEqual([]);
  });

  it('"SELECT * FROM " & tabla & " WHERE x=1" emits no table reference', () => {
    const joined = ['SELECT * FROM ', ' ', ' WHERE x=1'].join(' ');
    const rows = scanSqlTables(joined);
    expect(rows).toEqual([]);
  });

  it('"UPDATE " & tabla & " SET a=1" emits no table reference', () => {
    const joined = ['UPDATE ', ' ', ' SET a=1'].join(' ');
    const rows = scanSqlTables(joined);
    expect(rows).toEqual([]);
  });

  it('a literal table name still emits when present (regression guard)', () => {
    const joined = 'SELECT * FROM tblCustomers WHERE Id = 1';
    const rows = scanSqlTables(joined);
    expect(rows).toEqual([{ table: 'tblCustomers', clause: 'FROM', access: 'read' }]);
  });

  it('two adjacent table fragments across the dropped operand gap are NOT bridged (defensive)', () => {
    // `FROM tblA ` <dropped> ` tblB` — the dropped operand is a single
    // variable. There is NO keyword between tblA and the dropped
    // operand, so tblA is captured (real FROM table). tblB has no
    // preceding keyword, so it is NOT captured (no false edge).
    const joined = 'FROM tblA   tblB';
    const rows = scanSqlTables(joined);
    expect(rows).toEqual([{ table: 'tblA', clause: 'FROM', access: 'read' }]);
  });
});

describe('scanSqlTables — empty / malformed input', () => {
  it('returns [] on empty string', () => {
    expect(scanSqlTables('')).toEqual([]);
  });

  it('returns [] on a SQL-less statement', () => {
    expect(scanSqlTables('SELECT 1 AS One')).toEqual([]);
  });

  it('returns [] on whitespace only', () => {
    expect(scanSqlTables('   ')).toEqual([]);
  });
});