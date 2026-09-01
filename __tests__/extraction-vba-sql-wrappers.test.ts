/**
 * Issue #244 — which VBA receivers count as a SQL execution site.
 *
 * Before this change the scanner recognised a database handle only when the
 * receiver identifier ENDED at `db`, so every per-backend accessor a real
 * multi-database Access project uses (`getdbHPS()`, `getdbExpedientes()`,
 * `dbToUse`, `m_dbLanzadera`, and every DAO `QueryDef` receiver) dropped its
 * SQL out of the graph — 36% of execution sites in one measured corpus, 26%
 * in another. A second defect dropped the parenthesised literal form
 * (`getdb().OpenRecordset("SELECT * FROM TbX")`) from BOTH the literal path
 * (which demanded whitespace before the quote) and the variable path (a
 * literal is not an identifier).
 *
 * The suites below pin, in order:
 *
 *   1. every MISS line the issue lists now produces a `vba-sql-table` edge;
 *   2. every MATCH line still produces exactly the same edge, once;
 *   3. the parenthesised literal form resolves;
 *   4. `vba.sqlWrappers` extends the defaults instead of replacing them, and
 *      a malformed value warns and falls back without throwing;
 *   5. the receiver heuristic stays narrow — `dbg.Print` must not match;
 *   6. the shared reserved-word gate in `sql-table-scan.ts` still rejects a
 *      SQL keyword exposed by dynamic table-name concatenation.
 *
 * Real files on disk for the config path, the real extractor everywhere. No
 * mocking.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { loadVbaConfig, clearProjectConfigCache } from '../src/project-config';
import { DEFAULT_SQL_WRAPPERS, compileSqlWrappers } from '../src/extraction/vba/sql-wrapper';
import type { Edge } from '../src/types';

const FILE = 'src/modules/SqlWrapperProbe.bas';

/** Wrap `body` in a procedure — the SQL sweep only runs inside one. */
function moduleWith(body: string[]): string {
  return [
    'Attribute VB_Name = "SqlWrapperProbe"',
    'Public Sub Probe()',
    ...body.map((l) => `    ${l}`),
    'End Sub',
    '',
  ].join('\n');
}

function sqlTableEdges(result: { edges: Edge[] }): Edge[] {
  return result.edges.filter((e) => e.metadata?.synthesizedBy === 'vba-sql-table');
}

/** Distinct table names reached through `vba-sql-table` for `body`. */
function tablesFor(body: string[], sqlWrappers?: readonly string[]): string[] {
  const result = new VbaExtractor(FILE, moduleWith(body), { sqlWrappers }).extract();
  const byId = new Map(result.nodes.map((n) => [n.id, n.name]));
  const names = sqlTableEdges(result).map((e) => byId.get(e.target) ?? e.target);
  return [...new Set(names)].sort();
}

describe('Issue #244 — every missed execution receiver now reaches its table', () => {
  // One case per MISS line in the issue body, each with a distinct table name
  // so a passing assertion can only come from the line under test.
  const misses: ReadonlyArray<readonly [string, string[], string]> = [
    ['getdbHPS().OpenRecordset', ['Set rs = getdbHPS().OpenRecordset(strSQL)'], 'TbSolicitudes'],
    ['getdbExpedientes().OpenRecordset', ['Set rs = getdbExpedientes().OpenRecordset(strSQL)'], 'TbExpedientes'],
    ['getdbLanzadera().OpenRecordset', ['Set rs = getdbLanzadera().OpenRecordset(strSQL)'], 'TbLanzadera'],
    ['dbUse.OpenRecordset', ['Set rs = dbUse.OpenRecordset(strSQL)'], 'TbUsuarios'],
    ['m_dbLanzadera.OpenRecordset', ['Set rs = m_dbLanzadera.OpenRecordset(strSQL)'], 'TbAplicaciones'],
    ['dbToUse.Execute', ['dbToUse.Execute strSQL'], 'TbRiesgos'],
    ['qdf.Execute (DAO QueryDef)', ['qdf.Execute strSQL'], 'TbAuditoria'],
    ['qd.Execute (DAO QueryDef)', ['qd.Execute strSQL'], 'TbCorreos'],
  ];

  for (const [label, callLines, table] of misses) {
    it(`captures ${table} through ${label}`, () => {
      expect(tablesFor([`strSQL = "SELECT * FROM ${table}"`, ...callLines])).toEqual([table]);
    });
  }

  it('captures the ADO pair configured by default', () => {
    expect(
      tablesFor([
        'strSQL = "SELECT * FROM TbAdo"',
        'Connection.Execute strSQL',
      ]),
    ).toEqual(['TbAdo']);
    expect(
      tablesFor([
        'strSQL = "SELECT * FROM TbAdoRs"',
        'Recordset.Open strSQL',
      ]),
    ).toEqual(['TbAdoRs']);
  });
});

describe('Issue #244 — the receivers that already worked are unchanged', () => {
  // The MATCH lines from the issue body. Each must still produce EXACTLY one
  // edge for its table: the unified scanner must not double-emit now that the
  // literal and variable paths share one regex pair.
  const matches: ReadonlyArray<readonly [string, string[], string]> = [
    ['getdb().Execute strSQL', ['strSQL = "DELETE FROM TbTemporal"', 'getdb().Execute strSQL'], 'TbTemporal'],
    ['p_db.OpenRecordset sql', ['sql = "SELECT * FROM TbPersonas"', 'p_db.OpenRecordset sql'], 'TbPersonas'],
    ['CurrentDb.Execute sql', ['sql = "UPDATE TbCambios SET x = 1"', 'CurrentDb.Execute sql'], 'TbCambios'],
    ['DoCmd.RunSQL literal', ['DoCmd.RunSQL "DELETE FROM TbBorrar"'], 'TbBorrar'],
    ['DoCmd.RunSQL variable', ['strSQL = "DELETE FROM TbBorrar2"', 'DoCmd.RunSQL strSQL'], 'TbBorrar2'],
    ['db.Execute literal', ['db.Execute "INSERT INTO TbLog (x) VALUES (1)"'], 'TbLog'],
    ['suffix receiver (pre-#244 shape)', ['MiBaseDatosdb.Execute "DELETE FROM TbLegacy"'], 'TbLegacy'],
    ['DBEngine receiver', ['strSQL = "SELECT * FROM TbMotor"', 'DBEngine.Execute strSQL'], 'TbMotor'],
  ];

  for (const [label, body, table] of matches) {
    it(`${label} still emits exactly one edge for ${table}`, () => {
      const result = new VbaExtractor(FILE, moduleWith(body), {}).extract();
      const byId = new Map(result.nodes.map((n) => [n.id, n.name]));
      const hits = sqlTableEdges(result).filter((e) => byId.get(e.target) === table);
      expect(hits).toHaveLength(1);
    });
  }

  it('keeps the access direction the shared scanner assigns', () => {
    const result = new VbaExtractor(
      FILE,
      moduleWith(['getdb().Execute "DELETE FROM TbTemporal"']),
      {},
    ).extract();
    expect(sqlTableEdges(result).map((e) => e.metadata?.access)).toEqual(['write']);
  });
});

describe('Issue #244 — the parenthesised literal form', () => {
  it('resolves getdb().OpenRecordset("SELECT * FROM TbX")', () => {
    expect(tablesFor(['Set rs = getdb().OpenRecordset("SELECT * FROM TbX")'])).toEqual(['TbX']);
  });

  it('resolves the parenthesised Execute form', () => {
    expect(tablesFor(['CurrentDb.Execute ("DELETE FROM TbY")'])).toEqual(['TbY']);
  });

  it('still resolves the no-paren literal form it used to be confused with', () => {
    expect(tablesFor(['getdb().Execute "DELETE FROM TbTemporal"'])).toEqual(['TbTemporal']);
  });
});

describe('Issue #244 — vba.sqlWrappers extends the defaults', () => {
  it('a bare identifier entry makes a project accessor resolve', () => {
    const body = ['conn.Execute "SELECT * FROM TbConfig"'];
    expect(tablesFor(body)).toEqual([]);
    expect(tablesFor(body, ['conn'])).toEqual(['TbConfig']);
  });

  it('a bare identifier entry matches the whole accessor family by prefix', () => {
    expect(
      tablesFor(['abrirConexionRiesgos.Execute "DELETE FROM TbFamilia"'], ['abrirConexion']),
    ).toEqual(['TbFamilia']);
  });

  it('a receiver.method entry only matches that method', () => {
    expect(tablesFor(['cnn.Execute "SELECT * FROM TbPair"'], ['cnn.Execute'])).toEqual(['TbPair']);
    // `Open` is not the configured method, so the pair form must not reach it.
    expect(tablesFor(['cnn.Open "SELECT * FROM TbPair"'], ['cnn.Execute'])).toEqual([]);
  });

  it('does NOT disable the defaults', () => {
    expect(
      tablesFor(
        [
          'conn.Execute "SELECT * FROM TbConfig"',
          'CurrentDb.Execute "DELETE FROM TbDefault"',
        ],
        ['conn'],
      ),
    ).toEqual(['TbConfig', 'TbDefault']);
  });

  it('compiles the defaults plus the project entries, de-duplicated', () => {
    const compiled = compileSqlWrappers(['conn', 'CurrentDb', 'cnn.Execute']);
    expect(compiled.matchers).toHaveLength(DEFAULT_SQL_WRAPPERS.length + 2);
    expect(compiled.matchers).toContainEqual({ receiver: 'conn', method: null });
    expect(compiled.matchers).toContainEqual({ receiver: 'cnn', method: 'execute' });
  });
});

describe('Issue #244 — codegraph.json validation warns and falls back', () => {
  let dir: string;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-sqlwrappers-'));
    clearProjectConfigCache();
  });

  afterEach(() => {
    clearProjectConfigCache();
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // ignore cleanup EPERM on Windows
    }
  });

  function writeConfig(vba: unknown): void {
    fs.writeFileSync(path.join(dir, 'codegraph.json'), JSON.stringify({ vba }));
  }

  it('returns undefined when the key is absent', () => {
    expect(loadVbaConfig(dir).sqlWrappers).toBeUndefined();
  });

  it('loads a well-formed list', () => {
    writeConfig({ sqlWrappers: ['getdbHPS', 'cnn.Execute'] });
    expect(loadVbaConfig(dir).sqlWrappers).toEqual(['getdbHPS', 'cnn.Execute']);
  });

  it('warns and ignores a non-array value without throwing', () => {
    writeConfig({ sqlWrappers: 42 });
    expect(() => loadVbaConfig(dir)).not.toThrow();
    expect(loadVbaConfig(dir).sqlWrappers).toBeUndefined();
  });

  it('drops entries that are not an identifier or a receiver.method pair', () => {
    writeConfig({ sqlWrappers: ['getdb.*', '(a+)+b', '', 7, 'cnn.Execute'] });
    expect(loadVbaConfig(dir).sqlWrappers).toEqual(['cnn.Execute']);
  });

  it('leaves the built-in wrappers in force when every entry is rejected', () => {
    writeConfig({ sqlWrappers: ['getdb.*'] });
    expect(loadVbaConfig(dir).sqlWrappers).toBeUndefined();
    expect(tablesFor(['CurrentDb.Execute "DELETE FROM TbDefault"'])).toEqual(['TbDefault']);
  });

  it('a rejected regex entry never reaches the scanner', () => {
    // `compileSqlWrappers` is the last line of defence: even handed a raw
    // regex directly it treats the string as an identifier form and drops it,
    // so no user-supplied pattern is ever compiled into the per-line scan.
    const compiled = compileSqlWrappers(['(a+)+b', 'getdb.*']);
    expect(compiled.matchers).toHaveLength(DEFAULT_SQL_WRAPPERS.length);
  });
});

describe('Issue #244 — the receiver heuristic stays narrow', () => {
  it('dbg.Print does not become a SQL execution site', () => {
    expect(
      tablesFor(['strSQL = "SELECT * FROM TbNunca"', 'dbg.Print strSQL']),
    ).toEqual([]);
  });

  it('Debug.Print does not become a SQL execution site', () => {
    expect(
      tablesFor(['strSQL = "SELECT * FROM TbNunca"', 'Debug.Print strSQL']),
    ).toEqual([]);
  });

  it('a receiver that merely starts with the pattern mid-word is ignored', () => {
    // `db_test` and `dbase` continue the word instead of starting a new
    // segment, so the `db` wrapper must not claim them.
    expect(
      tablesFor(['db_test.Execute "SELECT * FROM TbNunca"']),
    ).toEqual([]);
    expect(
      tablesFor(['dbase.Execute "SELECT * FROM TbNunca"']),
    ).toEqual([]);
  });

  it('matches the receiver case-insensitively at a segment boundary', () => {
    expect(tablesFor(['GETDBHps().Execute "DELETE FROM TbCase"'])).toEqual(['TbCase']);
  });

  it('an unrelated receiver with a SQL-ish method is ignored', () => {
    expect(
      tablesFor(['strSQL = "SELECT * FROM TbNunca"', 'objInforme.Execute strSQL']),
    ).toEqual([]);
  });

  it('an unresolved variable emits nothing rather than guessing', () => {
    expect(tablesFor(['getdbHPS().Execute strNuncaAsignada'])).toEqual([]);
  });
});

describe('Issue #244 — the shared reserved-word gate still holds', () => {
  it('a concatenated table name never emits the following SQL keyword', () => {
    expect(
      tablesFor(['getdbHPS().Execute "DELETE FROM " & tabla & " WHERE activo = 1"']),
    ).toEqual([]);
  });

  it('a concatenated name in the variable form is equally silent', () => {
    expect(
      tablesFor([
        'strSQL = "DELETE FROM " & tabla & " WHERE activo = 1"',
        'getdbExpedientes().Execute strSQL',
      ]),
    ).toEqual([]);
  });
});
