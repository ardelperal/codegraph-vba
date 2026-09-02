/**
 * Issue #252 — SQL assigned to a binding property at runtime.
 *
 * This corpus binds its data in code, not in the form designer: there are
 * 161 `.RowSource =` assignments across the `.cls` / `.bas` modules, while
 * `RecordSource` and `ControlSource` appear ZERO times in `00_EXPEDIENTES`'s
 * `.form.txt` files. The static `.form.txt` sweep (`vba-row-source` /
 * `vba-record-source`) was therefore scanning the empty half of the problem
 * and every table named in a runtime binding was invisible to the graph.
 *
 * The fix feeds the right-hand side of
 * `<anything>.(RowSource|RecordSource|ControlSource|Filter|OrderBy) =`
 * through the SAME pipeline the SQL wrapper path uses —
 * `collectConcatFragments` then `scanSqlTables` — so the `?` sentinel for
 * non-literal operands and the shared reserved-word reject list apply
 * unchanged. The edges are tagged `vba-row-source-dynamic` so they stay
 * separable from the static `vba-row-source` in audits.
 *
 * The suites below pin, in order:
 *
 *   1. the literal form on every one of the five binding properties;
 *   2. the variable form, including #13 accumulate and #204 procedure scoping;
 *   3. the negative half — a filter expression is not a FROM clause, a
 *      dynamically concatenated table name yields nothing, and a non-SQL
 *      payload (bare query name, `Value List`) is left alone;
 *   4. an exhaustive guard that no `SQL_RESERVED_TABLE_TOKENS` word can ever
 *      be captured through the new path.
 *
 * Real extractor, no mocking.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { SQL_RESERVED_TABLE_TOKENS } from '../src/extraction/sql-table-scan';
import type { Edge, Node } from '../src/types';

const FILE = 'src/forms/Form_Probe.cls';

const DYNAMIC = 'vba-row-source-dynamic';
const STATIC = 'vba-row-source';
const WRAPPER = 'vba-sql-table';

/** Wrap `body` in one procedure — the SQL sweep only runs inside one. */
function moduleWith(body: string[]): string {
  return [
    'Attribute VB_Name = "Form_Probe"',
    'Public Sub Probe()',
    ...body.map((l) => `    ${l}`),
    'End Sub',
    '',
  ].join('\n');
}

function extract(source: string): { nodes: Node[]; edges: Edge[] } {
  return new VbaExtractor(FILE, source).extract();
}

/** Distinct target names of `references` edges carrying `synthesizedBy`. */
function tablesFor(
  result: { nodes: Node[]; edges: Edge[] },
  synthesizedBy: string,
): string[] {
  const names = new Set<string>();
  for (const edge of result.edges) {
    if (edge.kind !== 'references') continue;
    if (edge.metadata?.synthesizedBy !== synthesizedBy) continue;
    const name = result.nodes.find((n) => n.id === edge.target)?.name;
    if (name) names.add(name);
  }
  return [...names].sort();
}

/** Table names reached by a single-procedure body, via the dynamic path. */
function dynamicTables(body: string[]): string[] {
  return tablesFor(extract(moduleWith(body)), DYNAMIC);
}

describe('Issue #252 — literal SQL assigned to a binding property', () => {
  it('captures the table of a RowSource literal', () => {
    expect(
      dynamicTables(['Me.cboUsuario.RowSource = "SELECT Id FROM TbUsuarios"']),
    ).toEqual(['TbUsuarios']);
  });

  it.each([
    ['RowSource', 'Me.cboUsuario.RowSource'],
    ['RecordSource', 'Me.RecordSource'],
    ['ControlSource', 'Me.txtNombre.ControlSource'],
    ['Filter', 'Me.Filter'],
    ['OrderBy', 'Me.OrderBy'],
  ])('covers the %s property', (_prop, lhs) => {
    expect(dynamicTables([`${lhs} = "SELECT Id FROM TbUsuarios"`])).toEqual([
      'TbUsuarios',
    ]);
  });

  it('resolves the bare `.Property =` form inside a With block', () => {
    expect(
      dynamicTables([
        'With Me.cboUsuario',
        '    .RowSource = "SELECT Id FROM TbUsuarios"',
        'End With',
      ]),
    ).toEqual(['TbUsuarios']);
  });

  it('captures every table of a multi-table statement', () => {
    expect(
      dynamicTables([
        'Me.RecordSource = "SELECT * FROM TbExpedientes INNER JOIN TbUsuarios ON A.Id = B.Id"',
      ]),
    ).toEqual(['TbExpedientes', 'TbUsuarios']);
  });

  it('joins same-line `&` literal chains before scanning', () => {
    expect(
      dynamicTables([
        'Me.RowSource = "SELECT Id " & "FROM TbUsuarios " & "WHERE Activo = 1"',
      ]),
    ).toEqual(['TbUsuarios']);
  });

  it('tags the edge `vba-row-source-dynamic`, never the static `vba-row-source`', () => {
    const result = extract(
      moduleWith(['Me.cboUsuario.RowSource = "SELECT Id FROM TbUsuarios"']),
    );
    expect(tablesFor(result, DYNAMIC)).toEqual(['TbUsuarios']);
    expect(tablesFor(result, STATIC)).toEqual([]);
    expect(tablesFor(result, WRAPPER)).toEqual([]);
  });

  it('leaves the wrapper path untouched on the same file', () => {
    const result = extract(
      moduleWith([
        'getdb().Execute "DELETE FROM TbLog"',
        'Me.RowSource = "SELECT Id FROM TbUsuarios"',
      ]),
    );
    expect(tablesFor(result, WRAPPER)).toEqual(['TbLog']);
    expect(tablesFor(result, DYNAMIC)).toEqual(['TbUsuarios']);
  });
});

describe('Issue #252 — the variable form', () => {
  it('resolves `.RecordSource = strSQL` to the same tables as the literal form', () => {
    const viaVariable = dynamicTables([
      'strSQL = "SELECT Id FROM TbUsuarios"',
      'Me.RecordSource = strSQL',
    ]);
    const viaLiteral = dynamicTables([
      'Me.RecordSource = "SELECT Id FROM TbUsuarios"',
    ]);
    expect(viaVariable).toEqual(viaLiteral);
    expect(viaVariable).toEqual(['TbUsuarios']);
  });

  it('keeps the #13 accumulate semantics — `sql = sql & "..."` adds, never replaces', () => {
    expect(
      dynamicTables([
        'strSQL = "SELECT * FROM TbExpedientes"',
        'strSQL = strSQL & " INNER JOIN TbUsuarios ON A.Id = B.Id"',
        'Me.RecordSource = strSQL',
      ]),
    ).toEqual(['TbExpedientes', 'TbUsuarios']);
  });

  it('keeps the #204 procedure scoping — a name from another procedure does not leak', () => {
    const source = [
      'Attribute VB_Name = "Form_Probe"',
      'Public Sub Builder()',
      '    strSQL = "SELECT * FROM TbExpedientes"',
      'End Sub',
      'Public Sub Binder()',
      '    Me.RecordSource = strSQL',
      'End Sub',
      '',
    ].join('\n');
    expect(tablesFor(extract(source), DYNAMIC)).toEqual([]);
  });

  it('falls back to a module-level assignment, exactly as the wrapper path does', () => {
    const source = [
      'Attribute VB_Name = "Form_Probe"',
      'Private strSQL As String',
      'strSQL = "SELECT * FROM TbExpedientes"',
      'Public Sub Binder()',
      '    Me.RecordSource = strSQL',
      'End Sub',
      '',
    ].join('\n');
    expect(tablesFor(extract(source), DYNAMIC)).toEqual(['TbExpedientes']);
  });

  it('never guesses at an unknown identifier', () => {
    expect(dynamicTables(['Me.RecordSource = strDesconocido'])).toEqual([]);
  });
});

describe('Issue #252 — what must NOT become a table reference', () => {
  it('a Filter expression is not a FROM clause', () => {
    expect(dynamicTables(['Me.Filter = "Id = 1"'])).toEqual([]);
  });

  it('an OrderBy expression is not a FROM clause', () => {
    expect(dynamicTables(['Me.OrderBy = "Nombre DESC"'])).toEqual([]);
  });

  it('a dynamically concatenated table name yields nothing — and never `FROM` or `WHERE`', () => {
    const tables = dynamicTables([
      'Me.cboUsuario.RowSource = "SELECT " & campo & " FROM " & tabla',
    ]);
    expect(tables).toEqual([]);
    expect(tables).not.toContain('FROM');
    expect(tables).not.toContain('WHERE');
  });

  it('a dynamic table name followed by a WHERE clause never captures `WHERE`', () => {
    const tables = dynamicTables([
      'Me.RecordSource = "SELECT * FROM " & tabla & " WHERE Activo = 1"',
    ]);
    expect(tables).toEqual([]);
    expect(tables).not.toContain('WHERE');
  });

  it('a bare saved-query name is out of scope — nothing is guessed', () => {
    expect(
      dynamicTables(['Me.cboUsuario.RowSource = "qryUsuariosActivos"']),
    ).toEqual([]);
  });

  it('a Value List payload is out of scope', () => {
    expect(
      dynamicTables([
        'Me.cboEstado.RowSourceType = "Value List"',
        'Me.cboEstado.RowSource = "Alta;Baja;Pendiente"',
      ]),
    ).toEqual([]);
  });

  it('a ControlSource bound to a plain field name yields nothing', () => {
    expect(dynamicTables(['Me.txtNombre.ControlSource = "Nombre"'])).toEqual([]);
  });

  it('a comparison is not an assignment', () => {
    expect(
      dynamicTables([
        'If Me.Filter = "SELECT Id FROM TbUsuarios" Then',
        'End If',
      ]),
    ).toEqual([]);
  });

  it('a commented-out binding is not scanned', () => {
    expect(
      dynamicTables(['\' Me.cboUsuario.RowSource = "SELECT Id FROM TbUsuarios"']),
    ).toEqual([]);
  });

  it('a sibling property whose name merely starts with a binding name is ignored', () => {
    expect(dynamicTables(['Me.FilterOn = True'])).toEqual([]);
  });
});

describe('Issue #252 — no reserved word can ever be captured', () => {
  const RESERVED_SHAPES: string[] = [
    'Me.RowSource = "SELECT " & campo & " FROM " & tabla',
    'Me.RowSource = "SELECT * FROM " & tabla & " WHERE Activo = 1"',
    'Me.RecordSource = "SELECT * FROM " & tabla & " ORDER BY Nombre"',
    'Me.RecordSource = "SELECT * FROM " & tabla & " GROUP BY Id"',
    'Me.RecordSource = "SELECT * FROM " & t1 & " INNER JOIN " & t2 & " ON A.Id = B.Id"',
    'Me.RecordSource = "UPDATE " & tabla & " SET Activo = 1"',
    'Me.RecordSource = "INSERT INTO " & tabla & " VALUES (1)"',
    'Me.RecordSource = "SELECT * FROM " & tabla & " UNION SELECT * FROM " & otra',
    'Me.Filter = "Id = 1"',
    'Me.OrderBy = "Nombre DESC"',
    'Me.cbo.RowSource = "SELECT TOP 10 * FROM " & tabla',
  ];

  it.each(RESERVED_SHAPES)('captures no reserved word from: %s', (line) => {
    for (const table of dynamicTables([line])) {
      expect(SQL_RESERVED_TABLE_TOKENS.has(table.toUpperCase())).toBe(false);
    }
  });

  it('captures no reserved word across every shape at once', () => {
    const captured = dynamicTables(RESERVED_SHAPES);
    const reserved = captured.filter((t) =>
      SQL_RESERVED_TABLE_TOKENS.has(t.toUpperCase()),
    );
    expect(reserved).toEqual([]);
  });
});
