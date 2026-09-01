/**
 * extraction-vba-statement-call-kind.test.ts
 *
 * Issue #265: an unresolved statement-form Sub call used to land in
 * `unqualified-ident` — the bucket for ambiguous bare-identifier reads —
 * regardless of whether the syntax could actually have been a `Const` read.
 *
 * Only ONE of the three statement shapes is genuinely ambiguous:
 *
 *   | Form            | Can it be a `Const` read?                | Kind               |
 *   |-----------------|------------------------------------------|--------------------|
 *   | `Escribir`      | YES — a bare identifier read             | `unqualified-ident`|
 *   | `Call Escribir` | No — `Call` is only valid on a procedure | `calls`            |
 *   | `Escribir 1, 2` | No — a constant takes no argument list   | `calls`            |
 *
 * The bare form deliberately stays in `unqualified-ident`: the const-first
 * disambiguation rule of issue #108 (FR-3.1) depends on it.
 *
 * Regression guards carried along:
 *   - a bare identifier that resolves to a `Const` stays `unqualified-ident`
 *     (FR-3.1);
 *   - `MsgBox "x"` and other statement-form built-ins still end up
 *     `declined-runtime`, not an actionable `failed` user call (#192/#195).
 *
 * Pattern: real files, no mocking — `VbaExtractor` on inline sources for the
 * extractor-level assertions, and a real `CodeGraph.indexAll()` round-trip on
 * a temp directory for the resolver-status guard.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { UnresolvedReference } from '../src/types';
import { CodeGraph, DatabaseConnection, getDatabasePath } from '../src';
import { detectStatementCall } from '../src/extraction/vba/calls';
import { maskStringContent } from '../src/extraction/vba/text-utils';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

function findRefByName(
  r: { unresolvedReferences: UnresolvedReference[] },
  referenceName: string,
): UnresolvedReference | undefined {
  return r.unresolvedReferences.find((u) => u.referenceName === referenceName);
}

const openProjects: Array<{ cg: CodeGraph; dir: string }> = [];

afterEach(async () => {
  while (openProjects.length > 0) {
    const { cg, dir } = openProjects.pop()!;
    try {
      await cg.close();
    } catch {
      // ignore close errors
    }
    if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  }
});

async function buildProject(files: Record<string, string>): Promise<CodeGraph> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-stmt-call-kind-'));
  for (const [rel, src] of Object.entries(files)) {
    const full = path.join(dir, rel);
    fs.mkdirSync(path.dirname(full), { recursive: true });
    fs.writeFileSync(full, src);
  }
  const cg = await CodeGraph.init(dir, { index: false });
  openProjects.push({ cg, dir });
  await cg.indexAll();
  return cg;
}

function unresolvedRowsFor(cg: CodeGraph, names: string[]): Array<{
  reference_name: string;
  reference_kind: string;
  status: string;
}> {
  const projectRoot = (cg as unknown as { projectRoot: string }).projectRoot;
  const connection = DatabaseConnection.open(getDatabasePath(projectRoot));
  const placeholders = names.map(() => '?').join(',');
  const rows = connection.getDb().prepare(
    `SELECT reference_name, reference_kind, status FROM unresolved_refs WHERE reference_name IN (${placeholders}) ORDER BY reference_name, line`,
  ).all(...names) as Array<{
    reference_name: string;
    reference_kind: string;
    status: string;
  }>;
  connection.close();
  return rows;
}

describe('detectStatementCall unambiguity flag (issue #265)', () => {
  it('reports the Call keyword and an argument list as unambiguous, a bare ident as ambiguous', () => {
    expect(detectStatementCall('    Escribir')).toEqual({
      name: 'Escribir',
      unambiguous: false,
    });
    expect(detectStatementCall('    Call Escribir')).toEqual({
      name: 'Escribir',
      unambiguous: true,
    });
    expect(detectStatementCall('    Escribir 1, 2')).toEqual({
      name: 'Escribir',
      unambiguous: true,
    });
    expect(detectStatementCall('    Call Escribir 1, 2')).toEqual({
      name: 'Escribir',
      unambiguous: true,
    });
  });

  it('trailing whitespace alone does not make a bare identifier unambiguous', () => {
    expect(detectStatementCall('    Escribir   ')).toEqual({
      name: 'Escribir',
      unambiguous: false,
    });
  });

  it('an argument list made only of string literals survives the `_` mask', () => {
    // `maskStringContent(line)` (the space mask the column-sensitive
    // scanners use) would flatten this to `Escribir` + blanks. The
    // statement-call path masks with `_` precisely so it does not.
    expect(detectStatementCall(maskStringContent('    Escribir "texto"', '_'))).toEqual({
      name: 'Escribir',
      unambiguous: true,
    });
    expect(detectStatementCall(maskStringContent('    Escribir "texto"'))).toEqual({
      name: 'Escribir',
      unambiguous: false,
    });
  });

  it('still returns null for the shapes it never handled (paren form, assignment, declaration)', () => {
    expect(detectStatementCall('    Escribir(1)')).toBeNull();
    expect(detectStatementCall('    Escribir = 1')).toBeNull();
    expect(detectStatementCall('    Dim Escribir As Long')).toBeNull();
    expect(detectStatementCall("    ' Escribir 1, 2")).toBeNull();
  });
});

describe('unresolved statement-form calls: syntactically unambiguous shapes (issue #265)', () => {
  const SRC = `Attribute VB_Name = "modCaller"
Public Sub ConCall()
    Call Escribir
End Sub

Public Sub ConArgumentos()
    Anotar 1, 2
End Sub

Public Sub Desnudo()
    Registrar
End Sub
`;

  it('`Call Escribir` (unresolved) → referenceKind = "calls"', () => {
    const r = extract('src/modCaller.bas', SRC);
    const ref = findRefByName(r, 'Escribir');
    expect(ref, 'expected an unresolved ref for Escribir').toBeDefined();
    expect(ref?.referenceKind).toBe('calls');
  });

  it('`Anotar 1, 2` (unresolved, argument list) → referenceKind = "calls"', () => {
    const r = extract('src/modCaller.bas', SRC);
    const ref = findRefByName(r, 'Anotar');
    expect(ref, 'expected an unresolved ref for Anotar').toBeDefined();
    expect(ref?.referenceKind).toBe('calls');
  });

  it('`Registrar` (bare, no Call keyword, no arguments) stays "unqualified-ident"', () => {
    const r = extract('src/modCaller.bas', SRC);
    const ref = findRefByName(r, 'Registrar');
    expect(ref, 'expected an unresolved ref for Registrar').toBeDefined();
    expect(ref?.referenceKind).toBe('unqualified-ident');
  });

  it('moves rows between the two kinds without adding or dropping any', () => {
    const r = extract('src/modCaller.bas', SRC);
    const moving = r.unresolvedReferences.filter(
      (u) => u.referenceKind === 'calls' || u.referenceKind === 'unqualified-ident',
    );
    expect(moving.map((u) => u.referenceName).sort()).toEqual([
      'Anotar',
      'Escribir',
      'Registrar',
    ]);
    // Every row still carries the statement-call stamp — only the kind moved.
    for (const u of moving) {
      expect(u.metadata?.synthesizedBy).toBe('vba-statement-call-unresolved');
    }
  });
});

describe('resolved statement-form calls are unaffected (issue #265)', () => {
  it('`Call Escribir` where Escribir exists in-file still emits a resolved calls edge', () => {
    const src = `Attribute VB_Name = "modLocal"
Public Sub Escribir()
End Sub

Public Sub Caller()
    Call Escribir
End Sub
`;
    const r = extract('src/modLocal.bas', src);
    // No unresolved row at all — the call resolved same-file.
    expect(findRefByName(r, 'Escribir')).toBeUndefined();

    const target = r.nodes.find((n) => n.kind === 'function' && n.name === 'Escribir');
    const caller = r.nodes.find((n) => n.kind === 'function' && n.name === 'Caller');
    expect(target).toBeDefined();
    expect(caller).toBeDefined();
    const edge = r.edges.find(
      (e) => e.kind === 'calls' && e.source === caller!.id && e.target === target!.id,
    );
    expect(edge, 'expected a resolved calls edge Caller -> Escribir').toBeDefined();
  });
});

describe('FR-3.1 const-first disambiguation regression guard (issue #108)', () => {
  it('a bare identifier that names a module Const stays "unqualified-ident", never "calls"', () => {
    const src = `Attribute VB_Name = "modRiesgo"
Private Const HAY_ERROR As Boolean = False

Public Sub Foo()
    HAY_ERROR
End Sub
`;
    const r = extract('src/modRiesgo.bas', src);
    const ref = findRefByName(r, 'HAY_ERROR');
    expect(ref, 'expected the bare Const read to reach unresolved_refs').toBeDefined();
    expect(ref?.referenceKind).toBe('unqualified-ident');
  });

  it('the bare form is unaffected even when unambiguous calls sit on neighbouring lines', () => {
    const src = `Attribute VB_Name = "modMezcla"
Private Const HAY_ERROR As Boolean = False

Public Sub Foo()
    Call Escribir
    HAY_ERROR
    Anotar 1
End Sub
`;
    const r = extract('src/modMezcla.bas', src);
    expect(findRefByName(r, 'HAY_ERROR')?.referenceKind).toBe('unqualified-ident');
    expect(findRefByName(r, 'Escribir')?.referenceKind).toBe('calls');
    expect(findRefByName(r, 'Anotar')?.referenceKind).toBe('calls');
  });
});

describe('statement-form stdlib built-ins stay declined-runtime (#192/#195 regression guard)', () => {
  const FIXTURE: Record<string, string> = {
    'src/modBuiltins.bas': [
      'Attribute VB_Name = "modBuiltins"',
      'Option Explicit',
      '',
      'Public Sub Foo()',
      // Statement-form built-ins with arguments — now `calls`, and the
      // stdlib gate must still park them as declined-runtime.
      '    MsgBox "x"',
      '    Shell "calc"',
      '    Call MsgBox("y")',
      // Bare statement-form built-in — stays `unqualified-ident`.
      '    DoEvents',
      // A genuine missing user callee written with the Call keyword must
      // stay actionable.
      '    Call FaltaEsteProcedimiento',
      'End Sub',
      '',
    ].join('\n'),
  };

  it('MsgBox / Shell with arguments classify as runtime noise, not failed user calls', async () => {
    const cg = await buildProject(FIXTURE);
    const rows = unresolvedRowsFor(cg, ['MsgBox', 'Shell', 'DoEvents']);
    for (const name of ['MsgBox', 'Shell', 'DoEvents']) {
      const matches = rows.filter((row) => row.reference_name === name);
      expect(matches.length, `expected >=1 unresolved row for ${name}`).toBeGreaterThanOrEqual(1);
      for (const match of matches) {
        expect(match.status, `${name} must be declined-runtime`).toBe('declined-runtime');
      }
    }
    // Shape check: the argument-bearing built-ins now carry `calls` (their
    // string-literal argument list is what makes them unambiguous calls),
    // while the bare `DoEvents` keeps `unqualified-ident`.
    expect(
      rows.filter((row) => row.reference_name === 'DoEvents').map((row) => row.reference_kind),
    ).toEqual(['unqualified-ident']);
    for (const name of ['MsgBox', 'Shell']) {
      expect(
        rows.filter((row) => row.reference_name === name).every((row) => row.reference_kind === 'calls'),
        `${name} argument-bearing rows must carry the calls kind`,
      ).toBe(true);
    }
  });

  it('a genuine missing callee written `Call Foo` stays actionable as a failed call', async () => {
    const cg = await buildProject(FIXTURE);
    const rows = unresolvedRowsFor(cg, ['FaltaEsteProcedimiento']);
    expect(rows.length).toBeGreaterThanOrEqual(1);
    for (const row of rows) {
      expect(row.reference_kind).toBe('calls');
      expect(row.status).toBe('failed');
    }
  });
});
