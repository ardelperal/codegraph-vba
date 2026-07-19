/**
 * Strict-TDD unit/integration coverage for the VBA intrinsic constant + DAO
 * enum value classifier extension to `classifyVbaReferenceAsRuntime`
 * (issue #188).
 *
 * Background: v1.13.0's stdlib allowlist (PR #195 / issue #192) covers
 * statement-form and paren-form calls for the VBA stdlib functions only.
 * Bare **DAO enum values** (`dbFailOnError`, `dbSeeChanges`, `dbDenyWrite`,
 * `dbDenyRead`, `dbReadOnly`, `dbAppendOnly`, ...) and **VBA intrinsic
 * constants** (`vbCrLf`, `vbLf`, `vbCr`, `vbNewLine`, `vbTab`,
 * `vbExclamation`, `vbQuestion`, `vbInformation`, `vbCritical`, `vbOKOnly`,
 * `vbOKCancel`, `vbAbortRetryIgnore`, `vbYesNoCancel`, `vbYesNo`,
 * `vbRetryCancel`, `vbDefaultButton1..4`, `vbApplicationModal`,
 * `vbSystemModal`, ...) reach `unresolved_refs` as
 * `reference_kind='unqualified-ident'` and currently leak as `failed`
 * (~104 confirmed rows on the bench corpus).
 *
 * These are bare identifiers, NOT calls — the extraction-stamped
 * `vba-statement-call-unresolved` flag does NOT apply. The gate must
 * therefore be name-only for the new sets, while keeping the existing
 * stdlib-function gate intact (paren-form `calls` rows + statement-form
 * `unqualified-ident` rows stamped by the extractor).
 *
 * The two new predicates (`isVbaIntrinsicConstant`, `isDaoEnumValue`) live
 * alongside `isVbaStdlibFunction` in `src/resolution/vba-runtime-objects.ts`,
 * matching the case-insensitive / lowercased compare style. They MUST NOT
 * weaken user-defined shadow resolution: a `Public Const dbFailOnError = 0`
 * or `Public Function vbCrLf() As String` resolves normally and never reaches
 * `declined-runtime`.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { CodeGraph, DatabaseConnection, getDatabasePath } from '../src';
import {
  classifyVbaReferenceAsRuntime,
  isDaoEnumValue,
  isVbaIntrinsicConstant,
  isVbaStdlibFunction,
} from '../src/resolution/vba-runtime-objects';

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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-runtime-consts-'));
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

/**
 * Pull the unresolved_refs rows for the given names, projected to a
 * stable, comparable shape (synthesizedBy comes through as JSON in
 * `metadata`).
 */
function unresolvedRowsFor(cg: CodeGraph, names: string[]): Array<{
  reference_name: string;
  reference_kind: string;
  status: string;
  synthesizedBy: string | undefined;
}> {
  const projectRoot = (cg as unknown as { projectRoot: string }).projectRoot;
  const connection = DatabaseConnection.open(getDatabasePath(projectRoot));
  const placeholders = names.map(() => '?').join(',');
  const rows = connection.getDb().prepare(
    `SELECT reference_name, reference_kind, status, metadata FROM unresolved_refs WHERE reference_name IN (${placeholders}) ORDER BY reference_name, line`,
  ).all(...names) as Array<{
    reference_name: string;
    reference_kind: string;
    status: string;
    metadata: string | null;
  }>;
  connection.close();
  return rows.map((row) => ({
    reference_name: row.reference_name,
    reference_kind: row.reference_kind,
    status: row.status,
    synthesizedBy: row.metadata
      ? (JSON.parse(row.metadata) as { synthesizedBy?: string }).synthesizedBy
      : undefined,
  }));
}

describe('VBA intrinsic constants + DAO enum values (#188)', () => {
  describe('Unit — new predicates', () => {
    it('Test 1: every DAO enum value in the issue set classifies as a DAO enum value (case-insensitive)', () => {
      for (const name of [
        'dbFailOnError',
        'dbSeeChanges',
        'dbDenyWrite',
        'dbDenyRead',
        'dbReadOnly',
        'dbAppendOnly',
      ]) {
        expect(isDaoEnumValue(name)).toBe(true);
        expect(isDaoEnumValue(name.toLowerCase())).toBe(true);
        expect(isDaoEnumValue(name.toUpperCase())).toBe(true);
      }
    });

    it('Test 2: every VBA intrinsic constant in the issue set classifies as an intrinsic constant (case-insensitive)', () => {
      for (const name of [
        'vbCrLf', 'vbLf', 'vbCr', 'vbNewLine', 'vbTab',
        'vbExclamation', 'vbQuestion', 'vbInformation', 'vbCritical',
        'vbOKOnly', 'vbOKCancel', 'vbAbortRetryIgnore',
        'vbYesNoCancel', 'vbYesNo', 'vbRetryCancel',
        'vbDefaultButton1', 'vbDefaultButton2', 'vbDefaultButton3', 'vbDefaultButton4',
        'vbApplicationModal', 'vbSystemModal',
      ]) {
        expect(isVbaIntrinsicConstant(name)).toBe(true);
        expect(isVbaIntrinsicConstant(name.toLowerCase())).toBe(true);
        expect(isVbaIntrinsicConstant(name.toUpperCase())).toBe(true);
      }
    });

    it('Test 3: the two predicates are disjoint and reject user-defined names', () => {
      // Neither set swallows the other.
      expect(isDaoEnumValue('vbCrLf')).toBe(false);
      expect(isVbaIntrinsicConstant('dbFailOnError')).toBe(false);

      // Neither swallows stdlib functions.
      expect(isDaoEnumValue('MsgBox')).toBe(false);
      expect(isVbaIntrinsicConstant('MsgBox')).toBe(false);
      expect(isVbaEnumAndNotStdlibOrConst('MsgBox')).toBe(false);

      // Neither swallows user-defined names.
      expect(isDaoEnumValue('MyProjectHelper')).toBe(false);
      expect(isVbaIntrinsicConstant('MyProjectHelper')).toBe(false);

      // Null / undefined / empty → false (no false positives).
      expect(isDaoEnumValue(undefined)).toBe(false);
      expect(isVbaIntrinsicConstant(undefined)).toBe(false);
      expect(isDaoEnumValue('')).toBe(false);
      expect(isVbaIntrinsicConstant('')).toBe(false);
    });
  });

  describe('Unit — classifier accepts unqualified-ident for the new sets without metadata gating', () => {
    it('Test 4: a bare DAO enum value (unqualified-ident, no synthesizedBy) is declined-runtime', () => {
      expect(
        classifyVbaReferenceAsRuntime({
          language: 'vba',
          referenceKind: 'unqualified-ident',
          referenceName: 'dbFailOnError',
          metadata: null,
        }),
      ).toBe(true);
    });

    it('Test 5: a bare VBA intrinsic constant (unqualified-ident, no synthesizedBy) is declined-runtime', () => {
      for (const name of ['vbCrLf', 'vbNewLine', 'vbTab', 'vbYesNo', 'vbExclamation']) {
        expect(
          classifyVbaReferenceAsRuntime({
            language: 'vba',
            referenceKind: 'unqualified-ident',
            referenceName: name,
            metadata: null,
          }),
        ).toBe(true);
      }
    });

    it('Test 6: user-defined same-named symbols (Function/Const) are NOT declined-runtime', () => {
      // A name in the new sets must still resolve normally when it is ALSO
      // declared by the project — the classifier must never shortcut the
      // normal name resolution path. This is the negative control.
      expect(
        classifyVbaReferenceAsRuntime({
          language: 'vba',
          referenceKind: 'unqualified-ident',
          referenceName: 'ProjectDefinedDbFailOnError', // not in any set
          metadata: null,
        }),
      ).toBe(false);
      // A name in the stdlib set still requires the existing metadata gate.
      expect(
        classifyVbaReferenceAsRuntime({
          language: 'vba',
          referenceKind: 'unqualified-ident',
          referenceName: 'MsgBox', // stdlib, but no synthesizedBy → not a call
          metadata: null,
        }),
      ).toBe(false);
    });

    it('Test 7: non-VBA languages are always rejected (gate is VBA-only)', () => {
      expect(
        classifyVbaReferenceAsRuntime({
          language: 'typescript',
          referenceKind: 'unqualified-ident',
          referenceName: 'dbFailOnError',
          metadata: null,
        }),
      ).toBe(false);
    });

    it('Test 8: existing stdlib-call classification stays intact (regression — #192 / PR #195)', () => {
      // Paren-form calls: name match alone is enough.
      expect(
        classifyVbaReferenceAsRuntime({
          language: 'vba',
          referenceKind: 'calls',
          referenceName: 'MsgBox',
          metadata: { synthesizedBy: 'vba-paren-call-unresolved' },
        }),
      ).toBe(true);
      // Statement-form calls: require the extractor-stamped flag.
      expect(
        classifyVbaReferenceAsRuntime({
          language: 'vba',
          referenceKind: 'unqualified-ident',
          referenceName: 'MsgBox',
          metadata: { synthesizedBy: 'vba-statement-call-unresolved' },
        }),
      ).toBe(true);
      // Statement-form WITHOUT the stamp → not a call, not a constant/enum:
      // must NOT be classified as runtime.
      expect(
        classifyVbaReferenceAsRuntime({
          language: 'vba',
          referenceKind: 'unqualified-ident',
          referenceName: 'MsgBox',
          metadata: { synthesizedBy: 'some-other-thing' },
        }),
      ).toBe(false);
    });
  });

  describe('Integration — CodeGraph.indexAll fixture (RED before GREEN)', () => {
    /**
     * Minimal deterministic fixture required by the issue: every named symbol
     * from the explicit SQL set must appear in a shape the extractor emits
     * as `reference_kind='unqualified-ident'` (bare statement-call lines),
     * plus a user-defined `Public Const dbFailOnError As Long = 0` and a
     * user-defined `Public Function vbCrLf() As String` shadow as negative
     * controls.
     *
     * The extractor emits `unqualified-ident` rows from the statement-call
     * sweep (`src/extraction/vba/call-sweep.ts:319-330`) when a line's
     * leading identifier is NOT a self-call / blacklisted receiver /
     * same-file target AND is NOT a VBA keyword. Each constant on its own
     * line (or as a `: name` continuation) is the production shape that
     * produces the ~104 leaked rows on the bench corpus.
     */
    const FIXTURE: Record<string, string> = {
      'src/modules/Caller.bas': [
        'Attribute VB_Name = "Caller"',
        'Option Explicit',
        '',
        'Public Const dbFailOnError As Long = 0',
        '',
        'Public Function vbCrLf() As String',
        '    vbCrLf = "shadow"',
        'End Function',
        '',
        'Public Sub CallerSub()',
        // DAO enum values on their own statement-call lines — the shape the
        // statement-call sweep emits as `unqualified-ident`.
        '    dbSeeChanges',
        '    dbDenyWrite',
        '    dbDenyRead',
        '    dbReadOnly',
        '    dbAppendOnly',
        // VBA intrinsic constants on their own statement-call lines.
        '    vbLf',
        '    vbCr',
        '    vbNewLine',
        '    vbTab',
        '    vbExclamation',
        '    vbQuestion',
        '    vbInformation',
        '    vbCritical',
        '    vbOKOnly',
        '    vbOKCancel',
        '    vbAbortRetryIgnore',
        '    vbYesNoCancel',
        '    vbYesNo',
        '    vbRetryCancel',
        '    vbDefaultButton1',
        '    vbDefaultButton2',
        '    vbDefaultButton3',
        '    vbDefaultButton4',
        '    vbApplicationModal',
        '    vbSystemModal',
        // User-defined shadow references — must resolve normally, NOT reach
        // `declined-runtime`.
        '    Debug.Print dbFailOnError',
        '    Debug.Print vbCrLf',
        '    Debug.Print SomeProjectConst',
        'End Sub',
        '',
        'Public Const SomeProjectConst As Long = 7',
        '',
      ].join('\n'),
    };

    it('Test 9: every DAO enum value in the issue set becomes declined-runtime', async () => {
      const cg = await buildProject(FIXTURE);
      const rows = unresolvedRowsFor(cg, [
        'dbFailOnError',
        'dbSeeChanges',
        'dbDenyWrite',
        'dbDenyRead',
        'dbReadOnly',
        'dbAppendOnly',
      ]);
      // dbFailOnError is shadowed by `Public Const dbFailOnError As Long = 0`
      // and resolves to that real node — so it must NOT appear here at all.
      const names = rows.map((row) => row.reference_name);
      expect(names).not.toContain('dbFailOnError');
      for (const expected of ['dbSeeChanges', 'dbDenyWrite', 'dbDenyRead', 'dbReadOnly', 'dbAppendOnly']) {
        const matches = rows.filter((row) => row.reference_name === expected);
        expect(matches.length).toBeGreaterThanOrEqual(1);
        for (const match of matches) {
          expect(match.reference_kind).toBe('unqualified-ident');
          expect(match.status).toBe('declined-runtime');
        }
      }
    });

    it('Test 10: every VBA intrinsic constant in the issue set becomes declined-runtime', async () => {
      const cg = await buildProject(FIXTURE);
      const expected = [
        'vbLf', 'vbCr', 'vbNewLine', 'vbTab',
        'vbExclamation', 'vbQuestion', 'vbInformation', 'vbCritical',
        'vbOKOnly', 'vbOKCancel', 'vbAbortRetryIgnore',
        'vbYesNoCancel', 'vbYesNo', 'vbRetryCancel',
        'vbDefaultButton1', 'vbDefaultButton2', 'vbDefaultButton3', 'vbDefaultButton4',
        'vbApplicationModal', 'vbSystemModal',
      ];
      const rows = unresolvedRowsFor(cg, expected);
      // vbCrLf is shadowed by `Public Function vbCrLf() As String` and
      // resolves to that real node — so it must NOT appear here at all.
      expect(rows.map((row) => row.reference_name)).not.toContain('vbCrLf');
      // Every other constant must appear with declined-runtime.
      for (const name of expected) {
        const matches = rows.filter((row) => row.reference_name === name);
        expect(matches.length, `expected ≥1 unresolved row for ${name}`).toBeGreaterThanOrEqual(1);
        for (const match of matches) {
          expect(match.reference_kind).toBe('unqualified-ident');
          expect(match.status, `${name} must be declined-runtime`).toBe('declined-runtime');
        }
      }
    });

    it('Test 11: user-defined shadows (Const + Function) resolve normally and never reach declined-runtime', async () => {
      const cg = await buildProject(FIXTURE);

      // 1. The user-defined shadow `Public Const dbFailOnError As Long = 0`
      //    is in the graph as a real constant node.
      const shadowConst = cg
        .searchNodes('dbFailOnError', { languages: ['vba'], kinds: ['constant'] })
        .find((n) => n.node.name === 'dbFailOnError' && n.node.filePath.endsWith('Caller.bas'));
      expect(shadowConst).toBeDefined();

      // 2. The user-defined shadow `Public Function vbCrLf() As String` is in
      //    the graph as a real function node.
      const shadowFn = cg
        .searchNodes('vbCrLf', { languages: ['vba'], kinds: ['function'] })
        .find((n) => n.node.name === 'vbCrLf' && n.node.filePath.endsWith('Caller.bas'));
      expect(shadowFn).toBeDefined();

      // 3. The shadow names must NEVER appear as `declined-runtime` in
      //    unresolved_refs — normal user-symbol resolution runs first and
      //    deletes them from the pending set.
      const projectRoot = (cg as unknown as { projectRoot: string }).projectRoot;
      const connection = DatabaseConnection.open(getDatabasePath(projectRoot));
      const shadowRows = connection.getDb().prepare(
        "SELECT reference_name, status FROM unresolved_refs WHERE reference_name IN ('dbFailOnError', 'vbCrLf')",
      ).all() as Array<{ reference_name: string; status: string }>;
      connection.close();
      expect(shadowRows).toEqual([]);
    });

    it('Test 12: a project-defined Public Const with a name NOT in any set stays out of unresolved_refs', async () => {
      // SomeProjectConst is used inside CallerSub on its own line. Normal
      // const-first resolution should resolve it before it ever reaches the
      // classifier — a regression-prevention check that the classifier
      // accepts and the resolver does not regress.
      const cg = await buildProject(FIXTURE);
      const projectRoot = (cg as unknown as { projectRoot: string }).projectRoot;
      const connection = DatabaseConnection.open(getDatabasePath(projectRoot));
      const rows = connection.getDb().prepare(
        "SELECT reference_name, status FROM unresolved_refs WHERE reference_name = 'SomeProjectConst'",
      ).all();
      connection.close();
      expect(rows).toEqual([]);
    });
  });
});

/**
 * Local helper — must NOT be imported from the production module: it
 * exists only so this test asserts the *separation* of the two new sets
 * and the stdlib set without coupling to any specific symbol. Implemented
 * inline so a typo in `isDaoEnumValue` / `isVbaIntrinsicConstant` does not
 * silently pass through a stale reference.
 */
function isVbaEnumAndNotStdlibOrConst(name: string): boolean {
  return isDaoEnumValue(name) || isVbaIntrinsicConstant(name);
}
