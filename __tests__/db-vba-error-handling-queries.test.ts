/**
 * Integration coverage for every SQL query published in
 * `docs/vba-error-handling.md` (issue #262).
 *
 * The fixture is indexed through the real `CodeGraph` pipeline. The tests then
 * read the generated SQLite database and execute the documentation verbatim,
 * so schema drift or stale examples fail here instead of reaching consumers.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { DatabaseConnection, getDatabasePath } from '../src/db';
import CodeGraph from '../src/index';

const DOC_PATH = path.resolve(__dirname, '../docs/vba-error-handling.md');

const VBA_FIXTURE = [
  'Attribute VB_Name = "ErrorQueryFixture"',
  'Option Explicit',
  '',
  'Public Sub RiskyWrite()',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    CurrentDb.Execute "UPDATE RiskTable SET Value = 1"',
  'End Sub',
  '',
  'Public Sub PaddedWrite()',
  '    value = 1',
  '',
  "    ' blank and comment-only lines must not raise the statement count",
  '    value = value + 1',
  '',
  '    value = value + 1',
  "    ' still only five executable statements",
  '    value = value + 1',
  '    CurrentDb.Execute "UPDATE PaddedTable SET Value = 1"',
  'End Sub',
  '',
  'Public Sub FilesystemOnly()',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    FileCopy "source.txt", "target.txt"',
  'End Sub',
  '',
  'Public Sub KillRisk()',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    Kill "obsolete.txt"',
  'End Sub',
  '',
  'Public Sub OpenRisk()',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    Open "output.txt" For Output As #1',
  'End Sub',
  '',
  'Public Sub CloseRisk()',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    Close #1',
  'End Sub',
  '',
  'Public Sub ShortKill()',
  '    Kill "short.txt"',
  'End Sub',
  '',
  'Public Sub MalformedOpen()',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    Open "missing-mode.txt"',
  'End Sub',
  '',
  'Public Sub UserNamedKill(ByVal pathName As String)',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    Debug.Print pathName',
  'End Sub',
  '',
  'Public Sub CallsUserNamedKill()',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    Call UserNamedKill("not-runtime.txt")',
  'End Sub',
  '',
  'Public Sub DoCmdRisk()',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    DoCmd.OpenForm "RiskForm"',
  'End Sub',
  '',
  'Public Sub ShortDoCmd()',
  '    DoCmd.OpenForm "ShortForm"',
  'End Sub',
  '',
  'Public Sub UnrelatedCall()',
  '    value = 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    value = value + 1',
  '    AuditSomething "not risky"',
  'End Sub',
  '',
  'Public Sub OpenSuppression()',
  '    On Error Resume Next',
  '    RiskyWrite',
  'End Sub',
  '',
  'Public Sub ReachDisplay()',
  '    ShowFailure',
  'End Sub',
  '',
  'Public Sub ShowFailure()',
  '    On Error GoTo failed',
  '    Exit Sub',
  'failed:',
  '    MsgBox Err.Description',
  'End Sub',
  '',
  'Public Sub BrokenTarget()',
  '    On Error GoTo missingHandler',
  '    Exit Sub',
  'End Sub',
  '',
].join('\n');

type QueryRow = Record<string, string | number | null>;

function publishedQueries(): string[] {
  const markdown = fs.readFileSync(DOC_PATH, 'utf8');
  return [...markdown.matchAll(/```sql\s*\r?\n([\s\S]*?)```/g)].map((match) => match[1]!.trim());
}

describe('published VBA error-handling SQL', () => {
  let projectDir: string;
  let db: DatabaseConnection;
  let queries: string[];

  beforeAll(async () => {
    projectDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-error-queries-'));
    fs.writeFileSync(path.join(projectDir, 'ErrorQueryFixture.bas'), VBA_FIXTURE);

    const graph = CodeGraph.initSync(projectDir);
    await graph.indexAll();
    graph.close();

    db = DatabaseConnection.open(getDatabasePath(projectDir));
    queries = publishedQueries();
  }, 60_000);

  afterAll(() => {
    db?.close();
    try {
      fs.rmSync(projectDir, { recursive: true, force: true });
    } catch {
      // Windows may release SQLite sidecar handles shortly after close.
    }
  });

  it('publishes exactly the four promised runnable queries', () => {
    expect(queries).toHaveLength(4);
  });

  it('finds exactly-six-statement SQL and filesystem risks without padding or unrelated calls', () => {
    const rows = db.getDb().prepare(queries[0]!).all() as QueryRow[];
    expect(rows.map((row) => row.procedure).sort()).toEqual([
      'CloseRisk',
      'DoCmdRisk',
      'FilesystemOnly',
      'KillRisk',
      'OpenRisk',
      'RiskyWrite',
    ]);
    expect(rows).toEqual(expect.arrayContaining([
      expect.objectContaining({
        procedure: 'RiskyWrite',
        executable_statement_count: 6,
        risk_target: 'RiskTable',
        risk_target_kind: 'class',
        detected_by: 'vba-sql-table',
      }),
      expect.objectContaining({
        procedure: 'FilesystemOnly',
        executable_statement_count: 6,
        risk_target: 'FileCopy',
        risk_target_kind: 'runtime-call',
        detected_by: 'vba-statement-call-unresolved',
      }),
          expect.objectContaining({
            procedure: 'DoCmdRisk',
            executable_statement_count: 6,
            risk_target: 'RiskForm',
            risk_target_kind: 'form-layout',
            detected_by: 'vba-opens-form',
          }),
          ...(['Kill', 'Open', 'Close'] as const).map((riskTarget) =>
            expect.objectContaining({
              procedure: `${riskTarget}Risk`,
              executable_statement_count: 6,
              risk_target: riskTarget,
              risk_target_kind: 'runtime-call',
              detected_by: 'vba-filesystem-statement',
            }),
          ),
    ]));
    expect(rows.some((row) => row.procedure === 'PaddedWrite')).toBe(false);
    expect(rows.some((row) => row.procedure === 'ShortDoCmd')).toBe(false);
    expect(rows.some((row) => row.procedure === 'UnrelatedCall')).toBe(false);
    expect(rows.some((row) => row.procedure === 'ShortKill')).toBe(false);
    expect(rows.some((row) => row.procedure === 'MalformedOpen')).toBe(false);
    expect(rows.some((row) => row.procedure === 'UserNamedKill')).toBe(false);
    expect(rows.some((row) => row.procedure === 'CallsUserNamedKill')).toBe(false);
  });

  it('finds Resume Next scopes that remain open', () => {
    const rows = db.getDb().prepare(queries[1]!).all() as QueryRow[];
    expect(rows.map((row) => row.procedure)).toEqual(['OpenSuppression']);
    expect(rows[0]).toMatchObject({ resume_next_open: 1 });
  });

  it('finds display handlers and the procedures that reach them', () => {
    const rows = db.getDb().prepare(queries[2]!).all() as QueryRow[];
    expect(rows).toEqual([
      expect.objectContaining({
        display_procedure: 'ShowFailure',
        behavior: 'display',
        reaches_display: 'ReachDisplay',
      }),
    ]);
  });

  it('finds handler targets that are missing from their procedure', () => {
    const rows = db.getDb().prepare(queries[3]!).all() as QueryRow[];
    expect(rows.map((row) => row.procedure)).toEqual(['BrokenTarget']);
    expect(rows[0]).toMatchObject({ dangling_target: 'missingHandler' });
  });
});
