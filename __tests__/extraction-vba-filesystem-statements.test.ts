import { afterEach, describe, expect, it } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { DatabaseConnection, getDatabasePath } from '../src/db';
import CodeGraph from '../src/index';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { isExactVbaFilesystemStatementReference } from '../src/extraction/vba/filesystem-statements';
import type { UnresolvedReference } from '../src/types';

const SYNTHESIZED_BY = 'vba-filesystem-statement';

function extract(body: string[]): ReturnType<VbaExtractor['extract']> {
  return new VbaExtractor(
    'src/modules/FilesystemFixture.bas',
    ['Attribute VB_Name = "FilesystemFixture"', 'Option Explicit', ...body].join('\n'),
  ).extract();
}

function filesystemRefs(refs: UnresolvedReference[]): UnresolvedReference[] {
  return refs.filter((ref) => ref.metadata?.synthesizedBy === SYNTHESIZED_BY);
}

describe('VBA filesystem statement extraction', () => {
  it('emits one typed unresolved call for each exact Kill, Open, and Close statement shape', () => {
    const result = extract([
      'Public Sub ExerciseFilesystem()',
      '    Kill "C:\\temp\\old.txt"',
      '    Open pathName For Append As #1',
      '    Open pathName For Binary Access Read Write Lock Read Write As fileNo Len = 128',
      '    Open pathName For Input As #inputFile',
      '    Open pathName For Output As outputFile',
      '    Open pathName For Random As #5 Len = recordLength',
      '    Close',
      '    Close #1, fileNo, #outputFile',
      'End Sub',
    ]);

    const refs = filesystemRefs(result.unresolvedReferences);
    expect(refs.map((ref) => ref.referenceName)).toEqual([
      'Kill', 'Open', 'Open', 'Open', 'Open', 'Open', 'Close', 'Close',
    ]);
    expect(refs.map((ref) => ref.referenceKind)).toEqual(Array(8).fill('calls'));
    expect(refs.map((ref) => ref.metadata)).toEqual([
      { synthesizedBy: SYNTHESIZED_BY, runtimeFamily: 'filesystem', operation: 'kill' },
      ...Array(5).fill(null).map(() => ({
        synthesizedBy: SYNTHESIZED_BY,
        runtimeFamily: 'filesystem',
        operation: 'open',
      })),
      ...Array(2).fill(null).map(() => ({
        synthesizedBy: SYNTHESIZED_BY,
        runtimeFamily: 'filesystem',
        operation: 'close',
      })),
    ]);
    expect(new Set(refs.map((ref) => ref.fromNodeId)).size).toBe(1);
  });

  it('rejects prefix collisions, qualification, call syntax, malformed grammar, strings, and comments', () => {
    const result = extract([
      'Public Sub NegativeShapes()',
      '    KillFile pathName',
      '    VBA.Kill pathName',
      '    Kill(pathName)',
      '    Call Kill(pathName)',
      '    OpenFile pathName',
      '    VBA.Open pathName For Input As #1',
      '    Open(pathName)',
      '    Open pathName',
      '    Open pathName For Read As #1',
      '    Open pathName For Input',
      '    Open pathName For Input As',
      '    CloseAll',
      '    obj.Close',
      '    Close()',
      '    Call Close()',
      '    Debug.Print "Kill x: Open y For Output As #1: Close #1"',
      "    ' Kill pathName",
      'End Sub',
    ]);

    expect(filesystemRefs(result.unresolvedReferences)).toEqual([]);
  });

  it('rejects inconsistent or forged provenance at the shared exact-reference gate', () => {
    const exact = {
      language: 'vba',
      referenceKind: 'calls',
      referenceName: 'Open',
      metadata: {
        synthesizedBy: SYNTHESIZED_BY,
        runtimeFamily: 'filesystem',
        operation: 'open',
      },
    };
    expect(isExactVbaFilesystemStatementReference(exact)).toBe(true);
    expect(isExactVbaFilesystemStatementReference({
      ...exact,
      referenceName: 'Kill',
    })).toBe(false);
    expect(isExactVbaFilesystemStatementReference({
      ...exact,
      metadata: { ...exact.metadata, runtimeFamily: 'database' },
    })).toBe(false);
    expect(isExactVbaFilesystemStatementReference({
      ...exact,
      metadata: { ...exact.metadata, synthesizedBy: 'vba-statement-call-unresolved' },
    })).toBe(false);
    expect(isExactVbaFilesystemStatementReference({
      ...exact,
      referenceKind: 'unqualified-ident',
    })).toBe(false);
  });

  it('does not classify declarations of user procedures named Kill, Open, or Close as statements', () => {
    const result = extract([
      'Public Sub Kill(ByVal pathName As String)',
      '    Debug.Print pathName',
      'End Sub',
      'Public Sub Open(ByVal pathName As String)',
      '    Debug.Print pathName',
      'End Sub',
      'Public Sub Close()',
      '    Debug.Print "user"',
      'End Sub',
    ]);

    expect(filesystemRefs(result.unresolvedReferences)).toEqual([]);
  });
});

describe('VBA filesystem statement resolution gate', () => {
  const tempDirs: string[] = [];

  afterEach(() => {
    for (const dir of tempDirs.splice(0)) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('keeps exact statement refs declined-runtime even when user procedures share the intrinsic names', async () => {
    const projectDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-fs-statements-'));
    tempDirs.push(projectDir);
    fs.writeFileSync(path.join(projectDir, 'Filesystem.bas'), [
      'Attribute VB_Name = "Filesystem"',
      'Option Explicit',
      'Public Sub Caller()',
      '    Kill "old.txt"',
      '    Open "out.txt" For Output As #1',
      '    Close #1',
      'End Sub',
      'Public Sub Kill(ByVal value As String)',
      'End Sub',
      'Public Sub Open(ByVal value As String)',
      'End Sub',
      'Public Sub Close()',
      'End Sub',
    ].join('\n'));

    const graph = CodeGraph.initSync(projectDir);
    await graph.indexAll();
    graph.close();

    const db = DatabaseConnection.open(getDatabasePath(projectDir));
    const rows = db.getDb().prepare(`
      SELECT reference_name, reference_kind, status,
             json_extract(metadata, '$.runtimeFamily') AS runtime_family,
             json_extract(metadata, '$.operation') AS operation
      FROM unresolved_refs
      WHERE json_extract(metadata, '$.synthesizedBy') = ?
      ORDER BY line
    `).all(SYNTHESIZED_BY) as Array<Record<string, string>>;
    db.close();

    expect(rows).toEqual([
      { reference_name: 'Kill', reference_kind: 'calls', status: 'declined-runtime', runtime_family: 'filesystem', operation: 'kill' },
      { reference_name: 'Open', reference_kind: 'calls', status: 'declined-runtime', runtime_family: 'filesystem', operation: 'open' },
      { reference_name: 'Close', reference_kind: 'calls', status: 'declined-runtime', runtime_family: 'filesystem', operation: 'close' },
    ]);
  });
});
