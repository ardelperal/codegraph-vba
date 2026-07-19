/**
 * Tests for the human-readable `codegraph status` output (issue #193).
 *
 * Issue #189 added the authoritative JSON shape (`reindexReasons`) and the
 * `reindexRecommended` boolean, but the human-readable prose branch still
 * printed `Index is up to date` whenever there were no pending source-file
 * changes — even when the on-disk index was stamped by an older engine and a
 * re-index was recommended. That presented a stale index as healthy.
 *
 * Required behavior:
 *  - Fresh index + no source changes => "No source changes detected" (and the
 *    old "Index is up to date" phrase is gone).
 *  - `reindexReasons` non-empty => a one-line re-index recommendation, and
 *    NEVER the clean-health sentence.
 *  - JSON output remains identical (the JSON contract is owned by #189).
 *
 * Exercised end-to-end against the built binary so the wiring (not just the
 * library) is covered, matching the approach of status-json.test.ts.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { execFileSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph } from '../src';
import { codeGraphDirName } from '../src/directory';

const BIN = path.resolve(__dirname, '../dist/bin/codegraph.js');

function runStatusHuman(cwd: string): string {
  return execFileSync(process.execPath, [BIN, 'status'], {
    cwd,
    encoding: 'utf-8',
    env: { ...process.env, CODEGRAPH_NO_DAEMON: '1', FORCE_COLOR: '0' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

function runStatusJson(cwd: string): Record<string, unknown> {
  const stdout = execFileSync(process.execPath, [BIN, 'status', '--json'], {
    cwd,
    encoding: 'utf-8',
    env: { ...process.env, CODEGRAPH_NO_DAEMON: '1' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  const line = stdout.trim().split('\n').filter(Boolean).pop()!;
  return JSON.parse(line);
}

function runCodegraph(args: string[], cwd: string): string {
  return execFileSync(process.execPath, [BIN, ...args], {
    cwd,
    encoding: 'utf-8',
    env: { ...process.env, CODEGRAPH_NO_DAEMON: '1' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

describe('codegraph status — human-readable prose (#193)', () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-status-human-'));
  });
  afterEach(() => {
    fs.rmSync(tempDir, { recursive: true, force: true });
  });

  it('a fresh full index with no source changes prints "No source changes detected" and never the old "Index is up to date" phrase', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    const cg = CodeGraph.initSync(tempDir);
    await cg.indexAll();
    cg.close();

    const out = runStatusHuman(tempDir);
    expect(out).toContain('No source changes detected');
    expect(out).not.toContain('Index is up to date');
  });

  it('an index stamped with the previous extraction-version constant never prints the clean-health sentence; the prose flags the re-index instead', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    const cg = CodeGraph.initSync(tempDir);
    await cg.indexAll();
    cg.close();

    // Simulate an index built by the previous engine (constant 24). The bumped
    // binary (constant 25) must surface the mismatch in the human-readable
    // prose as a re-index recommendation, NOT as a healthy "No source changes
    // detected" line.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { DatabaseSync } = require('node:sqlite');
    const db = new DatabaseSync(path.join(tempDir, codeGraphDirName(), 'codegraph.db'));
    db.prepare(
      "INSERT INTO project_metadata (key, value, updated_at) VALUES ('indexed_with_extraction_version', '24', 0) " +
        "ON CONFLICT(key) DO UPDATE SET value = '24'"
    ).run();
    db.close();

    const out = runStatusHuman(tempDir);
    expect(out).not.toContain('Index is up to date');
    expect(out).not.toContain('No source changes detected');
    expect(out).toMatch(/Re-index recommended:.*extraction-version/);
  });

  it('JSON reindexReasons remains identical after the human-prose change (#193 must not touch the JSON contract)', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    const cg = CodeGraph.initSync(tempDir);
    await cg.indexAll();
    cg.close();

    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { DatabaseSync } = require('node:sqlite');
    const db = new DatabaseSync(path.join(tempDir, codeGraphDirName(), 'codegraph.db'));
    db.prepare(
      "INSERT INTO project_metadata (key, value, updated_at) VALUES ('indexed_with_extraction_version', '24', 0) " +
        "ON CONFLICT(key) DO UPDATE SET value = '24'"
    ).run();
    db.close();

    const json = runStatusJson(tempDir);
    const index = json.index as Record<string, unknown>;
    expect(index.reindexRecommended).toBe(true);
    expect(Array.isArray(index.reindexReasons)).toBe(true);
    expect((index.reindexReasons as string[])).toContain('extraction-version');
    expect(index.builtWithExtractionVersion).toBe(24);
    expect(index.currentExtractionVersion).toBe(25);
  });

  it('a fresh full index keeps the JSON reindexReasons empty and reindexRecommended=false (#193 must not touch the JSON contract)', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    const cg = CodeGraph.initSync(tempDir);
    await cg.indexAll();
    cg.close();

    const json = runStatusJson(tempDir);
    const index = json.index as Record<string, unknown>;
    expect(index.reindexRecommended).toBe(false);
    expect(index.reindexReasons).toEqual([]);
  });

  it('with pending source changes, the human output keeps the existing Pending Changes section (no healthy-sentence regression)', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    runCodegraph(['init'], tempDir);

    // Add a new file after init to trigger "Pending Changes".
    fs.writeFileSync(path.join(tempDir, 'b.ts'), 'export const y = 2;\n');

    const out = runStatusHuman(tempDir);
    expect(out).toContain('Pending Changes:');
    // Healthy prose must not appear when source changes are pending either.
    expect(out).not.toContain('Index is up to date');
    expect(out).not.toContain('No source changes detected');
  });
});