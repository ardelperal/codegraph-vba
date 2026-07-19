/**
 * Tests for the CI/scripting fields `codegraph status --json` exposes (issue
 * #329): the `version`, `indexPath`, and `lastIndexed` fields, plus the
 * matching `CodeGraph.getLastIndexedAt()` library method.
 *
 * The CLI itself is exercised end-to-end against the built binary so the JSON
 * field names survive future refactors of the underlying plumbing.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { execFileSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { CodeGraph } from '../src';
import { codeGraphDirName } from '../src/directory';

const BIN = path.resolve(__dirname, '../dist/bin/codegraph.js');
const PKG_VERSION = JSON.parse(
  fs.readFileSync(path.resolve(__dirname, '../package.json'), 'utf-8'),
).version as string;

function runStatusJson(cwd: string): Record<string, unknown> {
  const stdout = execFileSync(process.execPath, [BIN, 'status', '--json'], {
    cwd,
    encoding: 'utf-8',
    env: { ...process.env, CODEGRAPH_NO_DAEMON: '1' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  // JSON mode prints exactly one line to stdout; be defensive about any stray
  // leading output by parsing the last non-empty line.
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

describe('codegraph status --json — CI fields (#329)', () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-status-json-'));
  });
  afterEach(() => {
    fs.rmSync(tempDir, { recursive: true, force: true });
  });

  it('getLastIndexedAt() is null before indexing and a recent ms timestamp after', async () => {
    const cg = CodeGraph.initSync(tempDir);
    expect(cg.getLastIndexedAt()).toBeNull();

    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    const before = Date.now();
    await cg.indexAll();
    const after = Date.now();

    const last = cg.getLastIndexedAt();
    expect(last).not.toBeNull();
    expect(typeof last).toBe('number');
    expect(last!).toBeGreaterThanOrEqual(before - 1000);
    expect(last!).toBeLessThanOrEqual(after + 1000);
    cg.close();
  });

  it('status --json on an UNINITIALIZED project reports version + indexPath + lastIndexed:null', () => {
    const out = runStatusJson(tempDir);
    expect(out.initialized).toBe(false);
    expect(out.version).toBe(PKG_VERSION);
    expect(typeof out.indexPath).toBe('string');
    expect(out.indexPath as string).toContain('.codegraph');
    expect(out.lastIndexed).toBeNull();
  });

  it('status --json on an INDEXED project reports version + indexPath + a round-trippable lastIndexed', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    const before = Date.now();
    const cg = CodeGraph.initSync(tempDir);
    await cg.indexAll();
    const after = Date.now();
    cg.close();

    const out = runStatusJson(tempDir);
    expect(out.initialized).toBe(true);
    expect(out.version).toBe(PKG_VERSION);
    expect(out.indexPath as string).toContain('.codegraph');
    expect(typeof out.lastIndexed).toBe('string');
    // ISO string that round-trips back into the index window.
    const ms = Date.parse(out.lastIndexed as string);
    expect(ms).toBeGreaterThanOrEqual(before - 1000);
    expect(ms).toBeLessThanOrEqual(after + 1000);
  });
});

describe('index completeness marker (index_state)', () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-index-state-'));
  });

  afterEach(() => {
    fs.rmSync(tempDir, { recursive: true, force: true });
  });

  it('a clean full index stamps state=complete with reconciled counts', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export function f(): number { return 1; }\n');
    fs.writeFileSync(path.join(tempDir, 'b.ts'), 'import { f } from "./a";\nexport const y = f();\n');
    const cg = CodeGraph.initSync(tempDir);
    const result = await cg.indexAll();

    // The scan's ground truth is reported and fully accounted for.
    expect(result.filesDiscovered).toBeDefined();
    expect(result.filesIndexed + result.filesSkipped + result.filesErrored).toBe(
      result.filesDiscovered
    );
    expect(result.errors.filter((e) => e.code === 'index_partial')).toHaveLength(0);
    expect(cg.getIndexState()).toBe('complete');
    cg.close();

    const out = runStatusJson(tempDir);
    expect((out.index as Record<string, unknown>).state).toBe('complete');
  });

  it('a run killed mid-index leaves state=indexing, and status --json surfaces it', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    const cg = CodeGraph.initSync(tempDir);
    await cg.indexAll();
    cg.close();

    // Simulate a kill between the start-marker write and completion: the
    // marker a dead process leaves behind is exactly 'indexing'. Written
    // straight into the DB — the process that died can't have cleaned it up.
    // (require, not import: vite tries to bundle a dynamic import specifier.)
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { DatabaseSync } = require('node:sqlite');
    const db = new DatabaseSync(path.join(tempDir, codeGraphDirName(), 'codegraph.db'));
    db.prepare(
      "INSERT INTO project_metadata (key, value, updated_at) VALUES ('index_state', 'indexing', 0) " +
        "ON CONFLICT(key) DO UPDATE SET value = 'indexing'"
    ).run();
    db.close();

    const out = runStatusJson(tempDir);
    expect((out.index as Record<string, unknown>).state).toBe('indexing');

    const reopened = await CodeGraph.open(tempDir);
    expect(reopened.getIndexState()).toBe('indexing');
    reopened.close();
  });
});

// ---------------------------------------------------------------------------
// reindex reasons — runtime signal so CI can detect a missed extraction-version
// bump without parsing the prose warning (issue #189).
// ---------------------------------------------------------------------------

describe('reindexReasons in codegraph status --json (#189)', () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-reindex-reasons-'));
  });
  afterEach(() => {
    fs.rmSync(tempDir, { recursive: true, force: true });
  });

  it('reports an empty reindexReasons array for a fresh full index', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    const cg = CodeGraph.initSync(tempDir);
    await cg.indexAll();
    cg.close();

    const out = runStatusJson(tempDir);
    const index = out.index as Record<string, unknown>;
    expect(index.reindexRecommended).toBe(false);
    expect(Array.isArray(index.reindexReasons)).toBe(true);
    expect(index.reindexReasons as unknown[]).toEqual([]);
  });

  it('reports currentExtractionVersion=25, reindexRecommended=true, and reindexReasons=[extraction-version] for an index stamped with the previous extraction-version constant (24)', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    const cg = CodeGraph.initSync(tempDir);
    await cg.indexAll();
    cg.close();

    // Simulate an index built by the previous engine (constant 24). The bumped
    // binary (constant 25) must surface the mismatch as `extraction-version` in
    // the reindexReasons array, alongside the boolean reindexRecommended=true.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { DatabaseSync } = require('node:sqlite');
    const db = new DatabaseSync(path.join(tempDir, codeGraphDirName(), 'codegraph.db'));
    db.prepare(
      "INSERT INTO project_metadata (key, value, updated_at) VALUES ('indexed_with_extraction_version', '24', 0) " +
        "ON CONFLICT(key) DO UPDATE SET value = '24'"
    ).run();
    db.close();

    const out = runStatusJson(tempDir);
    const index = out.index as Record<string, unknown>;
    expect(index.currentExtractionVersion).toBe(25);
    expect(index.builtWithExtractionVersion).toBe(24);
    expect(index.reindexRecommended).toBe(true);
    expect(Array.isArray(index.reindexReasons)).toBe(true);
    expect((index.reindexReasons as string[])).toContain('extraction-version');
  });
});

// ---------------------------------------------------------------------------
// Smoke probe — a fresh full index through the built CLI stamps the current
// EXTRACTION_VERSION. Catches a future missed bump: if someone changes the
// constant but the test was run against a stale binary, this fails. Runs
// against `dist/bin/codegraph.js` so the wiring (not just the library) is
// covered (issue #189).
// ---------------------------------------------------------------------------

describe('extraction-version bump smoke probe (#189)', () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-ext-ver-probe-'));
  });
  afterEach(() => {
    fs.rmSync(tempDir, { recursive: true, force: true });
  });

  it('a fresh full index through the built CLI stamps the current EXTRACTION_VERSION on disk', async () => {
    fs.writeFileSync(path.join(tempDir, 'a.ts'), 'export const x = 1;\n');
    fs.writeFileSync(path.join(tempDir, 'b.ts'), 'import { x } from "./a";\nexport const y = x;\n');
    runCodegraph(['init'], tempDir);

    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { DatabaseSync } = require('node:sqlite');
    const db = new DatabaseSync(path.join(tempDir, codeGraphDirName(), 'codegraph.db'));
    const row = db
      .prepare("SELECT value FROM project_metadata WHERE key = 'indexed_with_extraction_version'")
      .get() as { value: string } | undefined;
    db.close();
    expect(row).toBeDefined();
    const stamped = Number(row!.value);
    expect(Number.isFinite(stamped)).toBe(true);
    // The constant is captured at build time, so we read it from the same
    // binary we just ran. If the bump was forgotten, the dist binary stamps
    // the OLD constant and this assertion fires.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { EXTRACTION_VERSION } = require('../dist/extraction/extraction-version');
    expect(stamped).toBe(EXTRACTION_VERSION);
    expect(EXTRACTION_VERSION).toBe(25);
  });
});
