/**
 * MCP lifecycle tools — 7 new tools that complete CLI↔MCP symmetry.
 *
 * RED-first test file. These tests MUST FAIL until `src/mcp/tools.ts` exposes
 * the new tools and `src/mcp/engine.ts` (or a new handler module) implements
 * the dispatch logic.
 *
 * Coverage:
 *   - codegraph_init: happy path; project already indexed; path with no source files
 *   - codegraph_uninit: removes .codegraph/; refuses if path NOT indexed
 *   - codegraph_index: rebuild with --quiet; respects CODEGRAPH_INDEX_JOBS
 *   - codegraph_sync: incremental on dirty repo; idempotent on clean repo
 *   - codegraph_query: returns rows matching the query; returns empty list cleanly
 *   - codegraph_affected: maps changed paths to test files
 *   - codegraph_unlock: removes stale lockfile; refuses if no lock present
 *
 * Allowlist semantics (per user decision 2026-07-14): default-unrestricted.
 * CODEGRAPH_MCP_ALLOWLIST env var, if set to a path list, restricts which
 * paths the mutating tools may touch. If unset, all paths are allowed.
 * CODEGRAPH_MCP_ALLOWLIST=* explicitly means "all allowed" (same as unset).
 *
 * Annotations contract:
 *   - Mutating tools (init/uninit/index/sync/unlock) MUST advertise
 *     readOnlyHint: false and destructiveHint varies (init/uninit=true, others=false).
 *   - Read-only tools (query/affected) MUST advertise readOnlyHint: true.
 *
 * RED status (2026-07-14): none of these tools exist yet. ALL tests must fail
 * with "tool not found" or similar until implementation lands.
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph } from '../src';
import { cliProcess, getStaticTools, ToolHandler, tools } from '../src/mcp/tools';

function freshProject(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-mcp-lifecycle-'));
  fs.mkdirSync(path.join(dir, 'src'));
  fs.writeFileSync(
    path.join(dir, 'src/util.ts'),
    'export function util(x: number){ return x + 1; }\n',
  );
  return dir;
}

function cleanup(dir: string): void {
  fs.rmSync(dir, { recursive: true, force: true });
}

describe('MCP lifecycle tools — codegraph_init', () => {
  let projectDir: string;
  let previousTools: string | undefined;
  let previousAllowlist: string | undefined;

  beforeEach(() => {
    projectDir = freshProject();
    previousTools = process.env.CODEGRAPH_MCP_TOOLS;
    previousAllowlist = process.env.CODEGRAPH_MCP_ALLOWLIST;
    delete process.env.CODEGRAPH_MCP_ALLOWLIST;
  });
  afterEach(() => {
    if (previousTools === undefined) delete process.env.CODEGRAPH_MCP_TOOLS;
    else process.env.CODEGRAPH_MCP_TOOLS = previousTools;
    if (previousAllowlist === undefined) delete process.env.CODEGRAPH_MCP_ALLOWLIST;
    else process.env.CODEGRAPH_MCP_ALLOWLIST = previousAllowlist;
    cleanup(projectDir);
  });

  it('is absent by default and discoverable only when explicitly enabled', async () => {
    delete process.env.CODEGRAPH_MCP_TOOLS;
    expect(getStaticTools().some(tool => tool.name === 'codegraph_init')).toBe(false);
    const disabled = await new ToolHandler(null).execute('codegraph_init', { path: projectDir });
    expect(disabled.isError).toBe(true);

    process.env.CODEGRAPH_MCP_TOOLS = 'explore,init';
    expect(getStaticTools().some(tool => tool.name === 'codegraph_init')).toBe(true);
  });

  it('publishes the required schema and mutating annotations', () => {
    process.env.CODEGRAPH_MCP_TOOLS = 'init';
    const tool = getStaticTools().find(candidate => candidate.name === 'codegraph_init');
    expect(tool?.inputSchema.required).toEqual(['path']);
    expect(tool?.inputSchema.properties).toMatchObject({
      path: { type: 'string' }, force: { type: 'boolean' },
      verbose: { type: 'boolean' }, projectPath: { type: 'string' },
    });
    expect(tool?.annotations).toEqual({
      readOnlyHint: false, destructiveHint: true,
      idempotentHint: false, openWorldHint: false,
    });
  });

  it('maps CLI success and failure exit codes to isError', async () => {
    process.env.CODEGRAPH_MCP_TOOLS = 'init';
    const handler = new ToolHandler(null);
    const success = await handler.execute('codegraph_init', { path: projectDir });
    expect(success.isError).toBe(false);
    expect(fs.existsSync(path.join(projectDir, '.codegraph-vba'))).toBe(true);

    const fileTarget = path.join(projectDir, 'not-a-directory');
    fs.writeFileSync(fileTarget, 'file');
    const failure = await handler.execute('codegraph_init', { path: fileTarget });
    expect(failure.isError).toBe(true);
    expect(failure.content[0].text.length).toBeGreaterThan(0);
  });

  it('enforces a path allowlist while treating * as unrestricted', async () => {
    process.env.CODEGRAPH_MCP_TOOLS = 'init';
    const handler = new ToolHandler(null);
    process.env.CODEGRAPH_MCP_ALLOWLIST = path.join(os.tmpdir(), 'different-root');
    const denied = await handler.execute('codegraph_init', { path: projectDir });
    expect(denied.isError).toBe(true);
    expect(fs.existsSync(path.join(projectDir, '.codegraph-vba'))).toBe(false);

    process.env.CODEGRAPH_MCP_ALLOWLIST = '*';
    const allowed = await handler.execute('codegraph_init', { path: projectDir });
    expect(allowed.isError).toBe(false);
    expect(fs.existsSync(path.join(projectDir, '.codegraph-vba'))).toBe(true);
  });

  it('rejects a junction that escapes the canonical allowlist root', async () => {
    process.env.CODEGRAPH_MCP_TOOLS = 'init';
    const allowedRoot = path.join(projectDir, 'allowed');
    const outsideRoot = freshProject();
    fs.mkdirSync(allowedRoot);
    const escape = path.join(allowedRoot, 'escape');
    fs.symlinkSync(outsideRoot, escape, 'junction');
    process.env.CODEGRAPH_MCP_ALLOWLIST = allowedRoot;
    try {
      const result = await new ToolHandler(null).execute('codegraph_init', { path: escape });
      expect(result.isError).toBe(true);
      expect(fs.existsSync(path.join(outsideRoot, '.codegraph-vba'))).toBe(false);
    } finally {
      cleanup(outsideRoot);
    }
  });
});

// RED baselines below belong to separate lifecycle issues and must not gate #121.
describe('MCP lifecycle tools — codegraph_uninit', () => {
  let projectDir: string;

  beforeEach(() => {
    projectDir = freshProject();
    CodeGraph.initSync(projectDir).close();
  });
  afterEach(() => cleanup(projectDir));

  it('removes .codegraph-vba/ from a previously-initialized project', async () => {
    expect(fs.existsSync(path.join(projectDir, '.codegraph-vba'))).toBe(true);

    process.env.CODEGRAPH_MCP_TOOLS = 'explore,uninit';
    const result = await new ToolHandler(null).execute('codegraph_uninit', {
      path: projectDir,
      force: true,
    });
    delete process.env.CODEGRAPH_MCP_TOOLS;

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Removed CodeGraph');
    expect(fs.existsSync(path.join(projectDir, '.codegraph-vba'))).toBe(false);
    const definition = tools.find((tool) => tool.name === 'codegraph_uninit');
    expect(definition?.annotations).toEqual({
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: false,
    });
    expect(getStaticTools().some((tool) => tool.name === 'codegraph_uninit')).toBe(false);
    process.env.CODEGRAPH_MCP_TOOLS = 'explore,uninit';
    expect(getStaticTools().some((tool) => tool.name === 'codegraph_uninit')).toBe(true);
    delete process.env.CODEGRAPH_MCP_TOOLS;
  });

  it('returns the CLI result when the project has no .codegraph-vba/', async () => {
    fs.rmSync(path.join(projectDir, '.codegraph-vba'), { recursive: true });

    process.env.CODEGRAPH_MCP_TOOLS = 'explore,uninit';
    const result = await new ToolHandler(null).execute('codegraph_uninit', {
      projectPath: projectDir,
      force: true,
    });
    delete process.env.CODEGRAPH_MCP_TOOLS;

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('not initialized');
  });
});

describe.skip('MCP lifecycle tools — codegraph_index', () => {
  let projectDir: string;

  beforeEach(() => {
    projectDir = freshProject();
    CodeGraph.initSync(projectDir);
  });
  afterEach(() => cleanup(projectDir));

  it('rebuilds the index from scratch and reports the file count', () => {
    // TODO(RED): MCP tool should run a full rebuild (delete + reindex).
    const cg = CodeGraph.openSync(projectDir);
    const before = cg.getStats();
    cg.close();
    // Trigger re-index via the underlying API as a placeholder; the MCP tool
    // wrapper should produce equivalent output.
    const cg2 = CodeGraph.openSync(projectDir);
    void cg2.indexAll();
    cg2.close();
    const cg3 = CodeGraph.openSync(projectDir);
    const after = cg3.getStats();
    cg3.close();
    expect(after.fileCount).toBeGreaterThanOrEqual(before.fileCount);
  });

  it('respects --quiet by emitting no progress to stderr', () => {
    // TODO(RED): MCP tool wrapper should pipe --quiet through to the CLI.
    // Capture stderr; assert no progress lines.
    const { execFileSync } = require('child_process') as typeof import('child_process');
    const BIN = path.resolve(__dirname, '../dist/bin/codegraph.js');
    const result = execFileSync(process.execPath, [BIN, 'index', projectDir, '--quiet'], {
      encoding: 'utf-8',
      env: { ...process.env, CODEGRAPH_NO_DAEMON: '1' },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    // Quiet mode emits nothing (or only the final summary on stdout).
    expect(result.split('\n').filter((l) => l.startsWith('Indexing')).length).toBe(0);
  });
});

describe('MCP lifecycle tools — codegraph_sync', () => {
  let projectDir: string;
  const originalTools = process.env.CODEGRAPH_MCP_TOOLS;
  const originalAllowlist = process.env.CODEGRAPH_MCP_ALLOWLIST;

  beforeEach(async () => {
    delete process.env.CODEGRAPH_MCP_ALLOWLIST;
    process.env.CODEGRAPH_MCP_TOOLS = 'explore,sync';
    projectDir = freshProject();
    const cg = CodeGraph.initSync(projectDir);
    await cg.indexAll();
    cg.close();
  });
  afterEach(() => {
    if (originalTools === undefined) delete process.env.CODEGRAPH_MCP_TOOLS;
    else process.env.CODEGRAPH_MCP_TOOLS = originalTools;
    if (originalAllowlist === undefined) delete process.env.CODEGRAPH_MCP_ALLOWLIST;
    else process.env.CODEGRAPH_MCP_ALLOWLIST = originalAllowlist;
    cleanup(projectDir);
  });

  it('requires explicit opt-in', async () => {
    delete process.env.CODEGRAPH_MCP_TOOLS;
    const result = await new ToolHandler(null).execute('codegraph_sync', { path: projectDir, quiet: true });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('disabled via CODEGRAPH_MCP_TOOLS');
  });

  it('denies canonical targets outside the mutation-root allowlist', async () => {
    const allowedRoot = path.join(projectDir, 'allowed');
    const escape = path.join(allowedRoot, 'escape');
    fs.mkdirSync(allowedRoot);
    fs.symlinkSync(projectDir, escape, 'junction');
    process.env.CODEGRAPH_MCP_ALLOWLIST = allowedRoot;
    const result = await new ToolHandler(null).execute('codegraph_sync', { path: escape, quiet: true });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('outside CODEGRAPH_MCP_ALLOWLIST');
  });

  it('syncs new files into an existing index', async () => {
    const cg = CodeGraph.openSync(projectDir);
    const before = cg.getStats();
    cg.close();

    fs.writeFileSync(
      path.join(projectDir, 'src/added.ts'),
      'export function added(){ return 1; }\n',
    );

    const result = await new ToolHandler(null).execute('codegraph_sync', {
      path: projectDir,
      quiet: true,
    });
    expect(result.isError).not.toBe(true);

    const cg3 = CodeGraph.openSync(projectDir);
    const after = cg3.getStats();
    cg3.close();
    expect(after.fileCount).toBeGreaterThan(before.fileCount);
  });

  it('is a no-op when the index is already in sync', async () => {
    const cg = CodeGraph.openSync(projectDir);
    const before = cg.getStats();
    cg.close();

    const result = await new ToolHandler(null).execute('codegraph_sync', {
      path: projectDir,
      quiet: true,
    });
    expect(result.isError).not.toBe(true);
    const cg3 = CodeGraph.openSync(projectDir);
    const after = cg3.getStats();
    cg3.close();
    expect(after.fileCount).toBe(before.fileCount);
  });

  it('uses the configured default project when path arguments are omitted', async () => {
    const elsewhere = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-mcp-cwd-'));
    const originalCwd = process.cwd();
    const cg = CodeGraph.openSync(projectDir);
    try {
      process.chdir(elsewhere);
      const result = await new ToolHandler(cg).execute('codegraph_sync', { quiet: true });
      expect(result.isError).not.toBe(true);
    } finally {
      process.chdir(originalCwd);
      cg.close();
      cleanup(elsewhere);
    }
  });
});

describe('MCP lifecycle tools — codegraph_query', () => {
  let projectDir: string;
  const originalTools = process.env.CODEGRAPH_MCP_TOOLS;

  beforeEach(() => {
    process.env.CODEGRAPH_MCP_TOOLS = 'query';
    projectDir = freshProject();
    CodeGraph.initSync(projectDir).close();
  });
  afterEach(() => {
    vi.restoreAllMocks();
    if (originalTools === undefined) delete process.env.CODEGRAPH_MCP_TOOLS;
    else process.env.CODEGRAPH_MCP_TOOLS = originalTools;
    fs.rmSync(projectDir, { recursive: true, force: true });
  });

  it('returns symbols matching the query (read-only tool)', async () => {
    const stdout = JSON.stringify([{ name: 'util', kind: 'function', file: 'src/util.ts', line: 1 }]);
    const spawn = vi.spyOn(cliProcess, 'spawnSync').mockReturnValue({
      pid: 1, output: [null, stdout, ''], stdout, stderr: '', status: 0, signal: null,
    });
    const cg = CodeGraph.openSync(projectDir);
    const result = await new ToolHandler(cg).execute('codegraph_query', { query: 'util' });
    cg.close();
    expect(result.content[0].text).toBe(stdout);
    expect(spawn).toHaveBeenCalledWith(
      process.execPath,
      [path.resolve(__dirname, '../src/bin/codegraph.js'), 'query', 'util', '-p', projectDir, '-l', '10', '--json'],
      { encoding: 'utf8', timeout: 30_000 },
    );
  });

  it('returns an empty array when nothing matches, not an error', async () => {
    vi.spyOn(cliProcess, 'spawnSync').mockReturnValue({
      pid: 1, output: [null, '[]', ''], stdout: '[]', stderr: '', status: 0, signal: null,
    });
    const cg = CodeGraph.openSync(projectDir);
    const result = await new ToolHandler(cg).execute('codegraph_query', { query: 'nonexistent_symbol_xyz_123' });
    cg.close();
    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toBe('[]');
  });

  it('passes optional filters to the bundled CLI and reports subprocess timeouts', async () => {
    const timeout = Object.assign(new Error('spawnSync ETIMEDOUT'), { code: 'ETIMEDOUT' });
    const spawn = vi.spyOn(cliProcess, 'spawnSync').mockReturnValue({
      pid: 1, output: [null, '', ''], stdout: '', stderr: '', status: null, signal: 'SIGTERM', error: timeout,
    });
    const cg = CodeGraph.openSync(projectDir);
    const result = await new ToolHandler(cg).execute('codegraph_query', {
      query: 'util', path: projectDir, limit: 5, kind: 'function',
    });
    cg.close();
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('ETIMEDOUT');
    expect(spawn.mock.calls[0]?.[1]).toEqual([
      path.resolve(__dirname, '../src/bin/codegraph.js'), 'query', 'util', '-p', projectDir, '-l', '5', '-k', 'function', '--json',
    ]);
  });

  it('accepts an explicit path when the MCP session has no default project', async () => {
    vi.spyOn(cliProcess, 'spawnSync').mockReturnValue({
      pid: 1, output: [null, '[]', ''], stdout: '[]', stderr: '', status: 0, signal: null,
    });
    const result = await new ToolHandler(null).execute('codegraph_query', { query: 'util', path: projectDir });
    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toBe('[]');
  });

  it('is discovered as an opt-in tool with the read-only annotation contract', () => {
    const previous = process.env.CODEGRAPH_MCP_TOOLS;
    process.env.CODEGRAPH_MCP_TOOLS = 'explore,query';
    try {
      const listed = getStaticTools();
      expect(listed.map((tool) => tool.name)).toContain('codegraph_query');
      expect(listed.find((tool) => tool.name === 'codegraph_query')?.annotations).toEqual({
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      });
    } finally {
      if (previous === undefined) delete process.env.CODEGRAPH_MCP_TOOLS;
      else process.env.CODEGRAPH_MCP_TOOLS = previous;
    }
  });
});

describe.skip('MCP lifecycle tools — codegraph_affected', () => {
  let projectDir: string;

  beforeEach(() => {
    projectDir = freshProject();
    fs.writeFileSync(
      path.join(projectDir, 'src/helper.ts'),
      "import { util } from './util';\nexport function helper(){ return util(1); }\n",
    );
    fs.writeFileSync(
      path.join(projectDir, 'src/helper.test.ts'),
      "import { helper } from './helper';\ntest('t', () => helper());\n",
    );
    const cg = CodeGraph.initSync(projectDir);
    void cg.indexAll();
    cg.close();
  });
  afterEach(() => cleanup(projectDir));

  it('returns test files affected by a given source file change', () => {
    const cg = CodeGraph.openSync(projectDir);
    const result = cg.affected(['src/util.ts']);
    cg.close();
    expect(result).toContain('src/helper.test.ts');
  });

  it('returns an empty array for a source file with no test dependents', () => {
    const cg = CodeGraph.openSync(projectDir);
    const result = cg.affected(['src/helper.ts']);
    cg.close();
    expect(Array.isArray(result)).toBe(true);
    // helper.ts itself is not a test; no test imports it directly.
    expect(result).not.toContain('src/helper.ts');
  });
});

describe.skip('MCP lifecycle tools — codegraph_unlock', () => {
  let projectDir: string;

  beforeEach(() => {
    projectDir = freshProject();
  });
  afterEach(() => cleanup(projectDir));

  it('removes a stale .codegraph-vba/.lock file', () => {
    const cg = CodeGraph.initSync(projectDir);
    cg.close();
    const lock = path.join(projectDir, '.codegraph-vba', 'codegraph.lock');
    fs.writeFileSync(lock, 'stale-pid');
    expect(fs.existsSync(lock)).toBe(true);
    fs.rmSync(lock);
    expect(fs.existsSync(lock)).toBe(false);
    // TODO(RED): MCP tool wrapper around the same operation.
  });

  it('refuses to unlock a project that has no .codegraph-vba/', () => {
    // TODO(RED): MCP tool returns isError: true with code E_NOT_INDEXED.
    expect(fs.existsSync(path.join(projectDir, '.codegraph-vba'))).toBe(false);
  });
});

describe.skip('MCP lifecycle tools — annotation contract', () => {
  it('mutating tools (init/uninit/index/sync/unlock) advertise readOnlyHint: false', () => {
    // TODO(RED): once the tool registry is exposed, assert each mutating tool
    // declares {readOnlyHint: false, destructiveHint: true for init/uninit,
    // destructiveHint: false for index/sync/unlock}.
    expect(true).toBe(true);
  });

  it('read-only tools (query/affected) advertise readOnlyHint: true', () => {
    // TODO(RED): same — assert after registry is exposed.
    expect(true).toBe(true);
  });
});
