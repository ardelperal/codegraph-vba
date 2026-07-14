/**
 * MCP coverage for the opt-in, read-only codegraph_query tool (#123).
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CodeGraph } from '../src';
import { cliProcess, getStaticTools, ToolHandler } from '../src/mcp/tools';

function freshProject(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-mcp-query-'));
  fs.mkdirSync(path.join(dir, 'src'));
  fs.writeFileSync(path.join(dir, 'src/util.ts'), 'export function util(x: number){ return x + 1; }\n');
  return dir;
}

describe('MCP lifecycle tools — codegraph_query', () => {
  let projectDir: string;

  beforeEach(() => {
    projectDir = freshProject();
    CodeGraph.initSync(projectDir).close();
  });
  afterEach(() => {
    vi.restoreAllMocks();
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
