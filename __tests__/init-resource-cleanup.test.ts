import { afterEach, describe, expect, it, vi } from 'vitest';

const cli = vi.hoisted(() => ({
  initAction: undefined as undefined | ((pathArg: string | undefined, options: Record<string, boolean>) => Promise<void>),
  actionPromise: undefined as undefined | Promise<void>,
}));

const graph = vi.hoisted(() => ({
  init: vi.fn(),
  destroy: vi.fn(),
  indexAll: vi.fn(),
}));

vi.mock('commander', () => {
  class FakeCommand {
    private current = '';

    name() { return this; }
    description() { return this; }
    version() { return this; }
    option() { return this; }
    aliases() { return this; }
    hook() { return this; }
    command(name: string) { this.current = name; return this; }
    action(fn: (...args: any[]) => Promise<void>) {
      if (this.current.startsWith('init ')) cli.initAction = fn;
      return this;
    }
    parse() {
      if (!cli.initAction) throw new Error('init action was not registered');
      cli.actionPromise = cli.initAction('C:/safe-project', {});
    }
  }
  return { Command: FakeCommand };
});

vi.mock('../src/mcp/early-ppid', () => ({}));
vi.mock('../src/bin/fatal-handler', () => ({ installFatalHandlers: vi.fn() }));
vi.mock('../src/extraction/wasm-runtime-flags', () => ({ relaunchWithWasmRuntimeFlagsIfNeeded: vi.fn() }));
vi.mock('../src/bin/command-supervision', () => ({ installCommandSupervision: () => ({ stop: vi.fn() }) }));
vi.mock('../src/directory', () => ({
  getCodeGraphDir: (root: string) => `${root}/.codegraph`,
  isInitialized: () => false,
  unsafeIndexRootReason: () => null,
  findNearestCodeGraphRoot: () => null,
  planFrontload: vi.fn(),
  hasStructuralKeyword: vi.fn(),
  extractCodeTokens: vi.fn(),
}));
vi.mock('../src/index', () => ({
  default: { init: graph.init },
  getDatabasePath: (root: string) => `${root}/.codegraph/codegraph.db`,
}));
vi.mock('../src/telemetry', () => ({
  TELEMETRY_DOCS: '',
  getTelemetry: () => ({ recordUsage: vi.fn(), maybeFlush: vi.fn(), flushNow: vi.fn(async () => undefined) }),
  recordIndexEvent: vi.fn(),
}));
vi.mock('../src/ui/shimmer-progress', () => ({
  createShimmerProgress: () => ({ onProgress: vi.fn(), stop: vi.fn(async () => undefined) }),
}));
vi.mock('../src/ui/glyphs', () => ({ getGlyphs: () => ({ err: 'x', ok: 'ok', dash: '-', rail: '|', info: 'i', warn: '!' }) }));
vi.mock('../src/bin/daemon-release', () => ({ registerDaemonStopCommand: vi.fn() }));
vi.mock('../src/sync/worktree', () => ({ detectWorktreeIndexMismatch: vi.fn(), worktreeMismatchWarning: vi.fn() }));
vi.mock('../src/search/identifier-segments', () => ({ extractProseCandidates: vi.fn() }));
vi.mock('../src/bin/node-version-check', () => ({ buildNodeTooOldBanner: vi.fn(), isBelowMinimumNodeVersion: () => false }));

const successfulResult = {
  success: true,
  filesIndexed: 1,
  filesSkipped: 0,
  filesErrored: 0,
  nodesCreated: 1,
  edgesCreated: 0,
  errors: [],
  durationMs: 1,
};

async function runInit(): Promise<void> {
  vi.resetModules();
  cli.initAction = undefined;
  cli.actionPromise = undefined;
  process.argv = ['node', 'codegraph', 'init', 'C:/safe-project'];
  process.env.CODEGRAPH_WASM_RELAUNCHED = '1';
  const log = { error: vi.fn(), info: vi.fn(), warn: vi.fn(), success: vi.fn() };
  const clack = { intro: vi.fn(), outro: vi.fn(), log, note: vi.fn(), confirm: vi.fn(), isCancel: () => false };
  vi.stubGlobal('Function', vi.fn(() => async () => clack));
  await import('../src/bin/codegraph');
  await cli.actionPromise;
}

describe('init resource cleanup (#241)', () => {
  const originalArgv = [...process.argv];
  const originalExitCode = process.exitCode;

  afterEach(() => {
    process.argv = [...originalArgv];
    process.exitCode = originalExitCode;
    graph.init.mockReset();
    graph.destroy.mockReset();
    graph.indexAll.mockReset();
    vi.restoreAllMocks();
  });

  it('releases the initialized graph after a successful initial index', async () => {
    graph.init.mockResolvedValue({ indexAll: graph.indexAll, destroy: graph.destroy });
    graph.indexAll.mockResolvedValue(successfulResult);

    await runInit();

    expect(graph.destroy).toHaveBeenCalledTimes(1);
  });

  it('releases the initialized graph when indexing fails unexpectedly', async () => {
    graph.init.mockResolvedValue({ indexAll: graph.indexAll, destroy: graph.destroy });
    graph.indexAll.mockRejectedValue(new Error('synthetic index failure'));
    vi.spyOn(process, 'exit').mockImplementation((() => undefined) as never);

    await runInit();

    expect(graph.destroy).toHaveBeenCalledTimes(1);
  });

  it('does not attempt cleanup when graph construction fails before ownership is acquired', async () => {
    graph.init.mockRejectedValueOnce(new Error('synthetic construction failure'));
    vi.spyOn(process, 'exit').mockImplementation((() => undefined) as never);

    await runInit();

    expect(graph.destroy).not.toHaveBeenCalled();
  });
});
