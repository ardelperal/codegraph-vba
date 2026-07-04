import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: true,
    environment: 'node',
    include: ['__tests__/**/*.test.ts'],
    /**
     * The integration suite starts many real Node/SQLite/indexing processes.
     * On Windows, Vitest's default worker fan-out can exhaust V8 isolate memory
     * and make a plain `npx vitest run` fail even when the individual tests are
     * green. Keep local and CI runs deterministic by capping concurrent workers.
     */
    maxWorkers: 1,
    minWorkers: 1,
    pool: 'forks',
    poolOptions: {
      forks: {
        execArgv: ['--liftoff-only', '--max-old-space-size=8192'],
        isolate: true,
      },
    },
    /**
     * Issue #53 follow-up: bump the repo-wide test timeout from vitest's
     * default 5000ms to 30s. Several integration / e2e tests do real
     * work that occasionally crosses 5s on the CI Windows runner under
     * variable load (cold-start Node subprocesses + SQLite WAL
     * checkpoint overhead + `node dist/bin/codegraph.js init`/`index`
     * e2e). The earlier fix attempt (PR #70) added a sibling
     * `vitest.config.ts` but vitest 3.x prefers the ESM `.mts` config
     * when both are present, so the `.ts` was a silent no-op. This
     * keeps the override on the actively-loaded `.mts` file.
     *
     * Per-test explicit timeouts (`it(name, fn, 15_000)`) STILL take
     * precedence; this only lifts the default for tests that never set
     * one. No assertion changes. No source-code changes.
     */
    testTimeout: 30_000,
    /**
     * Several MCP integration tests (mcp-daemon, mcp-initialize, mcp-ppid-watchdog,
     * mcp-roots) spawn `dist/bin/codegraph.js serve --mcp` with `process.execPath`
     * and rely on the child inheriting `process.env`. On a Node >= 25 dev machine
     * the CLI's hard-block (src/bin/codegraph.ts) would otherwise exit the child
     * before it ever responds, so every spawn-based test times out — see #478.
     *
     * Setting the override here keeps the CLI's runtime guard intact for end
     * users (it's still enforced when `codegraph` is invoked directly) while
     * letting the test suite run on whatever Node the contributor happens to
     * have installed. CI on Node 22/23 is unaffected — the guard doesn't fire
     * there, so the variable is a no-op.
     */
    env: {
      CODEGRAPH_ALLOW_UNSAFE_NODE: '1',
      /**
       * The suite spawns real CLI/MCP processes; without this they would write
       * telemetry state into the contributor's real ~/.codegraph and count test
       * tool calls as real usage. The telemetry unit tests are unaffected —
       * they inject their own `env` via the Telemetry constructor.
       */
      CODEGRAPH_TELEMETRY: '0',
    },
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html'],
    },
  },
});
