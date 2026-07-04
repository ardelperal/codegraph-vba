/**
 * Vitest config — sets a single repo-wide per-test timeout of 30s.
 *
 * Why this exists: vitest's default `testTimeout` is 5000ms. Several of
 * our tests do real work that occasionally crosses 5s on the CI Windows
 * runner (variable load on shared runners + cold-start Node subprocesses
 * for the e2e CLI tests + SQLite WAL checkpoint overhead) but well under
 * 30s. We have already hit multiple post-merge windows CI red runs from
 * this:
 *
 *   - run 28701721960 (post-merge #48): failed
 *       __tests__/index-command.test.ts:97
 *         (codegraph index — --quiet path) → 5000ms timeout
 *   - run 28702399021 (PR #70 retry): failed
 *       __tests__/db-perf.test.ts:116
 *         (deleteResolvedReferences chunking) → 5000ms timeout
 *
 * Both tests pass locally on dev Windows in 1.3–2.1s and 1–3s respectively
 * — the bottleneck is the CI Windows runner, not the test or the code.
 *
 * Per-test explicit timeouts (`it(name, fn, 15_000)`) STILL take
 * precedence; this global config only lifts the default for tests that
 * never set one. No assertion changes. No source-code changes.
 */
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    testTimeout: 30_000,
  },
});