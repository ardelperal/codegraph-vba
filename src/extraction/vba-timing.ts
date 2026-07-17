/**
 * Issue #156 — per-stage timing instrumentation for the VBA extraction
 * pipeline.
 *
 * Gates
 * -----
 * The instrumentation is opt-in via the `CODEGRAPH_VBA_TIMING` env var:
 *   - unset (or any value other than `1` / `2`) → no allocations, no
 *     stderr output, zero hot-path cost.
 *   - `1` → per-file timing block to stderr.
 *   - `2` → per-file block + a per-process aggregate line on shutdown.
 *
 * The default path keeps the `VbaExtractorContext.timings` Map `null`
 * (see `context.ts`). `recordStage` short-circuits when null, so the
 * closure passed to `withStage` is the only cost in the hot path —
 * one `Map.get` + one `Map.set` per stage boundary.
 *
 * Output format
 * -------------
 *   [vba-timing] <basename>            ← per-file header
 *     preprocess: <stage> <N>ms
 *     classifiers: <name> (n=<N>) <N>ms
 *     walk: <stage> <N>ms
 *
 *   [vba-timing-aggregate] files=<N> total=<N>ms
 *     preprocess: <stage> <N>ms (n=<N>)
 *     classifiers: <name> <N>ms (n=<N>)
 *     walk: <stage> <N>ms
 *
 * Format is plain-text, one-statement-per-line, prefixed with the stage
 * bucket so the test can `grep` for `preprocess:`, `classifiers:`,
 * `walk:` to validate coverage.
 */
import * as path from 'path';
import { VbaExtractorContext, VbaClassifier } from './vba/context';

export type TimingMode = 'off' | 'per-file' | 'aggregate';

/**
 * Read the timing mode from the env var. Reads are cached per call —
 * the env var is consulted on every `extract()` but the result is just
 * a number/string compare.
 */
export function readTimingMode(): TimingMode {
  const v = process.env.CODEGRAPH_VBA_TIMING;
  if (v === '1') return 'per-file';
  if (v === '2') return 'aggregate';
  return 'off';
}

/**
 * Run `fn`, time the wall-clock, and record into `ctx.timings` under
 * `stageName`. Returns the value `fn` returned. No-op when timings
 * are disabled (the Map is null) — except the `performance.now()` pair,
 * which is the irreducible cost of measuring anything.
 *
 * `performance.now()` is the right primitive here (monotonic, sub-ms
 * resolution on every supported Node version). `Date.now()` would round
 * to whole milliseconds and lose the per-stage signal on small files.
 */
export function withStage<T>(
  ctx: VbaExtractorContext,
  stageName: string,
  fn: () => T,
): T {
  if (ctx.timings === null) return fn();
  const t0 = performance.now();
  const result = fn();
  ctx.recordStage(stageName, performance.now() - t0);
  return result;
}

/**
 * Time a single `VbaClassifier.classifyLine` invocation. The accumulated
 * time is reported under the bucket `classifier.<name>` so the
 * aggregate stage totals match the per-classifier view.
 */
export function withClassifier(
  ctx: VbaExtractorContext,
  classifier: VbaClassifier,
  line: string,
  index: number,
): void {
  if (ctx.timings === null) {
    classifier.classifyLine(line, index, ctx);
    return;
  }
  const t0 = performance.now();
  classifier.classifyLine(line, index, ctx);
  const elapsed = performance.now() - t0;
  ctx.recordStage(`classifier.${classifier.name}`, elapsed);
  if (ctx.classifierInvokeCounts) {
    ctx.classifierInvokeCounts.set(
      classifier.name,
      (ctx.classifierInvokeCounts.get(classifier.name) ?? 0) + 1,
    );
  }
}

/**
 * Per-file stderr block. Cheap string-build; the per-extract call cost
 * is bounded by the small number of stages (~10).
 */
export function emitPerFileTimings(
  ctx: VbaExtractorContext,
  filePath: string,
): void {
  if (ctx.timings === null) return;
  const basename = path.basename(filePath);
  const lines: string[] = [`[vba-timing] ${basename}`];
  // Sort by stage name so the output is stable across runs and easy to
  // diff between two files.
  const entries = [...ctx.timings.entries()].sort(([a], [b]) =>
    a < b ? -1 : a > b ? 1 : 0,
  );
  for (const [name, ms] of entries) {
    const bucket = stageBucket(name);
    const invokeCount = ctx.classifierInvokeCounts?.get(name.replace(/^classifier\./, '')) ?? 0;
    const suffix =
      bucket === 'classifiers' && invokeCount > 0 ? ` (n=${invokeCount})` : '';
    lines.push(`  ${bucket}: ${name} ${ms.toFixed(3)}ms${suffix}`);
  }
  process.stderr.write(lines.join('\n') + '\n');
}

/**
 * Stage name → output-bucket mapping. Everything is bucketed by the
 * leading path component so the user can `grep -E '^(preprocess|
 * classifiers|walk):'` against the output.
 */
function stageBucket(name: string): string {
  if (name.startsWith('preprocess.')) return 'preprocess';
  if (name.startsWith('classifier.')) return 'classifiers';
  if (name.startsWith('walk.')) return 'walk';
  return 'other';
}

// ---------------------------------------------------------------------------
// Aggregate mode (CODEGRAPH_VBA_TIMING=2)
// ---------------------------------------------------------------------------

/**
 * Per-process aggregate accumulator. The process is the natural unit
 * for an `index` run — every `extract()` call inside the same Node
 * process contributes to the same totals. We expose a single
 * `flushAggregate()` so callers (orchestrator or test harness) can
 * print the totals on demand.
 */
interface AggregateTotals {
  files: number;
  perStage: Map<string, { totalMs: number; count: number }>;
}

const aggregate: AggregateTotals = {
  files: 0,
  perStage: new Map<string, { totalMs: number; count: number }>(),
};

export function recordAggregateTimings(
  ctx: VbaExtractorContext,
  _filePath: string,
): void {
  if (ctx.timings === null) return;
  aggregate.files += 1;
  for (const [name, ms] of ctx.timings) {
    const entry = aggregate.perStage.get(name) ?? { totalMs: 0, count: 0 };
    entry.totalMs += ms;
    entry.count += 1;
    aggregate.perStage.set(name, entry);
  }
}

export function flushAggregate(): void {
  if (aggregate.files === 0) return;
  let total = 0;
  for (const v of aggregate.perStage.values()) total += v.totalMs;
  const lines: string[] = [
    `[vba-timing-aggregate] files=${aggregate.files} total=${total.toFixed(3)}ms`,
  ];
  const entries = [...aggregate.perStage.entries()].sort(([a], [b]) =>
    a < b ? -1 : a > b ? 1 : 0,
  );
  for (const [name, v] of entries) {
    const bucket = stageBucket(name);
    lines.push(`  ${bucket}: ${name} ${v.totalMs.toFixed(3)}ms (n=${v.count})`);
  }
  process.stderr.write(lines.join('\n') + '\n');
  // NOTE: do NOT reset the accumulator. Subsequent flushes report the
  // running cumulative totals across every `extract()` call in this
  // process — that's the whole point of an aggregate over an index run.
  // Tests use `_resetAggregateForTests()` to clean up between cases.
}

/**
 * Test-only: clear the aggregate between test cases so they do not
 * pollute each other. Not exported in the public API.
 */
export function _resetAggregateForTests(): void {
  aggregate.files = 0;
  aggregate.perStage.clear();
}
