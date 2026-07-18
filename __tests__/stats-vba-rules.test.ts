/**
 * `codegraph stats vba-rules` — issue #168.
 *
 * `VBA_RULE_TABLES` (declared in `src/extraction/vba-extractor.ts`) is the
 * orchestrator's aggregated view of every per-concern classifier's
 * declarative rule table. The issue resolved here is "what consumer, if
 * any, exercises this public surface?" — answer: a `codegraph stats
 * vba-rules [--json]` CLI command that lists every rule with its
 * concern, id, description, serialized pattern, per-concern count, and
 * the cross-concern total.
 *
 * IMPORTANT — scope of this test (and of the feature itself):
 *
 *  - The output reports **static rule counts** (the size of each
 *    classifier's `RULES` array), NOT per-rule runtime emission counts.
 *    `vba-extractor.ts` does NOT track per-rule emission counts at the
 *    orchestrator level — the only instrumentation in the pipeline is
 *    `vba-timing.ts` per-stage wall-clock timings and `VbaClassifier.count`
 *    per-classifier totals; neither attribute "rule X fired N times". So
 *    any "current emission count" column in the output would be invented
 *    data and is intentionally absent. The label `ruleCount` keeps that
 *    explicit.
 *
 *  - The CLI surface is `codegraph stats vba-rules` (subcommand of a new
 *    `stats` parent) with `--json` for machine consumption and a
 *    pretty-printed default. It mirrors `codegraph status --json`'s
 *    shape (one JSON object on stdout; pretty mode for humans).
 */

import { describe, it, expect } from 'vitest';
import { execFileSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { buildStatsVbaRules } from '../src/cli/stats-vba-rules';

const BIN = path.resolve(__dirname, '../dist/bin/codegraph.js');

/**
 * Wrap the binary call so every assertion sees the same env (skip the
 * daemon + wasm re-exec — this command does no graph work, so a fresh
 * process is fine).
 */
function runCli(
  args: string[],
  opts: { cwd?: string } = {},
): { stdout: string; stderr: string; status: number } {
  try {
    const stdout = execFileSync(process.execPath, [BIN, ...args], {
      cwd: opts.cwd ?? __dirname,
      encoding: 'utf-8',
      env: { ...process.env, CODEGRAPH_NO_DAEMON: '1', CODEGRAPH_WASM_RELAUNCHED: '1' },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    return { stdout, stderr: '', status: 0 };
  } catch (err: unknown) {
    const e = err as { stdout?: string; stderr?: string; status?: number };
    return {
      stdout: e.stdout ?? '',
      stderr: e.stderr ?? '',
      status: typeof e.status === 'number' ? e.status : 1,
    };
  }
}

describe('buildStatsVbaRules() — pure unit shape (issue #168)', () => {
  it('returns the canonical envelope: { concerns, totalRules }', () => {
    const out = buildStatsVbaRules();
    expect(out).toBeTypeOf('object');
    expect(Array.isArray(out.concerns)).toBe(true);
    expect(typeof out.totalRules).toBe('number');
  });

  it('exposes every concern the orchestrator dispatches with the live ruleCount', () => {
    // Mirrors `src/extraction/vba-extractor.ts:VBA_RULE_TABLES` exactly —
    // a refactor that adds/removes a concern here must update both sides
    // or this test will catch it.
    const out = buildStatsVbaRules();
    const byName = Object.fromEntries(out.concerns.map((c) => [c.concern, c]));
    expect(byName.implements?.ruleCount).toBe(1);
    expect(byName.procedures?.ruleCount).toBe(1);
    expect(byName.declarations?.ruleCount).toBe(5);
    expect(byName.dims?.ruleCount).toBe(2);
    expect(byName['enums-consts']?.ruleCount).toBe(6);
    expect(byName['call-sweep']?.ruleCount).toBe(4);
  });

  it('totalRules equals the sum of every concern\'s ruleCount', () => {
    const out = buildStatsVbaRules();
    const sum = out.concerns.reduce((acc, c) => acc + c.ruleCount, 0);
    expect(out.totalRules).toBe(sum);
    expect(out.totalRules).toBe(19);
  });

  it('preserves VBA_RULE_TABLES key order in concerns[]', () => {
    // Order matters: `vba-extractor.ts` exports the aggregate as an
    // object literal whose insertion order is the canonical order this
    // command should respect, so the JSON output is stable across runs.
    const out = buildStatsVbaRules();
    expect(out.concerns.map((c) => c.concern)).toEqual([
      'implements',
      'procedures',
      'declarations',
      'dims',
      'enums-consts',
      'call-sweep',
    ]);
  });

  it('every rule exposes id, description, and a serialized pattern', () => {
    const out = buildStatsVbaRules();
    for (const concern of out.concerns) {
      expect(concern.rules.length, `${concern.concern} rules length`).toBe(
        concern.ruleCount,
      );
      for (const rule of concern.rules) {
        expect(typeof rule.id, `${concern.concern}/${rule.id} id`).toBe('string');
        expect(rule.id.length, `${concern.concern}/${rule.id} id non-empty`).toBeGreaterThan(0);
        expect(
          typeof rule.description,
          `${concern.concern}/${rule.id} description`,
        ).toBe('string');
        expect(
          rule.description.length,
          `${concern.concern}/${rule.id} description non-empty`,
        ).toBeGreaterThan(0);
      }
    }
  });

  it('single-RegExp rules serialize their pattern to a string (the .source)', () => {
    const out = buildStatsVbaRules();
    // Pick a rule known to use a single RegExp: `implements.implements`
    // is a one-liner with no `defineRuleAlternatives` wrapping.
    const impl = out.concerns.find((c) => c.concern === 'implements');
    expect(impl).toBeDefined();
    expect(impl!.rules).toHaveLength(1);
    const rule = impl!.rules[0]!;
    expect(rule.id).toBe('implements');
    expect(typeof rule.pattern).toBe('string');
    expect(rule.pattern).toContain('Implements');
  });

  it('does NOT invent emission counts (no per-rule "fired N times" data exists)', () => {
    // The output must only describe what `VBA_RULE_TABLES` knows
    // statically. Anything resembling a runtime emission count would be
    // fabrication — this test pins that boundary so a future contributor
    // is forced to either wire real instrumentation or leave it out.
    const out = buildStatsVbaRules();
    for (const concern of out.concerns) {
      for (const rule of concern.rules) {
        const bag = rule as unknown as Record<string, unknown>;
        expect(bag.emissionCount, `${concern.concern}/${rule.id} emissionCount absent`).toBeUndefined();
        expect(bag.matchCount, `${concern.concern}/${rule.id} matchCount absent`).toBeUndefined();
        expect(bag.firedCount, `${concern.concern}/${rule.id} firedCount absent`).toBeUndefined();
      }
    }
    // Envelope-level: the same prohibition.
    const env = out as unknown as Record<string, unknown>;
    expect(env.emissionCount).toBeUndefined();
    expect(env.matchCount).toBeUndefined();
  });
});

describe('codegraph stats vba-rules --json (CLI integration, issue #168)', () => {
  it('the `stats` parent command exists and lists `vba-rules` in its --help', () => {
    const help = runCli(['stats', '--help']).stdout;
    expect(help).toContain('vba-rules');
  });

  it('the new subcommand is listed in the top-level --help', () => {
    const help = runCli(['--help']).stdout;
    // `stats` shows up as a top-level Commands entry. commander renders
    // it as `stats [options] [command]` — match that shape.
    expect(help).toMatch(/^\s+stats(\s|\[)/m);
  });

  it('--json prints exactly one valid JSON envelope to stdout', () => {
    const { stdout, status } = runCli(['stats', 'vba-rules', '--json']);
    expect(status).toBe(0);
    expect(stdout.trim()).not.toBe('');
    const parsed = JSON.parse(stdout.trim()) as {
      concerns: Array<{
        concern: string;
        ruleCount: number;
        rules: Array<{ id: string; description: string; pattern: string | string[] }>;
      }>;
      totalRules: number;
    };
    expect(parsed.totalRules).toBe(19);
    expect(parsed.concerns).toHaveLength(6);
    // Every concern has matching static count.
    const byName = Object.fromEntries(parsed.concerns.map((c) => [c.concern, c]));
    expect(byName.implements!.ruleCount).toBe(1);
    expect(byName.procedures!.ruleCount).toBe(1);
    expect(byName.declarations!.ruleCount).toBe(5);
    expect(byName.dims!.ruleCount).toBe(2);
    expect(byName['enums-consts']!.ruleCount).toBe(6);
    expect(byName['call-sweep']!.ruleCount).toBe(4);
  });

  it('JSON patterns serialize as string (single RegExp) or string[] (alternatives)', () => {
    const { stdout, status } = runCli(['stats', 'vba-rules', '--json']);
    expect(status).toBe(0);
    const parsed = JSON.parse(stdout.trim()) as {
      concerns: Array<{
        rules: Array<{ id: string; pattern: string | string[] }>;
      }>;
    };
    let sawString = false;
    for (const concern of parsed.concerns) {
      for (const rule of concern.rules) {
        if (typeof rule.pattern === 'string') sawString = true;
        // Pattern is either a string or a non-empty array of strings.
        expect(
          typeof rule.pattern === 'string' ||
            (Array.isArray(rule.pattern) &&
              rule.pattern.length > 0 &&
              rule.pattern.every((p) => typeof p === 'string')),
          `${rule.id} pattern must be string | string[]`,
        ).toBe(true);
      }
    }
    expect(sawString, 'at least one rule must serialize as a string').toBe(true);
  });

  it('pretty (no --json) output mentions every concern name and the total', () => {
    const { stdout, status } = runCli(['stats', 'vba-rules']);
    expect(status).toBe(0);
    expect(stdout).toContain('VBA');
    for (const concern of ['implements', 'procedures', 'declarations', 'dims', 'enums-consts', 'call-sweep']) {
      expect(stdout).toContain(concern);
    }
    // The pretty mode surfaces the total somewhere on stdout.
    expect(stdout).toMatch(/\b19\b/);
  });
});

describe('module surface (issue #168)', () => {
  it('the source module ships compiled to dist/cli/stats-vba-rules.js', () => {
    // CLI integration tests depend on the build copying src/cli/*.ts
    // through tsc to dist/cli/*.js. If a contributor refactors tsconfig
    // and accidentally narrows the include list, this test catches the
    // breakage before the integration tests above fall over with
    // `Cannot find module`.
    expect(fs.existsSync(path.resolve(__dirname, '../dist/cli/stats-vba-rules.js'))).toBe(true);
  });
});
