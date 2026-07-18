/**
 * vba-extractor module-load invariant (Issue #164).
 *
 * The orchestrator at `src/extraction/vba-extractor.ts` runs a self-check
 * IIFE at module load. If any required per-concern `RULES` array is empty
 * (procedures, dims, implements, declarations, enums-consts, or
 * call-sweep), the IIFE must FAIL LOUDLY — a silent `console.error`
 * means a consumer running `codegraph index` without first running
 * `npm test` can lose an entire concern without noticing.
 *
 * `extraction-vba-rule-table.test.ts` already covers the validator's
 * return-value contract (`validateVbaRuleTables(...)` with a fake
 * tables arg). THIS suite covers the OTHER half of the contract: the
 * IIFE itself refuses to leave module-load alive when one of the
 * tables is empty.
 *
 * Test seam: this project compiles to CommonJS, so the same dynamic
 * `import()` flow Vitest uses for `vi.doMock` lets us re-import the
 * orchestrator with a stubbed classifier module AFTER clearing the
 * module cache (`vi.resetModules()`). The IIFE runs synchronously on
 * import, so the rejection from a stubbed-out classifier surfaces as
 * an `import(...)` rejection — exactly the behavior the issue's
 * "spawn a Node child process" suggestion aimed to exercise, but
 * deterministic and in-process.
 */
import { afterEach, describe, it, expect, vi } from 'vitest';

afterEach(() => {
  vi.doUnmock('../src/extraction/vba/dims');
  vi.resetModules();
});

describe('vba-extractor module-load invariant (#164)', () => {
  it('refuses to load the orchestrator when a classifier RULES table is empty', async () => {
    // Empty out the dims classifier's RULES. The other tables are
    // imported live and stay non-empty, so the validator's
    // REQUIRED_DISPATCH_TABLES sweep must report exactly `dims`.
    vi.doMock('../src/extraction/vba/dims', () => ({
      RULES: [],
      createDimsClassifier: () => ({
        name: 'dims',
        classifyLine: () => undefined,
        count: 0,
      }),
    }));
    vi.resetModules();

    await expect(import('../src/extraction/vba-extractor')).rejects.toThrow(
      /VBA_RULE_TABLES is missing rules for: dims/,
    );

  });

  it('loads the orchestrator cleanly when every RULES table is non-empty', async () => {
    // Sanity check: with NO mock in place the live tables must all
    // validate, and the import must resolve without throwing. This
    // pins the live happy-path the production refactor must keep
    // intact (acceptance criterion #3 — "all 30 existing VBA tests
    // still pass").
    vi.resetModules();
    await expect(import('../src/extraction/vba-extractor')).resolves.toBeDefined();
  });
});
