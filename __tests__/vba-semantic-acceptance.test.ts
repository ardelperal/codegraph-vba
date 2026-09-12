/**
 * Issue #301 — the semantic acceptance layer, and proof that it can fail.
 *
 * This file IS the documented command:
 *
 *   npm run acceptance:vba
 *
 * It builds a fresh isolated index over a throwaway copy of the corpus, asks
 * the public consumer surfaces the ground truth's questions, compares the
 * answers and exits non-zero on anything missing or unexpected. It also runs
 * on the normal `npm test` / CI path — no extra workflow hosts it.
 *
 * The corpus run is the acceptance gate. The three cases around it evaluate
 * the EVALUATOR: a harness that cannot fail is not evidence, so each feeds the
 * comparison a deliberately wrong answer through the `perturb` hook and
 * asserts the exact missing/unexpected item it should report. Those three are
 * tied to the checked-in fixture's case ids, so they stand down when the run
 * is pointed at a local export copy (see `VBA_ACCEPTANCE_CORPUS` in
 * `evaluation/vba-semantic-acceptance.ts`) — the gate itself still runs.
 */
import { describe, expect, it } from 'vitest';
import {
  AcceptanceAnswer,
  formatAcceptanceReport,
  runSemanticAcceptance,
  USING_LOCAL_CORPUS,
} from './evaluation/vba-semantic-acceptance';

const LONG = 120_000;

describe('issue #301 — semantic acceptance on the Access corpus', () => {
  it(
    'semantic acceptance passes on canonical export corpus',
    async () => {
      const report = await runSemanticAcceptance();

      // Surface the whole report on failure — a bare count tells nobody which
      // expected identity went missing or which unrelated one showed up.
      if (report.failed > 0) throw new Error(formatAcceptanceReport(report));
      expect(report.failed).toBe(0);
      expect(report.passed).toBe(report.environment.caseCount);
      // Every consumer question in the issue is actually covered.
      expect(new Set(report.cases.map((c) => c.kind))).toEqual(
        new Set(['behavior', 'ambiguity', 'impact', 'layout']),
      );
    },
    LONG,
  );

  it.skipIf(USING_LOCAL_CORPUS)(
    'acceptance detects missing expected path',
    async () => {
      const report = await runSemanticAcceptance({
        perturb: (id, answer): AcceptanceAnswer =>
          id === 'orders-save-control'
            ? {
                ...answer,
                // The consumer stopped at the handler and never reached the callee.
                callPaths: [['btnSave_Click']],
                callPathMembers: ['btnSave_Click'],
                tables: [],
              }
            : answer,
      });

      const result = report.cases.find((c) => c.id === 'orders-save-control')!;
      expect(result.status).toBe('failed');
      expect(result.missing).toContain('callPaths: btnSave_Click -> SaveOrderTotals');
      expect(result.missing).toContain('tables: tblOrderLines');
      expect(report.failed).toBe(1);
    },
    LONG,
  );

  it.skipIf(USING_LOCAL_CORPUS)(
    'acceptance rejects unrelated cross-form path',
    async () => {
      const report = await runSemanticAcceptance({
        perturb: (id, answer): AcceptanceAnswer =>
          id === 'orders-save-control'
            ? {
                ...answer,
                // The consumer answered with the OTHER form's btnSave.
                callPaths: [
                  ...answer.callPaths,
                  ['btnSave_Click', 'SaveInvoiceTotals'],
                ],
                callPathMembers: [...answer.callPathMembers, 'SaveInvoiceTotals'],
                tables: [...answer.tables, 'tblInvoiceLines'],
              }
            : answer,
      });

      const result = report.cases.find((c) => c.id === 'orders-save-control')!;
      expect(result.status).toBe('failed');
      expect(result.unexpected).toContain('callPathMembers: SaveInvoiceTotals');
      expect(result.unexpected).toContain('tables: tblInvoiceLines');
      // The answer still contained everything expected — only the extra is wrong.
      expect(result.missing).toEqual([]);
    },
    LONG,
  );

  it.skipIf(USING_LOCAL_CORPUS)(
    'acceptance preserves unresolved evidence',
    async () => {
      const report = await runSemanticAcceptance({
        perturb: (id, answer): AcceptanceAnswer =>
          id === 'orders-export-unresolved'
            ? {
                ...answer,
                // The consumer "resolved" the dynamic call by guessing.
                unresolved: [],
                callPaths: [['btnExport_Click', 'ExportOrdersToExcel']],
                callPathMembers: ['btnExport_Click', 'ExportOrdersToExcel'],
              }
            : answer,
      });

      const result = report.cases.find((c) => c.id === 'orders-export-unresolved')!;
      expect(result.status).toBe('failed');
      expect(result.missing).toContain('unresolved: ExportOrdersToExcel');
      expect(result.unexpected).toContain('callPathMembers: ExportOrdersToExcel');
    },
    LONG,
  );

  it(
    'reports reproduction metadata without timestamps or temporary paths',
    async () => {
      const report = await runSemanticAcceptance();

      expect(report.environment.node).toBe(process.version);
      expect(report.environment.package).toMatch(/@\d+\.\d+\.\d+$/);
      expect(report.environment.corpus).toBe(
        USING_LOCAL_CORPUS
          ? 'local export copy (path not recorded)'
          : '__tests__/fixtures/vba-consumer-semantics',
      );
      expect(report.environment.index).toContain('fresh throwaway copy');

      const rendered = formatAcceptanceReport(report);
      expect(rendered).not.toMatch(/vba-semantic-acceptance-[A-Za-z0-9]{6}/);
      expect(rendered).not.toMatch(/\d{4}-\d{2}-\d{2}T/);
      // Deterministic: a second run renders byte-identically.
      expect(formatAcceptanceReport(await runSemanticAcceptance())).toBe(rendered);
    },
    LONG,
  );
});
