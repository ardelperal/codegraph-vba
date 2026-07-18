/**
 * Issue #166 — bench the VBA extractor on the real `00_VBA_TOOLKIT_BENCH`
 * corpus and pin the report's required structure so the perf numbers
 * can't silently lose their evidence (date, commit SHA, run count,
 * named fixtures, per-stage breakdown, honest v1.6.2 baseline claim).
 *
 * The acceptance test is two-layered:
 *
 *   1. **Fixtures exist** — every file the issue cites as a measurement
 *      target must be on disk where the report says it was measured.
 *      `ACAuditoriaOperaciones.cls` and `ARAuditoria.cls` live under
 *      `__tests__/fixtures/vba/src/classes/` (committed in this repo).
 *      The other two cited files (`Form_FormGestionRiesgos.cls`,
 *      `mdlCursor.bas`) live only in the external bench corpus
 *      `00_VBA_TOOLKIT_BENCH`; we do NOT depend on that path here so
 *      this test stays portable — the doc itself documents bench
 *      provenance for those two files.
 *
 *   2. **Report structure** — `docs/vba-extraction-perf.md` must be a
 *      real, reproducible evidence document, not just numbers:
 *        - the date stamp 2026-07-18 (today, per the issue),
 *        - the *labeled measurement source SHA* pinned to a known
 *          historical commit (NOT the doc's own self-referential tip
 *          SHA, which drifts every time the doc is amended),
 *        - the 4 required files cited by name,
 *        - a "3 runs" / "median" methodology note (timing is noisy;
 *          n=1 is the issue's explicit anti-pattern),
 *        - every documented timing bucket (preprocess / classifiers /
 *          walk) — the test pin is at the bucket level, not the
 *          individual stage, so the report can include any subset of
 *          stages under each bucket without breaking the contract,
 *        - an honest v1.6.2 baseline comparison — either concrete
 *          numbers OR an explicit "v1.6.2 raw baseline is unavailable,
 *          compared only against documented claim" sentence. Inventing
 *          a number is the cardinal sin; the report must pick one of
 *          the two honest paths.
 *
 * Test isolation: this test reads files only, never writes. It does NOT
 * depend on the external bench corpus being installed at any path —
 * the doc cites bench provenance, this test only verifies the doc + the
 * committed fixtures. Portable across machines and CI.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';

const REPO_ROOT = path.resolve(__dirname, '..');
const REPORT_PATH = path.join(REPO_ROOT, 'docs', 'vba-extraction-perf.md');

/**
 * Labeled measurement source SHA — the exact commit whose
 * `dist/bin/codegraph.js` was running when the stderr logs were
 * captured. This is the historical, stable identifier for "the code
 * that produced the numbers". It is pinned here so a future SHA bump
 * (e.g. a re-measurement against a newer tip) is an explicit, reviewable
 * change — not a silent drift.
 *
 * Source: parent of the commit that introduced this report
 * (`5b501c3b…`) — see `git log --first-parent`.
 */
const MEASUREMENT_SOURCE_SHA = '425e33ad2ba86af1e26279c42ce14fe8cd107589';

const REQUIRED_FIXTURE_FILES = [
  'ACAuditoriaOperaciones.cls',
  'ARAuditoria.cls',
] as const;

describe('Issue #166 — vba-extraction-perf.md report structure', () => {
  describe('fixture files exist where the report says they were measured', () => {
    it('ACAuditoriaOperaciones.cls and ARAuditoria.cls exist in the committed fixture corpus', () => {
      // These two files are the fixture-only measurement targets — the
      // bench corpus does NOT contain them. They live in this repo so
      // the test is portable across machines.
      const classesDir = path.join(
        REPO_ROOT,
        '__tests__',
        'fixtures',
        'vba',
        'src',
        'classes',
      );
      for (const f of REQUIRED_FIXTURE_FILES) {
        const p = path.join(classesDir, f);
        expect(fs.existsSync(p), `${p} must exist`).toBe(true);
      }
    });
  });

  describe('docs/vba-extraction-perf.md report contract', () => {
    it('the report file exists and is non-empty', () => {
      expect(fs.existsSync(REPORT_PATH), `${REPORT_PATH} must exist`).toBe(true);
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      expect(body.length).toBeGreaterThan(200);
    });

    it('the report carries the date stamp 2026-07-18', () => {
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      expect(body).toContain('2026-07-18');
    });

    it('the report pins the labeled measurement source SHA (not a self-referential tip SHA)', () => {
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      // The "labeled measurement source SHA" is the commit whose source
      // tree produced the data. We pin it explicitly here so an amend
      // of the doc can't silently drift the data provenance. The doc
      // is allowed to reference the full 40-char SHA or any 7+-char
      // unambiguous prefix; we check the full one to be strict.
      expect(
        body,
        `report must pin the labeled measurement source SHA ${MEASUREMENT_SOURCE_SHA}`,
      ).toContain(MEASUREMENT_SOURCE_SHA);
      // And the 7-char short form (defends against an accidental
      // 40→7 truncation silently passing).
      expect(body, 'report must also include the 7-char short SHA 425e33a').toContain(
        MEASUREMENT_SOURCE_SHA.slice(0, 7),
      );
    });

    it('the report cites every required measurement target by name', () => {
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      for (const name of [
        'ACAuditoriaOperaciones.cls',
        'ARAuditoria.cls',
        'Form_FormGestionRiesgos.cls',
        'mdlCursor.bas',
      ]) {
        expect(body, `report must cite ${name}`).toContain(name);
      }
    });

    it('the report distinguishes physical line counts from extractor split slots', () => {
      // The vba-timing `(n=N)` annotation reports the post-preprocessing
      // `source.split('\n').length` count (= CRLF terminators + 1 trailing
      // slot). It is NOT the same as the file's physical line count from
      // `Get-Content | Measure-Object -Line`. The report must call this
      // out so a reader does not confuse "lines" between the two.
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      const mentionsSplit =
        /split\s*\(\s*['"]\\?n['"]\s*\)/i.test(body) ||
        /split\s*slots/i.test(body) ||
        /CRLF terminators?\s*\+\s*1/i.test(body);
      expect(
        mentionsSplit,
        'report must explain that the per-file line count in the vba-timing output is the post-preprocessing split slot count, not the physical line count',
      ).toBe(true);
    });

    it('the report identifies the exact byte provenance for mdlCursor.bas (bench vs fixture)', () => {
      // The fixture copy of mdlCursor.bas is 40 split slots (39 CRLF)
      // and 1331 bytes; the bench copy is 39 split slots (38 CRLF) and
      // 1329 bytes. The measured `(n=39)` matches the BENCH copy. The
      // report must state this explicitly so a re-measurement against
      // the fixture copy would surface as a different number, not as a
      // silent inconsistency.
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      const pinBench =
        /mdlCursor\.bas[^.\n]*?\b(?:bench|00_VBA_TOOLKIT_BENCH)\b/i.test(body) ||
        /\bbench\b[^.\n]*?mdlCursor\.bas/i.test(body) ||
        /mdlCursor\.bas[^.\n]*?\b(?:1329|38)\b/.test(body);
      expect(
        pinBench,
        'report must explicitly identify the mdlCursor.bas source as the bench corpus copy (1329 bytes / 38 CRLF, not the fixture copy 1331 bytes / 39 CRLF)',
      ).toBe(true);
    });

    it('the report documents the run count (≥3 runs, median)', () => {
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      const hasRunCount = /\b3\s*(?:runs?|measurements?|passes)\b/i.test(body) ||
        /\b[4-9]\s*(?:runs?|measurements?|passes)\b/i.test(body) ||
        /\b\d{2,}\s*(?:runs?|measurements?|passes)\b/i.test(body);
      expect(hasRunCount, 'report must cite at least 3 runs').toBe(true);
      const hasMedianOrRange = /\bmedian\b/i.test(body) || /\brange\b/i.test(body) ||
        /\bmin\b.*\bmax\b/i.test(body);
      expect(hasMedianOrRange, 'report must summarize runs by median or range').toBe(true);
    });

    it('the report breaks timings into the documented buckets (preprocess / classifiers / walk)', () => {
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      for (const bucket of ['preprocess', 'classifiers', 'walk']) {
        expect(body, `report must mention the ${bucket} bucket`).toContain(bucket);
      }
    });

    it('the report compares to v1.6.2 honestly — numbers OR an explicit unavailability note', () => {
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      expect(body, 'report must mention v1.6.2').toMatch(/v?1\.6\.2/);
      const hasHonestNumbers = /\b\d+(\.\d+)?\s*(?:ms|s)\b/.test(body);
      const hasExplicitBaselineNote =
        /baseline.*(?:unavailable|not available|missing|rebuilt|not retained)/i.test(body) ||
        /(?:compared|comparison).*(?:only|documented|claim|published)/i.test(body);
      expect(
        hasHonestNumbers || hasExplicitBaselineNote,
        'report must either publish concrete ms numbers OR state explicitly that the v1.6.2 raw baseline is unavailable',
      ).toBe(true);
    });

    it('the report has a Conclusion / verdict section', () => {
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      const headingRe = /^##\s+.*$/gm;
      const headings = [...body.matchAll(headingRe)].map((m) => m[0]);
      const hasConclusionHeading = headings.some((h) =>
        /conclusion|verdict|summary|result|findings/i.test(h),
      );
      expect(hasConclusionHeading, 'report must have a ## Conclusion / Verdict section').toBe(
        true,
      );
    });

    it('the report points at the committed parser script under scripts/', () => {
      // The reproduction block must reference a parser script that is
      // actually checked in, not a hardcoded temp-dir path. Otherwise
      // the "how to reproduce" claim is a lie.
      const body = fs.readFileSync(REPORT_PATH, 'utf8');
      expect(
        body,
        'report must reference the committed parser under scripts/ (e.g. scripts/parse-vba-timing-stderr.mjs)',
      ).toMatch(/scripts\/parse-vba-timing-stderr\.mjs/);
      const scriptPath = path.join(REPO_ROOT, 'scripts', 'parse-vba-timing-stderr.mjs');
      expect(
        fs.existsSync(scriptPath),
        `${scriptPath} must exist on disk for the reproduction claim to be honest`,
      ).toBe(true);
    });
  });
});
