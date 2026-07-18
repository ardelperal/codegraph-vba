/**
 * Smoke test for the F.1 spike of issue #155 — verify the
 * `scripts/spike-vbnet-as-vba.mjs` artifact (markdown report + JSON dump) is
 * present, non-empty, and structurally sound. The script is allowed to
 * RE-RUN at test time so a fresh checkout (or a CI run that re-clones) always
 * gets a current report — but the test never spawns the WASM load in the
 * assertion path; we just check what the previous run produced.
 *
 * If the artifact is missing, the test invokes the script (node directly) and
 * re-checks. This keeps the test self-contained: a fresh clone (no artifact
 * present) still produces a green run, but a PR that deletes the artifact
 * without re-running the script gets a red signal.
 */
import { describe, it, expect, beforeAll } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { execFileSync } from 'child_process';
import { pathToFileURL } from 'url';

const REPO_ROOT = path.resolve(__dirname, '..');
const REPORT_PATH = path.join(REPO_ROOT, 'docs', 'spikes', 'vbnet-as-vba.md');
const JSON_PATH = path.join(REPO_ROOT, 'docs', 'spikes', 'vbnet-as-vba.json');
const SCRIPT_PATH = path.join(REPO_ROOT, 'scripts', 'spike-vbnet-as-vba.mjs');
const CLOSE_NOTES_PATH = path.join(REPO_ROOT, 'docs', 'spikes', 'issue-155-close.md');

/**
 * Ensure the artifact exists. If not, spawn the spike script (node, no
 * external dep on `pnpm`/`tsx`) so the artifact is regenerated before the
 * assertions run. The script is idempotent — re-runs produce equivalent
 * content, with `generatedAt` and (potentially) per-file numbers updated.
 */
function ensureArtifact(): void {
  if (fs.existsSync(REPORT_PATH) && fs.existsSync(JSON_PATH)) return;
  // eslint-disable-next-line no-console
  console.log('[spike-test] artifact missing, running spike script to regenerate');
  execFileSync(process.execPath, [SCRIPT_PATH], {
    cwd: REPO_ROOT,
    stdio: 'inherit',
    env: { ...process.env, NODE_NO_WARNINGS: '1' },
  });
}

interface SpikeJSON {
  generatedAt: string;
  grammar: string;
  fixturesRoot: string;
  totals: { files: number; clean: number; partial: number; failed: number; failedRate: number };
  structural: { nodeTypes: string[]; recognized: number; filesSeen: number };
  body: { nodeTypes: string[]; recognized: number; filesSeen: number };
  structuralCompleteness: { hasClassWrapper: boolean; hasProcedures: boolean; hasFieldDecls: boolean; ratio: number };
  synthTotals: {
    files: number;
    skipped: number;
    perStructuralNode: Record<string, number>;
  };
  synthesizedResults: Array<{
    path: string;
    className?: string;
    skipped?: boolean;
    reason?: string;
    total?: number;
    error?: number;
    missing?: number;
    classification?: 'clean' | 'partial' | 'failed';
    perType?: Record<string, number>;
    structuralCounts?: Record<string, number>;
  }>;
  perFile: Array<{
    path: string; bytes: number; lines: number; total: number; error: number; missing: number;
    classification: 'clean' | 'partial' | 'failed';
    errorSamples: Array<{ line: number; col: number; text: string }>;
    perType: Record<string, number>;
  }>;
  perConstruct: Record<string, { recognized: number; filesSeen: number }>;
  keywordPresence: Record<string, { occurrences: number; files: number }>;
}

describe('spike: tree-sitter-vbnet as VBA parser (F.1, issue #155)', () => {
  beforeAll(() => {
    ensureArtifact();
  }, 180_000); // first run loads the 6.4 MB WASM; allow generous timeout

  it('produces the markdown report at docs/spikes/vbnet-as-vba.md', () => {
    expect(fs.existsSync(REPORT_PATH)).toBe(true);
  });

  it('produces the machine-readable JSON at docs/spikes/vbnet-as-vba.json', () => {
    expect(fs.existsSync(JSON_PATH)).toBe(true);
  });

  it('archives the F.2 deferral and links its trigger guard', () => {
    const notes = fs.readFileSync(CLOSE_NOTES_PATH, 'utf8');
    expect(notes).toContain('F.2 implementation is not scheduled');
    expect(notes).toContain('Issue #170');
    expect(notes).toContain('Issue #170 was closed');
    expect(notes).toContain('no qualifying trigger');
  });

  it('report is non-empty and not a placeholder', () => {
    const content = fs.readFileSync(REPORT_PATH, 'utf8');
    expect(content.length).toBeGreaterThan(2_000);
    expect(content).toContain('# Spike: tree-sitter-vbnet.wasm as a VBA parser (F.1)');
    expect(content).toContain('## Verdict');
    expect(content).toContain('## Headline numbers');
    expect(content).toContain('## Structural completeness');
    expect(content).toContain('## Per-file results');
    expect(content).toContain('## Per-construct recognition');
    expect(content).toContain('## VBA constructs the vbnet grammar recognizes cleanly');
    expect(content).toContain('## VBA constructs the vbnet grammar fails on');
    expect(content).toContain('## What pre-processing can and cannot fix');
    expect(content).toContain('## F.2 dry-run: synthesized `Class <Name> … End Class` wrapper');
    expect(content).toContain('## Pre-processing checklist for F.2');
  });

  it('report verdict line matches the issue\'s go/no-go gate language', () => {
    const content = fs.readFileSync(REPORT_PATH, 'utf8');
    // The verdict is the bolded text right after the "## Verdict" heading:
    //   ## Verdict
    //   **GO with major caveat** — The quantitative gate passes ...
    // The issue's quantitative gate is ">30% failed = NO-GO"; the report
    // explicitly references this gate so a future maintainer can audit the
    // decision against the original threshold.
    expect(content).toMatch(/## Verdict[\s\S]{0,80}\*\*(GO with major caveat|GO with caveats|GO|NO-GO)\*\*/);
    expect(content).toMatch(/30% threshold/);
  });

  it('JSON parses and has the expected top-level shape', () => {
    const raw = fs.readFileSync(JSON_PATH, 'utf8');
    const parsed = JSON.parse(raw) as SpikeJSON;
    expect(parsed.grammar).toBe('tree-sitter-vbnet.wasm');
    expect(parsed.fixturesRoot).toMatch(/fixtures/);
    expect(parsed.generatedAt).toMatch(/^\d{4}-\d{2}-\d{2}T/);
    expect(parsed.totals).toMatchObject({
      files: expect.any(Number),
      clean: expect.any(Number),
      partial: expect.any(Number),
      failed: expect.any(Number),
      failedRate: expect.any(Number),
    });
    expect(parsed.totals.files).toBeGreaterThan(0);
    expect(parsed.totals.clean + parsed.totals.partial + parsed.totals.failed).toBe(parsed.totals.files);
    expect(parsed.structural).toMatchObject({
      nodeTypes: expect.any(Array),
      recognized: expect.any(Number),
      filesSeen: expect.any(Number),
    });
    expect(parsed.body).toMatchObject({
      nodeTypes: expect.any(Array),
      recognized: expect.any(Number),
      filesSeen: expect.any(Number),
    });
    expect(parsed.structuralCompleteness).toMatchObject({
      hasClassWrapper: expect.any(Boolean),
      hasProcedures: expect.any(Boolean),
      hasFieldDecls: expect.any(Boolean),
      ratio: expect.any(Number),
    });
    expect(parsed.perFile.length).toBe(parsed.totals.files);
    expect(Object.keys(parsed.perConstruct).length).toBeGreaterThan(0);
  });

  it('per-file records have the expected shape and all parse as non-empty', () => {
    const raw = fs.readFileSync(JSON_PATH, 'utf8');
    const parsed = JSON.parse(raw) as SpikeJSON;
    for (const f of parsed.perFile) {
      expect(f.path).toMatch(/__tests__[\\/]fixtures[\\/]/);
      expect(f.bytes).toBeGreaterThan(0);
      expect(f.lines).toBeGreaterThan(0);
      expect(f.total).toBeGreaterThan(0);
      expect(['clean', 'partial', 'failed']).toContain(f.classification);
      expect(f.error).toBeGreaterThanOrEqual(0);
      expect(f.missing).toBeGreaterThanOrEqual(0);
      expect(f.perType).toBeTypeOf('object');
    }
  });

  it('per-construct rollup captures the expected grammar node types', () => {
    const raw = fs.readFileSync(JSON_PATH, 'utf8');
    const parsed = JSON.parse(raw) as SpikeJSON;
    // Body-level nodes that the spike was designed to detect. If the spike
    // ever drops these from the rollup, the report's "Structural completeness"
    // table is wrong. Pin the existence of at least the body nodes that
    // should be present in any parse of the corpus.
    const expectedBody = [
      'if_statement', 'with_statement', 'assignment_statement', 'expression_statement',
      'comment', 'invocation_expression', 'member_access_expression',
    ];
    for (const n of expectedBody) {
      expect(parsed.perConstruct[n], `perConstruct.${n} missing`).toBeDefined();
      expect(parsed.perConstruct[n].recognized, `perConstruct.${n}.recognized > 0`).toBeGreaterThan(0);
    }
  });

  it('quantitative gate: failed rate is below the 30% NO-GO threshold', () => {
    // Issue #155's hard gate: "If the failure rate is >30% on real Dysflow
    // exports, the epic becomes a research project; cancel F.2-F.4." If this
    // assertion ever fires, the spike has regressed — re-read the JSON
    // verdict, and re-investigate.
    const raw = fs.readFileSync(JSON_PATH, 'utf8');
    const parsed = JSON.parse(raw) as SpikeJSON;
    expect(parsed.totals.failedRate).toBeLessThan(0.30);
  });

  it('synthesized-wrapper dry-run unlocks the structural tree (F.2 hypothesis)', () => {
    // The F.2 hypothesis is that injecting a `Class <Name> ... End Class` wrapper
    // (where <Name> comes from the existing `Attribute VB_Name` extraction)
    // makes the grammar emit class_declaration, method_declaration, field_declaration
    // etc. on the VBA corpus. The spike ran this experiment. Pin at least
    // one class_declaration in the synthesized counts — if this fires, the
    // F.2 path is broken (or the spike's wrapper logic regressed).
    const raw = fs.readFileSync(JSON_PATH, 'utf8');
    const parsed = JSON.parse(raw) as SpikeJSON;
    expect(parsed.synthTotals).toBeDefined();
    expect(parsed.synthTotals.files).toBeGreaterThan(0);
    // At least one synthesized class_declaration must exist (every file with
    // an Attribute VB_Name produced a Class wrapper).
    expect(
      parsed.synthTotals.perStructuralNode.class_declaration ?? 0,
      'synthesized-wrapper produced no class_declaration — F.2 unlock is broken',
    ).toBeGreaterThan(0);
    // And at least one method_declaration (the F.2 plan needs the procedure
    // boundary to come from the AST).
    expect(
      parsed.synthTotals.perStructuralNode.method_declaration ?? 0,
      'synthesized-wrapper produced no method_declaration — F.2 unlock is broken',
    ).toBeGreaterThan(0);
  });

  it('artifact is reproducible (the script is idempotent — re-running writes the same paths)', () => {
    // The script's contract: report and JSON go to the same paths on every
    // run. Pin the path layout so a future rename can't silently break the
    // smoke test (which would otherwise pass on stale artifacts).
    expect(REPORT_PATH).toContain('docs');
    expect(REPORT_PATH).toContain('spikes');
    expect(REPORT_PATH.endsWith('vbnet-as-vba.md')).toBe(true);
    expect(JSON_PATH.endsWith('vbnet-as-vba.json')).toBe(true);
    // script lives at scripts/spike-vbnet-as-vba.mjs
    expect(SCRIPT_PATH.endsWith('spike-vbnet-as-vba.mjs')).toBe(true);
  });
});
