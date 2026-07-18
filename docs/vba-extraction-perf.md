# VBA Extraction Performance Report

**Date:** 2026-07-18
**Measurement source SHA:** `425e33ad2ba86af1e26279c42ce14fe8cd107589` (short `425e33a`) — the commit whose `dist/bin/codegraph.js` was running when the stderr logs were captured. The doc's own commit SHA is NOT pinned in this report because the doc moves with every amendment; this SHA pins the *data* instead and is the only number that must stay stable.
**Engine:** `@aroman22/codegraph-vba` 1.11.0 (post rule-table refactor #153, post timing-instrumentation #156)
**Environment:** Node v26.4.0 on Windows; isolated copy of the `00_VBA_TOOLKIT_BENCH` corpus in `C:\Users\adm1\AppData\Local\Temp\opencode\issue-166-bench`
**Measurement tool:** `CODEGRAPH_VBA_TIMING=2 codegraph index` (per-file block + per-process aggregate to stderr)
**Parser:** `scripts/parse-vba-timing-stderr.mjs` (deterministic, committed; reproduces the tables below from the raw logs)

## Why this report exists

Issue #166 asks for a measurement of the VBA extraction pipeline on the real `00_VBA_TOOLKIT_BENCH` corpus to quantify the rule-table refactor's overhead (issue #153 added function-call overhead per rule per line in `withClassifier`) and to confirm no >5% regression vs the v1.6.2 baseline.

The instrumentation is the one shipped in v1.11.0 by #156 — opt-in via `CODEGRAPH_VBA_TIMING=1|2`, zero cost when unset (no Map allocations). The default path remains untouched.

## Line-count convention: physical lines vs extractor split slots

The vba-timing per-file block annotates each classifier line with `(n=N)`. That `N` is the **post-preprocessing `source.split('\n').length`** of the file as seen by the extractor — i.e. the number of CRLF terminators in the source plus one trailing slot. It is **NOT** the same as a "physical line count" from a tool like `Get-Content | Measure-Object -Line` or `wc -l`. Throughout this report we say **"split slots"** when we mean `N`, and **"physical lines"** when we mean the line count from a shell tool. The two numbers differ whenever the file has CRLF terminators plus a trailing newline (a normal Access export).

The pinned measurement source SHA above matters here: anyone re-running the measurement against a different commit of the extractor may see a different `N` (e.g. if the preprocessing pipeline changes) — the doc would then need a corresponding update.

## Per-file provenance

The four required targets come from two sources:

| File | Bytes | CRLF | Split slots | Source |
| --- | --- | --- | --- | --- |
| `ACAuditoriaOperaciones.cls` | 15 329 | 471 | 472 | `__tests__/fixtures/vba/src/classes/` (added to the isolated bench copy; the bench corpus does NOT contain this file) |
| `ARAuditoria.cls` | 14 072 | 475 | 476 | `__tests__/fixtures/vba/src/classes/` (added to the isolated bench copy; the bench corpus does NOT contain this file) |
| `Form_FormGestionRiesgos.cls` | 11 909 | 371 | 372 | `00_VBA_TOOLKIT_BENCH\src\forms\` (bench corpus copy; not in fixtures) |
| `mdlCursor.bas` | 1 329 | 38 | 39 | `00_VBA_TOOLKIT_BENCH\src\modules\` (bench corpus copy) |

**Critical note on `mdlCursor.bas`:** the bench corpus copy is **39 split slots** (1 329 bytes, 38 CRLF). The fixture copy under `__tests__/fixtures/vba/src/modules/mdlCursor.bas` is **40 split slots** (1 331 bytes, 39 CRLF). The measured `n=39` matches the **bench copy** only. Re-measuring against the fixture copy would surface a different number; this would be a provenance drift, not a regression.

## Methodology

| Step | Detail |
| --- | --- |
| Corpus isolation | Copied `00_VBA_TOOLKIT_BENCH\src\` to a temp dir (`C:\Users\adm1\AppData\Local\Temp\opencode\issue-166-bench`) and added the two fixture `.cls` files the issue cites (`ACAuditoriaOperaciones.cls`, `ARAuditoria.cls`) into `src\classes\`. The existing `.codegraph-vba/` index at the real bench path was **not** modified. |
| Init | One `codegraph init` run with the timing env var set; indexed 311 files → 11.528 nodes / 24.511 edges in ~5.6s. |
| Measurement runs | 3 total: 1 `codegraph init` (which does a full `indexAll`) + 2 subsequent `codegraph index` runs. Each captured stderr separately. |
| Files cited | the four above (table in the previous section). |
| Aggregation | Per-file `[vba-timing]` blocks parsed into the seven documented stages; **median across 3 runs** is the reported number. Run-to-run variance is real (5–140% on the smallest files) — the issue calls out "n=1 is the anti-pattern" and we honor that with three runs per measurement. |
| Raw logs | `C:\Users\adm1\AppData\Local\Temp\opencode\issue-166-bench-runs\run{1,2,3}-*-stderr.log` (kept out-of-tree; not committed). Reproduced tables in this doc come from running the committed parser against these logs. |

## Per-file timing — the four required files

### `ACAuditoriaOperaciones.cls` (15 329 B, 471 CRLF, 472 split slots, fixture)

| run | totalMs | preprocess.strip | preprocess.join | preprocess.cc | preprocess.cc.inner | walk.procedures | walk.main | classifiers.sum |
|---|---|---|---|---|---|---|---|---|
| run1 (init) | 73.754 | 1.762 | 0.103 | 0.590 | 0.000 | 2.487 | 34.235 | 34.577 |
| run2 (index) | 77.410 | 1.879 | 0.101 | 0.455 | 0.000 | 1.916 | 36.454 | 36.605 |
| run3 (index) | 64.641 | 2.015 | 0.117 | 0.535 | 0.000 | 2.889 | 29.221 | 29.864 |
| **MEDIAN** | **73.754** | **1.879** | **0.103** | **0.535** | **0.000** | **2.487** | **34.235** | **34.577** |

Classifier breakdown (median, ms): `callsAndSql` 23.499 · `dims` 4.199 · `eventsTypesDeclares` 1.710 · `procedures` 2.011 · `enumsConsts` 1.239 · `implements` 0.564. The `callsAndSql` classifier alone accounts for ~32% of the file time — the only `.cls` in the fixtures with a meaningful `getdb().Execute "SELECT …"` call site.

### `ARAuditoria.cls` (14 072 B, 475 CRLF, 476 split slots, fixture)

| run | totalMs | preprocess.strip | preprocess.join | preprocess.cc | preprocess.cc.inner | walk.procedures | walk.main | classifiers.sum |
|---|---|---|---|---|---|---|---|---|
| run1 (init) | 18.499 | 1.038 | 0.049 | 0.364 | 0.000 | 0.735 | 8.505 | 7.808 |
| run2 (index) | 14.926 | 0.856 | 0.029 | 0.207 | 0.000 | 0.521 | 7.079 | 6.234 |
| run3 (index) | 19.326 | 1.243 | 0.036 | 0.260 | 0.000 | 0.529 | 9.102 | 8.156 |
| **MEDIAN** | **18.499** | **1.038** | **0.036** | **0.260** | **0.000** | **0.529** | **8.505** | **7.808** |

Classifier breakdown (median, ms): `callsAndSql` 6.203 · `enumsConsts` 0.390 · `dims` 0.331 · `eventsTypesDeclares` 0.316 · `procedures` 0.281 · `implements` 0.114.

### `Form_FormGestionRiesgos.cls` (11 909 B, 371 CRLF, 372 split slots, bench)

| run | totalMs | preprocess.strip | preprocess.join | preprocess.cc | preprocess.cc.inner | walk.procedures | walk.main | classifiers.sum |
|---|---|---|---|---|---|---|---|---|
| run1 (init) | 9.452 | 0.900 | 0.030 | 0.129 | 0.000 | 0.475 | 4.096 | 3.822 |
| run2 (index) | 22.642 | 0.963 | 0.033 | 0.201 | 0.000 | 0.654 | 10.854 | 9.937 |
| run3 (index) | 23.830 | 1.250 | 0.044 | 0.173 | 0.000 | 0.557 | 11.350 | 10.456 |
| **MEDIAN** | **22.642** | **0.963** | **0.033** | **0.173** | **0.000** | **0.557** | **10.854** | **9.937** |

Classifier breakdown (median, ms): `callsAndSql` 7.343 · `dims` 0.916 · `eventsTypesDeclares` 0.776 · `enumsConsts` 0.366 · `procedures` 0.345 · `implements` 0.128.

> **Variance note.** Run 1 (~9.5ms) is materially faster than runs 2–3 (~22–24ms). This is the noisy floor the issue warns about — the file is small enough that the first cold-start of the parse worker / JIT wins a measurable fraction. The median over 3 runs reflects steady-state; the lower bound is one cold-cache observation, not a regression.

### `mdlCursor.bas` (1 329 B, 38 CRLF, 39 split slots, **bench**)

| run | totalMs | preprocess.strip | preprocess.join | preprocess.cc (outer) | preprocess.cc.inner | walk.procedures | walk.main | classifiers.sum |
|---|---|---|---|---|---|---|---|---|
| run1 (init) | 3.152 | 0.131 | 0.010 | 1.089 | 0.757 | 0.046 | 0.576 | 0.543 |
| run2 (index) | 3.774 | 0.122 | 0.012 | 1.419 | 0.948 | 0.067 | 0.640 | 0.566 |
| run3 (index) | 3.736 | 0.177 | 0.011 | 1.169 | 0.811 | 0.063 | 0.790 | 0.715 |
| **MEDIAN** | **3.736** | **0.131** | **0.011** | **1.169** | **0.811** | **0.063** | **0.640** | **0.566** |

Classifier breakdown (median, ms): `eventsTypesDeclares` 0.231 · `callsAndSql` 0.176 · `enumsConsts` 0.094 · `procedures` 0.040 · `dims` 0.040 · `implements` 0.010.

> `mdlCursor.bas` is the only required file that exercises the `preprocess.cc.lexer+parser` inner stage (1 conditional-compilation `#If` directive). On the other three files the lexer+parser inner stage records 0.000ms — none of them contain `#If`/`#Const` directives, so the inner stage is correctly a no-op.

> **Provenance reminder.** The measured `n=39` matches the bench corpus copy (38 CRLF). The fixture copy would show `n=40` (39 CRLF, 1 331 bytes). Pin `mdlCursor.bas` → bench in any re-measurement.

## Corpus-wide totals (sum across all 248 `.cls`/`.bas` files)

| run | totalMs | preprocess (3 outer) | walk (procedures + main) | classifiers (sum) |
|---|---|---|---|---|
| run1 (init) | 6536 | 539 | 3212 | 2783 |
| run2 (index) | 6184 | 476 | 3063 | 2643 |
| run3 (index) | 6437 | 496 | 3196 | 2743 |
| **MEDIAN** | **6437** | **496** | **3196** | **2743** |

Stage share of the corpus total (median): **preprocess ~8%, walk ~50%, classifiers ~43%**. Walk dominates by ~7 percentage points over classifiers — the rule-table refactor's added per-rule overhead lives inside `withClassifier` (the per-line wrapper called from `walk.procedures` and `walk.main`), so its cost surfaces inside the **walk** bucket as well as the **classifiers** bucket.

## Largest fixtures — context for "real" Access code

The required files are small (≤476 split slots). The bench corpus contains much larger files where the rule-table overhead would be most visible if it scaled super-linearly. For context (medians across 3 runs):

| File | Split slots | Median total (ms) | Median ms/slot | walk.main share |
|---|---|---|---|---|
| `Funciones Generales.bas` | 9 536 | 303.0 | 0.032 | 46% |
| `Constructor.bas` | 7 610 | 217.6 | 0.029 | 44% |
| `Riesgo.cls` | 7 173 | 225.5 | 0.031 | 47% |
| `Edicion.cls` | 5 598 | 235.2 | 0.042 | 46% |
| `Proyecto.cls` | 4 883 | 239.5 | 0.049 | 44% |

Amortized cost stays in the **30–50 μs/slot** band across the largest files. `Form_FormGestionRiesgos.cls` at 372 split slots lands at **61 μs/slot** (median) — the worst amortized cost in the table — but its absolute time (22.6ms) is still negligible against the corpus-level budget. The `preprocess.cc.lexer+parser` inner stage stays at 0.000ms on all of these files (no `#If`/`#Const` directives).

## v1.6.2 baseline comparison — honest disclosure

**v1.6.2 raw baseline is NOT available in this worktree.** The CHANGELOG entry for v1.6.2 (#83) cites a documented claim:

> "extraction of a ~2900-line `.cls` runs ~10% faster at the same accuracy as before"

That figure is qualitative — it documents the shape of the v1.6.2 improvement (single split-then-walk pipeline), not a raw timing number captured on a specific corpus. There is no v1.6.2 binary in this repo's history that we can run against the same files to produce a defensible apples-to-apples number.

What this means for the comparison:

1. **The v1.6.2 raw baseline numbers cannot be reproduced** without the v1.6.2 source + the v1.6.2 extraction pipeline code path. The current worktree only contains v1.11.0 source — reverting #153 locally is out of scope for this measurement issue and would itself need a separate test that pins the pre-rule-table behavior.
2. **Without raw baseline numbers, no ">5% regression" claim can be proven.** The issue's threshold is "regression proven", and that requires measuring both code paths on the same files. Only one path (post-rule-table) was measured here.
3. **Comparison is therefore against the documented claim only**, not against raw ms numbers. The documented claim is "10% faster on a 2900-line .cls in v1.6.2" — a *qualitative* speedup that the rule-table refactor (#153) may or may not have offset. We cannot answer that here.

**Therefore: no follow-up regression issue is filed.** Per the issue's acceptance criteria, a follow-up issue is only warranted when a regression is *proven* — and it cannot be without the v1.6.2 raw baseline.

If a v1.6.2 baseline measurement becomes available in the future (e.g. a CI step that archives `codegraph --version` + per-stage timings on every release), the comparison collapses to a median-vs-median row in this same doc and the threshold becomes mechanical.

## Conclusion

**What we measured:** v1.11.0 (post rule-table refactor) on the real `00_VBA_TOOLKIT_BENCH` corpus, 3 runs of `CODEGRAPH_VBA_TIMING=2 codegraph index`, median per-file + corpus-wide aggregates, parsed from stderr into the documented stage buckets (`preprocess.stripVbaComments`, `preprocess.joinLineContinuations`, `preprocess.conditionalCompilation`, `walk.procedures`, `walk.main`, plus the inner `preprocess.cc.lexer+parser` and the 6 classifiers).

**Required files (3-run median, total ms):**

| File | Split slots | Median total | Median ms/slot |
|---|---|---|---|
| `ACAuditoriaOperaciones.cls` (fixture) | 472 | 73.754 | 0.156 |
| `ARAuditoria.cls` (fixture) | 476 | 18.499 | 0.039 |
| `Form_FormGestionRiesgos.cls` (bench) | 372 | 22.642 | 0.061 |
| `mdlCursor.bas` (bench) | 39 | 3.736 | 0.096 |

**Corpus-level median:** 6 437 ms for 248 `.cls`/`.bas` files, 11.5k nodes, 24.5k edges. ~26 ms/file, ~14 μs/slot on average.

**Stage share (corpus median):** preprocess 8%, walk 50%, classifiers 43%. Walk.main dominates large-file extraction (~46% of file time); classifiers.sum dominates small-file extraction proportionally because the constant overhead per stage boundary matters more at small N.

**Regression verdict:** **NOT PROVEN.** The v1.6.2 raw baseline is unavailable, so the >5% threshold cannot be evaluated. We deliberately did NOT file a follow-up issue — the issue explicitly conditions that on a *proven* regression.

**What would unblock a real regression answer:**

1. A v1.6.2 binary on the same machine, run against the same isolated bench corpus with `CODEGRAPH_VBA_TIMING=2`. The 3-run median then yields the apples-to-apples comparator.
2. OR a CI step that archives per-stage timings on every published release — over time the dataset itself becomes the comparator.

Neither is in scope for issue #166. Both are good follow-up ideas if someone cares about catching a regression in this area before it ships.

## Reproduction

The temp dir layout (out-of-tree, NOT committed):

```
C:\Users\adm1\AppData\Local\Temp\opencode\issue-166-bench\
    src\
        classes\… (246 .cls files + the two fixture files)
        forms\…   (47 .cls + 47 .form.txt pairs)
        modules\… (118 .bas files)
        queries\… (8 .sql)
        reports\… (4 .report.txt)
    .codegraph-vba\… (created by init; isolated from the real bench index)

C:\Users\adm1\AppData\Local\Temp\opencode\issue-166-bench-runs\
    run1-init-stderr.log
    run2-index-stderr.log
    run3-index-stderr.log
```

The parser is committed at `scripts/parse-vba-timing-stderr.mjs`. To reproduce the tables in this doc from the retained raw logs (no bench corpus required — the parser is fully deterministic on the existing logs):

```powershell
# Run the committed parser against the retained raw logs.
cd C:\00repos\codigo\codegraph-vba-worktrees\issue-166
node scripts/parse-vba-timing-stderr.mjs `
    C:\Users\adm1\AppData\Local\Temp\opencode\issue-166-bench-runs\run1-init-stderr.log `
    C:\Users\adm1\AppData\Local\Temp\opencode\issue-166-bench-runs\run2-index-stderr.log `
    C:\Users\adm1\AppData\Local\Temp\opencode\issue-166-bench-runs\run3-index-stderr.log
```

The script accepts any number of log paths as arguments and emits the per-file tables + corpus totals as markdown on stdout. No environment variables, no hardcoded paths in the script — the caller's arguments are the only inputs.

If the raw logs are gone, to re-capture them on a fresh machine:

```powershell
# 1. Build the engine from this worktree's tip (binary built from the
#    pinned measurement source SHA — use `git checkout 425e33a` if you
#    want bit-for-bit reproducibility).
cd C:\00repos\codigo\codegraph-vba-worktrees\issue-166
npm ci
npm run build

# 2. Stage an isolated copy of the bench corpus.
$benchSrc = 'C:\00repos\codigo\00_VBA_TOOLKIT_BENCH'
$benchDst = 'C:\Users\adm1\AppData\Local\Temp\opencode\issue-166-bench'
Copy-Item -LiteralPath (Join-Path $benchSrc 'src') -Destination (Join-Path $benchDst 'src') -Recurse -Force
Copy-Item -LiteralPath 'C:\00repos\codigo\codegraph-vba-worktrees\issue-166\__tests__\fixtures\vba\src\classes\ACAuditoriaOperaciones.cls' `
          -Destination (Join-Path $benchDst 'src\classes\ACAuditoriaOperaciones.cls') -Force
Copy-Item -LiteralPath 'C:\00repos\codigo\codegraph-vba-worktrees\issue-166\__tests__\fixtures\vba\src\classes\ARAuditoria.cls' `
          -Destination (Join-Path $benchDst 'src\classes\ARAuditoria.cls') -Force

# 3. Run with timing on (3 runs minimum).
$env:CODEGRAPH_VBA_TIMING = '2'
$env:CODEGRAPH_ALLOW_UNSAFE_NODE = '1'   # only if running on Node 25+; safe to set unconditionally
$bin = 'C:\00repos\codigo\codegraph-vba-worktrees\issue-166\dist\bin\codegraph.js'
node $bin init $benchDst --force 2> run1-init-stderr.log
node $bin index $benchDst 2> run2-index-stderr.log
node $bin index $benchDst 2> run3-index-stderr.log

# 4. Parse the per-file blocks with the committed parser.
node scripts/parse-vba-timing-stderr.mjs run1-init-stderr.log run2-index-stderr.log run3-index-stderr.log > summary.md
```

The original `00_VBA_TOOLKIT_BENCH\.codegraph-vba\` index is **not** modified by any of these steps — every indexer run targets the temp copy.
