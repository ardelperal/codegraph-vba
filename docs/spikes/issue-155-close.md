## F.1 done — GO verdict, PR #163 merged, v1.11.0 released

**Status**: F.1 spike **complete**, verdict **GO**, shipped in **v1.11.0** (https://github.com/ardelperal/codegraph-vba/releases/tag/v1.11.0).

### What landed

- **PR #163** (https://github.com/ardelperal/codegraph-vba/pull/163) merged to main.
- **Spike script**: `scripts/spike-vbnet-as-vba.mjs` (~50 KB, idempotent, regenerable via `pnpm run spike:vbnet`).
- **Report**: `docs/spikes/vbnet-as-vba.md` + `docs/spikes/vbnet-as-vba.json`.
- **Smoke test**: `__tests__/spike-vbnet-as-vba.test.ts` — 10 tests pin the artifact, structure, verdict format, and the F.2-unlock dry-run result.
- **`package.json`**: added `spike:vbnet` npm script (no dep changes).

### Verdict: GO (both gates pass)

| Gate | Threshold | Measured | Pass? |
|---|---|---|---|
| Quantitative (issue #155 hard gate) | <30% files failed | 0% (16/16) | ✅ |
| Structural (new in this spike) | body parses AND structure unlocks via synthesized wrapper | body=1681 nodes; synth=13 class + 26 method + 36 field + 36 param_list + 2 event + 2 raiseevent | ✅ |

### Headline finding (the F.2 unlock)

VBA module files (`.bas`/`.cls`) have **no `Class X` / `Module X` opener** — the module name lives in the Access class header (`Attribute VB_Name = "X"`), which the grammar does not recognize. The spike ran a synthesized-wrapper dry-run that injects `Class <Name> ... End Class` around every file (where `<Name>` comes from the existing `Attribute VB_Name` extraction). **Result**: the grammar emits the full structural tree (class / method / field / property / parameter / event). F.2 is feasible — see the report's "F.2 dry-run" section for the data.

### Closing this issue (per the epic's own structure)

The original issue framed this as a 4-phase epic "with go/no-go at F.1" — and the closing clause was:

> "Epic can be CANCELLED at F.1 with no further work if the report is negative."

The report is **positive (GO)**, not negative, so the epic does NOT cancel. F.2, F.3, F.4 remain pending work. Closing this tracking issue because:

1. **F.1 acceptance criterion is met** ("F.1 spike produces a go/no-go report").
2. **F.2-F.4 are scoped sub-efforts**, not sub-issues of this one — they should be tracked as their own issues when work begins, not as a checklist hanging off a parent issue that says "F.1 done" on it.
3. **The spike report is the durable artifact** — `docs/spikes/vbnet-as-vba.md` captures the verdict, the per-construct data, the F.2 unlock dry-run, and the pre-processing checklist. Anyone picking up F.2 should start there.

### F.2-F.4 status (open follow-up work, not closed)

These are **not** done. F.2 is the next step and the report spells out exactly what it needs:

- **F.2** — Hybrid: vbnet AST for the language skeleton, regex for Access-specific.
  - **Pre-processing checklist for F.2** (in the report):
    1. Blank `VERSION 1.0 CLASS` + `BEGIN ... END` block.
    2. Strip `Attribute VB_*` lines.
    3. Blank `Option Compare Database` / `Option Explicit` lines.
    4. Rewrite `Wend` → `End While`.
    5. Append a trailing newline.
    6. **Inject a synthetic `Class <Name>` opener + `End Class` closer** — the unlock.
  - **F.2 implementation is feasible** but needs revision of the original "walk the AST for procedures/classes" plan to include step 6.
- **F.3** — Coverage matrix: every existing test must pass without modification. Run delta.
- **F.4** — Performance + memory characterization. Compare AST walk vs regex pipeline.

Recommend opening separate issues for F.2 / F.3 / F.4 when work begins, each referencing the F.1 spike report and the F.2 pre-processing checklist.

Refs: PR #163, v1.11.0.
