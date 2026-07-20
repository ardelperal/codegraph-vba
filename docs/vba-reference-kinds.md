# VBA unresolved-ref taxonomy — `reference_kind` semantics (shipped in v1.7.0+)

This document is the canonical reference for the seven `reference_kind` literals
that survive round-3 (issue #108) and that consumer-side lints, audit scripts
(`audit_v113_compare.py` and its successors), and the post-extraction stub
resolver all key on. It complements — but does not duplicate — the sibling doc
[`docs/vba-stub-repoint-decision.md`](vba-stub-repoint-decision.md) (which
covers `repointDecision`) and [`docs/vba-extraction-perf.md`](vba-extraction-perf.md)
(which covers per-stage timing). The single TypeScript source of truth is the
JSDoc block above `export type ReferenceKind` in
[`src/types.ts`](../src/types.ts) (around line 362–389); every row in the table
below is anchored there.

## TL;DR

The seven `reference_kind` literals a consumer sees in `unresolved_refs` after
the v1.7.0 round-4 round-trip:

| `reference_kind`     | Syntactic shape                                              | Example                                            | Resolves to user code?       | `stdlib allowlist` catches? |
|----------------------|--------------------------------------------------------------|----------------------------------------------------|------------------------------|-----------------------------|
| `unqualified-ident`  | bare identifier (no `()`), default for `If HayError…`-style reads | `If HayErrorEnRiesgo Then …`                 | Only via Const-read disambiguation (FR-3.1) | No — name-only gate required (issue #188) |
| `calls`              | paren-form `Name(...)` or statement-form Sub call            | `MiFuncion(x)` / `Call LimpiaBuffer`              | Yes, same-file; cross-module via resolver | No — stdlib lives in `declined-runtime` |
| `qualified-call`     | `Receiver.Member(...)` (paren) or `Receiver.Member args` (statement) | `getdb().Execute "SELECT …"`              | Yes, when receiver type matches user module | No — receiver may be runtime (`DAO`, `fso`) |
| `member-with`        | `.Member` inside a `With <receiver>` block                   | `With rs: .MoveFirst: .EOF`                       | Yes, when receiver is local/typed | No — same caveat as `qualified-call` |
| `property-get`       | dot read: `Me.Name`, `obj.Prop`                              | `Debug.Print Me.Name`                             | Yes, Me-control reads         | No (controls layer)         |
| `property-set`       | dot assignment: `Me.Name = value`                            | `Me.Name = "X"`                                   | Yes, Me-control writes        | No (controls layer)         |
| `references`         | legacy literal — anything the round did not reclassify (e.g. Implements) | `Implements IClass`                       | Yes, interface name            | No — preserved for back-compat |

Two additional literals (`function_ref`, `bang-get`, `bang-set`, `dao-query`)
are emitted by the extractor but are out of scope for the consumer-side filters
in this doc — see `src/types.ts:362-389` for the full union and rationale.

## When `declined-runtime` appears

`declined-runtime` is **never** emitted by the extractor. It is a status the
post-extraction resolver at `src/resolution/index.ts` (lines 1184, 1235, 1435)
stamps on a row when **`classifyVbaReferenceAsRuntime(reference)`** returns
`true`. The classifier consults two allowlists in
[`src/resolution/vba-runtime-objects.ts`](../src/resolution/vba-runtime-objects.ts):

1. **`RUNTIME_OBJECTS`** — the receiver is a known VBA/Access runtime object
   (`DAO`, `fso`, `err`, `ListBox`, `Collection`, `VBA`, `Application`, `Screen`,
   `DoCmd`, `CurrentDb`, `Forms`, `Reports`, `Debug`, `Modules`, `References`,
   `CommandBars`, `SysCmd`, `CurrentProject`, `CodeData`, `CodeProject`,
   `TextBox`, `ComboBox`, `CreateObject`, `GetObject`). These never link to
   user code by definition — the runtime ships them.
2. **`VBA_STDLIB_FUNCTIONS`** and **`DAO_ENUM_VALUES`** — the call site is a
   bare built-in (`CStr`, `IsNull`, `MsgBox`, `Shell`, `vbCrLf`, `dbFailOnError`,
   …). Same reasoning.

When either allowlist matches, the row's `status` flips from `failed` to
`declined-runtime`; the row stays in `unresolved_refs` (it does **not** turn
into an edge) but the consumer filter `WHERE status = 'failed'` no longer
catches it. That is the entire point — `failed` becomes the actionable bucket,
`declined-runtime` is the documented noise.

### Consumer-side SQL: distinguishing `declined-runtime` from `failed`

```sql
-- Diagnostic: distribution of unresolved-ref statuses.
-- A healthy corpus shows `declined-runtime` dominating, `failed` small.
SELECT r.status,
       r.reference_kind,
       COUNT(*) AS n
FROM unresolved_refs r
GROUP BY r.status, r.reference_kind
ORDER BY n DESC;

-- Actionable: the genuine-missing-callees list. THIS is the bucket a
-- parser-bug audit should iterate over.
SELECT r.reference_kind,
       COUNT(*) AS n
FROM unresolved_refs r
WHERE r.status = 'failed'
GROUP BY r.reference_kind
ORDER BY n DESC;
```

The first query should show `declined-runtime` at the top (DAO / `fso` / stdlib
noise); the second should show only the kinds where a real callee is
unfindable.

## Noise-ratio benchmarks (v1.13.0 bench corpus)

These are the expected **`failed` rate per `reference_kind`** on the
`00_VBA_TOOLKIT_BENCH` corpus at v1.13.0 — i.e. what fraction of each kind
*survives* the `declined-runtime` classifier and lands in the actionable
`failed` bucket. They are the lower-bound on residual noise for any new
parser-bug audit; if a future change drops a kind's `failed` count below these
numbers, that is **not** a win — it almost certainly means the classifier
swallowed a real callee.

| `reference_kind`     | Total emitted | `failed` count | `failed` / total | Source PRs                             |
|----------------------|--------------:|---------------:|-----------------:|----------------------------------------|
| `unqualified-ident`  |           649 |             65 |            ~10%  | #183 (DAO enum gate) + #188 (VBA intrinsics gate) |
| `calls`              |           689 |            207 |            ~30%  | #184 (stdlib statement-form gate) + #195 (statement-form stdlib) |
| `qualified-call`     |           560 |             56 |            ~80%  | #185 (qualified-call DAO escape) — most resolve via `vba-name-resolution` |
| `member-with`        |           702 |             35 |            ~95%  | #186 (With-block runtime-object escape) — almost all `With rs`/`With fso` patterns |
| `property-get`       |           140 |             70 |            ~50%  | inherits from cross-line property-set fix; residual is dynamic `Me.Controls!X` |

Two derived facts from this table:

- **`declined-runtime` is the dominant bucket** on every kind that reaches the
  classifier. The aggregate `failed` rate is well under 20%; the rest is
  documented runtime noise. An audit that reports `unresolved_refs` without
  filtering on `status='failed'` is reporting noise.
- **`member-with` and `qualified-call` resolve best** because they preserve the
  receiver type and the resolver can use it. `unqualified-ident` and `calls`
  resolve worst because they strip the receiver — disambiguation falls back to
  Const-read priority (FR-3.1) and name matching, which the runtime allowlist
  catches the rest of.

### How the numbers move

Each PR above is **additive**: it added a new allowlist entry or
classifier branch and shrank the `failed` bucket without touching the
emitters. Future PRs that want to improve resolution should follow the same
shape (add a classifier → measure before/after on this corpus → cite the
delta in the changelog) — never modify an emitter without re-running this
table.

## Cross-reference

- **TypeScript source of truth** — `src/types.ts:362-389` (JSDoc above
  `export type ReferenceKind`). The union literal there is the canonical
  shape; this doc mirrors it for human readers.
- **Extractors that emit each kind:**
  - `src/extraction/vba/calls.ts` — emits `calls` (line 122) and `qualified-call`
    (lines 149, 175)
  - `src/extraction/vba/call-sweep.ts` — emits `unqualified-ident` (line 323)
    and `member-with` (line 373)
  - `src/extraction/vba/controls.ts` — emits `property-get`/`property-set`
    (line 106) and the legacy `references` (line 131)
  - `src/extraction/vba/context.ts`, `src/extraction/vba/dims.ts`,
    `src/extraction/vba/rules.ts` — the supporting state, type map, and rule
    table the emissions above consult; no `referenceKind` literals of their own.
- **Runtime-object + stdlib allowlists** —
  `src/resolution/vba-runtime-objects.ts` (`RUNTIME_OBJECTS`,
  `VBA_STDLIB_FUNCTIONS`, `DAO_ENUM_VALUES`, and the `isRuntimeObject` /
  `isVbaStdlibFunction` predicates). These are what `declined-runtime` keys on.
- **Three `declined-runtime` gates in the resolver** —
  `src/resolution/index.ts:1184`, `:1235`, `:1435`. Each is the
  `status: classifyVbaReferenceAsRuntime(r) ? 'declined-runtime' : 'failed'`
  ternary at the point where unresolved refs are persisted to
  `markReferencesFailed`.
- **Sibling doc on `repointDecision`** —
  [`docs/vba-stub-repoint-decision.md`](vba-stub-repoint-decision.md). The
  `reponted-to-real` / `declined-runtime` / `declined-ambiguous` /
  `declined-not-found` taxonomy is orthogonal to `reference_kind` but shares
  the `declined-runtime` label — read both when debugging "why didn't this
  edge get repointed?"
- **Sibling doc on per-stage timing** —
  [`docs/vba-extraction-perf.md`](vba-extraction-perf.md). The classifiers that
  emit the seven kinds each have a per-stage ms cost in the
  `CODEGRAPH_VBA_TIMING=2` instrumentation; the `callsAndSql` classifier alone
  is ~32% of `ACAuditoriaOperaciones.cls` total time.
- **PRs that produced the bench numbers** — #183 (DAO enum gate, v1.13.0),
  #184 (stdlib statement-form gate, v1.13.0), #185 (qualified-call DAO escape,
  v1.13.0), #186 (With-block runtime-object escape, v1.13.0), #188 (VBA
  intrinsic constants — closed by #196), #195 (statement-form stdlib,
  v1.13.0).