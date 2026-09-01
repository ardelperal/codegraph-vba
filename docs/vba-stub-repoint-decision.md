# VBA call-stub resolver — `repointDecision` semantics (shipped in v1.7.0)

This document is the canonical reference for the post-extraction stub resolver
introduced in [#110](https://github.com/ardelperal/codegraph-vba/issues/110) and
shipped in v1.7.0. It supersedes the original round-5 prompt acceptance
criterion (`stub_true_count < 500`) — see [#115](https://github.com/ardelperal/codegraph-vba/issues/115).

## TL;DR

| `metadata.repointDecision` | `metadata.stub` | Meaning | Consumer action |
|---|---|---|---|
| `reponted-to-real` | `false` | The synthetic function node was repointed to a real `nodes.id`. This was the original goal of round-4. | None — treat as a normal call edge. |
| `declined-runtime` | `true` | The receiver is a known runtime object (`DAO.*`, `fso.*`, `ListBox.*`, `Collection.*`, `err.*`, `VBA.*`, `Application.*`, `Screen.*`, `DoCmd.*`, `CurrentDb.*`, `Forms`, `Reports`, `Debug`, `Modules`, `References`, `CommandBars`, `SysCmd`, `CreateObject`, `GetObject`, `Fields`). It can never link to user code. **Since #245 this bucket is empty in practice** — see "What #245 changed". | Filter OUT — noise. |
| `declined-ambiguous` | `true` | The qualified name resolves to multiple real `nodes.id` candidates (e.g. two classes named `DAO` in different modules). Repointing would be unsafe. | Investigate — flag for manual review. |
| `declined-not-found` | `true` | The qualified name doesn't match any real `nodes.id` — this is a genuinely missing callee. | Filter IN — actionable signal. |

## Why the criterion shifted

The original round-5 prompt expected the resolver to **repoint** every
`stub=true` edge whose target exists in `nodes`, reducing the global
`stub=true` count from ~5000 to <500. In practice that wasn't possible
because runtime-object stubs (`DAO.Execute`, `fso.GetFileName`, etc.) have
no user-code target at all — they're not "found", they're "impossible".

So the shipped fix (v1.7.0) instead **annotates** each stub with the
diagnostic `repointDecision` field. The `stub=true` count was unchanged
(~5000 in `gestion_riesgos`), but the actionable noise collapses to
`repointDecision='declined-not-found'` (~451 in `gestion_riesgos`).

Issue #245 then removed the runtime-object stubs at the source — see
"What #245 changed" below — so the `stub=true` count itself drops sharply
while `declined-not-found` stays exactly where it was.

## Consumer-side filter (the actual contract)

```sql
-- What consumer-side lints SHOULD use, per round-5 shipped semantics:
SELECT json_extract(e.metadata, '$.receiverType') || '.' ||
       json_extract(e.metadata, '$.member') AS qualified,
       COUNT(*) AS n
FROM edges e
WHERE e.kind='calls'
  AND json_extract(e.metadata, '$.synthesizedBy')='vba-name-resolution'
  AND json_extract(e.metadata, '$.repointDecision') = 'declined-not-found'
GROUP BY qualified ORDER BY n DESC;
-- This is the genuine-missing-callees list. Round-5's "<500 total" gate is
-- replaced by "<500 declined-not-found" (per the actual round-5 final design).
```

```sql
-- Diagnostic / verification query — should show a clean distribution:
SELECT json_extract(e.metadata, '$.repointDecision') AS decision,
       json_extract(e.metadata, '$.stub') AS stub,
       COUNT(*) AS n
FROM edges e
WHERE e.kind='calls'
  AND json_extract(e.metadata, '$.synthesizedBy')='vba-name-resolution'
GROUP BY decision, stub ORDER BY n DESC;
-- Expected shape on a healthy codebase (post-#245):
--   reponted-to-real, false → largest bucket; previously-stubbed, now linked
--   declined-not-found, true → small bucket (genuine typos / missing callees)
--   declined-ambiguous, true → rare; investigate manually
--   declined-runtime, true  → ~0; the extractor no longer creates these stubs
```

## What #245 changed

`declined-runtime` used to be the largest bucket in the query above. It was
annotating stub nodes that should never have existed: on the reference corpus
6,132 of 9,972 VBA `function` nodes were synthetic call stubs, and about 70% of
those named a VBA/Access runtime member. Edge consumers could filter them on
`repointDecision`; symbol search, `codegraph query` results and node counts
could not.

The extractor now refuses to synthesize the node. Two allowlists that had
drifted apart — the resolver's `RUNTIME_OBJECTS` and the extractor's
`RUNTIME_RECEIVER_BLACKLIST` — are derived from one literal in
`src/extraction/vba/runtime-objects.ts`, and the qualified-call sweep gates
node creation on `isRuntimeObject(receiverType)`.

Consumer impact:

- **`unresolved_refs` is unchanged.** This removed nodes, not references, so
  every `reference_kind` count and the `declined-runtime` classification of
  unresolved references (see
  [`docs/vba-reference-kinds.md`](vba-reference-kinds.md)) is byte-identical.
- **`edges.kind='calls'` loses exactly one edge per removed stub** — 8,322 →
  3,787 on the reference corpus. Any consumer SQL joining stub call edges to
  runtime members returns fewer rows; the `declined-not-found` query above,
  which is the actual contract, is unaffected.
- **Real procedure counts are unchanged** (3,840 on the reference corpus).
  The discriminator is `metadata.stub === true`; declarations never carry it.
- **A user class named after a runtime object is no longer call-linked
  through that name.** VBA extraction is per-file, so it cannot know a
  `DAO.cls` exists elsewhere when it parses `DAO.Execute`; the stub the
  resolver's shadow bypass (FR-2.1) needed is gone before the resolver runs.
  This is the trade the extractor had already made for the fourteen
  `RUNTIME_RECEIVER_BLACKLIST` names — a class named `DoCmd` or `Application`
  was never linked either — now extended to the rest of the canonical set.

## Migration from the round-5 prompt

| Round-5 prompt expected | Round-5 actual (shipped) | Consumer SQL change |
|---|---|---|
| `stub=true AND stub_true_count < 500` | `repointDecision='declined-not-found'` count <500 | Replace `WHERE stub=true` with `WHERE repointDecision='declined-not-found'`. |
| `declared_targets = 0` | `declared_targets = ~0`; runtime noise is gone from `nodes` entirely (#245) | OK; nothing to filter. |
| `stub=true` filter for "missing callee" | `declined-not-found` filter | Switch the bucket. |

## Reference

- Issue #110 — original specification
- PR #113 — shipped implementation
- src/resolution/index.ts:resolveVbaCallStubs — producer side
- src/extraction/vba/runtime-objects.ts — runtime-object allowlist source (re-exported by src/resolution/vba-runtime-objects.ts)
- src/extraction/vba/calls.ts — the stub-node gate (#245)
- server-instructions.ts (MCP) — AI-facing summary (same values, terser)

## See also

- [`docs/vba-reference-kinds.md`](vba-reference-kinds.md) — the full 7-kind `reference_kind` taxonomy that gates whether a row becomes `declined-runtime` here.