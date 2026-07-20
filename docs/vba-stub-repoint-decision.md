# VBA call-stub resolver — `repointDecision` semantics (shipped in v1.7.0)

This document is the canonical reference for the post-extraction stub resolver
introduced in [#110](https://github.com/ardelperal/codegraph-vba/issues/110) and
shipped in v1.7.0. It supersedes the original round-5 prompt acceptance
criterion (`stub_true_count < 500`) — see [#115](https://github.com/ardelperal/codegraph-vba/issues/115).

## TL;DR

| `metadata.repointDecision` | `metadata.stub` | Meaning | Consumer action |
|---|---|---|---|
| `reponted-to-real` | `false` | The synthetic function node was repointed to a real `nodes.id`. This was the original goal of round-4. | None — treat as a normal call edge. |
| `declined-runtime` | `true` | The receiver is a known runtime object (`DAO.*`, `fso.*`, `ListBox.*`, `Collection.*`, `err.*`, `VBA.*`, `Application.*`, `Screen.*`, `DoCmd.*`, `CurrentDb.*`, `Forms`, `Reports`, `Debug`, `Modules`, `References`, `CommandBars`, `SysCmd`, `CreateObject`, `GetObject`, `Fields`). It can never link to user code. | Filter OUT — noise. |
| `declined-ambiguous` | `true` | The qualified name resolves to multiple real `nodes.id` candidates (e.g. two classes named `DAO` in different modules). Repointing would be unsafe. | Investigate — flag for manual review. |
| `declined-not-found` | `true` | The qualified name doesn't match any real `nodes.id` — this is a genuinely missing callee. | Filter IN — actionable signal. |

## Why the criterion shifted

The original round-5 prompt expected the resolver to **repoint** every
`stub=true` edge whose target exists in `nodes`, reducing the global
`stub=true` count from ~5000 to <500. In practice that wasn't possible
because runtime-object stubs (`DAO.Execute`, `fso.GetFileName`, etc.) have
no user-code target at all — they're not "found", they're "impossible".

So the shipped fix (v1.7.0) instead **annotates** each stub with the
diagnostic `repointDecision` field. The `stub=true` count is unchanged
(~5000 in `gestion_riesgos`), but the actionable noise collapses to
`repointDecision='declined-not-found'` (~451 in `gestion_riesgos`).

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
-- Expected shape on a healthy codebase:
--   declined-runtime, true  → largest bucket (DAO/fso/etc.)
--   declined-not-found, true → small bucket (genuine typos)
--   reponted-to-real, false → previously-stubbed, now linked
--   declined-ambiguous, true → rare; investigate manually
```

## Migration from the round-5 prompt

| Round-5 prompt expected | Round-5 actual (shipped) | Consumer SQL change |
|---|---|---|
| `stub=true AND stub_true_count < 500` | `repointDecision='declined-not-found'` count <500 | Replace `WHERE stub=true` with `WHERE repointDecision='declined-not-found'`. |
| `declared_targets = 0` | `declared_targets = ~0` AND `declined-runtime` is large | OK; runtime noise is `declined-runtime`. |
| `stub=true` filter for "missing callee" | `declined-not-found` filter | Switch the bucket. |

## Reference

- Issue #110 — original specification
- PR #113 — shipped implementation
- src/resolution/index.ts:resolveVbaCallStubs — producer side
- src/extraction/vba/vba-runtime-objects.ts — runtime-object allowlist source
- server-instructions.ts (MCP) — AI-facing summary (same values, terser)

## See also

- [`docs/vba-reference-kinds.md`](vba-reference-kinds.md) — the full 7-kind `reference_kind` taxonomy that gates whether a row becomes `declined-runtime` here.