# Query VBA error handling

CodeGraph records each VBA procedure's error policy in the procedure node's
`metadata.errorPolicy` object. Run the four queries below against the generated
`.codegraph-vba/codegraph.db` SQLite database to find risky unprotected code,
open suppression scopes, user-facing error displays, and missing handlers.

> If `CODEGRAPH_DIR` is set, use that directory instead of `.codegraph-vba`.
> Build or refresh the index before querying so these fields reflect the current
> source. SQLite's JSON functions (`json_extract`) are required.

## Find unprotected procedures with risky outgoing work

This query narrows `protection = 'none'` to procedures with more than five
executable statements and an outgoing SQL-table, `DoCmd`, or filesystem-call
relationship. The exact statement count excludes blank/comment-only lines,
declarations, labels, and procedure boundary markers. Some SQL and `DoCmd`
references are initially owned by the module node, so the line-range join is
essential. Filesystem function calls are included by exact intrinsic name, and
reserved `Kill`, `Open ... For ... As ...`, and `Close` statement forms are
included only with their exact extractor provenance. This works both before and
after resolution because runtime references persist as `declined-runtime` rows;
unrelated, malformed, qualified, and user-procedure calls are not treated as risk.

```sql
WITH risky_relationships AS (
  SELECT
    e.source AS source_id,
    e.line,
    target.name AS risk_target,
    target.kind AS risk_target_kind,
    json_extract(e.metadata, '$.synthesizedBy') AS detected_by
  FROM edges AS e
  JOIN nodes AS target ON target.id = e.target
  WHERE json_extract(e.metadata, '$.synthesizedBy') = 'vba-sql-table'
    OR json_extract(e.metadata, '$.docmdVerb') IS NOT NULL
    OR json_extract(e.metadata, '$.synthesizedBy') IN (
      'vba-opens-form',
      'vba-opens-report',
      'vba-opens-query',
      'vba-closes-form',
      'vba-runs-macro',
      'vba-opens-table',
      'vba-applies-filter',
      'vba-copies-object',
      'vba-deletes-object',
      'vba-renames-object',
      'vba-selects-object',
      'vba-browses-to',
      'vba-outputs-to',
      'vba-sends-object',
      'vba-transfers-spreadsheet',
      'vba-transfers-text',
      'vba-transfers-database'
    )

  UNION ALL

  SELECT
    r.from_node_id AS source_id,
    r.line,
    r.reference_name AS risk_target,
    'runtime-call' AS risk_target_kind,
    json_extract(r.metadata, '$.synthesizedBy') AS detected_by
  FROM unresolved_refs AS r
      WHERE r.language = 'vba'
        AND (
          (
            r.reference_kind IN ('calls', 'unqualified-ident')
            AND json_extract(r.metadata, '$.synthesizedBy') IN (
              'vba-paren-call-unresolved',
              'vba-statement-call-unresolved'
            )
            AND lower(r.reference_name) IN (
              'chdir', 'chdrive', 'dir', 'filecopy', 'filelen',
              'getattr', 'mkdir', 'rmdir', 'setattr'
            )
          )
          OR
          (
            r.reference_kind = 'calls'
            AND json_extract(r.metadata, '$.synthesizedBy') =
              'vba-filesystem-statement'
            AND json_extract(r.metadata, '$.runtimeFamily') = 'filesystem'
            AND lower(r.reference_name) IN ('kill', 'open', 'close')
            AND json_extract(r.metadata, '$.operation') = lower(r.reference_name)
          )
        )
)
SELECT
  p.name AS procedure,
  p.qualified_name,
  p.file_path,
  p.start_line,
  p.end_line,
  json_extract(
    p.metadata,
    '$.errorPolicy.executableStatementCount'
  ) AS executable_statement_count,
  risk.risk_target,
  risk.risk_target_kind,
  risk.detected_by
FROM nodes AS p
JOIN risky_relationships AS risk
  ON risk.source_id = p.id
  OR (
    risk.line BETWEEN p.start_line AND p.end_line
    AND EXISTS (
      SELECT 1
      FROM nodes AS relationship_owner
      WHERE relationship_owner.id = risk.source_id
        AND relationship_owner.file_path = p.file_path
    )
  )
WHERE p.kind = 'function'
  AND p.language = 'vba'
  AND json_extract(p.metadata, '$.errorPolicy.protection') = 'none'
  AND json_extract(
    p.metadata,
    '$.errorPolicy.executableStatementCount'
  ) > 5
ORDER BY executable_statement_count DESC, p.file_path, p.start_line, risk_target;
```

A row is a review candidate, not proof of a defect. Exactly six executable
statements meet the threshold; blank and comment-only padding does not. An
unresolved VBA table or Access-object reference can retain the synthetic `class`
placeholder kind; `detected_by`, rather than `risk_target_kind`, is the reliable
reason it appears here. Filesystem statement rows deliberately report
`risk_target_kind = 'runtime-call'`; their exact `vba-filesystem-statement`
provenance is the precision gate.

**Snapshot validation.** The current reproducible corpus identifies 27 distinct
procedures: 21 in `00_EXPEDIENTES`, 0 in `00_GESTION_RIESGOS`, and 6 in
`HPS_SOLICITUDES`. This intentionally narrower result requires a concrete SQL-table,
`DoCmd`, or precision-gated filesystem relationship; unlike the historical broad
185-procedure source-body census and later 103-procedure marker probe, it is snapshot
evidence rather than a stable cardinality assertion. Rows may repeat a procedure for
multiple risk targets, so compare corpora by distinct procedure identity.

## Find `On Error Resume Next` scopes left open

`resumeNextOpen` is true when the last suppression event in a procedure opens a
scope and no later `On Error GoTo 0` or `On Error GoTo -1` closes it.

```sql
SELECT
  p.name AS procedure,
  p.qualified_name,
  p.file_path,
  p.start_line,
  p.end_line,
  json_extract(p.metadata, '$.errorPolicy.resumeNextOpen') AS resume_next_open
FROM nodes AS p
WHERE p.kind = 'function'
  AND p.language = 'vba'
  AND json_extract(p.metadata, '$.errorPolicy.resumeNextOpen') = 1
ORDER BY p.file_path, p.start_line;
```

## Find where errors surface to a person

`display` means the handler displays or prints the error. `mixed` is included
because that handler displays the error and also records or re-raises it. The
left join keeps display procedures that have no resolved caller; when a caller
is known, `reaches_display` names it.

```sql
SELECT
  display.name AS display_procedure,
  display.qualified_name AS display_qualified_name,
  display.file_path,
  display.start_line,
  json_extract(display.metadata, '$.errorPolicy.behavior') AS behavior,
  caller.name AS reaches_display,
  caller.qualified_name AS caller_qualified_name
FROM nodes AS display
LEFT JOIN edges AS incoming
  ON incoming.target = display.id
  AND incoming.kind = 'calls'
LEFT JOIN nodes AS caller ON caller.id = incoming.source
WHERE display.kind = 'function'
  AND display.language = 'vba'
  AND json_extract(display.metadata, '$.errorPolicy.behavior') IN ('display', 'mixed')
ORDER BY display.file_path, display.start_line, caller.qualified_name;
```

This is one call hop, not a whole-program error-propagation claim. Repeat the
incoming-call traversal, or use a recursive CTE, when you need every transitive
caller that can reach a display point.

## Find handler targets that do not exist

A non-null `danglingTarget` means the procedure contains `On Error GoTo <label>`
but does not define that label inside the same procedure.

```sql
SELECT
  p.name AS procedure,
  p.qualified_name,
  p.file_path,
  p.start_line,
  json_extract(p.metadata, '$.errorPolicy.danglingTarget') AS dangling_target
FROM nodes AS p
WHERE p.kind = 'function'
  AND p.language = 'vba'
  AND json_extract(p.metadata, '$.errorPolicy.danglingTarget') IS NOT NULL
ORDER BY p.file_path, p.start_line;
```

Fix or inspect these first: VBA cannot route the error to a label that is absent
from that procedure, even if another procedure defines a label with the same name.
