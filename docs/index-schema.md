# codegraph-vba index — SQL schema reference

This document is the published SQL contract for the `.codegraph-vba/codegraph.db` index. It is **auto-generated** by `scripts/dump-index-schema.ts` from the canonical schema in `src/db/schema.sql` (issue #200, Deliverable 2).

Regenerate with `npm run schema:dump` after any schema change.

## How to read this doc

Each table is tagged with a **Stability** line in its header. Per-column stability is in the third column of every table.

- **`stable`** — part of the public contract. External tools may rely on these column names and types within the v1.x line. Removing or retyping is a breaking change and requires a major version bump.
- **`implementation detail`** — internal to the indexer. May be added, removed, or retyped between minor releases. External tools MUST NOT depend on these.

## Canonical (public) tables

The three tables below are the contract surface for external SQL consumers (see `docs/external-integration.md` §3 for the pattern, and [Dysflow#1015](https://github.com/DysTelefonica/dysflow/issues/1015) for the cross-tool motivation). The column list mirrors the row keys `CodeGraph.searchNodes`, `getCallers`, `getCallees`, and `getImpactRadius` expose on the JSON surface.

## edges

**Stability: stable.** Part of the public contract — columns and types here are pinned across the v1.x line. Adding columns is allowed; removing or retyping is a breaking change.

| Column | Type | Stability |
|---|---|---|
| `id` | INTEGER NULL PRIMARY KEY (part 1) | stable |
| `source` | TEXT NOT NULL | stable |
| `target` | TEXT NOT NULL | stable |
| `kind` | TEXT NOT NULL | stable |
| `metadata` | TEXT NULL | stable |
| `line` | INTEGER NULL | implementation detail |
| `col` | INTEGER NULL | implementation detail |
| `provenance` | TEXT NULL DEFAULT "NULL" | stable |

### Indexes

| Name | Unique | Origin | Partial |
|---|---|---|---|
| `idx_edges_provenance` | no | c | no |
| `idx_edges_identity` | yes | c | no |
| `idx_edges_target_kind` | no | c | no |
| `idx_edges_source_kind` | no | c | no |
| `idx_edges_kind` | no | c | no |

## nodes

**Stability: stable.** Part of the public contract — columns and types here are pinned across the v1.x line. Adding columns is allowed; removing or retyping is a breaking change.

| Column | Type | Stability |
|---|---|---|
| `id` | TEXT NULL PRIMARY KEY (part 1) | stable |
| `kind` | TEXT NOT NULL | stable |
| `name` | TEXT NOT NULL | stable |
| `qualified_name` | TEXT NOT NULL | stable |
| `file_path` | TEXT NOT NULL | stable |
| `language` | TEXT NOT NULL | stable |
| `start_line` | INTEGER NOT NULL | stable |
| `end_line` | INTEGER NOT NULL | stable |
| `start_column` | INTEGER NOT NULL | implementation detail |
| `end_column` | INTEGER NOT NULL | implementation detail |
| `docstring` | TEXT NULL | implementation detail |
| `signature` | TEXT NULL | implementation detail |
| `visibility` | TEXT NULL | implementation detail |
| `is_exported` | INTEGER NULL DEFAULT "0" | implementation detail |
| `is_async` | INTEGER NULL DEFAULT "0" | implementation detail |
| `is_static` | INTEGER NULL DEFAULT "0" | implementation detail |
| `is_abstract` | INTEGER NULL DEFAULT "0" | implementation detail |
| `decorators` | TEXT NULL | implementation detail |
| `type_parameters` | TEXT NULL | implementation detail |
| `return_type` | TEXT NULL | implementation detail |
| `metadata` | TEXT NULL | stable |
| `updated_at` | INTEGER NOT NULL | implementation detail |

### Indexes

| Name | Unique | Origin | Partial |
|---|---|---|---|
| `idx_nodes_lower_name` | no | c | no |
| `idx_nodes_file_line` | no | c | no |
| `idx_nodes_language` | no | c | no |
| `idx_nodes_file_path` | no | c | no |
| `idx_nodes_qualified_name` | no | c | no |
| `idx_nodes_name` | no | c | no |
| `idx_nodes_kind` | no | c | no |
| `sqlite_autoindex_nodes_1` | yes | pk | no |

## unresolved_refs

**Stability: stable.** Part of the public contract — columns and types here are pinned across the v1.x line. Adding columns is allowed; removing or retyping is a breaking change.

| Column | Type | Stability |
|---|---|---|
| `id` | INTEGER NULL PRIMARY KEY (part 1) | stable |
| `from_node_id` | TEXT NOT NULL | stable |
| `reference_name` | TEXT NOT NULL | stable |
| `reference_kind` | TEXT NOT NULL | stable |
| `line` | INTEGER NOT NULL | stable |
| `col` | INTEGER NOT NULL | stable |
| `candidates` | TEXT NULL | implementation detail |
| `file_path` | TEXT NOT NULL DEFAULT "''" | implementation detail |
| `language` | TEXT NOT NULL DEFAULT "'unknown'" | implementation detail |
| `status` | TEXT NOT NULL DEFAULT "'pending'" | stable |
| `name_tail` | TEXT NOT NULL DEFAULT "''" | implementation detail |
| `metadata` | TEXT NULL | stable |

### Indexes

| Name | Unique | Origin | Partial |
|---|---|---|---|
| `idx_unresolved_failed_tail` | no | c | yes |
| `idx_unresolved_status` | no | c | no |
| `idx_unresolved_from_name` | no | c | no |
| `idx_unresolved_file_path` | no | c | no |
| `idx_unresolved_name` | no | c | no |
| `idx_unresolved_from_node` | no | c | no |

## Infrastructure tables

Internal bookkeeping for the indexer — versions, project metadata, the prompt-hook segment vocabulary. Documented here for completeness; no external tool should query them.

## name_segment_vocab

**Stability: implementation detail.** Internal to the indexer; may change between minor releases without notice.

| Column | Type | Stability |
|---|---|---|
| `segment` | TEXT NOT NULL PRIMARY KEY (part 1) | implementation detail |
| `name` | TEXT NOT NULL PRIMARY KEY (part 2) | implementation detail |

### Indexes

| Name | Unique | Origin | Partial |
|---|---|---|---|
| `sqlite_autoindex_name_segment_vocab_1` | yes | pk | no |

## project_metadata

**Stability: implementation detail.** Internal to the indexer; may change between minor releases without notice.

| Column | Type | Stability |
|---|---|---|
| `key` | TEXT NULL PRIMARY KEY (part 1) | implementation detail |
| `value` | TEXT NOT NULL | implementation detail |
| `updated_at` | INTEGER NOT NULL | implementation detail |

### Indexes

| Name | Unique | Origin | Partial |
|---|---|---|---|
| `sqlite_autoindex_project_metadata_1` | yes | pk | no |

## schema_versions

**Stability: implementation detail.** Internal to the indexer; may change between minor releases without notice.

| Column | Type | Stability |
|---|---|---|
| `version` | INTEGER NULL PRIMARY KEY (part 1) | implementation detail |
| `applied_at` | INTEGER NOT NULL | implementation detail |
| `description` | TEXT NULL | implementation detail |

## Other tables

## files

**Stability: implementation detail.** Internal to the indexer; may change between minor releases without notice.

| Column | Type | Stability |
|---|---|---|
| `path` | TEXT NULL PRIMARY KEY (part 1) | implementation detail |
| `content_hash` | TEXT NOT NULL | implementation detail |
| `language` | TEXT NOT NULL | implementation detail |
| `size` | INTEGER NOT NULL | implementation detail |
| `modified_at` | INTEGER NOT NULL | implementation detail |
| `indexed_at` | INTEGER NOT NULL | implementation detail |
| `node_count` | INTEGER NULL DEFAULT "0" | implementation detail |
| `errors` | TEXT NULL | implementation detail |

### Indexes

| Name | Unique | Origin | Partial |
|---|---|---|---|
| `idx_files_modified_at` | no | c | no |
| `idx_files_language` | no | c | no |
| `sqlite_autoindex_files_1` | yes | pk | no |

## nodes_fts

**Stability: implementation detail.** Auto-generated by SQLite for the `nodes_fts` virtual table. External tools MUST NOT query this table directly.

| Column | Type |
|---|---|
| `id` | ANY |
| `name` | ANY |
| `qualified_name` | ANY |
| `docstring` | ANY |
| `signature` | ANY |

## nodes_fts_config

**Stability: implementation detail.** Auto-generated by SQLite for the `nodes_fts` virtual table. External tools MUST NOT query this table directly.

| Column | Type |
|---|---|
| `k` | ANY |
| `v` | ANY |

## nodes_fts_data

**Stability: implementation detail.** Auto-generated by SQLite for the `nodes_fts` virtual table. External tools MUST NOT query this table directly.

| Column | Type |
|---|---|
| `id` | INTEGER |
| `block` | BLOB |

## nodes_fts_docsize

**Stability: implementation detail.** Auto-generated by SQLite for the `nodes_fts` virtual table. External tools MUST NOT query this table directly.

| Column | Type |
|---|---|
| `id` | INTEGER |
| `sz` | BLOB |

## nodes_fts_idx

**Stability: implementation detail.** Auto-generated by SQLite for the `nodes_fts` virtual table. External tools MUST NOT query this table directly.

| Column | Type |
|---|---|
| `segid` | ANY |
| `term` | ANY |
| `pgno` | ANY |

## Notes for consumers

- **`metadata` columns are JSON.** `nodes.metadata`, `edges.metadata`, `unresolved_refs.metadata`, and `unresolved_refs.candidates` carry a JSON string. Use SQLite's `json_extract()` to read individual keys; the keys inside the JSON object are NOT covered by the stable contract — see the per-feature doc (e.g. `docs/vba-stub-repoint-decision.md` for `edges.metadata.repointDecision`).
- **`status` on `unresolved_refs`** is the lifecycle field (`pending` → resolved-and-deleted, or `failed` / `declined-runtime` / `declined-ambiguous` / `declined-not-found`). See `docs/vba-stub-repoint-decision.md` for the resolution taxonomy and `docs/vba-reference-kinds.md` for the `reference_kind` taxonomy.
- **FTS5 queries** go through the `nodes_fts` virtual table; the shadow tables (`nodes_fts_config`, `nodes_fts_data`, `nodes_fts_docsize`, `nodes_fts_idx`) are managed by SQLite. External tools should treat the `nodes_fts_*` shadow tables as off-limits.
- **`provenance`** on `edges` is `'static'` (default, null in old rows) or `'heuristic'` for synthesized edges. Stable: the column is part of the contract, but the value space may grow with new synthesizer types — see `docs/design/callback-edge-synthesis.md`.

## Cross-references

- [`docs/external-integration.md`](external-integration.md) — Pattern A (subprocess spawn) and the JSON contracts.
- [`docs/vba-stub-repoint-decision.md`](vba-stub-repoint-decision.md) — `repointDecision` taxonomy for `edges.metadata`.
- [`docs/vba-reference-kinds.md`](vba-reference-kinds.md) — `reference_kind` taxonomy for `unresolved_refs`.
- Sister issue [Dysflow#1015](https://github.com/DysTelefonica/dysflow/issues/1015) — this doc unblocks its Path 1 fix.
- Issue #200 — this issue.
