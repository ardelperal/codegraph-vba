# Proposal: VBA SQL Impact Analysis (vba-sql-impact)

## Intent
Trace Microsoft Access saved queries (QueryDefs) to their VBA caller modules, extract form bindings (`RecordSource`/`RowSource` properties), and analyze target tables/columns touched including SQL alias resolution.

## Scope

### In Scope
- Create a new agent skill (`.agents/skills/vba-sql-impact/SKILL.md`) to dynamically resolve SQL impact at query time.
- Implement dynamic SQL alias resolution (e.g. mapping `alias.column` back to its original table) and column touch analysis by reading workspace files on demand.
- Parse `RecordSource` and `RowSource` properties directly from form configuration files (`.form.txt`/`.report.txt`) on the fly.
- Integrate call-graph trace to find VBA callers invoking saved queries via recordsets or query references.

### Out of Scope
- Modifying the core indexer SQLite schema or adding new database tables/edges.
- Adding third-party SQL parser dependencies to `package.json`.
- Full static parsing of complex non-standard Access SQL dialects in the core codebase.

## Capabilities

### New Capabilities
- `vba-sql-impact`: A dynamic agent skill that maps saved queries, form control bindings, and resolves aliases/columns touched on demand.

### Modified Capabilities
- None.

## Approach
Implement Approach A (Dynamic Agent Skill). Rather than modifying the indexer backend, the skill orchestrates execution at query time:
1. Query SQLite using `codegraph_explore` to retrieve query name references, form modules, and basic call structures.
2. Read raw query definition `.sql` files and form layout files (`.form.txt`/`.report.txt`) directly from the workspace on demand.
3. Perform regex-based and semantic parsing locally to extract bindings (`RecordSource`/`RowSource`) and resolve table aliases to columns.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `.agents/skills/vba-sql-impact/SKILL.md` | Created | New dynamic agent skill containing the impact analysis logic |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Workspace file reads scale poorly | Low | Scope file reads to specific form/query files resolved from initial SQLite queries |
| Access SQL dialect parses incorrectly | Med | Rely on agent reasoning and regex fallbacks within the skill |

## Rollback Plan
Revert changes in git (remove the created skill file). Since no database schema or indexer code changes are made, rollback is completely clean and instantaneous.

## Dependencies
- None.

## Success Criteria
- [ ] New `vba-sql-impact` skill successfully traces saved queries to callers.
- [ ] Skill extracts `RecordSource` and `RowSource` form bindings.
- [ ] Skill resolves tables/columns touched including alias resolution.
- [ ] Existing vitest test suite passes.
