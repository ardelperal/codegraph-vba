# Proposal: VBA Event Tracer

## Intent
Enhance the VBA extraction parser and introduce a `vba-event-tracer` skill to parse event signatures and dynamically bind event handlers for custom `WithEvents` variables.

## Scope
### In Scope
- Modify `src/extraction/vba-extractor.ts` to populate the `signature` column on `event` nodes.
- Resolve custom class `WithEvents` event handlers at query time using graph-based query/execution.
- Implement the `vba-event-tracer` skill to trace event declarations, raise sites, and handlers.
- Update Vitest suites for extraction validation.

### Out of Scope
- Binding handlers at extraction time (handled dynamically at query time).
- Supporting non-standard/untyped `WithEvents` bindings.

## Capabilities
### New Capabilities
- `vba-event-tracer`: Ability to query and list event declarations, raise sites, and custom `WithEvents` event handlers.

## Approach
Based on exploration findings:
1. **Event Signatures**: Store the exact signature line on `event` nodes in `vba-extractor.ts` under the `signature` column.
2. **Graph-Based SQLite Queries (Approach A)**: Resolve event handlers at query time by matching `<VariableName>_<EventName>` procedures within modules holding `subscribes-event` edges. This keeps query execution under 100ms.
3. **Trace Skill**: Implement `vba-event-tracer` querying logic to traverse events, raise sites (`raises-event`), and handlers.

## Affected Areas
| Area | Impact |
|---|---|
| `src/extraction/vba-extractor.ts` | Modified |
| `__tests__/extraction-vba.test.ts` | Modified |
| `openspec/changes/vba-event-tracer/proposal.md` | Created |

## Risks
| Risk | Likelihood | Mitigation |
|---|---|---|
| Mismatched/unresolved unqualified event names | Med | Return `EVENT_AMBIGUOUS` with candidate details |
| Extractor signature mismatch | Low | Rely on clean source line extraction during AST parsing |

## Rollback Plan
Revert code changes in git.

## Success Criteria
- [ ] `event` nodes populate the `signature` column in the database.
- [ ] Graph-based queries resolve `WithEvents` handlers correctly under 100ms.
- [ ] `vba-event-tracer` skill successfully traces events, raises, and handlers.
- [ ] Vitest test suite passes.
