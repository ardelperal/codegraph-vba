# Design: VBA Event Tracer

This design document outlines the technical approach to implement the `vba-event-tracer` capability, including event signature extraction, query-time dynamic handler resolution, and the tracer skill integration.

## Technical Approach

We will modify the VBA extraction parser to capture and persist the exact event declaration signatures. At query time, we resolve custom `WithEvents` event handlers dynamically by walking the graph's `subscribes-event` edges and matching subroutine naming patterns inside subscriber modules. The trace capabilities are shipped as a custom agent skill (`vba-event-tracer`) that queries the index via `codegraph_explore` and formats the results into the required JSON schema.

## Architecture Decisions

### Decision: Parse and populate the event signature during extraction
- **Choice**: Extract the trimmed source declaration line in `src/extraction/vba-extractor.ts` and set it on the `signature` field of `event` nodes during the AST extraction pass.
- **Rationale**: The AST parser has direct, synchronous access to the raw source lines. Populating it at extraction time avoids the overhead of reading file contents at query time. It ensures that the database is self-contained for queries.

### Decision: Resolve event handlers dynamically at query time
- **Choice**: Resolve event handlers on-the-fly rather than persisting explicit `event-handler` edges in the database at indexing time.
- **Rationale**: Saving transient handler bindings at extraction time is error-prone and pollutes the database with stub dependencies. If a user refactors or renames a handler or a variable, static edges quickly drift out of sync. Dynamic resolution leverages existing `subscribes-event` edges and matches handler procedures of the form `<VariableName>_<EventName>` dynamically, ensuring high accuracy and low database maintenance.

### Decision: Implement the tracer capability as an agent skill
- **Choice**: Write a custom agent skill at `.agents/skills/vba-event-tracer/SKILL.md`. The skill describes instructions for the agent to query the graph using `codegraph_explore`, perform the resolution logic, and return the structured JSON schema.
- **Rationale**: The spec lists `codegraph_explore` as the required tool. Using an agent skill avoids polluting the core MCP tool registry with language-specific tracer tools, keeps the indexer codebase clean, and leverages the agent's capability to process the retrieved subgraph.

## Detailed Design & Data Flow

### 1. Extractor Changes (`src/extraction/vba-extractor.ts`)
In `sweepEventsTypesAndDeclares`, update the `eventNode` creation block:
```typescript
const eventNode: Node = {
  id: eventId,
  kind: 'event',
  name,
  qualifiedName: this.classNamePrefix ? `${this.classNamePrefix}.${name}` : name,
  filePath: this.filePath,
  language: 'vba',
  startLine: lineNum,
  endLine: lineNum,
  startColumn: 0,
  endColumn: line.length,
  visibility,
  signature: line.trim(), // Populate signature with trimmed source line
  updatedAt: Date.now(),
};
```

### 2. Resolution Logic & Ambiguity Handling
When a trace request for an event (e.g., `DataChanged`) is initiated:
1. **Ambiguity Check**: Query `event` nodes matching the name. If multiple event nodes exist (e.g., `PublisherA.DataChanged` and `PublisherB.DataChanged`) and the query is unqualified, return `EVENT_AMBIGUOUS` along with candidate qualifiers.
2. **Declaration Site**: Retrieve the single matching `event` node's `filePath`, `startLine`, and `signature`.
3. **Raise Sites**: Find incoming `raises-event` edges targeting the event node. Trace back to the calling procedures (`module`, `line`, and parent subroutine context).
4. **Handler Resolution**:
   - Query all subscriber modules targeting the event's parent class module via `subscribes-event` edges.
   - For each subscriber module, retrieve the variable name from `edge.metadata.variableName`.
   - Scan the procedures inside the subscriber module for subroutines named `<VariableName>_<EventName>` (case-insensitive).
   - Map each matching subroutine to a handler entry.
5. **Circular Dependencies**: To prevent infinite recursion during traversal, keep a `Set` of visited `(module, event)` pairs. Stop traversing a branch if it has already been visited.
6. **Warnings**: If the event has no incoming `raises-event` edges, add `NO_RAISERS` to the warnings array.

### 3. Skill Integration (`.agents/skills/vba-event-tracer/SKILL.md`)
The skill will guide the agent to perform the following steps:
1. Call `codegraph_explore` with the event name.
2. Execute the resolution logic on the returned nodes/edges.
3. Output the exact JSON schema:
```json
{
  "event_declarations": [{"module": "PedidoPublisher.cls", "line": 2, "signature": "Public Event PedidoGuardado(ByVal id As Long)"}],
  "raise_sites": [{"module": "PedidoPublisher.cls", "line": 15, "context": "Public Sub Guardar()"}],
  "handlers": [{"form": "PedidoSubscriber.cls", "handler": "m_Publisher_PedidoGuardado", "via": "m_Publisher"}],
  "warnings": []
}
```

## Testing Strategy

- **Unit Tests (`__tests__/extraction-vba-roadmap-25-26.test.ts`)**: Add test cases to assert that the `signature` column on the extracted `event` nodes matches the trimmed declaration line.
- **E2E Skill Tests**: Verify that querying `vba-event-tracer` correctly resolves handlers, detects ambiguity, rejects circular loops, and completes in <100ms.
