## Exploration: vba-event-tracer

### Current State
VBA event modeling was introduced in `codegraph-vba` via roadmap item #26. Specifically:
- **Event Declarations**: Extracted in [vba-extractor.ts](file:///C:/00repos/codigo/00_codegraph_main/src/extraction/vba-extractor.ts#L573-L598) using `EVENT_DECL_RE`. It creates a node of kind `event` with a qualified name prefixed by the containing class (e.g. `PedidoPublisher.PedidoGuardado`). A `contains` edge is added from the module/class to the event node. However, the exact signature is not currently saved in the `signature` column of the event node.
- **Raise Event Sites**: Scanned in [vba-extractor.ts](file:///C:/00repos/codigo/00_codegraph_main/src/extraction/vba-extractor.ts#L1412-L1429) using `RAISE_EVENT_RE`. It produces a `raises-event` edge from the calling procedure's node to the event node. The edge contains metadata such as the event name.
- **WithEvents Subscriptions**: Checked in [vba-extractor.ts](file:///C:/00repos/codigo/00_codegraph_main/src/extraction/vba-extractor.ts#L810-L847) using `WITHEVENTS_RE`. It produces a `references` edge (synthesized by `vba-withevents`) and a `subscribes-event` edge from the subscriber class to the target type, carrying the subscription's `variableName` in the edge metadata.
- **Event Handlers**: The extractor currently has a heuristic to bind Access form control event handlers (e.g. `ComandoAltaPM_Click`) to `form-instance-control` nodes using `event-handler` edges in [vba-extractor.ts](file:///C:/00repos/codigo/00_codegraph_main/src/extraction/vba-extractor.ts#L443-L480). However, custom class `WithEvents` event handlers (e.g. `Private Sub m_Form_SomethingChanged()` for `Private WithEvents m_Form As Form_Pedido`) are *not* explicitly linked to the event node or target declaration in the database at extraction time.

### Affected Areas
- `src/extraction/vba-extractor.ts` — The extractor needs to store the event signature (the declaring source line) in the `signature` column of the `event` node. Additionally, we need to decide if `WithEvents` event handlers themselves should be annotated during extraction or resolved at query time.
- `~/.config/opencode/skills/vba-event-tracer/SKILL.md` — The location of the new skill that handles the queries and format mapping.

### Approaches
1. **Approach A: Graph-Based SQLite Queries (Recommended)**
   - Description: Query the existing SQLite database via the codegraph engine. First, find the matching `event` node by its name (or qualified name). Then, trace outgoing `raises-event` edges to find raise sites, and query `subscribes-event` edges to identify subscriber modules. For each subscriber, check for procedures named `<VariableName>_<EventName>` inside the module to find handlers.
   - Pros:
     - Extremely fast (<100ms), meeting the `<2s` constraint.
     - Accurate: correctly maps typed `WithEvents` instances across files.
     - Reuses the existing AST/extraction data without duplicate parsing or slow file-system scans.
   - Cons:
     - Requires the `signature` field to be populated on event nodes (or fallback to reading source files at the declaration line).
   - Effort: Low

2. **Approach B: File-Based Regex Scan**
   - Description: Bypass the SQLite database entirely. Use regex-based grep tools (or Node.js `fs` file reads) to scan all source files for event declarations, `RaiseEvent` sites, `WithEvents` variables, and handler-pattern subroutines.
   - Pros:
     - Independent of the codegraph DB indexing status.
   - Cons:
     - Very slow: reading and regex-matching 300+ files on every query can easily exceed the 2-second limit.
     - Inaccurate: cannot resolve type-based subscriptions. If multiple variables share a name or if multiple classes define the same event name, simple regex will misattribute handlers.
   - Effort: High

### Recommendation
We recommend **Approach A (Graph-Based SQLite Queries)**. The codegraph database already indexes the structural relationships (events, raise sites, WithEvents subscriptions, files) and handles indexing/comment-stripping perfectly. Resolving handlers by combining `subscribes-event` edge metadata and subscriber procedure names is robust, simple, and takes advantage of the existing index, ensuring sub-second performance. We can easily fetch event signatures by reading the line directly from the file snapshot if needed, or by minor enhancements to `VbaExtractor`.

### Risks
- **Extractor Signature Gaps**: If the database doesn't store the exact event signature, the skill must read the source code file at the declaration's `startLine`. If the file is modified or out of sync, the line contents might mismatch.
- **Ambiguity in Unqualified Event Names**: If multiple classes define an event with the same name, querying without a class prefix is ambiguous. We must ensure the tool returns `EVENT_AMBIGUOUS` with candidate choices.

### Ready for Proposal
Yes
