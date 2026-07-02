# Design: VBA Handler Backtrace Custom Agent Skill

## 1. Architecture Decision & Rationale

We will implement the `vba-handler-backtrace` capability as an **agent-side custom skill** located at `.agents/skills/vba-handler-backtrace/SKILL.md`.

### Rationale
- **Decoupled Core:** Avoids bloating the core indexer codebase and keeps the SQLite database schema clean.
- **Lightweight Fallback Reads:** The agent can dynamically read local source files along the trace path (typically only 3–10 files) to extract runtime SQL strings and UDT parameters. This prevents having to store full SQL concatenations and parameter lists in the DB, avoiding expensive schema migrations.
- **Runtime Flexibility:** Traversal depths, cycle rules, and extraction patterns can be quickly modified directly in the skill without recompiling or redeploying the core binaries.

---

## 2. Detailed Traversal Algorithm

The backtrace capability is executed dynamically via the agent skill following these steps:

### Step 1: Starting Node Lookup
- Locate the starting `form-instance-control` node by name (e.g., `cmdConfirm`) in the database.
- Retrieve the corresponding event-handler procedure (e.g., `cmdConfirm_Click` in `Form_Orders.cls`) by querying for incoming `event-handler` edges to the control node matching the specified `eventName` (e.g., `"Click"`).

### Step 2: Graph Retrieval
- Query the graph database using `codegraph_explore` to load all relevant nodes, plus their outgoing `calls` and `references` edges.
- Extract the metadata, `file_path`, and `start_line` for all candidates.

### Step 3: Recursive Traversal
Traverse downstream call paths recursively starting from the handler procedure node, passing:
1. `currentNode`: The node currently being traversed.
2. `currentDepth`: The current level of recursion (starts at 0).
3. `visited`: An ordered list or set of node IDs in the current path.

For each traversed procedure node:
- **Cycle Detection:** If `currentNode.id` is in `visited`:
  - Set `cycle_detected: true` on the current trace node metadata.
  - Terminate traversal of this branch to prevent infinite loops.
- **Depth Capping:** If `currentDepth >= maxDepth` (default 5):
  - Append `"MAX_DEPTH_EXCEEDED"` to the top-level `warnings` array.
  - Terminate traversal of this branch.
- **UDT Parameter Extraction:**
  - Read the procedure signature from the source file starting at `currentNode.startLine`.
  - Parse the parameters using a regex matching pattern:
    `/(?:ByVal|ByRef)?\s*(\w+)\s+As\s+(\w+)/gi`
  - Identify and map custom UDT parameter types (filtering out primitive types like `String`, `Integer`, `Long`, `Boolean`, `Double`, `Variant`, `Object`). Add them to the node's `parameters` metadata.
- **DAO SQL Extraction:**
  - For outgoing `calls`/`references` targeting database operations (e.g., `.Execute` or `.OpenRecordset` calls on DAO variables or helper functions like `getdb()`), read the source file lines starting at the call site.
  - If the line ends with a line continuation character `_` or contains string concatenation `&`, parse and accumulate adjacent lines to reconstruct the full SQL query string.
  - Truncate the reconstructed string to a maximum of 200 characters and format it as `sql_hint`.
- **Recursive Step:** Add `currentNode.id` to `visited` and invoke traversal on each child node targeted by outgoing `calls` or `references` edges.

---

## 3. Output Schema

The custom skill will output the trace in a recursive JSON tree structure:

```json
{
  "trace": {
    "id": "Form_Orders.cls::cmdConfirm_Click",
    "name": "cmdConfirm_Click",
    "kind": "method",
    "file": "Form_Orders.cls",
    "line": 12,
    "parameters": [],
    "calls": [
      {
        "id": "OrderHelper.bas::ProcessOrder",
        "name": "ProcessOrder",
        "kind": "function",
        "file": "OrderHelper.bas",
        "line": 40,
        "parameters": [
          { "name": "ctx", "type": "OrderContext" }
        ],
        "calls": [
          {
            "id": "OrderHelper.bas::db.Execute@42",
            "name": "db.Execute",
            "kind": "database_operation",
            "file": "OrderHelper.bas",
            "line": 42,
            "sql_hint": "INSERT INTO Log (Msg) VALUES ('Order Processed')"
          }
        ]
      }
    ]
  },
  "warnings": []
}
```
