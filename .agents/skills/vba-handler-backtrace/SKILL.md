---
name: vba-handler-backtrace
description: Traces VBA control event handlers to construct call graph trees, resolves multiline SQL queries, and extracts custom UDT parameters.
---

# Skill: VBA Handler Backtrace

This skill enables the agent to trace a VBA control event handler back to the methods and subroutines it calls, parsing custom UDT signatures, reconstructing multiline SQL string concatenations, and reporting results in a structured JSON tree format.

## Triggering Context
This skill is triggered when tracing VBA event flows, diagnosing execution flow from form controls (e.g., buttons, textboxes) to SQL query executions, or resolving method calls from event handler signatures.

## Step-by-Step Execution Guide

### Step 1: Trace Control Event Handlers (Graph Traversal)
1. Resolve the starting node FIRST, scoped to its layout. Control nodes are
   `form-instance-control`, forms and reports are `form-layout` /
   `report-layout`. The same control name (`btnSave`) exists on several forms,
   so match on the layout file too — never on the name alone.
2. Use the `traverseGraph` helper with that node id. It follows the two
   directions the extractor actually stores:
   - an `event-handler` edge is stored HANDLER -> control/layout, so the
     handler is reached by following it BACKWARDS from the control;
   - a call is followed forwards, through `calls` edges and through the
     `references` edge the resolver stores for VBA's statement-form Sub call
     (a bare `SaveRecord` on its own line is ambiguous at parse time, so it is
     resolved as a reference to the procedure, not as `calls`).
   Containment, typing and data references (tables, saved queries) are NOT
   call steps and never appear as children.
3. Handle the result envelope:
   - Node attributes `id`, `name`, and `kind` must be retrieved.
   - Set maximum search depth to prevent excessive execution (default is `10` unless custom is specified); an exceeded depth reports `MAX_DEPTH_EXCEEDED`.
   - An unknown start node reports `tree: null` and `START_NODE_NOT_FOUND` — that is a lookup miss, not proof the control has no handler.
   - Trace circular dependencies using the `visited` node set tracking logic. If a cycle is detected, flag `cycle_detected: true` and terminate branch expansion.
4. The trace is STATIC evidence from exported source. It does not prove the
   handler ran, and it does not say whether the `.accdb` binary matches the
   export.

### Step 2: Extract Signature Custom UDT Parameters
1. Use the `parseSignatureParams` helper to parse subroutine or function signatures.
2. The helper extracts variables and their types using the regex `/(?:ByVal|ByRef)?\s*(\w+)\s+As\s+(\w+)/gi`.
3. Filter out VBA primitive types (case-insensitive):
   - `Long`, `Integer`, `String`, `Boolean`, `Double`, `Single`, `Byte`, `Currency`, `Date`, `Variant`, `Object`, `LongLong`, `LongPtr`, `Decimal`.
4. Keep only custom/user-defined type parameters (e.g. custom classes, structs).

### Step 3: Reconstruct Multiline SQL Statements
1. Locate files/lines containing multiline SQL query string concatenations (using VBA `_` and `&`).
2. Pass the sequence of line strings to `reconstructSQL`.
3. The helper extracts and cleans string literals inside double quotes `"`, stripping escapes.
4. Limit the resulting SQL query string to a maximum of `200` characters to maintain compact context logs.

## Output JSON Formatting

Format the final trace tree and extraction metadata into the following schema:

```json
{
  "tree": {
    "id": "Form_Orders.form.txt::btnSave",
    "name": "btnSave",
    "kind": "form-instance-control",
    "children": [
      {
        "id": "Form_Orders.cls::btnSave_Click",
        "name": "btnSave_Click",
        "kind": "function",
        "children": [
          {
            "id": "SaveRecord",
            "name": "SaveRecord",
            "kind": "function",
            "children": []
          }
        ]
      }
    ]
  },
  "cycle_detected": false,
  "warnings": [],
  "extracted_parameters": [
    {
      "name": "ctx",
      "type": "OrderContext"
    }
  ],
  "reconstructed_sql": "INSERT INTO Log (Msg) VALUES ('Order Processed')"
}
```
