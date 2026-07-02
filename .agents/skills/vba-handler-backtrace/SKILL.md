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
1. Use the `traverseGraph` helper to trace call paths starting from a control node (e.g., `btnSave`) or event handler.
2. Recursively trace outgoing relationship edges (`defines-event`, `calls`, etc.) in the SQLite database:
   - Node attributes `id`, `name`, and `kind` must be retrieved.
   - Set maximum search depth to prevent excessive execution (default is `10` unless custom is specified).
   - Trace circular dependencies using the `visited` node set tracking logic. If a cycle is detected, flag `cycle_detected: true` and terminate branch expansion.

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
    "id": "btnSave",
    "name": "btnSave",
    "kind": "control",
    "children": [
      {
        "id": "btnSave_Click",
        "name": "btnSave_Click",
        "kind": "event",
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
