---
name: vba-event-tracer
description: Traces VBA events to find declaration sites, raise sites, and dynamic handler resolutions using the codegraph index.
---

# Skill: VBA Event Tracer

This skill enables the agent to trace a custom VBA event, locate its declaration and raise sites, resolve subscriber handlers dynamically, and report warnings (e.g., ambiguity, no raisers) in a structured JSON schema.

## Triggering Context
This skill is triggered when the user requests to trace or explore a VBA event (e.g., "Trace VBA event PedidoGuardado", "Find event handlers for DataChanged").

## Step-by-Step Execution Guide

### Step 1: Run codegraph_explore Query
1. Identify the absolute path of the project workspace (`projectPath`).
2. Invoke `codegraph_explore` using:
   - `query`: The event name to trace (e.g., `"PedidoGuardado"`) or the event name plus class/module qualifiers if provided (e.g., `"PedidoPublisher PedidoGuardado"`).
   - `projectPath`: The absolute path of the project root.
3. The tool returns the source code and relationship edges (e.g., `raises-event`, `subscribes-event`) in the target event graph.

### Step 2: Ambiguity Detection & Verification
1. From the returned nodes, filter for nodes with `kind: "event"` whose simple `name` matches the target event name (case-insensitive).
2. If multiple event nodes exist in different classes/modules (e.g., `PublisherA.DataChanged` and `PublisherB.DataChanged`), and the query is unqualified (e.g., only `"DataChanged"`):
   - Format a warning string: `"EVENT_AMBIGUOUS: Event name '<EventName>' is ambiguous. Candidates: <Candidate1>, <Candidate2>, ..."`
   - Immediately stop and return the JSON output with this warning in the `warnings` array, leaving `event_declarations`, `raise_sites`, and `handlers` empty.
3. Otherwise, proceed with the single resolved event node.

### Step 3: Event Declaration Site Extraction
1. Extract the resolved event node's details:
   - `module`: The file basename of `filePath`.
   - `line`: The `startLine` number.
   - `signature`: The `signature` field of the node (which stores the trimmed source declaration line, e.g., `"Public Event PedidoGuardado(ByVal IdPedido As Long)"`).
2. Add this object to the `event_declarations` list.

### Step 4: Raise Sites Resolution
1. Scan the incoming edges in the returned graph targeting the resolved event node where `kind` is `"raises-event"`.
2. For each edge:
   - Locate the source node (the subroutine/function/property raising the event).
   - Retrieve:
     - `module`: The file basename of the source node's `filePath`.
     - `line`: The line number of the actual `RaiseEvent` statement (if available in edge metadata, or look inside the source procedure body).
     - `context`: The full declaration signature of the parent procedure raising the event (e.g., `"Public Sub Guardar()"`).
3. If no raise sites are found, append `"NO_RAISERS"` to the `warnings` list.

### Step 5: Dynamic Handler Resolution
1. Identify the event's parent class module (the class/module defining the event).
2. Find all `subscribes-event` edges in the graph targeting that class module.
3. For each subscribing module (source of the `subscribes-event` edge):
   - Retrieve the variable name used for subscription from the edge's metadata: `edge.metadata.variableName` (e.g., `"m_Form"`).
   - Scan the procedures inside the subscribing module for a subroutine named `<VariableName>_<EventName>` (case-insensitive, e.g., `Sub m_Form_PedidoGuardado()`).
   - If a matching handler procedure is found:
     - Add an entry to the `handlers` list:
       - `form`: File basename of the subscribing module.
       - `handler`: The exact name of the handler subroutine (e.g., `"m_Form_PedidoGuardado"`).
       - `via`: The subscription variable name (e.g., `"m_Form"`).

### Step 6: Loop & Circular Reference Avoidance
1. Maintain a set of visited `(module, event)` tuples during trace resolution.
2. Before traversing any event-handler path or nested event raises, verify if it has already been visited.
3. If a loop is detected, abort further traversal down that path to prevent infinite recursion, ensuring the trace completes under 100ms.

### Step 7: Output Formatting
Output the final result in the exact JSON schema:

```json
{
  "event_declarations": [
    {
      "module": "PedidoPublisher.cls",
      "line": 2,
      "signature": "Public Event PedidoGuardado(ByVal IdPedido As Long)"
    }
  ],
  "raise_sites": [
    {
      "module": "PedidoPublisher.cls",
      "line": 15,
      "context": "Public Sub Guardar()"
    }
  ],
  "handlers": [
    {
      "form": "FormListener.cls",
      "handler": "m_Form_PedidoGuardado",
      "via": "m_Form"
    }
  ],
  "warnings": []
}
```
