# vba-handler-backtrace

## Purpose

Specifies the requirements and testable scenarios for the `vba-handler-backtrace` capability. This custom skill resolves the call path from a UI control event handler through subroutines to DAO database operations, extracting query string hints and custom types dynamically.

## Requirements

### Requirement: Call Trace Resolution

The backtrace tool MUST dynamically traverse the call graph from a Form Control event to its handler, downstream subroutines/functions, and eventual database operations.

#### Scenario: Happy path call trace resolution
- GIVEN a form control event on a button `cmdConfirm` in module `Form_Orders.cls`
- AND an event handler `cmdConfirm_Click` calling subroutine `ProcessOrder` in helper module `OrderHelper.bas`
- AND `ProcessOrder` performs a DAO query on table `Orders`
- WHEN the backtrace is initiated for `cmdConfirm.Click`
- THEN the system MUST return the complete path: `cmdConfirm_Click` -> `ProcessOrder` -> DAO database call
- AND the execution time SHOULD be less than 150ms

---

### Requirement: DAO Query String Extraction

When traversing a DAO call, the backtrace tool MUST dynamically read the source file line and extract the SQL query string as a hint, truncated to a maximum of 200 characters. It SHALL reconstruct simple multiline string concatenations.

#### Scenario: SQL hint extraction from DAO call
- GIVEN a DAO call in `OrderHelper.bas` at line 42: `db.Execute "INSERT INTO Log (Msg) " & _` and line 43: `"VALUES ('Order Processed')"`
- WHEN the backtrace tool reaches the DAO call site at line 42
- THEN it MUST parse lines 42 and 43 to reconstruct the SQL string
- AND extract the SQL hint: `"INSERT INTO Log (Msg) VALUES ('Order Processed')"`
- AND truncate the extracted hint if it exceeds 200 characters

---

### Requirement: UDT Parameter Mapping

For each procedure in the trace path, the tool MUST read its signature from the source file, extract parameter names and their User-Defined Types (UDT), and map them.

#### Scenario: Extracting UDT parameter types from signature
- GIVEN a subroutine declaration in `OrderHelper.bas`:
  `Public Sub ProcessOrder(ByRef ctx As OrderContext, ByVal flags As Long)`
- WHEN the procedure is traversed in the backtrace path
- THEN the system MUST parse the signature to extract the parameter `ctx` of UDT type `OrderContext`
- AND include this parameter mapping in the trace node metadata

---

### Requirement: Cycle Detection and Termination

The call graph traversal MUST detect cyclic calls (e.g., A -> B -> A) and terminate traversal for that branch immediately to prevent infinite recursion.

#### Scenario: Circular call trace termination
- GIVEN a subroutine `SubA` that calls `SubB`
- AND `SubB` calls `SubA`
- WHEN the call graph is traversed
- THEN the system MUST detect that `SubA` has already been visited
- AND stop traversing downstream calls for that branch
- AND append a `CYCLE_DETECTED` warning to the trace metadata for the cyclic node

---

### Requirement: Depth Capping and Truncation Warning

The call graph traversal MUST NOT exceed a configurable maximum depth `N` (default 5). If truncated, a warning MUST be added to the trace warnings array.

#### Scenario: Depth truncation warning
- GIVEN a deep call chain `Sub1 -> Sub2 -> Sub3 -> Sub4 -> Sub5 -> Sub6`
- AND a maximum traversal depth set to 5
- WHEN the call graph is traversed
- THEN the system MUST truncate the traversal at `Sub5`
- AND add a warning object containing `MAX_DEPTH_EXCEEDED` to the trace's `warnings` array
