# vba-extraction-enhancements

## Purpose

Enhance the VBA extraction parser to support external DLL declarations, preprocessor conditional compilation, customized database SQL execution, and constant resolution in `DoCmd.OpenForm` calls.

## Requirements

### Requirement: DLL API Declarations Extraction

The extractor MUST parse external DLL `Declare` and `Declare PtrSafe` statements as procedure nodes of kind `function`.
1. The declaration MUST be mapped to a single-line range where `startLine` equals `endLine`.
2. Visibility MUST be normalized to `'public'` (for `Public` or default) or `'private'` (for `Private`).
3. The node `metadata` MUST include `{ isDeclare: true }`.

#### Scenario: Extract Public PtrSafe Sub Declaration
- GIVEN the declaration `Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)` on line 5
- WHEN the module is extracted
- THEN it emits a `function` node named `Sleep` with:
  - `node.visibility === 'public'`
  - `startLine === 5`
  - `endLine === 5`
  - `metadata.isDeclare === true`

---

### Requirement: Preprocessor Evaluation of Conditional Compilation

The preprocessor MUST evaluate `#If`, `#ElseIf`, `#Else`, and `#End If` directives before extraction sweeps.
1. The preprocessor MUST evaluate conditions using the constants `VBA7 = true`, `Win64 = true`, and `Mac = false`.
2. Directives and inactive branch lines MUST be replaced with empty lines (`""`) to preserve line number indexing.

#### Scenario: Preprocessing conditional DLL imports
- GIVEN a module containing:
  ```vba
  #If Win64 Then
  Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dw As Long)
  #Else
  Public Declare Sub Sleep Lib "kernel32" (ByVal dw As Long)
  #End If
  ```
- WHEN the source is preprocessed
- THEN the `#Else` branch line is replaced with an empty line, and directive lines are blanked
- AND only the `PtrSafe` declaration is visible to downstream extraction sweeps

---

### Requirement: Custom DB Variables for SQL Extraction

The extractor MUST support arbitrary variable identifiers ending with `db` (case-insensitive) in `.OpenRecordset` and `.Execute` SQL queries.
1. The extractor MUST match these variables and emit `references` edges to target tables.

#### Scenario: Extract SQL query from parameter p_db
- GIVEN a procedure using `p_db.Execute "SELECT * FROM Employees"`
- WHEN the extractor scans for SQL statements
- THEN it detects the query and emits a `references` edge to the `Employees` table

---

### Requirement: Constant Resolution in DoCmd.OpenForm

The extractor MUST track constant declarations and resolve constant identifiers used in `DoCmd.OpenForm` calls.
1. Constants declared in the module MUST be tracked in a local constants map.
2. `DoCmd.OpenForm` calls with constant arguments MUST resolve the argument to the literal value.
3. If the constant is not locally defined, it MUST fall back to the constant name as the target form name.

#### Scenario: Resolve known constant in OpenForm call
- GIVEN a module declaring `Const FORM_EMPLOYEES = "frmEmployees"`
- AND a call `DoCmd.OpenForm FORM_EMPLOYEES`
- WHEN the module is extracted
- THEN it emits an `opens-form` edge targeting `frmEmployees`

#### Scenario: Fall back to constant name when unresolved
- GIVEN a call `DoCmd.OpenForm FORM_UNKNOWN` without a local definition for `FORM_UNKNOWN`
- WHEN the module is extracted
- THEN it emits an `opens-form` edge targeting `FORM_UNKNOWN`
