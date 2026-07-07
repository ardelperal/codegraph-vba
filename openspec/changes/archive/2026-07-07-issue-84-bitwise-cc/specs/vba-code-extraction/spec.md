# Delta for vba-code-extraction

## NEW Requirements

### Requirement: Bitwise-Precise Conditional-Compilation

To support accurate extraction of VBA code containing conditional compilation directives, the preprocessor MUST parse and evaluate conditional compilation expressions using bitwise-precise evaluation, signed 32-bit integer arithmetic, and VBA's `-1` (True) / `0` (False) truthiness semantics. The evaluation MUST NOT use unsafe JavaScript `eval` or `Function` execution.

The preprocessor MUST support:
- Logical/bitwise operators: `And`, `Or`, `Not`, `Xor` (case-insensitive, evaluated as bitwise operations on signed 32-bit integers, without short-circuiting).
- Comparison operators: `=`, `<>`, `<`, `>`, `<=`, `>=` (returning `-1` for true, `0` for false).
- Hardcoded constants: `True` -> `-1`, `False` -> `0`, `VBA7` -> `-1`, `Win64` -> `-1`, `Win32` -> `-1`, `Win16` -> `0`, `Mac` -> `0`.
- File-scoped constants declared with `#Const`.
- Unresolved/undefined identifiers falling back to `0`.
- Out-of-grammar/unhandled syntax falling back safely to false (evaluating to `0` / blanking the inactive branches).

#### Scenario: Bitwise AND evaluates to false

- GIVEN a VBA source file containing:
  ```vba
  #Const FLAGS = 3
  #If (FLAGS And 4) Then
  Public Sub ActiveSub()
  End Sub
  #Else
  Public Sub InactiveSub()
  End Sub
  #End If
  ```
- WHEN the preprocessor evaluates the conditional expression `(FLAGS And 4)`
- THEN `(3 And 4)` evaluates to `0`
- AND the `#If` branch is blanked (deactivated) while the `#Else` branch is kept active.

#### Scenario: Bitwise NOT

- GIVEN a VBA source file containing:
  ```vba
  #Const FLAGS = 0
  #If Not FLAGS Then
  Public Sub ActiveSub()
  End Sub
  #End If
  ```
- WHEN the preprocessor evaluates the conditional expression `Not FLAGS`
- THEN `Not 0` evaluates to `-1`
- AND the `#If` branch is kept active.

#### Scenario: Bitwise XOR

- GIVEN a VBA source file containing:
  ```vba
  #Const FLAGS = 1
  #If FLAGS Xor 1 Then
  Public Sub ActiveSub()
  End Sub
  #Else
  Public Sub InactiveSub()
  End Sub
  #End If
  ```
- WHEN the preprocessor evaluates the conditional expression `FLAGS Xor 1`
- THEN `1 Xor 1` evaluates to `0`
- AND the `#If` branch is blanked while the `#Else` branch is kept active.

#### Scenario: Comparisons return -1 and 0

- GIVEN a VBA source file containing:
  ```vba
  #If Win64 = -1 Then
  Public Sub Win64Sub()
  End Sub
  #End If
  ```
- WHEN the preprocessor evaluates `Win64 = -1` (where `Win64` is pre-defined to `-1`)
- THEN it performs a comparison and returns `-1` (true)
- AND the `#If` branch is kept active.

#### Scenario: Unhandled or out-of-grammar syntax fallback

- GIVEN a VBA source file containing invalid conditional expression syntax:
  ```vba
  #If FLAGS InvalidSyntax @@@ Then
  Public Sub InvalidSub()
  End Sub
  #End If
  ```
- WHEN the preprocessor parses the expression
- THEN it falls back safely to `0` (false)
- AND the `#If` branch is blanked.
