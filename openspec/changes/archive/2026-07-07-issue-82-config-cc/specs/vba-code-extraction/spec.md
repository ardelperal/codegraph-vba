# Delta for vba-code-extraction (Issue #82)

## NEW Requirements

### Requirement: Configurable Conditional-Compilation Platform Targets

To support customizable preprocessing of VBA conditional compilation directives, the system MUST allow platform target constants (e.g., `VBA7`, `Win64`, `Win32`, `Win16`, `Mac`) to be configured within `codegraph.json` or local `config.json` under the `vba.targets` configuration block.

The preprocessor MUST evaluate conditional compilation expressions using these targets under the following precedence and fallback rules:
1. **File-scoped `#Const` definitions**: Explicit `#Const` declarations within the VBA file take top priority and override any configured platform targets for that file.
2. **Configured Preprocessor Targets**: Loaded configurations from `codegraph.json` / `config.json` define the platform targets.
3. **Hardcoded Defaults**: When targets are not configured, the preprocessor falls back to the standard VBA platform target defaults:
   - `VBA7 = true` (evaluated as `-1`)
   - `Win64 = true` (evaluated as `-1`)
   - `Win32 = true` (evaluated as `-1`)
   - `Win16 = false` (evaluated as `0`)
   - `Mac = false` (evaluated as `0`)

If a configured value is invalid (e.g., non-boolean values or malformed configuration structure), the system MUST warn about the malformed configuration and fall back gracefully to the default values.

---

#### Scenario: Predefined Platform Constants Configured

- GIVEN a project configuration `codegraph.json` specifying:
  ```json
  {
    "vba": {
      "targets": {
        "Win64": false,
        "Win32": true
      }
    }
  }
  ```
- AND a VBA source file containing:
  ```vba
  #If Win64 Then
  Public Sub Win64Sub()
  End Sub
  #Else
  Public Sub Win32Sub()
  End Sub
  #End If
  ```
- WHEN the extractor processes the source under this project configuration
- THEN `Win64` evaluates to `0` (false) and `Win32` evaluates to `-1` (true)
- AND the `#If` branch is deactivated (blanked) while the `#Else` branch is kept active
- AND the extractor emits the `Win32Sub` function node and zero `Win64Sub` function nodes.

#### Scenario: Predefined Platform Constants Default

- GIVEN no project configuration or an empty/absent `vba.targets` configuration
- AND a VBA source file containing:
  ```vba
  #If Win64 Then
  Public Sub Win64Sub()
  End Sub
  #Else
  Public Sub Win32Sub()
  End Sub
  #End If
  ```
- WHEN the extractor processes the source
- THEN `Win64` evaluates to `-1` (true) by default
- AND the `#If` branch is kept active while the `#Else` branch is deactivated (blanked)
- AND the extractor emits the `Win64Sub` function node and zero `Win32Sub` function nodes.

#### Scenario: Config Override Precedence

- GIVEN a project configuration `codegraph.json` specifying:
  ```json
  {
    "vba": {
      "targets": {
        "Win64": false
      }
    }
  }
  ```
- AND a VBA source file containing:
  ```vba
  #Const Win64 = -1
  #If Win64 Then
  Public Sub Win64Sub()
  End Sub
  #Else
  Public Sub Win32Sub()
  End Sub
  #End If
  ```
- WHEN the extractor processes the source under this project configuration
- THEN the file-scoped `#Const Win64 = -1` overrides the configured `Win64: false`
- AND the preprocessor evaluates `Win64` as `-1` (true), keeping the `#If` branch active.

#### Scenario: Malformed Configuration Handling

- GIVEN a project configuration `codegraph.json` with a malformed targets block:
  ```json
  {
    "vba": {
      "targets": {
        "Win64": "invalid-non-boolean"
      }
    }
  }
  ```
- AND a VBA source file containing:
  ```vba
  #If Win64 Then
  Public Sub Win64Sub()
  End Sub
  #Else
  Public Sub Win32Sub()
  End Sub
  #End If
  ```
- WHEN the extractor loads the project configuration and processes the source
- THEN the system warns about the invalid type for `Win64` configuration
- AND falls back to the default `Win64 = true` (-1)
- AND evaluates `Win64` as `-1` (true), keeping the `#If` branch active.
