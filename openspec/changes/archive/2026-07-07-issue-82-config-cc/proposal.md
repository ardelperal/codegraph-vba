# Proposal: Configurable VBA CC Platforms via codegraph.json (Issue #82)

## Intent
Make the VBA preprocessor platform targets (e.g., `VBA7`, `Win64`, `Win32`, `Win16`, `Mac`) configurable via `codegraph.json`. This replaces the hardcoded defaults with project-scoped configurable values, allowing teams to pre-process conditional compilation branches according to their specific target platforms (e.g., x86 vs. x64, Mac vs. Windows).

## Scope
- **In Scope**:
  - Extend the JSON config parser in [src/project-config.ts](file:///C:/00repos/codigo/codegraph-vba/src/project-config.ts) to support the `vba.targets` options object.
  - Thread the loaded target configuration (`vbaTargets`) through:
    - [src/extraction/index.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/index.ts) (`ExtractionOrchestrator`)
    - [src/extraction/parse-pool.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/parse-pool.ts) (`ParseWorkerPool`)
    - [src/extraction/parse-worker.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/parse-worker.ts) (Worker thread)
    - [src/extraction/tree-sitter.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/tree-sitter.ts) (`extractFromSource`)
    - [src/extraction/vba-extractor.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-extractor.ts) (`VbaExtractor`)
    - [src/extraction/vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts) (`preprocessConditionalCompilation`)
  - Update tests to verify configuration loading and preprocessor evaluation with custom targets.
- **Out of Scope**:
  - Implement full bitwise-precise preprocessor expression evaluation (this is handled under issue 84 / bitwise-cc).

## Capabilities
- **New**: None.
- **Modified**: `vba-code-extraction` (to support configurable pre-processing targets).

## Technical Approach: Thread-Safe Explicit Argument Passing
To maintain thread safety and project isolation (important for the multi-project daemon/MCP server), we will explicitly pass `vbaTargets` through function arguments and worker messages rather than storing it in global process state.

### Configuration Schema
```json
{
  "vba": {
    "targets": {
      "VBA7": true,
      "Win64": false,
      "Win32": true,
      "Win16": false,
      "Mac": false
    }
  }
}
```

### Precedence Order during Preprocessing Evaluation
When evaluating conditional compilation identifiers, we will look up values in the following order:
1. **File-scoped `#Const` definitions** (e.g., `#Const DEBUG_MODE = -1`).
2. **Configured Preprocessor Targets** (loaded from `codegraph.json` or local `config.json`).
3. **Hardcoded Defaults** (`VBA7 = true`, `Win64 = true`, `Win32 = true`, `Win16 = false`, `Mac = false`).

## Affected Areas

### 1. [src/project-config.ts](file:///C:/00repos/codigo/codegraph-vba/src/project-config.ts)
- Extend `ProjectConfig` and `ParsedConfig` interfaces to include `vba?: { targets?: Record<string, boolean> }`.
- Implement `extractVbaTargets(parsed: object, file: string)` to validate keys and values under `vba.targets`.
- Merge targets from local `config.json` and root `codegraph.json` in `loadParsedConfig`.
- Export `loadVbaTargets(rootDir: string): Record<string, boolean>`.

### 2. [src/extraction/index.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/index.ts)
- Load `vbaTargets` using `loadVbaTargets(this.rootDir)` inside `indexAll` and `extractSingleFile`.
- Forward `vbaTargets` to `extractFromSource` and `pool.requestParse`.

### 3. [src/extraction/parse-pool.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/parse-pool.ts)
- Add `vbaTargets?: Record<string, boolean>` to the `ParseTask` interface.
- Pass `vbaTargets` in the `postMessage` call inside the `dispatch` method.

### 4. [src/extraction/parse-worker.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/parse-worker.ts)
- Retrieve `vbaTargets` from the message payload.
- Forward `vbaTargets` to the `extractFromSource` function call.

### 5. [src/extraction/tree-sitter.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/tree-sitter.ts)
- Add `vbaTargets?: Record<string, boolean>` to the signature of `extractFromSource`.
- Pass it to the constructor of `VbaExtractor`.

### 6. [src/extraction/vba-extractor.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-extractor.ts)
- Update the constructor of `VbaExtractor` to accept `vbaTargets` and store it as a private property.
- Pass `this.vbaTargets` to the preprocessor: `preprocessConditionalCompilation(joined, this.vbaTargets)`.

### 7. [src/extraction/vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts)
- Modify `preprocessConditionalCompilation` to accept `customTargets?: Record<string, boolean>`.
- Pass `customTargets` down to `evaluateConditionalExpression`, `evaluateConstRhs`, and `tokenize`.
- In `tokenize`, search `customTargets` case-insensitively before falling back to the hardcoded constants.

### 8. Tests
- Extend/create tests to verify:
  - Configuration loader successfully validates and loads well-formed, case-insensitive VBA targets.
  - Preprocessor correctly evaluates conditional branches based on the configured custom targets.
  - Workers successfully receive and process files using the custom preprocessor targets.

## Risks & Mitigation
- **Missing or Invalid Configuration**:
  - *Mitigation*: The loader will fallback gracefully on parsing errors or type mismatches (e.g. non-boolean target values), dropping invalid targets and falling back to the hardcoded targets. Zero-config operations will continue to work exactly as they do today.

## Rollback Plan
Revert the edits using:
```bash
git checkout -- src/project-config.ts src/extraction/index.ts src/extraction/parse-pool.ts src/extraction/parse-worker.ts src/extraction/tree-sitter.ts src/extraction/vba-extractor.ts src/extraction/vba-preprocess.ts
```

## Success Criteria
- Validated by unit and integration tests under `vitest`.
- Successful configuration parsing and preprocessor path execution verified under multiple target layouts.
