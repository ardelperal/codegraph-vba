# Exploration: Configurable VBA CC Platforms via codegraph.json (Issue #82)

## Current State
Currently, the VBA preprocessor evaluates conditional compilation platform targets (such as `VBA7`, `Win64`, `Win32`, `Win16`, and `Mac`) using hardcoded boolean flags that map to `-1` (VBA True) and `0` (VBA False). These targets are not customizable on a per-project basis. 

The project configuration file `codegraph.json` is loaded and validated in `src/project-config.ts`, but it only supports:
- `extensions` (overrides custom file extension mapping)
- `includeIgnored` (force inclusion of gitignored folders)
- `exclude` (ignores specific files/directories)

During indexing, the `IndexBuilder` (`src/extraction/index.ts`) reads this configuration and manages parsing tasks, dispatching them to worker threads via `ParseWorkerPool` (`src/extraction/parse-pool.ts`). The tasks are executed inside `src/extraction/parse-worker.ts`, calling the entry point `extractFromSource` in `src/extraction/tree-sitter.ts`, which routes VBA files to `VbaExtractor` (`src/extraction/vba-extractor.ts`). The preprocessor `preprocessConditionalCompilation` in `src/extraction/vba-preprocess.ts` receives no configuration and relies entirely on hardcoded constants.

---

## Affected Areas

### 1. Project Configuration (`src/project-config.ts`)
- **Action**: Extend `ProjectConfig` and `ParsedConfig` interfaces to support `vba?: { targets?: Record<string, boolean> }`.
- **Validation**: Implement a validation helper (e.g., `extractVbaTargets`) that parses keys and boolean values under the `vba.targets` path, warning on invalid entries.
- **Export**: Export a new helper function `loadVbaTargets(rootDir: string): Record<string, boolean>`.

### 2. Extraction Indexer (`src/extraction/index.ts`)
- **Action**: Inside `indexAll`, load the configuration overrides using `loadVbaTargets(this.rootDir)`.
- **Action**: Pass the resolved targets to `extractFromSource` (for in-process fallback) and to `pool.requestParse({ ..., vbaTargets })`.

### 3. Parser Pool (`src/extraction/parse-pool.ts` & `src/extraction/parse-worker.ts`)
- **Action**: Add `vbaTargets?: Record<string, boolean>` to the `ParseTask` interface.
- **Action**: In `ParseWorkerPool.dispatch`, forward `vbaTargets` inside the `w.postMessage` payload.
- **Action**: In `parse-worker.ts`, extract `vbaTargets` from the message and pass it to `extractFromSource`.

### 4. Tree-Sitter Extractor Interface (`src/extraction/tree-sitter.ts`)
- **Action**: Update `extractFromSource` signature to accept `vbaTargets?: Record<string, boolean>`.
- **Action**: Route it to the `VbaExtractor` constructor: `new VbaExtractor(filePath, source, vbaTargets)`.

### 5. VBA Extractor (`src/extraction/vba-extractor.ts`)
- **Action**: Update constructor of `VbaExtractor` to accept `vbaTargets?: Record<string, boolean>` and store it on the class instance.
- **Action**: Pass it to the preprocessor: `preprocessConditionalCompilation(joined, this.vbaTargets)`.

### 6. VBA Preprocessor (`src/extraction/vba-preprocess.ts`)
- **Action**: Modify `preprocessConditionalCompilation` signature to accept `customTargets?: Record<string, boolean>`.
- **Action**: Pass the `customTargets` option down to `evaluateConditionalExpression`, `evaluateConstRhs`, and finally `tokenize`.
- **Action**: In `tokenize`, search `customTargets` case-insensitively for the upper-case identifier before falling back to the hardcoded environment constants.

---

## Approaches

### Approach 1: Thread-Safe Explicit Argument Passing (Recommended)
Pass `vbaTargets` explicitly through the function call and worker thread message chains.
- **Pros**:
  - Thread-safe and compatible with multi-threaded execution (the daemon uses multiple worker threads).
  - Explicit and fits the existing design of configuration options like `frameworkNames` and `language`.
- **Cons**:
  - Requires updating multiple method signatures across 6 files.

### Approach 2: Global Configuration / Node Environment Variables
Store configuration in global state or inject them into worker environments via `process.env`.
- **Pros**:
  - Fewer signature changes.
- **Cons**:
  - `process.env` values are strings, requiring extra serialization/parsing logic.
  - Global variables complicate multi-project setups (the CodeGraph daemon and multi-project MCP server must isolate configuration by project root).

---

## Recommendation
Implement **Approach 1**. Explicitly passing the config object guarantees project isolation, thread safety, and maintainable type definitions.

---

## Risks & Findings

### Test Assertion Bug in `srcOverflow`
During the investigation of `__tests__/extraction-vba-preprocess.test.ts`, we discovered a test bug in the `srcOverflow` case:
```typescript
#If 2147483647 + 1 = -2147483648 Then
  Debug.Print "active"
#Else
  Debug.Print "inactive"
#End If
```
The Pratt/Recursive Descent Parser inside `vba-preprocess.ts` supports unary operations (`+` and `-`) but **does not implement binary addition or subtraction**. When the parser encounters `2147483647 + 1`, it fails to parse the binary `+` operator, throws an exception, and falls back to evaluating the expression as `false`.

As a result, the preprocessor correctly selects the `#Else` branch and outputs `Debug.Print "inactive"`. However, the test asserted `expect(out).toContain('active')`. Because the string `"inactive"` contains the substring `"active"`, the test passed successfully despite the expression failing to parse. 

We recommend updating the tests to check for exact line contents or adjusting the parser if full binary arithmetic becomes a requirement.

### Case-Insensitivity
Users might write configuration in `codegraph.json` using different casing (e.g., `win64: false` or `WIN64: false`). The preprocessor must normalize and resolve these keys case-insensitively against VBA identifiers.

---

## Ready for Proposal
Yes.
