# Tasks: Configurable VBA CC Platforms via codegraph.json (Issue #82)

## Review Workload Forecast

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1    | Implement configuration loading, parsing, and caching in `src/project-config.ts` | PR 1 | Phase 1 changes |
| 2    | Thread configuration through orchestrator, worker pool, and AST extraction pipeline | PR 1 | Phase 2 changes |
| 3    | Update `vba-preprocess.ts` to consume custom platform targets case-insensitively | PR 1 | Phase 3 changes |
| 4    | Add unit and integration tests and verify the entire test suite passes | PR 1 | Phase 4 changes |

---

## Phase 1: Config loading

- [x] 1.1 Extend `ProjectConfig` interface in [project-config.ts](file:///C:/00repos/codigo/codegraph-vba/src/project-config.ts) to support the optional `vba` configuration block containing `targets?: Record<string, boolean>`.
- [x] 1.2 Extend `ParsedConfig` interface in [project-config.ts](file:///C:/00repos/codigo/codegraph-vba/src/project-config.ts) to include the optional `vba` configuration block.
- [x] 1.3 Implement validation helper `extractVbaTargets(parsed: object, file: string): Record<string, boolean> | undefined` in [project-config.ts](file:///C:/00repos/codigo/codegraph-vba/src/project-config.ts).
  - Reject non-objects for `vba` or `vba.targets` and log warning.
  - Warn and skip individual target keys with non-boolean values.
- [x] 1.4 Update `parseConfig` in [project-config.ts](file:///C:/00repos/codigo/codegraph-vba/src/project-config.ts) to extract `vbaTargets` using `extractVbaTargets` and return it as part of `ParsedConfig`.
- [x] 1.5 Update `loadParsedConfig` in [project-config.ts](file:///C:/00repos/codigo/codegraph-vba/src/project-config.ts) to merge `vba.targets` from local config file (`.codegraph/config.json` / `codeGraphDirName() + '/config.json'`) and project config file (`codegraph.json`), with the project-scoped file taking precedence.
- [x] 1.6 Export loader function `loadVbaConfig(rootDir: string): { targets?: Record<string, boolean> }` from [project-config.ts](file:///C:/00repos/codigo/codegraph-vba/src/project-config.ts).

## Phase 2: Threading

- [x] 2.1 Update `indexAll` in [index.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/index.ts) to load the VBA targets configuration via `loadVbaConfig` and pass `vbaTargets` to the `parseFile` closure.
- [x] 2.2 Update `ParseTask` interface in [parse-pool.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/parse-pool.ts) to accept optional `vbaTargets?: Record<string, boolean>`.
- [x] 2.3 Update `ParseWorkerPool.dispatch` in [parse-pool.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/parse-pool.ts) to include `vbaTargets` in the message posted to the worker.
- [x] 2.4 Update [parse-worker.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/parse-worker.ts) to destructure `vbaTargets` from the worker task message and pass it to `extractFromSource`.
- [x] 2.5 Update `extractFromSource` signature in [tree-sitter.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/tree-sitter.ts) to accept `vbaTargets?: Record<string, boolean>`, and forward it when initializing `VbaExtractor`.
- [x] 2.6 Update `VbaExtractor` constructor in [vba-extractor.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-extractor.ts) to accept and store `vbaTargets?: Record<string, boolean>`, and forward it to `preprocessConditionalCompilation`.

## Phase 3: Preprocessor

- [x] 3.1 Update `preprocessConditionalCompilation` signature in [vba-preprocess.ts](file:///C:/00repos/codigo/codegraph-vba/src/extraction/vba-preprocess.ts) to accept `customTargets?: Record<string, boolean>`.
- [x] 3.2 Thread `customTargets` down to helper functions `evaluateConditionalExpression`, `evaluateConstRhs`, and `tokenize`.
- [x] 3.3 Update resolution in `tokenize` to check custom targets case-insensitively:
  - First, search the `#Const` table (`constTable`).
  - Second, search `customTargets` case-insensitively, converting `true` to `-1` and `false` to `0`.
  - Third, fall back to hardcoded defaults (`VBA7`, `WIN64`, `WIN32`, `TRUE` to `-1` and `WIN16`, `MAC`, `FALSE` to `0`).
  - Fourth, fall back to `0` for any unknown identifier.

## Phase 4: Testing

- [x] 4.1 Add config parser tests to verify:
  - Valid targets block parses correctly.
  - Project config overrides local config file values.
  - Invalid types under `vba.targets` log warnings and fall back gracefully without crashing.
- [x] 4.2 Add preprocessor evaluation tests in [extraction-vba-preprocess.test.ts](file:///C:/00repos/codigo/codegraph-vba/__tests__/extraction-vba-preprocess.test.ts) to verify:
  - Precedence: local `#Const` overrides custom targets, which in turn override built-in defaults.
  - Case-insensitivity of custom target keys.
  - Falling back of undefined targets to `0`.
- [x] 4.3 Add integration test verifying the end-to-end threading via worker parser parses conditional compilation with custom targets.
- [x] 4.4 Run the full test suite (`npm test`) to ensure everything compiles and all tests pass.
