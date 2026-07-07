# Design: Configurable VBA CC Platforms via codegraph.json (Issue #82)

## Intent
Make the VBA preprocessor platform targets (e.g., `VBA7`, `Win64`, `Win32`, `Win16`, `Mac`) configurable via `codegraph.json` or local `config.json`. This replaces hardcoded default targets with project-scoped configurable values, allowing the preprocess conditional compilation stage to evaluate branches according to the project's target platforms.

---

## Configuration Schema

The project-scoped configuration file `codegraph.json` and the local `.codegraph/config.json` will support a new `vba` configuration block:

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

---

## Detailed Changes by File

### 1. `src/project-config.ts`

Extend configuration types and implement parsing and loading for `vba.targets`.

- **Type Definition Changes**:
  Extend `ProjectConfig` and `ParsedConfig` interfaces:
  ```typescript
  export interface ProjectConfig {
    // ... existing fields ...
    vba?: {
      targets?: Record<string, boolean>;
    };
  }

  interface ParsedConfig {
    // ... existing fields ...
    vba?: {
      targets?: Record<string, boolean>;
    };
  }
  ```

- **Validation & Parsing Helper**:
  Add `extractVbaTargets(parsed: object, file: string)` to validate keys and values under `vba.targets`. It will log warnings on type mismatches and ignore invalid entries gracefully:
  ```typescript
  function extractVbaTargets(parsed: object, file: string): Record<string, boolean> | undefined {
    const vba = (parsed as any).vba;
    if (vba === undefined) return undefined;
    if (!vba || typeof vba !== 'object' || Array.isArray(vba)) {
      logWarn(`Ignoring "vba" in ${PROJECT_CONFIG_FILENAME}: must be an object`, { file });
      return undefined;
    }

    const targets = vba.targets;
    if (targets === undefined) return undefined;
    if (!targets || typeof targets !== 'object' || Array.isArray(targets)) {
      logWarn(`Ignoring "vba.targets" in ${PROJECT_CONFIG_FILENAME}: must be an object`, { file });
      return undefined;
    }

    const out: Record<string, boolean> = {};
    for (const [rawKey, rawVal] of Object.entries(targets)) {
      if (typeof rawVal !== 'boolean') {
        logWarn(`Ignoring invalid target "${rawKey}" in ${PROJECT_CONFIG_FILENAME}: value must be a boolean`, { file });
        continue;
      }
      out[rawKey] = rawVal;
    }
    return out;
  }
  ```

- **Merging & Caching**:
  Update `loadParsedConfig` to load and merge VBA targets:
  ```typescript
  // Inside loadParsedConfig:
  const fileConfig = fileMtimeMs > 0 ? parseConfig(file) : EMPTY_CONFIG;
  const localConfig = localMtimeMs > 0 ? parseConfig(localFile) : EMPTY_CONFIG;

  // Merge VBA targets (fileConfig overrides localConfig)
  let vba: { targets?: Record<string, boolean> } | undefined;
  if (localConfig.vba?.targets || fileConfig.vba?.targets) {
    vba = {
      targets: {
        ...localConfig.vba?.targets,
        ...fileConfig.vba?.targets,
      }
    };
  }

  const config: ParsedConfig = {
    extensions,
    includeIgnored,
    exclude,
    vba,
  };
  ```

- **Exported Loader Function**:
  Expose `loadVbaConfig` (and a helper alias `loadVbaTargets`) for external consumers:
  ```typescript
  export function loadVbaConfig(rootDir: string): { targets?: Record<string, boolean> } {
    return loadParsedConfig(rootDir).vba || {};
  }
  ```

---

## 2. `src/extraction/index.ts` (ExtractionOrchestrator / IndexBuilder)

Thread the custom VBA targets configuration through the index building pipeline.

- Load custom VBA platform targets inside `indexAll`:
  ```typescript
  const vbaTargets = loadVbaConfig(this.rootDir).targets;
  ```
- Pass `vbaTargets` into the `parseFile` calls:
  ```typescript
  const parseFile = (filePath: string, content: string): Promise<ExtractionResult> => {
    const language = detectLanguage(filePath, content, overrides);
    if (!pool) return Promise.resolve(extractFromSource(filePath, content, language, frameworkNames, vbaTargets));
    return pool.requestParse({ filePath, content, language, frameworkNames, vbaTargets });
  };
  ```

---

## 3. `src/extraction/parse-pool.ts`

- Add `vbaTargets?: Record<string, boolean>` to the `ParseTask` interface:
  ```typescript
  export interface ParseTask {
    filePath: string;
    content: string;
    language: Language;
    frameworkNames?: string[];
    vbaTargets?: Record<string, boolean>;
  }
  ```
- Update `ParseWorkerPool.dispatch` to post `vbaTargets` in the message:
  ```typescript
  w.postMessage({
    type: 'parse',
    id: job.id,
    filePath: job.task.filePath,
    content: job.task.content,
    frameworkNames: job.task.frameworkNames,
    language: job.task.language,
    vbaTargets: job.task.vbaTargets,
  });
  ```

---

## 4. `src/extraction/parse-worker.ts`

- Add `vbaTargets?: Record<string, boolean>` to the incoming message type definition.
- Destructure `vbaTargets` from the message and pass it to `extractFromSource`:
  ```typescript
  const { id, filePath, content, frameworkNames, vbaTargets } = msg;
  // ...
  const result: ExtractionResult = extractFromSource(filePath!, content!, language, frameworkNames, vbaTargets);
  ```

---

## 5. `src/extraction/tree-sitter.ts`

- Update `extractFromSource` signature to accept `vbaTargets?: Record<string, boolean>`:
  ```typescript
  export function extractFromSource(
    filePath: string,
    source: string,
    language?: Language,
    frameworkNames?: string[],
    vbaTargets?: Record<string, boolean>
  ): ExtractionResult
  ```
- Pass `vbaTargets` to `VbaExtractor`:
  ```typescript
  } else if (detectedLanguage === 'vba') {
    const extractor = new VbaExtractor(filePath, source, vbaTargets);
    result = extractor.extract();
  }
  ```

---

## 6. `src/extraction/vba-extractor.ts`

- Accept `vbaTargets?: Record<string, boolean>` in the constructor:
  ```typescript
  export class VbaExtractor {
    private filePath: string;
    private source: string;
    private ctx: VbaExtractorContext;
    private vbaTargets?: Record<string, boolean>;

    constructor(filePath: string, source: string, vbaTargets?: Record<string, boolean>) {
      this.filePath = filePath;
      this.source = source;
      this.vbaTargets = vbaTargets;
      this.ctx = new VbaExtractorContext(filePath);
    }
    // ...
  }
  ```
- Pass it down when preprocessing:
  ```typescript
  const uncommented = preprocessConditionalCompilation(joined, this.vbaTargets);
  ```

---

## 7. `src/extraction/vba-preprocess.ts`

- Update `preprocessConditionalCompilation` to accept custom targets:
  ```typescript
  export function preprocessConditionalCompilation(src: string, customTargets?: Record<string, boolean>): string {
    // ...
  ```
- Pass `customTargets` down to helper functions `evaluateConstRhs` and `evaluateConditionalExpression`:
  ```typescript
  const value = evaluateConstRhs(rhs, constTable, customTargets);
  // ...
  const active = parentActive && evaluateConditionalExpression(ifMatch[1] ?? '', constTable, customTargets);
  // ...
  evaluateConditionalExpression(elseIfMatch[1] ?? '', constTable, customTargets);
  ```
- Update `evaluateConditionalExpression` and `evaluateConstRhs` signatures to accept and forward `customTargets`:
  ```typescript
  function evaluateConditionalExpression(
    expr: string,
    constTable: ReadonlyMap<string, string> = new Map(),
    customTargets?: Record<string, boolean>,
  ): boolean {
    // ...
    const tokens = tokenize(exprClean, constTable, customTargets);
    // ...
  }

  function evaluateConstRhs(
    rhs: string,
    constTable: ReadonlyMap<string, string>,
    customTargets?: Record<string, boolean>,
  ): string | null {
    // ...
    const tokens = tokenize(rhs.trim(), constTable, customTargets);
    // ...
  }
  ```
- Update `tokenize` to accept `customTargets` and check it case-insensitively before defaults:
  ```typescript
  function tokenize(
    expr: string,
    constTable: ReadonlyMap<string, string>,
    customTargets?: Record<string, boolean>
  ): Token[] {
    // ...
    // Inside the identifier token resolution loop:
    
    // 1. File-scoped #Const table first (case-insensitive)
    let resolvedValue: number | null = null;
    for (const [key, val] of constTable.entries()) {
      if (key.toUpperCase() === idUpper) {
        resolvedValue = parseInt(val, 10) | 0;
        break;
      }
    }

    // 2. Configured Preprocessor Targets (case-insensitive)
    if (resolvedValue === null && customTargets) {
      for (const [key, val] of Object.entries(customTargets)) {
        if (key.toUpperCase() === idUpper) {
          resolvedValue = val ? -1 : 0;
          break;
        }
      }
    }

    // 3. Hardcoded environment constants fallback
    if (resolvedValue === null) {
      switch (idUpper) {
        case 'VBA7':
        case 'WIN64':
        case 'WIN32':
        case 'TRUE':
          resolvedValue = -1;
          break;
        case 'WIN16':
        case 'MAC':
        case 'FALSE':
          resolvedValue = 0;
          break;
      }
    }

    // 4. Default fallback
    if (resolvedValue === null) {
      resolvedValue = 0;
    }
    // ...
  }
  ```

---

## Verification Plan

### Automated Unit Tests
1. **Config Validation Unit Tests**:
   - Verify that well-formed configs with various platforms are correctly parsed.
   - Verify that malformed configs (non-boolean values, wrong structure) emit warnings and degrade gracefully to default values.
   - Verify merging priority: `codegraph.json` (root) overrides `.codegraph/config.json` (local).
2. **Preprocessor Evaluation Tests**:
   - Verify precedence: `#Const` overrides loaded configs, and loaded configs override defaults.
   - Verify case-insensitivity: keys like `win64`, `WIN64`, `Win64` configured in `codegraph.json` resolve appropriately.
   - Verify fallback handling: unknown identifiers and empty targets default to `0`.
3. **Parse Worker Integration Tests**:
   - Verify worker parsing of files with custom targets.
