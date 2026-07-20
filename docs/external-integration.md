# External-tool integration — canonical spawn pattern + headless query contract

This document is the canonical reference for any external tool (the dysflow
runtime, CI scripts, in-house dashboards, or a future MCP-aware sidecar) that
needs to invoke `codegraph-vba` from a non-PowerShell, non-interactive shell
context. It is the second half of issue #200 — sibling docs cover
[VBA-specific concerns](vba-reference-kinds.md), and the
[stability semantics of the unresolved-ref resolver](vba-stub-repoint-decision.md).

Two patterns are documented here:

| Pattern | When to use | Overhead |
|---|---|---|
| **A. Spawn the CLI as a subprocess** (Path 2 of Dysflow#1015) | Ad-hoc lookups, CI, one-off diagnostics, anything that doesn't need millisecond-level latency. | ~50–150ms per call (Node startup + WASM init). |
| **B. Read the index SQLite DB directly** (Path 1 of Dysflow#1015 — see [`docs/index-schema.md`](index-schema.md)) | Tight loops, agent hot paths, dashboards that re-query on every keystroke. | One `node:sqlite` open; sub-millisecond per query. |

Pattern A is what the rest of this doc covers. Pattern B is covered by
[`docs/index-schema.md`](index-schema.md) and its companion dump script
`scripts/dump-index-schema.ts`.

## 1. The npm-shim layout (per platform)

The package's `package.json` declares a single `bin` entry:

```json
{
  "bin": {
    "codegraph-vba": "./dist/bin/codegraph.js"
  }
}
```

`npm install -g codegraph-vba` then generates three platform-specific shim
filenames from that one entry:

| Platform | Shim filenames in `$PATH` | Resolves to |
|---|---|---|
| Windows (PowerShell) | `codegraph-vba`, `codegraph-vba.cmd`, `codegraph-vba.ps1` | `codegraph-vba.ps1` (PowerShell-only — `child_process.spawn` does NOT resolve this name) |
| Windows (cmd.exe) | `codegraph-vba.cmd` | `codegraph-vba.cmd` |
| macOS / Linux | `codegraph-vba` | a single POSIX shell script that `exec`s Node with the bundled runtime |

The `npm`-managed shim naming convention is fixed — we do NOT control it, and
renaming would break the install on every existing consumer. What an external
tool MUST do is use a platform-aware command resolution: see the snippet in
§2. (This is the cross-tool pitfall surfaced by sister issue Dysflow#1015:
`child_process.spawn('codegraph-vba', ['--version'])` fails with `ENOENT` on
Windows because `spawn` doesn't go through PowerShell's path resolution.)

### Verifying the install layout (CI / smoke matrix)

```bash
# On the consumer machine, after `npm install -g codegraph-vba`:
which codegraph-vba          # POSIX: prints /path/to/shim
where.exe codegraph-vba.cmd  # Windows: prints C:\Users\...\AppData\Roaming\npm\codegraph-vba.cmd
node -e "console.log(require('child_process').execSync('codegraph-vba --version', {stdio:['ignore','pipe','pipe']}).toString().trim())"
# → 1.13.0 (or whichever version)
```

## 2. Canonical cross-platform spawn pattern (TypeScript / JavaScript)

The recommended way to invoke `codegraph-vba` from any Node-based tool. The
platform branch is the entire point — see the in-line comment for why:

```ts
import { spawn } from 'node:child_process';

/**
 * Run a `codegraph-vba` subcommand and collect stdout as a string.
 * Works on Windows, macOS, and Linux without changes.
 *
 * @param args  Subcommand + its arguments (e.g. ['status', '--json']).
 * @returns     Resolves with the trimmed stdout string. Rejects on non-zero exit.
 */
export function runCodeGraph(args: string[]): Promise<string> {
  // The npm-managed shim is `codegraph-vba` on POSIX (a single exec'able
  // shell script). On Windows, `child_process.spawn` does NOT search the
  // `PATHEXT` list — it does a literal lookup. The `.cmd` shim is the
  // only one `spawn` will resolve; the `.ps1` one is PowerShell-only and
  // would ENOENT here. (See Dysflow#1015 for the live repro.)
  const binName =
    process.platform === 'win32' ? 'codegraph-vba.cmd' : 'codegraph-vba';

  return new Promise((resolve, reject) => {
    const child = spawn(binName, args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      // Skip the daemon attach + the wasm-flag re-exec — both add 1-2s
      // per call and are irrelevant to a headless one-shot invocation.
      env: { ...process.env, CODEGRAPH_NO_DAEMON: '1', CODEGRAPH_WASM_RELAUNCHED: '1' },
    });
    const out: Buffer[] = [];
    const err: Buffer[] = [];
    child.stdout.on('data', (c) => out.push(c));
    child.stderr.on('data', (c) => err.push(c));
    child.on('error', reject);
    child.on('close', (code) => {
      if (code === 0) resolve(Buffer.concat(out).toString('utf-8'));
      else reject(new Error(`codegraph-vba exited ${code}: ${Buffer.concat(err).toString('utf-8')}`));
    });
  });
}
```

For a Promise-wrapped helper that parses JSON automatically, just
`JSON.parse(await runCodeGraph(['status', '--json', projectPath]))`. The
shape of that JSON is pinned in §3.

## 3. Headless query contract — the JSON subcommands

Every read-only subcommand below accepts `-j, --json` and prints exactly one
JSON value to stdout on success, with no progress / log noise mixed in. (The
shimmer progress UI is suppressed by setting `CODEGRAPH_NO_DAEMON=1` and
`CODEGRAPH_WASM_RELAUNCHED=1` in the spawn env, as the snippet above does.
Without those, the JSON output may be preceded by banner / progress lines on
older builds.)

### 3.1 `status --json [path]`

Shape (top-level keys; values elided):

```jsonc
{
  "initialized": true,
  "version": "1.13.0",                // npm package version (string)
  "projectPath": "C:\\...\\my-project",
  "indexPath": "C:\\...\\my-project\\.codegraph-vba",
  "lastIndexed": "2026-07-20T13:45:22.512Z",  // ISO-8601, null before first index
  "fileCount": 643,
  "nodeCount": 9289,
  "edgeCount": 27842,
  "dbSizeBytes": 4194304,
  "backend": "node-sqlite",            // sole backend (Node ≥22.5 built-in)
  "journalMode": "wal",                // 'wal' ⇒ readers never block on writer
  "nodesByKind": { "function": 4128, "class": 211, ... },
  "languages": ["typescript", "javascript"],
  "pendingChanges": { "added": 0, "modified": 3, "removed": 0 },
  "worktreeMismatch": null,            // or { worktreeRoot, indexRoot }
  "index": {
    "builtWithVersion": "1.13.0",      // may be null on a fresh DB
    "builtWithExtractionVersion": 24,
    "currentExtractionVersion": 25,
    "reindexRecommended": false,
    "reindexReasons": [],              // ['extraction-version', ...] when non-empty
    "state": "complete",               // 'complete' | 'partial' | 'indexing' | 'failed'
    "pendingRefs": 0                   // non-zero at rest ⇒ interrupted resolution
  }
}
```

Useful consumer queries:

- "Is this project indexed and healthy?" — read `initialized`, `index.state`, `pendingRefs`.
- "Does the index need a rebuild?" — `index.reindexRecommended` is the boolean; `reindexReasons` is the structured why.

### 3.2 `query --json <search> [-p path] [-l limit] [-k kind]`

Array of `{ node, score }` objects (ranked by BM25/FTS relevance; the
`score` field is unbounded and ranking-only — do NOT interpret as a
percentage):

```jsonc
[
  {
    "node": {
      "id": "function:src/auth.ts:loginUser:42",
      "kind": "function",
      "name": "loginUser",
      "qualifiedName": "auth.loginUser",
      "filePath": "src/auth.ts",
      "language": "typescript",
      "startLine": 42, "endLine": 78,
      "startColumn": 1, "endColumn": 2,
      "signature": "function loginUser(email: string, password: string): Promise<Session>",
      "visibility": "public",
      "isExported": true, "isAsync": true, "isStatic": false, "isAbstract": false,
      "updatedAt": 1784548247709
    },
    "score": 83.00000211132438
  }
]
```

### 3.3 `callers --json <symbol>` / `callees --json <symbol>`

```jsonc
{
  "symbol": "loginUser",
  "callers": [
    {
      "name": "handleLogin", "kind": "function",
      "filePath": "src/routes/session.ts", "startLine": 22,
      "edgeLabel": "calls"
    }
  ]
}
```

`callees` returns the same shape with the inner field named `callees`.

### 3.4 `impact --json <symbol> [-d depth]`

```jsonc
{
  "symbol": "loginUser",
  "depth": 2,
  "nodeCount": 18,
  "edgeCount": 27,
  "affected": [
    { "name": "handleLogin", "kind": "function", "filePath": "src/routes/session.ts", "startLine": 22 }
  ]
}
```

### 3.5 `affected --json [files...] [--stdin]`

```jsonc
{
  "changedFiles": ["src/auth.ts"],
  "affectedTests": ["__tests__/auth.test.ts"],
  "totalDependentsTraversed": 7
}
```

## 4. Stability contract

This section pins what external tools may rely on. The bar is "if you write
a CI script or a tight loop against this surface today, it must keep working
without changes for the lifetime of the v1.x line".

### Public API (stable)

| Surface | Stability guarantee |
|---|---|
| The `bin` name `codegraph-vba` (from `package.json`) | Stable — npm owns the shim convention; we never rename it. |
| The set of top-level subcommands listed in `codegraph --help` | Subcommands are ADDITIVE — new commands may appear, but existing ones will not be removed without a major version bump. A removed subcommand is a breaking change. |
| The `--json` flag on every read-only subcommand | Stable. The JSON top-level keys listed in §3 are stable; existing keys are never removed, only added. |
| The spawn snippet in §2 (Node `child_process.spawn` + platform branch) | Stable — this is the exact pattern the maintainer team uses internally. |
| The presence of `.codegraph-vba/codegraph.db` after `codegraph init` | Stable — the directory + file name are part of the contract. |
| The published schema in [`docs/index-schema.md`](index-schema.md) | **Stable columns** (tagged in the doc) are stable. **Implementation-detail columns** may change between minor releases. |

### Implementation detail (NOT stable)

These surfaces may change between minor releases without notice — do not
build external tools against them:

| Surface | Why it's unstable |
|---|---|
| Subcommand flags other than the documented `--json` | Flags are an internal commander concern; they may be renamed or reorganized. |
| The human-readable prose output of any subcommand | Cosmetic; subject to redesign. |
| The `metadata` JSON shape on individual edges / nodes beyond the top-level fields listed in §3 | Each extractor / synthesizer writes its own `metadata.synthesizedBy` keys; consult the relevant extractor doc (e.g. `docs/vba-stub-repoint-decision.md`) for the per-feature contract. |
| Anything in `package.json` other than the `bin` entry's NAME | Implementation detail of the packaging flow. |
| Subcommands marked `hidden: true` in commander (e.g. `serve`, `prompt-hook`) | Not part of the user-facing contract. `serve --mcp` is for the AI-agent installer; `prompt-hook` is for the Claude Code hook. |

### How to verify a contract change

The drift-detection tests at `__tests__/external-integration-doc.test.ts` and
`__tests__/dump-index-schema.test.ts` pin every claim in §3 and §4 against
the live CLI / DB. If a future PR changes the surface in a way that breaks a
contract, those tests fail at PR time — the maintainer can then either
revert the change or update the contract (with a changelog note) before
merging.

## 5. Cross-references

- [`docs/index-schema.md`](index-schema.md) — Pattern B (direct SQL). Generated
  by `scripts/dump-index-schema.ts`; ships as a committed golden.
- [`docs/vba-stub-repoint-decision.md`](vba-stub-repoint-decision.md) — the
  `metadata.repointDecision` taxonomy, when reading `edges.metadata` from
  Pattern B.
- [`docs/vba-reference-kinds.md`](vba-reference-kinds.md) — the
  `unresolved_refs.reference_kind` taxonomy, when reading `unresolved_refs`
  from Pattern B.
- Sister issue `DysTelefonica/dysflow#1015` — the dysflow side of this fix.
- Issue #200 — this issue.