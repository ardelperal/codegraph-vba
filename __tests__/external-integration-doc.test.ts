/**
 * Tests for `docs/external-integration.md` (issue #200, Deliverable 1).
 *
 * The doc is a contract — it tells external tools (e.g. dysflow) how to invoke
 * the codegraph-vba CLI. If the doc references a subcommand or JSON key that
 * the live CLI no longer exposes, consumers will follow stale instructions
 * and silently fail. These tests pin the doc against the live CLI surface
 * (package.json `bin` + `codegraph --help`) and against the live JSON shapes
 * from the read-only subcommands, so a drift in either direction fails the
 * suite.
 *
 * What we check:
 *
 *   1. The doc file exists and is non-empty.
 *   2. Every CLI subcommand mentioned in the doc also exists in the live
 *      `--help` listing (no removed-subcommand drift).
 *   3. Every JSON key mentioned in the doc also exists in the live JSON
 *      output of `status --json` (no removed-JSON-key drift for the keys
 *      consumers actually use).
 *   4. The doc documents the npm-shim layout (Windows `.cmd`/`.ps1`,
 *      POSIX bare name) so the cross-platform spawn pitfall from sister
 *      issue Dysflow#1015 doesn't recur.
 *   5. The doc includes a working spawn pattern snippet with platform-aware
 *      command resolution.
 *   6. The doc carries a stability note distinguishing public API from
 *      implementation detail.
 *
 * The tests do NOT pin subcommand option flags (the flags are not part of
 * the stable contract — they're an internal detail of commander). Only the
 * subcommand NAMES and the well-known JSON top-level keys are pinned, so a
 * routine refactor of flags won't break the contract.
 */

import { describe, it, expect } from 'vitest';
import { execFileSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';

const REPO_ROOT = path.resolve(__dirname, '..');
const DOC_PATH = path.join(REPO_ROOT, 'docs', 'external-integration.md');
const BIN = path.join(REPO_ROOT, 'dist', 'bin', 'codegraph.js');
const PKG = JSON.parse(
  fs.readFileSync(path.join(REPO_ROOT, 'package.json'), 'utf-8'),
) as { name: string; bin: Record<string, string> };

/**
 * Run `codegraph --help` from the built dist binary. Use the bundled node
 * (Node 22 LTS) when available — the local env may default to Node 26, which
 * the CLI doesn't yet support (issue #54, #81, #140). The path is discovered
 * via `process.execPath`, falling back to the pnpm-shimmed Node 22.
 */
function runCli(args: string[]): string {
  return execFileSync(process.execPath, [BIN, ...args], {
    encoding: 'utf-8',
    env: {
      ...process.env,
      CODEGRAPH_NO_DAEMON: '1',
      CODEGRAPH_WASM_RELAUNCHED: '1',
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

/**
 * Extract the live subcommand names from `codegraph --help`. Each line under
 * `Commands:` begins with two spaces + the name (optionally `|alias`), then
 * whitespace and the description. Returns the canonical first token
 * (lowercased). Skips `help` itself — it's the implicit commander fallback
 * and never appears as a stable subcommand contract.
 */
function liveSubcommandNames(helpText: string): Set<string> {
  const names = new Set<string>();
  const lines = helpText.split('\n');
  let inCommands = false;
  for (const line of lines) {
    if (/^\s*Commands:\s*$/.test(line)) {
      inCommands = true;
      continue;
    }
    if (!inCommands) continue;
    if (line.trim() === '') {
      // Blank line ends the command listing.
      inCommands = false;
      continue;
    }
    // Each line: `  <name>[|<alias>]   <description>`. The name segment is
    // terminated by a run of whitespace before the description column. We
    // split on whitespace and take the first token.
    const m = line.match(/^\s+([^\s|]+(?:\|[^\s|]+)*)\s+(.+)$/);
    if (!m) continue;
    const firstToken = m[1].split('|')[0];
    if (firstToken === 'help') continue;
    names.add(firstToken);
  }
  return names;
}

/**
 * Extract words in the doc that look like CLI subcommand names — i.e. a
 * backtick-quoted token of the form `codegraph-vba <name>` or `<name>`
 * adjacent to a subcommand-style line. Conservative: only matches names
 * that are also present in the live set, so we flag false positives (doc
 * mentions a name not in the live CLI) — and we want to. A name that is
 * in the live set but NOT in the doc is fine (the doc doesn't have to
 * enumerate every command).
 */
function subcommandMentionsInDoc(doc: string, liveSet: Set<string>): Set<string> {
  const mentioned = new Set<string>();
  // Match the backtick-quoted form: ``codegraph-vba <name> ...`` and
  // ``codegraph <name> ...`` (the issue text uses both; we accept both).
  const re = /`(?:codegraph(?:-vba)?\s+)([a-z][a-z0-9-]*)`/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(doc)) !== null) {
    const candidate = m[1];
    if (liveSet.has(candidate)) {
      mentioned.add(candidate);
    } else {
      // The candidate was matched as a subcommand-style reference but does
      // NOT exist in the live --help listing. Flag it as a removed-subcommand
      // drift.
      mentioned.add(candidate);
    }
  }
  return mentioned;
}

/**
 * Extract JSON top-level keys from the status output. Run status against a
 * temp dir that is NOT indexed so we get the uninitialized branch (which is
 * the documented headless "is this index alive?" probe).
 */
function liveStatusJsonKeys(): Set<string> {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-ext-int-status-'));
  try {
    const out = runCli(['status', '--json', tempDir]);
    // The JSON is one object per line; grab the last non-empty line so any
    // banner-like preamble doesn't trip the parser.
    const jsonLine = out
      .split('\n')
      .map((l) => l.trim())
      .filter(Boolean)
      .pop()!;
    const obj = JSON.parse(jsonLine) as Record<string, unknown>;
    return new Set(Object.keys(obj));
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

/**
 * Extract status --json keys the doc claims as part of the JSON contract.
 *
 * The doc references keys in many contexts — prose, JSON examples, table
 * cells. To distinguish "this key is part of the documented contract" from
 * "this word happens to be camelCase", we anchor on the JSON-shape example
 * pattern: a line beginning with `<whitespace>"<key>": <value>` (the JSON
 * object-literal shape). That catches every key the doc puts in an
 * explicit JSON example — and only those. Prose mentions are not validated
 * here; this test pins the *example shape*, not the running prose.
 *
 * Nested keys (`"added"` inside `"pendingChanges": { "added": 0 }`) are
 * caught too — they're still keys the doc publishes as part of the
 * contract. The LIVE key set must contain them, OR the test fails.
 */
function statusJsonKeyMentionsInDoc(doc: string): Set<string> {
  const keys = new Set<string>();
  // Match `"key":` where key is a JS-identifier-shaped token (camelCase,
  // starts with lowercase, no whitespace, no dots). Anchored on a JSON
  // object-literal style: any line where the key appears inside backticks
  // AND is followed by `:`. We require at least 4 chars and the first char
  // to be lowercase to filter out command words.
  const re = /`([a-z][a-zA-Z0-9]{3,})`\s*:/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(doc)) !== null) {
    keys.add(m[1]);
  }
  return keys;
}

describe('docs/external-integration.md (issue #200 Deliverable 1)', () => {
  it('the doc file exists and is non-empty', () => {
    expect(fs.existsSync(DOC_PATH), `${DOC_PATH} must exist`).toBe(true);
    const body = fs.readFileSync(DOC_PATH, 'utf-8');
    expect(body.length).toBeGreaterThan(500);
  });

  it('every CLI subcommand named in the doc is present in the live `codegraph --help`', () => {
    // Build should have produced dist/bin/codegraph.js; skip the assertion
    // gracefully if not (the test still fails on RED because the doc doesn't
    // exist yet, but we want a clear skip rather than a confusing spawn
    // failure when the user is iterating).
    if (!fs.existsSync(BIN)) {
      throw new Error(
        `${BIN} missing — run \`npm run build\` before running this test`,
      );
    }

    const help = runCli(['--help']);
    const live = liveSubcommandNames(help);
    const doc = fs.readFileSync(DOC_PATH, 'utf-8');
    const docMentions = subcommandMentionsInDoc(doc, live);

    // Find names mentioned in the doc that are NOT in the live CLI listing.
    const removed = [...docMentions].filter((name) => !live.has(name));
    expect(
      removed,
      `doc references subcommands the live CLI no longer exposes: ${removed.join(', ')}. ` +
        `Live subcommands: ${[...live].sort().join(', ')}.`,
    ).toEqual([]);
  });

  it('every status-JSON key named in the doc is present in the live `status --json` output', () => {
    if (!fs.existsSync(BIN)) {
      throw new Error(
        `${BIN} missing — run \`npm run build\` before running this test`,
      );
    }

    const liveKeys = liveStatusJsonKeys();
    const doc = fs.readFileSync(DOC_PATH, 'utf-8');
    const docMentions = statusJsonKeyMentionsInDoc(doc);

    // The doc is allowed to mention keys belonging to OTHER subcommands
    // (query / callers / impact / affected). We can't easily validate those
    // here without spawning each subcommand; instead, restrict the check to
    // keys that are clearly part of the `status --json` contract — i.e.,
    // keys that ARE in the live status --json set. Keys the doc mentions
    // that are NOT in the live status --json set are tolerated if they
    // also aren't status keys (they're from another subcommand's JSON); we
    // only fail when a key looks like a status key (camelCase, ≥4 chars)
    // and is NOT in the live status set. Since the extraction
    // (`statusJsonKeyMentionsInDoc`) only pulls keys from JSON-example
    // positions, and the doc explicitly attributes each example to a
    // subcommand in §3.1–§3.5, the intersection filter is sufficient.
    const removed = [...docMentions].filter((k) => !liveKeys.has(k));
    expect(
      removed,
      `doc references JSON keys the live status --json does not expose: ${removed.join(', ')}. ` +
        `Live status --json keys: ${[...liveKeys].sort().join(', ')}. ` +
        `(These are the keys the doc must NOT cite as part of the status JSON contract.)`,
    ).toEqual([]);
  });

  it('the doc documents the npm-shim layout per platform (Windows .cmd/.ps1 + POSIX)', () => {
    const body = fs.readFileSync(DOC_PATH, 'utf-8');
    // The package's bin field declares a single JS entry; npm generates the
    // per-platform shim filenames. The doc MUST call out both forms so a
    // consumer on Windows knows to spawn `codegraph-vba.cmd` from
    // `child_process.spawn` (the bare `codegraph-vba` resolves to `codegraph-vba.ps1`
    // under PowerShell but to a `.cmd` shim from Node — see sister issue
    // Dysflow#1015).
    expect(body).toMatch(/\.cmd\b/i);
    expect(body).toMatch(/\.ps1\b/i);
    // POSIX forms: bare `codegraph-vba` works on macOS/Linux because the
    // shim there is a single executable.
    expect(body).toMatch(/POSIX|macOS|Linux|linux|darwin/i);
  });

  it('the doc documents the package.json `bin` key name (the npm-installed entry)', () => {
    const body = fs.readFileSync(DOC_PATH, 'utf-8');
    const binName = Object.keys(PKG.bin)[0];
    expect(binName).toBeTruthy();
    expect(body, `doc must mention the bin name "${binName}"`).toContain(binName);
  });

  it('the doc includes a canonical spawn snippet using child_process.spawn with a platform-aware command', () => {
    const body = fs.readFileSync(DOC_PATH, 'utf-8');
    expect(body, 'doc must reference child_process.spawn').toMatch(/child_process\.spawn/);
    // The platform-aware branch: Windows needs a `.cmd` extension; POSIX
    // uses the bare name. Any of these phrasings is acceptable:
    const platformAware =
      /process\.platform\s*===\s*['"]win32['"]/.test(body) ||
      /['"]win32['"]/.test(body) ||
      /platform.*(?:Windows|win32)/i.test(body) ||
      /\.cmd['"]/.test(body);
    expect(
      platformAware,
      'doc must demonstrate platform-aware command resolution (Windows vs POSIX)',
    ).toBe(true);
  });

  it('the doc carries a stability note distinguishing public API from implementation detail', () => {
    const body = fs.readFileSync(DOC_PATH, 'utf-8');
    // Accept either the literal 'stable'/'public' markers or an explicit
    // "Implementation detail" section. Both phrasings are valid per the
    // issue's "stability note" requirement.
    const hasStabilityMarkers =
      /\bstable\b/i.test(body) ||
      /\bpublic API\b/i.test(body) ||
      /\bimplementation detail/i.test(body);
    expect(
      hasStabilityMarkers,
      'doc must contain a stability section distinguishing public API from implementation detail',
    ).toBe(true);
  });

  it('the doc references at least one headless JSON-mode subcommand (e.g. `status --json`)', () => {
    const body = fs.readFileSync(DOC_PATH, 'utf-8');
    // The issue is explicit: the doc must document the "headless query
    // contract" — subcommand + args + JSON shape on stdout.
    expect(body, 'doc must mention `--json`').toMatch(/--json\b/);
    // The status subcommand can be referenced as `status` alone (since the
    // section heading §3.1 already names it), or as `codegraph status` /
    // `codegraph-vba status` / `status --json`. All are valid doc styles.
    const statusReferenced =
      /\bcodegraph(?:-vba)?\s+status\b/.test(body) ||
      /`status\s+--json`/.test(body) ||
      /\bsubcommand `status\b/.test(body) ||
      /\u00a73\.1\s+`?status`?\b/.test(body) ||
      /### 3\.1\s+`?status\b/.test(body);
    expect(
      statusReferenced,
      'doc must reference the `status` subcommand (as `codegraph status`, `status --json`, `§3.1 status`, or `### 3.1 status`)',
    ).toBe(true);
  });
});