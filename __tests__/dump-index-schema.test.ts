/**
 * Tests for `scripts/dump-index-schema.ts` + the generated
 * `docs/index-schema.md` (issue #200, Deliverable 2).
 *
 * The schema doc is a generated artifact — its source of truth is the live
 * SQLite file produced by `src/db/schema.sql`. These tests pin the contract
 * so that:
 *
 *   1. The dump script exists and is runnable via `tsx`.
 *   2. The script produces a markdown file at `docs/index-schema.md`.
 *   3. Every table and every column present in the live schema (queried via
 *      `sqlite_master` + `PRAGMA table_info`) is also present in the doc.
 *   4. The canonical tables (`nodes`, `edges`, `unresolved_refs`) and their
 *      well-known "stable" columns are tagged as **stable** in the doc — the
 *      whole point of Deliverable 2 is to publish a stable contract for
 *      external tools.
 *
 * The dump is tested against a FRESH temp DB created from the canonical
 * `src/db/schema.sql` (NOT against the user's real `.codegraph-vba/codegraph.db`),
 * so the test is hermetic — it never reads or mutates project state.
 */

import { describe, it, expect, beforeAll } from 'vitest';
import { execFileSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

const REPO_ROOT = path.resolve(__dirname, '..');
const DOC_PATH = path.join(REPO_ROOT, 'docs', 'index-schema.md');
const SCRIPT_PATH = path.join(REPO_ROOT, 'scripts', 'dump-index-schema.ts');
const SCHEMA_SQL_PATH = path.join(REPO_ROOT, 'src', 'db', 'schema.sql');

/**
 * Canonical (issue-#200) "stable" tables — the ones the lookup queries
 * external tools depend on. The issue explicitly names `nodes`, `edges`,
 * and `unresolved_refs`. If a future PR drops one of these from the
 * contract surface, this test will fail and force an explicit decision.
 */
const CANONICAL_TABLES = ['nodes', 'edges', 'unresolved_refs'] as const;

/**
 * Canonical "stable" columns per table — the minimal contract an external
 * SQL consumer can rely on. Columns not in this list are still part of the
 * schema, but are NOT part of the stable contract (they may be added /
 * removed in any release without breaking the contract).
 *
 * The set below is intentionally conservative: it lists the columns a
 * basic node / edge / unresolved-ref query must read to be useful. Adding
 * columns here is a deliberate, reviewable change.
 */
const STABLE_COLUMNS: Record<string, readonly string[]> = {
  nodes: [
    'id',
    'kind',
    'name',
    'qualified_name',
    'file_path',
    'language',
    'start_line',
    'end_line',
    'metadata',
  ],
  edges: ['id', 'source', 'target', 'kind', 'metadata', 'provenance'],
  unresolved_refs: [
    'id',
    'from_node_id',
    'reference_name',
    'reference_kind',
    'line',
    'col',
    'status',
    'metadata',
  ],
};

/**
 * Run the dump script with `--out <path>` pointing at a temp file. The
 * script must accept a `--out` arg so the test can capture its output
 * without polluting the committed `docs/index-schema.md`.
 *
 * We invoke `npx tsx` (same command the `npm run schema:dump` script uses)
 * so the test pins the same execution path a real consumer would hit. If
 * a tsx binary exists in node_modules/.bin we prefer it for speed, but we
 * always have a working fallback via `npx`.
 */
function runDumpScript(outPath: string): void {
  const localTsxCmd = path.join(REPO_ROOT, 'node_modules', '.bin', 'tsx.cmd');
  const localTsx = path.join(REPO_ROOT, 'node_modules', '.bin', 'tsx');

  if (fs.existsSync(localTsxCmd) || fs.existsSync(localTsx)) {
    const tsxBin = fs.existsSync(localTsxCmd) ? localTsxCmd : localTsx;
    execFileSync(tsxBin, [SCRIPT_PATH, '--out', outPath], {
      cwd: REPO_ROOT,
      encoding: 'utf-8',
      stdio: ['ignore', 'pipe', 'pipe'],
      // Windows .cmd shims must run through a shell — Node refuses to
      // spawn them directly since the CVE-2024-27980 mitigation
      // (regression that motivated the npx fallback below keeping
      // `shell: true` on win32).
      shell: process.platform === 'win32',
    });
    return;
  }

  // Fallback: shell out to `npx tsx` which will resolve via npm's bin
  // cache. This mirrors what a fresh consumer running `npm run schema:dump`
  // would see (the eval script in package.json does the same). The
  // `shell: true` flag is required on Windows because `npx.cmd` is a batch
  // shim and Node's `execFileSync` rejects it without a shell.
  const npxCmd = process.platform === 'win32' ? 'npx.cmd' : 'npx';
  execFileSync(npxCmd, ['tsx', SCRIPT_PATH, '--out', outPath], {
    cwd: REPO_ROOT,
    encoding: 'utf-8',
    stdio: ['ignore', 'pipe', 'pipe'],
    shell: process.platform === 'win32',
  });
}

interface ColumnInfo {
  cid: number;
  name: string;
  type: string;
  notnull: number;
  dflt_value: unknown;
  pk: number;
}

interface TableInfo {
  name: string;
  columns: ColumnInfo[];
}

/**
 * Open a fresh `node:sqlite` database, apply the canonical schema, and
 * return the live table/column metadata. This is what the dump script must
 * produce output FOR — by reading the same `sqlite_master` + `PRAGMA`
 * source, the dump is pinned to the live schema.
 */
function liveSchema(): { tables: Map<string, TableInfo>; order: string[] } {
  // node:sqlite is loaded via the same require-dynamic trick the production
  // code uses (it's a built-in; require works at runtime but tsc doesn't see
  // the type when we go through `require` directly).
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const { DatabaseSync } = require('node:sqlite') as typeof import('node:sqlite');

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-dump-schema-'));
  const dbPath = path.join(tempDir, 'live.db');
  const db = new DatabaseSync(dbPath);
  const schemaSql = fs.readFileSync(SCHEMA_SQL_PATH, 'utf-8');
  db.exec(schemaSql);

  const tableRows = db
    .prepare(
      `SELECT name FROM sqlite_master
       WHERE type = 'table' AND name NOT LIKE 'sqlite_%'
       ORDER BY name`,
    )
    .all() as Array<{ name: string }>;
  const tables = new Map<string, TableInfo>();
  const order: string[] = [];
  for (const { name } of tableRows) {
    const cols = db.prepare(`PRAGMA table_info(${name})`).all() as ColumnInfo[];
    tables.set(name, { name, columns: cols });
    order.push(name);
  }
  db.close();
  fs.rmSync(tempDir, { recursive: true, force: true });
  return { tables, order };
}

/**
 * Parse the table/column sections of the generated markdown. The dump
 * script's output format is described in `scripts/dump-index-schema.ts`:
 * one `## <table>` heading per table, then a table whose first column is
 * the column name and second column carries a Stability tag (either
 * "stable" or "implementation detail"). We don't pin the exact prose —
 * just the column list and the stability tag for each column.
 */
function parseDocTables(markdown: string): Map<string, { stableColumns: Set<string> }> {
  const result = new Map<string, { stableColumns: Set<string> }>();
  const lines = markdown.split('\n');
  let currentTable: string | null = null;
  // Detect the markdown-table header row. The dump script writes a pipe-
  // delimited table with a `Stability` column. We look for a row like
  // `| Column | Type | Stability | … |` and capture the column indices.
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const h2 = line.match(/^##\s+`?([a-zA-Z_][a-zA-Z0-9_]*)`?\s*$/);
    if (h2) {
      currentTable = h2[1];
      if (!result.has(currentTable)) {
        result.set(currentTable, { stableColumns: new Set() });
      }
      continue;
    }
    if (!currentTable) continue;
    // Header row: find the column named "Stability" and "Column" (or
    // "Name"). The dump script uses `Column` as the column-name header.
    const headerMatch = line.match(/^\|\s*([^|]+)\s*\|\s*([^|]+)\s*\|\s*([^|]+)\s*\|/);
    if (
      headerMatch &&
      /^column$/i.test(headerMatch[1].trim()) &&
      /^stability$/i.test(headerMatch[3].trim())
    ) {
      // Next non-separator line(s) are the data rows.
      let j = i + 1;
      // Skip the markdown separator line (`|---|---|---|`).
      while (j < lines.length && /^\|[\s-:|]+\|$/.test(lines[j])) j++;
      for (; j < lines.length; j++) {
        const row = lines[j];
        if (!row.startsWith('|')) break;
        const cells = row
          .split('|')
          .map((c) => c.trim())
          .filter((c) => c !== '');
        if (cells.length < 3) continue;
        const colName = cells[0].replace(/`/g, '');
        const stability = cells[2].toLowerCase();
        if (colName === 'Column' || colName === '') continue;
        const entry = result.get(currentTable)!;
        if (stability.includes('stable')) {
          entry.stableColumns.add(colName);
        }
      }
    }
  }
  return result;
}

describe('scripts/dump-index-schema.ts + docs/index-schema.md (issue #200 Deliverable 2)', () => {
  let live: ReturnType<typeof liveSchema>;
  let doc: string;
  let parsed: ReturnType<typeof parseDocTables>;

  beforeAll(() => {
    // Read both the live schema and the committed doc ONCE for the suite so
    // every assertion can compare against the same snapshot. The live
    // schema is hermetic (temp DB); the committed doc is the artifact under
    // test.
    live = liveSchema();
    if (!fs.existsSync(DOC_PATH)) {
      doc = '';
      parsed = new Map();
      return;
    }
    doc = fs.readFileSync(DOC_PATH, 'utf-8');
    parsed = parseDocTables(doc);
  });

  it('the dump script and committed doc both exist', () => {
    expect(fs.existsSync(SCRIPT_PATH), `${SCRIPT_PATH} must exist`).toBe(true);
    expect(fs.existsSync(DOC_PATH), `${DOC_PATH} must exist`).toBe(true);
  });

  it('the committed docs/index-schema.md lists every table in the live schema', () => {
    const missing: string[] = [];
    for (const name of live.order) {
      if (!parsed.has(name)) missing.push(name);
    }
    expect(
      missing,
      `doc is missing tables that exist in the live schema: ${missing.join(', ')}. ` +
        `Live tables: ${live.order.join(', ')}.`,
    ).toEqual([]);
  });

  it('the committed docs/index-schema.md lists every column of every canonical table', () => {
    const failures: string[] = [];
    for (const tableName of CANONICAL_TABLES) {
      const liveTable = live.tables.get(tableName);
      if (!liveTable) {
        failures.push(`canonical table ${tableName} missing from live schema`);
        continue;
      }
      const docTable = parsed.get(tableName);
      if (!docTable) {
        failures.push(`canonical table ${tableName} missing from doc`);
        continue;
      }
      const liveColNames = new Set(liveTable.columns.map((c) => c.name));
      const docStableNames = docTable.stableColumns;
      // Every live column must be SOMEWHERE in the doc's table — but the
      // script tags some as "implementation detail" rather than "stable".
      // We accept both: a column is present iff it's tagged either stable
      // OR implementation-detail. Re-parse the raw doc to enumerate all
      // column names for this check.
      const allDocCols = allColumnsInTable(doc, tableName);
      for (const col of liveColNames) {
        if (!allDocCols.has(col)) {
          failures.push(`${tableName}.${col} present in live schema but missing from doc`);
        }
      }
    }
    expect(failures, failures.join('\n')).toEqual([]);
  });

  it('every canonical stable column is tagged "stable" in the doc', () => {
    const failures: string[] = [];
    for (const [tableName, expectedCols] of Object.entries(STABLE_COLUMNS)) {
      const docTable = parsed.get(tableName);
      if (!docTable) {
        failures.push(`canonical table ${tableName} missing from doc`);
        continue;
      }
      for (const col of expectedCols) {
        if (!docTable.stableColumns.has(col)) {
          failures.push(
            `${tableName}.${col} must be tagged "stable" in docs/index-schema.md`,
          );
        }
      }
    }
    expect(failures, failures.join('\n')).toEqual([]);
  });

  it('the dump script runs end-to-end via tsx and produces a non-empty markdown file', () => {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'codegraph-dump-run-'));
    const outPath = path.join(tempDir, 'schema.md');
    try {
      runDumpScript(outPath);
      expect(fs.existsSync(outPath), 'dump script must create the output file').toBe(true);
      const produced = fs.readFileSync(outPath, 'utf-8');
      expect(produced.length).toBeGreaterThan(200);
      // The output must include at least one H2 heading for each canonical
      // table — a silent failure (empty file or only-prose output) is the
      // worst kind of regression.
      for (const tableName of CANONICAL_TABLES) {
        expect(
          produced,
          `dump output must include a ## ${tableName} section`,
        ).toMatch(new RegExp(`^##\\s+\`?${tableName}\`?\\s*$`, 'm'));
      }
    } finally {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  });

  it('the schema:dump npm script is wired in package.json', () => {
    const pkg = JSON.parse(
      fs.readFileSync(path.join(REPO_ROOT, 'package.json'), 'utf-8'),
    ) as { scripts?: Record<string, string> };
    const dumpScript = pkg.scripts?.['schema:dump'];
    expect(
      dumpScript,
      'package.json must declare a `schema:dump` script',
    ).toBeTruthy();
    expect(dumpScript).toMatch(/dump-index-schema/);
  });
});

/**
 * Helper: extract every column name mentioned in the doc's table section
 * for `tableName` — both "stable" and "implementation detail" rows. The
 * parser above only collects "stable"; this one is the superset.
 */
function allColumnsInTable(doc: string, tableName: string): Set<string> {
  const out = new Set<string>();
  const lines = doc.split('\n');
  let inTable = false;
  let inData = false;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (new RegExp(`^##\\s+\`?${tableName}\`?\\s*$`).test(line)) {
      inTable = true;
      continue;
    }
    if (inTable && /^##\s+/.test(line)) {
      inTable = false;
      continue;
    }
    if (!inTable) continue;
    const headerMatch = line.match(/^\|\s*([^|]+)\s*\|\s*([^|]+)\s*\|\s*([^|]+)\s*\|/);
    if (headerMatch && /^column$/i.test(headerMatch[1].trim())) {
      inData = true;
      continue;
    }
    if (inData && /^\|[\s-:|]+\|$/.test(line)) continue;
    if (inData && line.startsWith('|')) {
      const cells = line
        .split('|')
        .map((c) => c.trim())
        .filter((c) => c !== '');
      if (cells.length >= 3 && cells[0] !== 'Column') {
        out.add(cells[0].replace(/`/g, ''));
      }
    } else if (inData && !line.startsWith('|')) {
      inData = false;
    }
  }
  return out;
}