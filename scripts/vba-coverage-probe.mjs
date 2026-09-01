/**
 * VBA coverage probe — the committed measurement tool behind every
 * acceptance criterion in `docs/vba-node-discovery-plan.md`.
 *
 * Every task in the VBA node-discovery roadmap states its acceptance as a
 * number ("stub nodes drop by >= 3,800", "table references rise"). This
 * script is the only supported way to produce those numbers, so a PR can
 * prove it landed a win instead of asserting one.
 *
 * Usage (from the repo root):
 *   npm run probe:vba -- <root> [<root> ...]
 *   npx tsx scripts/vba-coverage-probe.mjs <root> [<root> ...]
 *   npx tsx scripts/vba-coverage-probe.mjs --json <root> [<root> ...]
 *
 * It imports `VbaExtractor`, `VbaFormExtractor` and `SqlQueryExtractor`
 * DIRECTLY. No index, no SQLite, no CLI, no `codegraph.json`. It walks the
 * given roots, dispatches each file by extension, aggregates, and prints a
 * human-readable report (or JSON with `--json`).
 *
 * ## The stub discriminator — read this before changing anything
 *
 * A synthesized call target and a declared procedure are BOTH
 * `kind: 'function'` and BOTH carry `visibility: 'public'`. The ONLY
 * discriminator is `metadata.stub === true`, set in
 * `src/extraction/vba/calls.ts`.
 *
 * Without it the reference corpus reads 9,972 procedures instead of 3,840
 * — a 2.6x inflation that silently corrupts every later measurement.
 *
 * ## Reported fields
 *
 * Adding fields later is fine; REMOVING one breaks comparability with
 * earlier runs. The contract is:
 *
 *   files                     total files dispatched, by extension
 *   nodesByKind               { [kind]: count }
 *   declaredProcedures        function nodes WITHOUT metadata.stub
 *   stubProcedures            function nodes WITH metadata.stub
 *   stubTargetsTop            top 30 stub node names by count
 *   edgesByKind               { [kind]: count }
 *   edgesBySynthesizer        { [metadata.synthesizedBy]: count }
 *   unresolvedByKind          { [referenceKind]: count }
 *   sqlTablesReferenced       distinct table names reached via vba-sql-table
 *   formsReferenced           distinct form names reached via opens-form / form refs
 *   proceduresWithNoOutgoing  declared procedures with zero edges and zero unresolved refs
 *   errors                    extraction errors by code
 *
 * `files` additionally reports `total` (every regular file walked, including
 * ones no extractor claims) and `skipped` so the walked-vs-dispatched split
 * is visible rather than inferred.
 */
import * as fs from 'fs';
import * as path from 'path';
import { pathToFileURL } from 'url';

/** `.form.txt` / `.report.txt` — Dysflow SaveAsText UI exports. */
const FORM_FILE_RE = /\.(form|report)\.txt$/i;
/** `.bas` / `.cls` / `.frm` / `.dsr` — VBA code modules. */
const CODE_FILE_RE = /\.(bas|cls|frm|dsr)$/i;
/** `queries/<Name>.sql` — Dysflow-exported saved Access queries. */
const SQL_FILE_RE = /\.sql$/i;

/** How many stub target names `stubTargetsTop` reports. */
const STUB_TARGETS_TOP_N = 30;

/**
 * The `metadata.synthesizedBy` tag `src/extraction/vba/sql-wrapper.ts` puts
 * on every `references` edge that reaches a table named inside VBA-embedded
 * SQL. `sqlTablesReferenced` counts distinct targets of exactly these edges.
 */
const SQL_TABLE_SYNTHESIZER = 'vba-sql-table';
/**
 * The `metadata.synthesizedBy` tag `src/extraction/vba-form-extractor.ts`
 * puts on the `.form.txt` -> sibling `.cls` unresolved reference. Together
 * with `opens-form` edges this is what `formsReferenced` counts.
 */
const FORM_BINDING_SYNTHESIZER = 'vba-form-binding';

/**
 * Two-segment extension for `.form.txt` / `.report.txt`, plain
 * `path.extname()` otherwise. `path.extname('X.form.txt')` is `.txt`, which
 * would collapse forms and reports into one meaningless bucket.
 */
export function fileExtensionKey(filePath) {
  const base = path.basename(filePath).toLowerCase();
  const twoSegment = base.match(/(\.(?:form|report)\.txt)$/);
  if (twoSegment) return twoSegment[1];
  const ext = path.extname(base);
  return ext === '' ? '(none)' : ext;
}

/** Which of the three extractors claims this file, or `null`. */
export function dispatchFor(filePath) {
  if (FORM_FILE_RE.test(filePath)) return 'form';
  if (CODE_FILE_RE.test(filePath)) return 'code';
  if (SQL_FILE_RE.test(filePath)) return 'sql';
  return null;
}

/**
 * Walk `roots` and return every regular file, sorted, deduplicated. Sorting
 * makes the aggregate order-independent numbers reproducible run to run and
 * keeps `stubTargetsTop` tie-breaking stable.
 */
export function collectFiles(roots) {
  const seen = new Set();
  for (const root of roots) {
    const stat = fs.statSync(root);
    if (stat.isFile()) {
      seen.add(path.resolve(root));
      continue;
    }
    const stack = [path.resolve(root)];
    while (stack.length > 0) {
      const dir = stack.pop();
      for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const p = path.join(dir, entry.name);
        if (entry.isDirectory()) stack.push(p);
        else if (entry.isFile()) seen.add(p);
      }
    }
  }
  return [...seen].sort();
}

/**
 * Resolve the three extractor classes. `src/` is the source of truth — a
 * stale `dist/` would silently change the measurement — so it wins whenever
 * it is present. `dist/` is the fallback for an installed package, where
 * `src/` does not ship.
 */
export async function loadExtractors(repoRoot) {
  const srcDir = path.join(repoRoot, 'src', 'extraction');
  const distDir = path.join(repoRoot, 'dist', 'extraction');
  const useSrc = fs.existsSync(path.join(srcDir, 'vba-extractor.ts'));
  const base = useSrc ? srcDir : distDir;
  const ext = useSrc ? '.ts' : '.js';
  const load = async (name) =>
    import(pathToFileURL(path.join(base, name + ext)).href);
  const [code, form, sql] = await Promise.all([
    load('vba-extractor'),
    load('vba-form-extractor'),
    load('sql-query-extractor'),
  ]);
  return {
    VbaExtractor: code.VbaExtractor,
    VbaFormExtractor: form.VbaFormExtractor,
    SqlQueryExtractor: sql.SqlQueryExtractor,
  };
}

function bump(map, key) {
  map.set(key, (map.get(key) ?? 0) + 1);
}

/** Map -> plain object, sorted by count desc then key asc. Deterministic. */
function sortedObject(map) {
  const out = {};
  for (const [k, v] of [...map.entries()].sort(
    (a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0),
  )) {
    out[k] = v;
  }
  return out;
}

/**
 * Run the probe over `roots`.
 *
 * @param {string[]} roots
 * @param {{ extractors?: object, repoRoot?: string }} [options]
 *   `extractors` injects the three classes directly (the test does this so
 *   it exercises the same aggregation the CLI does without a second module
 *   resolution path). Otherwise they are loaded from `repoRoot`.
 */
export async function runProbe(roots, options = {}) {
  const extractors =
    options.extractors ??
    (await loadExtractors(options.repoRoot ?? process.cwd()));
  const { VbaExtractor, VbaFormExtractor, SqlQueryExtractor } = extractors;

  const filesByExt = new Map();
  const skippedByExt = new Map();
  const nodesByKind = new Map();
  const edgesByKind = new Map();
  const edgesBySynthesizer = new Map();
  const unresolvedByKind = new Map();
  const stubTargets = new Map();
  const errorsByCode = new Map();
  const sqlTables = new Set();
  const forms = new Set();

  let dispatched = 0;
  let skipped = 0;
  let declaredProcedures = 0;
  let stubProcedures = 0;
  let proceduresWithNoOutgoing = 0;

  for (const absPath of collectFiles(roots)) {
    const extKey = fileExtensionKey(absPath);
    const dispatch = dispatchFor(absPath);
    if (dispatch === null) {
      skipped += 1;
      bump(skippedByExt, extKey);
      continue;
    }
    dispatched += 1;
    bump(filesByExt, extKey);

    let result;
    try {
      const source = fs.readFileSync(absPath, 'utf8');
      if (dispatch === 'form') {
        result = new VbaFormExtractor(absPath, source).extract();
      } else if (dispatch === 'sql') {
        result = new SqlQueryExtractor(absPath, source).extract();
      } else {
        result = new VbaExtractor(absPath, source).extract();
      }
    } catch {
      bump(errorsByCode, 'probe_extractor_threw');
      continue;
    }

    // Node id -> name, so edges (which carry ids) can be attributed to the
    // name a human reads in the roadmap tables.
    const nodeNameById = new Map();
    for (const node of result.nodes) nodeNameById.set(node.id, node.name);

    for (const node of result.nodes) {
      bump(nodesByKind, node.kind);
      if (node.kind !== 'function') continue;
      // THE stub discriminator. See the header comment.
      if (node.metadata?.stub === true) {
        stubProcedures += 1;
        bump(stubTargets, node.name);
      } else {
        declaredProcedures += 1;
      }
    }

    const outgoing = new Set();
    for (const edge of result.edges) {
      bump(edgesByKind, edge.kind);
      bump(edgesBySynthesizer, edge.metadata?.synthesizedBy ?? '(none)');
      if (edge.source) outgoing.add(edge.source);
      if (edge.metadata?.synthesizedBy === SQL_TABLE_SYNTHESIZER) {
        const name = nodeNameById.get(edge.target);
        if (name) sqlTables.add(name);
      }
      if (edge.kind === 'opens-form') {
        const name =
          edge.metadata?.targetFormName ?? nodeNameById.get(edge.target);
        if (name) forms.add(name);
      }
    }

    for (const ref of result.unresolvedReferences) {
      bump(unresolvedByKind, ref.referenceKind);
      if (ref.fromNodeId) outgoing.add(ref.fromNodeId);
      if (ref.metadata?.synthesizedBy === FORM_BINDING_SYNTHESIZER) {
        forms.add(ref.referenceName);
      }
    }

    for (const node of result.nodes) {
      if (node.kind !== 'function') continue;
      if (node.metadata?.stub === true) continue;
      if (!outgoing.has(node.id)) proceduresWithNoOutgoing += 1;
    }

    for (const error of result.errors) {
      bump(errorsByCode, error.code ?? '(none)');
    }
  }

  const stubTargetsTop = [...stubTargets.entries()]
    .sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0))
    .slice(0, STUB_TARGETS_TOP_N)
    .map(([name, count]) => ({ name, count }));

  return {
    roots: roots.map((r) => path.resolve(r)),
    files: {
      total: dispatched + skipped,
      dispatched,
      skipped,
      byExtension: sortedObject(filesByExt),
      skippedByExtension: sortedObject(skippedByExt),
    },
    nodesByKind: sortedObject(nodesByKind),
    declaredProcedures,
    stubProcedures,
    stubTargetsTop,
    edgesByKind: sortedObject(edgesByKind),
    edgesBySynthesizer: sortedObject(edgesBySynthesizer),
    unresolvedByKind: sortedObject(unresolvedByKind),
    sqlTablesReferenced: { count: sqlTables.size, names: [...sqlTables].sort() },
    formsReferenced: { count: forms.size, names: [...forms].sort() },
    proceduresWithNoOutgoing,
    errors: sortedObject(errorsByCode),
  };
}

function table(title, obj, keyHeader) {
  const lines = [`### ${title}`, '', `| ${keyHeader} | count |`, '|---|--:|'];
  const entries = Object.entries(obj);
  if (entries.length === 0) lines.push('| _(none)_ | 0 |');
  for (const [k, v] of entries) lines.push(`| \`${k}\` | ${v} |`);
  lines.push('');
  return lines.join('\n');
}

/** Render the report as markdown — the shape pasted into PR bodies. */
export function formatReport(report) {
  const out = ['# VBA coverage probe', ''];
  out.push('Roots:');
  for (const root of report.roots) out.push(`  - \`${root}\``);
  out.push('');

  out.push('### Files', '', '| extension | files |', '|---|--:|');
  for (const [k, v] of Object.entries(report.files.byExtension)) {
    out.push(`| \`${k}\` | ${v} |`);
  }
  for (const [k, v] of Object.entries(report.files.skippedByExtension)) {
    out.push(`| \`${k}\` _(no extractor)_ | ${v} |`);
  }
  out.push(`| **dispatched** | **${report.files.dispatched}** |`);
  out.push(`| **total walked** | **${report.files.total}** |`, '');

  out.push(table('Nodes by kind', report.nodesByKind, 'kind'));

  const fnTotal = report.declaredProcedures + report.stubProcedures;
  out.push(
    '### Procedures',
    '',
    '| | count |',
    '|---|--:|',
    `| declared procedures (no \`metadata.stub\`) | ${report.declaredProcedures} |`,
    `| stub function nodes (\`metadata.stub === true\`) | ${report.stubProcedures} |`,
    `| \`function\` nodes total | ${fnTotal} |`,
    `| declared procedures with no outgoing | ${report.proceduresWithNoOutgoing} |`,
    '',
  );

  out.push(
    `### Top ${STUB_TARGETS_TOP_N} stub targets`,
    '',
    '| stub name | count |',
    '|---|--:|',
  );
  if (report.stubTargetsTop.length === 0) out.push('| _(none)_ | 0 |');
  for (const { name, count } of report.stubTargetsTop) {
    out.push(`| \`${name}\` | ${count} |`);
  }
  out.push('');

  out.push(table('Edges by kind', report.edgesByKind, 'kind'));
  out.push(
    table('Edges by synthesizer', report.edgesBySynthesizer, 'synthesizedBy'),
  );
  out.push(
    table('Unresolved references by kind', report.unresolvedByKind, 'referenceKind'),
  );

  out.push(
    '### Reach',
    '',
    '| | count |',
    '|---|--:|',
    `| distinct SQL tables (\`${SQL_TABLE_SYNTHESIZER}\`) | ${report.sqlTablesReferenced.count} |`,
    `| distinct forms (\`opens-form\` / form refs) | ${report.formsReferenced.count} |`,
    '',
  );

  out.push(table('Extraction errors by code', report.errors, 'code'));
  return out.join('\n');
}

async function main(argv) {
  const asJson = argv.includes('--json');
  const roots = argv.filter((a) => !a.startsWith('--'));
  if (roots.length === 0) {
    console.error(
      'usage: npx tsx scripts/vba-coverage-probe.mjs [--json] <root> [<root> ...]',
    );
    process.exit(2);
  }
  for (const root of roots) {
    if (!fs.existsSync(root)) {
      console.error(`root does not exist: ${root}`);
      process.exit(2);
    }
  }
  const report = await runProbe(roots, { repoRoot: process.cwd() });
  console.log(asJson ? JSON.stringify(report, null, 2) : formatReport(report));
}

// Only run the CLI when invoked as a script, never when imported by a test.
const invokedPath = process.argv[1] ? pathToFileURL(process.argv[1]).href : '';
if (invokedPath === import.meta.url) {
  main(process.argv.slice(2)).catch((err) => {
    console.error(err);
    process.exit(1);
  });
}
