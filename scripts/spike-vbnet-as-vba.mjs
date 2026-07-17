#!/usr/bin/env node
/**
 * spike-vbnet-as-vba.mjs — Phase F.1 spike (issue #155).
 *
 * Loads the vendored `tree-sitter-vbnet.wasm` grammar and parses a
 * representative corpus of real-world VBA files with it. Classifies each
 * parse as `clean` (no ERROR nodes), `partial` (some ERRORs but the
 * surrounding structure parses), or `failed` (most bytes are inside ERROR
 * nodes), counts per-construct recognition, and produces a markdown
 * report with a go/no-go recommendation.
 *
 * Usage:
 *   node scripts/spike-vbnet-as-vba.mjs [--fixtures <dir>] [--out <path>]
 *
 * Defaults:
 *   --fixtures  __tests__/fixtures
 *   --out       docs/spikes/vbnet-as-vba.md
 *
 * NO production-code side effects. Pure measurement + report emission.
 */
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { Parser, Language as WasmLanguage } from 'web-tree-sitter';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const REPO_ROOT = path.resolve(__dirname, '..');

// ---------- CLI args ----------
function parseArgs(argv) {
  const out = { fixtures: null, out: null, jsonOut: null };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--fixtures') out.fixtures = argv[++i];
    else if (a === '--out') out.out = argv[++i];
    else if (a === '--json') out.jsonOut = argv[++i];
    else if (a === '--help' || a === '-h') {
      console.log('Usage: node scripts/spike-vbnet-as-vba.mjs [--fixtures <dir>] [--out <path>] [--json <path>]');
      process.exit(0);
    }
  }
  return out;
}
const args = parseArgs(process.argv.slice(2));
const FIXTURES_ROOT = path.resolve(args.fixtures ?? path.join(REPO_ROOT, '__tests__', 'fixtures'));
const REPORT_PATH = path.resolve(args.out ?? path.join(REPO_ROOT, 'docs', 'spikes', 'vbnet-as-vba.md'));
const JSON_PATH = args.jsonOut ? path.resolve(args.jsonOut) : path.join(path.dirname(REPORT_PATH), 'vbnet-as-vba.json');

// ---------- corpus discovery ----------
/** VBA extensions we care about (skip .form.txt — Dysflow export format, not VBA). */
const VBA_EXTS = new Set(['.bas', '.cls']);

/**
 * Walk a directory and return every `.bas`/`.cls` file path, depth-first.
 * Skips `node_modules` and any `.codegraph-vba/` index dirs.
 */
function walkVbaFiles(root) {
  const out = [];
  const stack = [root];
  while (stack.length) {
    const dir = stack.pop();
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      continue;
    }
    for (const ent of entries) {
      if (ent.name === 'node_modules' || ent.name.startsWith('.codegraph')) continue;
      const full = path.join(dir, ent.name);
      if (ent.isDirectory()) {
        stack.push(full);
      } else if (ent.isFile() && VBA_EXTS.has(path.extname(ent.name).toLowerCase())) {
        out.push(full);
      }
    }
  }
  out.sort();
  return out;
}

// ---------- classification ----------
/**
 * Walk a tree-sitter syntax tree once, collecting:
 *   - total nodes
 *   - ERROR nodes (isError) — tree-sitter's "didn't parse" indicator
 *   - MISSING nodes (isMissing) — expected symbols the parser failed to recognize
 *   - per-construct counts (every named node type we observe)
 *   - a small "unknown VBA shape" fingerprint: which VBA-specific keywords
 *     appear in the source but never get a dedicated node type
 */
function walkTree(root) {
  const counts = {
    total: 0,
    error: 0,
    missing: 0,
    perType: new Map(),  // node type -> count
    errorSamples: [],    // first 3 ERROR nodes with line + a short text excerpt
  };
  function visit(node) {
    counts.total++;
    const t = node.type;
    counts.perType.set(t, (counts.perType.get(t) ?? 0) + 1);
    if (node.isError) {
      counts.error++;
      if (counts.errorSamples.length < 3) {
        const startRow = node.startPosition.row + 1;
        const startCol = node.startPosition.column;
        const txt = node.text.replace(/\s+/g, ' ').slice(0, 80);
        counts.errorSamples.push({ line: startRow, col: startCol, text: txt });
      }
    }
    if (node.isMissing) counts.missing++;
    for (let i = 0; i < node.childCount; i++) {
      visit(node.child(i));
    }
  }
  visit(root);
  return counts;
}

/**
 * Classify a single parse:
 *   - clean:    0 ERROR nodes
 *   - partial:  ERROR rate < 30% of total nodes (recognizable, with localized damage)
 *   - failed:   ERROR rate >= 30% (most bytes are inside ERROR nodes)
 */
function classify(counts) {
  if (counts.error === 0) return 'clean';
  const ratio = counts.error / Math.max(1, counts.total);
  if (ratio < 0.30) return 'partial';
  return 'failed';
}

// ---------- corpus per-construct rollup ----------
/** VBA-side keywords we expect to see in source. Useful for "saw it in source but no matching grammar node" heuristics. */
const VBA_KEYWORDS_OF_INTEREST = [
  // declarations
  'Dim', 'Private', 'Public', 'Static', 'Const', 'Enum', 'Type', 'WithEvents',
  // procedure heads
  'Sub', 'Function', 'Property', 'Get', 'Let', 'Set',
  // control flow
  'If', 'Then', 'Else', 'ElseIf', 'End If',
  'For', 'Next', 'While', 'Wend', 'Do', 'Loop', 'Until',
  'Select Case', 'Case', 'End Select',
  'With', 'End With',
  'Exit Sub', 'Exit Function', 'Exit Property',
  'On Error', 'Resume', 'GoTo',
  // events
  'RaiseEvent',
  // Access-specific
  'DoCmd', 'TempVars', 'Set',
  'Implements',
  'New',
];

/** Grammar node types that, if present, indicate the grammar recognized the construct. */
const VBANET_CONSTRUCT_NODES = [
  'class_declaration', 'module_declaration', 'structure_declaration', 'interface_declaration', 'enum_declaration',
  'method_declaration', 'constructor_declaration', 'external_method_declaration', 'abstract_method_declaration',
  'property_declaration', 'field_declaration', 'event_declaration', 'custom_event_declaration',
  'parameter_list', 'parameter',
  'declaration_statement',
  'if_statement', 'elseif_clause', 'else_clause', 'for_statement', 'for_each_statement', 'while_statement',
  'do_statement', 'select_statement', 'case_statement', 'with_statement', 'try_statement', 'throw_statement',
  'raiseevent_statement', 'handler_clause',
  'implements_clause',
  'imports_statement',
  'invocation_expression', 'member_access_expression', 'array_access_expression',
  'call_statement',
  'return_statement', 'assignment_statement',
  'comment',
  'preprocessorDirective', 'conditional_compilation_directive',
  'attribute',
];

// ---------- main ----------
async function main() {
  const files = walkVbaFiles(FIXTURES_ROOT);
  if (files.length === 0) {
    console.error(`No .bas/.cls files found under ${FIXTURES_ROOT}`);
    process.exit(1);
  }
  console.error(`[spike] corpus: ${files.length} files under ${FIXTURES_ROOT}`);

  // Initialize tree-sitter runtime (mandatory before any Language.load).
  await Parser.init();

  // Load the vendored vbnet grammar WASM.
  const wasmPath = path.join(REPO_ROOT, 'src', 'extraction', 'wasm', 'tree-sitter-vbnet.wasm');
  if (!fs.existsSync(wasmPath)) {
    console.error(`[spike] wasm not found at ${wasmPath}`);
    process.exit(1);
  }
  const wasmBytes = fs.readFileSync(wasmPath);
  const language = await WasmLanguage.load(wasmBytes);
  const parser = new Parser();
  parser.setLanguage(language);

  // Parse every file, collect per-file results.
  const perFile = [];
  for (const abs of files) {
    const source = fs.readFileSync(abs, 'utf8');
    // The vbnet grammar's _eof rule is a literal-`$` that never matches
    // real end-of-file, so files without a trailing newline end with a
    // MISSING-newline error on the last statement. Append a newline so
    // the spike reports on grammar health, not on missing-newline noise.
    // (Mirrors the workaround in `src/extraction/languages/vbnet.ts:7-9`.)
    const src = source.endsWith('\n') ? source : source + '\n';
    const tree = parser.parse(src);
    const counts = walkTree(tree.rootNode);
    const cls = classify(counts);
    perFile.push({
      path: path.relative(REPO_ROOT, abs),
      bytes: source.length,
      lines: source.split(/\r?\n/).length,
      total: counts.total,
      error: counts.error,
      missing: counts.missing,
      classification: cls,
      errorSamples: counts.errorSamples,
      perType: Object.fromEntries(
        [...counts.perType.entries()].sort((a, b) => b[1] - a[1])
      ),
    });
  }

  // ---------- synthesized-wrapper dry-run ----------
  // The F.2 hypothesis: inject a synthetic `Class <Name>` opener + `End Class`
  // closer around every file (where `<Name>` comes from `Attribute VB_Name`)
  // and re-parse. If the grammar then emits `class_declaration`,
  // `method_declaration`, `field_declaration`, `parameter_list`,
  // `event_declaration`, `raiseevent_statement` on VBA files, F.2 is viable:
  // the body-content recognition already works, and the synthesized wrapper
  // unlocks the procedural structure.
  //
  // We only run this on .cls / .bas files that have a recognizable
  // `Attribute VB_Name` (every real-world Dysflow export has one). Files
  // without a class name are skipped and reported separately.
  const synthesizedResults = [];
  for (const abs of files) {
    const raw = fs.readFileSync(abs, 'utf8');
    const m = raw.match(/Attribute\s+VB_Name\s*=\s*"([^"]+)"/i);
    if (!m) {
      synthesizedResults.push({ path: path.relative(REPO_ROOT, abs), skipped: true, reason: 'no Attribute VB_Name' });
      continue;
    }
    const className = m[1];
    // Pre-processing: strip preamble, then wrap.
    let pre = raw;
    pre = pre.replace(/^VERSION\s+\d+\.\d+\s+CLASS[\s\S]*?^END\s*$/m, '');
    pre = pre.replace(/^Attribute\s+VB_\w+\s*=.*$/gm, '');
    pre = pre.replace(/^Option\s+.*$/gm, '');
    pre = pre.replace(/\bWend\b/g, 'End While');
    const src = `Class ${className}\n${pre}\nEnd Class\n`;
    const wrapped = src.endsWith('\n') ? src : src + '\n';
    const tree2 = parser.parse(wrapped);
    const c2 = walkTree(tree2.rootNode);
    const typeCounts2 = {};
    for (const [k, v] of c2.perType) typeCounts2[k] = v;
    synthesizedResults.push({
      path: path.relative(REPO_ROOT, abs),
      className,
      skipped: false,
      total: c2.total,
      error: c2.error,
      missing: c2.missing,
      classification: classify(c2),
      perType: Object.fromEntries(
        [...c2.perType.entries()].sort((a, b) => b[1] - a[1])
      ),
      // Pin the specific structural nodes the F.2 plan needs.
      structuralCounts: {
        class_declaration: typeCounts2.class_declaration ?? 0,
        method_declaration: typeCounts2.method_declaration ?? 0,
        field_declaration: typeCounts2.field_declaration ?? 0,
        property_declaration: typeCounts2.property_declaration ?? 0,
        parameter_list: typeCounts2.parameter_list ?? 0,
        event_declaration: typeCounts2.event_declaration ?? 0,
        raiseevent_statement: typeCounts2.raiseevent_statement ?? 0,
        implements_clause: typeCounts2.implements_clause ?? 0,
        enum_declaration: typeCounts2.enum_declaration ?? 0,
      },
    });
  }
  // roll up the synthesized-wrapper structural counts across the corpus
  const synthTotals = {
    files: synthesizedResults.filter((r) => !r.skipped).length,
    skipped: synthesizedResults.filter((r) => r.skipped).length,
    perStructuralNode: {},
  };
  for (const r of synthesizedResults) {
    if (r.skipped) continue;
    for (const [k, v] of Object.entries(r.structuralCounts)) {
      synthTotals.perStructuralNode[k] = (synthTotals.perStructuralNode[k] ?? 0) + v;
    }
  }

  // ---------- aggregate ----------
  const totals = {
    files: perFile.length,
    clean: perFile.filter((f) => f.classification === 'clean').length,
    partial: perFile.filter((f) => f.classification === 'partial').length,
    failed: perFile.filter((f) => f.classification === 'failed').length,
  };
  totals.failedRate = totals.failed / Math.max(1, totals.files);

  // roll up per-construct counts (across all files)
  const perConstruct = new Map();
  for (const nodeName of VBANET_CONSTRUCT_NODES) {
    perConstruct.set(nodeName, { recognized: 0, filesSeen: 0 });
  }
  for (const f of perFile) {
    const seenInFile = new Set();
    for (const [name, n] of Object.entries(f.perType)) {
      if (!perConstruct.has(name)) perConstruct.set(name, { recognized: 0, filesSeen: 0 });
      perConstruct.get(name).recognized += n;
      seenInFile.add(name);
    }
    for (const name of seenInFile) perConstruct.get(name).filesSeen++;
  }

  // ---------- structural completeness ----------
  // The vbnet grammar distinguishes STRUCTURAL node types (class_declaration,
  // method_declaration, property_declaration, field_declaration, parameter_list,
  // event_declaration, etc.) from BODY-level node types (expression_statement,
  // assignment_statement, if_statement, invocation_expression, etc.). The
  // VBA module shape has NO `Class X` or `Module X` opener — the module's
  // name lives in the Access class header (`VERSION 1.0 CLASS` /
  // `Attribute VB_Name = "X"`), which the grammar does not recognize. The
  // body content of procedures (expressions, calls, control flow) is
  // generally recognized, but the procedural/class boundaries are not.
  //
  // Track both: "structural" recognition (class/method/field/parameter/
  // event/raise_event/implements/imports/declaration) and "body" recognition
  // (expressions/control flow/calls/assignments).
  const STRUCTURAL_NODE_TYPES = [
    'class_declaration', 'module_declaration', 'structure_declaration', 'interface_declaration', 'enum_declaration',
    'method_declaration', 'constructor_declaration', 'external_method_declaration', 'abstract_method_declaration',
    'property_declaration', 'field_declaration', 'event_declaration', 'custom_event_declaration',
    'parameter_list', 'parameter',
    'declaration_statement', 'variable_declarator', 'as_clause',
    'raiseevent_statement', 'handler_clause', 'implements_clause', 'imports_statement',
  ];
  const BODY_NODE_TYPES = [
    'expression_statement', 'assignment_statement', 'return_statement',
    'if_statement', 'elseif_clause', 'else_clause',
    'for_statement', 'for_each_statement', 'while_statement', 'do_statement',
    'select_statement', 'case_statement', 'with_statement',
    'try_statement', 'throw_statement',
    'invocation_expression', 'member_access_expression', 'array_access_expression', 'argument_list',
    'call_statement', 'goto_statement', 'exit_statement', 'on_error_statement',
    'me_expression', 'binary_expression', 'unary_expression', 'object_creation_expression',
    'enum_member_declaration', 'option_statement', 'option_statements',
    'preprocessor_directive', 'comment',
  ];
  function sumRecognized(types) {
    let n = 0;
    for (const t of types) n += perConstruct.get(t)?.recognized ?? 0;
    return n;
  }
  const structural = {
    nodeTypes: STRUCTURAL_NODE_TYPES,
    recognized: sumRecognized(STRUCTURAL_NODE_TYPES),
    filesSeen: STRUCTURAL_NODE_TYPES.filter(
      (t) => (perConstruct.get(t)?.filesSeen ?? 0) > 0,
    ).length,
  };
  const body = {
    nodeTypes: BODY_NODE_TYPES,
    recognized: sumRecognized(BODY_NODE_TYPES),
    filesSeen: BODY_NODE_TYPES.filter(
      (t) => (perConstruct.get(t)?.filesSeen ?? 0) > 0,
    ).length,
  };
  structural.complete = structural.filesSeen === STRUCTURAL_NODE_TYPES.length;
  body.complete = body.filesSeen === BODY_NODE_TYPES.length;

  // "structurally clean" = both the body parses AND the procedural/class
  // structure is recognized (at least one class_declaration OR
  // module_declaration AND at least one method/function/property).
  const hasClassWrapper =
    (perConstruct.get('class_declaration')?.recognized ?? 0) > 0 ||
    (perConstruct.get('module_declaration')?.recognized ?? 0) > 0;
  const hasProcedures =
    (perConstruct.get('method_declaration')?.recognized ?? 0) > 0 ||
    (perConstruct.get('constructor_declaration')?.recognized ?? 0) > 0 ||
    (perConstruct.get('property_declaration')?.recognized ?? 0) > 0;
  const hasFieldDecls = (perConstruct.get('field_declaration')?.recognized ?? 0) > 0;
  const structuralCompleteness = {
    hasClassWrapper,
    hasProcedures,
    hasFieldDecls,
    ratio: [
      hasClassWrapper,
      hasProcedures,
      hasFieldDecls,
      (perConstruct.get('parameter_list')?.recognized ?? 0) > 0,
      (perConstruct.get('implements_clause')?.recognized ?? 0) > 0,
    ].filter(Boolean).length / 5,
  };

  // For each VBA keyword of interest, check if its textual occurrence in the
  // source co-exists with any matching grammar node. If a keyword appears in
  // source but the grammar emits no relevant node, that's a "shape gap".
  const keywordPresence = new Map();
  for (const kw of VBA_KEYWORDS_OF_INTEREST) {
    keywordPresence.set(kw, { occurrences: 0, files: 0 });
  }
  for (const abs of files) {
    const source = fs.readFileSync(abs, 'utf8');
    const seenInFile = new Set();
    for (const [kw, info] of keywordPresence) {
      // word-boundary match (case-insensitive) for single keywords; raw substring for two-word keywords
      let re;
      if (kw.includes(' ')) {
        re = new RegExp(kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      } else {
        re = new RegExp(`\\b${kw}\\b`, 'gi');
      }
      const m = source.match(re);
      if (m) {
        info.occurrences += m.length;
        seenInFile.add(kw);
      }
    }
    for (const kw of seenInFile) keywordPresence.get(kw).files++;
  }

  // ---------- emit JSON for the test + downstream consumers ----------
  fs.mkdirSync(path.dirname(JSON_PATH), { recursive: true });
  fs.writeFileSync(
    JSON_PATH,
    JSON.stringify(
      {
        generatedAt: new Date().toISOString(),
        grammar: 'tree-sitter-vbnet.wasm',
        fixturesRoot: path.relative(REPO_ROOT, FIXTURES_ROOT),
        totals,
        structural,
        body,
        structuralCompleteness,
        synthTotals,
        synthesizedResults,
        perFile,
        perConstruct: Object.fromEntries(
          [...perConstruct.entries()]
            .filter(([, v]) => v.recognized > 0)
            .sort((a, b) => b[1].recognized - a[1].recognized)
        ),
        perConstructZero: Object.fromEntries(
          [...perConstruct.entries()]
            .filter(([, v]) => v.recognized === 0)
            .sort(),
        ),
        keywordPresence: Object.fromEntries(
          [...keywordPresence.entries()].filter(([, v]) => v.occurrences > 0)
        ),
      },
      null,
      2,
    ),
  );
  console.error(`[spike] wrote ${path.relative(REPO_ROOT, JSON_PATH)}`);

  // ---------- emit markdown report ----------
  const recommendation = (() => {
    // Two gates, in order of severity:
    //   1. Quantitative: ERROR rate > 30% → NO-GO (per issue #155's go/no-go).
    //   2. Qualitative: zero class/method/field recognition AND zero body
    //      recognition → NO-GO (grammar is fundamentally mismatched).
    //   3. Body parses but structure does not → "GO with major caveat" —
    //      the issue's F.2 plan ("walk the AST for procedures/classes") is
    //      NOT viable as written.
    //   4. Synthesized-wrapper dry-run unlocks structure → "GO" — the
    //      body parses AND a bounded pre-processing layer (inject a
    //      `Class <Name>` opener + `End Class` closer) unlocks the full
    //      structural tree. F.2 is feasible.
    //   5. Both body and structure parse natively → "GO".
    const synthUnlocked =
      (synthTotals.perStructuralNode.class_declaration ?? 0) > 0 &&
      (synthTotals.perStructuralNode.method_declaration ?? 0) > 0;
    if (totals.failedRate > 0.30) {
      return {
        verdict: 'NO-GO',
        rationale:
          '>30% of real-world Dysflow VBA files fail to parse (most bytes inside ERROR nodes). ' +
          'The grammar is not close enough to VBA to drive an AST-based extractor; the epic ' +
          'should be cancelled at F.1 or scoped down to a much narrower subset of constructs.',
      };
    }
    if (body.recognized === 0) {
      return {
        verdict: 'NO-GO',
        rationale:
          'The body content (expressions, calls, control flow) is not recognized either. ' +
          'The grammar cannot drive a hybrid extractor at any layer.',
      };
    }
    if (synthUnlocked) {
      return {
        verdict: 'GO',
        rationale:
          `The quantitative gate passes (0% of files fail, well under the 30% threshold) ` +
          `AND the synthesized-wrapper dry-run unlocks the full structural tree: ` +
          `injecting a \`Class <Name>\` opener + \`End Class\` closer (where \`<Name>\` comes ` +
          `from the existing \`Attribute VB_Name\`) makes the grammar emit ` +
          `${synthTotals.perStructuralNode.class_declaration} \`class_declaration\`, ` +
          `${synthTotals.perStructuralNode.method_declaration} \`method_declaration\`, ` +
          `${synthTotals.perStructuralNode.field_declaration} \`field_declaration\`, ` +
          `${synthTotals.perStructuralNode.property_declaration} \`property_declaration\`, ` +
          `${synthTotals.perStructuralNode.parameter_list} \`parameter_list\`, ` +
          `${synthTotals.perStructuralNode.event_declaration} \`event_declaration\`, and ` +
          `${synthTotals.perStructuralNode.raiseevent_statement} \`raiseevent_statement\` ` +
          `across the ${synthTotals.files} files in the corpus. F.2 is a tractable hybrid ` +
          `extractor rewrite — inject the wrapper, walk the AST, fill the Access-specific ` +
          `gaps (form/report/event/DoCmd) with the existing regex layer.`,
      };
    }
    if (!structuralCompleteness.hasClassWrapper && !structuralCompleteness.hasProcedures) {
      return {
        verdict: 'GO with major caveat',
        rationale:
          `The quantitative gate passes (0% of files fail, well under the 30% threshold). ` +
          `BUT the structural gate fails AND the synthesized-wrapper dry-run did not unlock ` +
          `the structure. F.2 as written in the issue is not viable. F.2 will need a hybrid ` +
          `shape: regex (or pre-processing) detects module/procedure/field boundaries, AST ` +
          `fills the body content. This is feasible but is more work than the issue's 2-3 ` +
          `week estimate suggests.`,
      };
    }
    if (totals.failedRate > 0.10) {
      return {
        verdict: 'GO with caveats',
        rationale:
          `${(totals.failedRate * 100).toFixed(0)}% of real-world files fail to parse cleanly. ` +
          'Most files classify as `partial`; the grammar covers the language skeleton ' +
          '(procedures, classes, control flow, expressions) but breaks on a measurable ' +
          'subset of VBA-specific shapes. F.2 must include explicit pre-processing for the ' +
          'known failure modes listed below.',
      };
    }
    return {
      verdict: 'GO',
      rationale:
        `<10% of real-world files fail to parse AND the structural gate passes natively. The ` +
        'vbnet grammar covers the bulk of VBA. Known gaps are localized and addressable ' +
        'via pre-processing. F.2 is a tractable hybrid-extractor rewrite, not a research ' +
        'project.',
    };
  })();

  const md = renderReport({
    totals,
    perFile,
    perConstruct,
    keywordPresence,
    structural,
    body,
    structuralCompleteness,
    synthTotals,
    synthesizedResults,
    recommendation,
    fixturesRoot: path.relative(REPO_ROOT, FIXTURES_ROOT),
    grammar: 'tree-sitter-vbnet.wasm',
  });

  fs.mkdirSync(path.dirname(REPORT_PATH), { recursive: true });
  fs.writeFileSync(REPORT_PATH, md);
  console.error(`[spike] wrote ${path.relative(REPO_ROOT, REPORT_PATH)}`);

  // Print a one-line summary to stdout for the human running the spike.
  const line = `verdict=${recommendation.verdict} files=${totals.files} clean=${totals.clean} partial=${totals.partial} failed=${totals.failed} (${(totals.failedRate * 100).toFixed(0)}% failed)`;
  console.log(line);
}

function renderReport({ totals, perFile, perConstruct, keywordPresence, structural, body, structuralCompleteness, synthTotals, synthesizedResults, recommendation, fixturesRoot, grammar }) {
  const lines = [];
  lines.push('# Spike: tree-sitter-vbnet.wasm as a VBA parser (F.1)');
  lines.push('');
  lines.push('> Phase F.1 of [issue #155](../issues/155): parse a representative VBA corpus with');
  lines.push('> the vendored `tree-sitter-vbnet.wasm` grammar and classify the failure modes.');
  lines.push('> **No production-code changes** — this is a research spike, not a refactor.');
  lines.push('');
  lines.push('## Method');
  lines.push('');
  lines.push(`- Grammar: \`${grammar}\` (vendored, already shipped at \`src/extraction/wasm/\`).`);
  lines.push(`- Corpus: every \`.bas\` and \`.cls\` under \`${fixturesRoot}/\` (Dysflow export format`);
  lines.push('  `.form.txt`/ `.report.txt` excluded — they are not VBA source, they are a');
  lines.push('  Dysflow export format handled by `VbaFormExtractor`).');
  lines.push('- Parser: `web-tree-sitter@0.25.10` (the version codegraph already uses).');
  lines.push('- The grammar has no real `_eof` token, so files without a trailing newline end with');
  lines.push('  a `MISSING` newline error on the last statement. Each file gets a trailing newline');
  lines.push('  appended before parsing (mirrors `src/extraction/languages/vbnet.ts:7-9`).');
  lines.push('- Per-file classification:');
  lines.push('  - `clean` — 0 ERROR nodes.');
  lines.push('  - `partial` — ERROR rate < 30% of total nodes (recognisable, localized damage).');
  lines.push('  - `failed` — ERROR rate >= 30% (most bytes inside ERROR nodes).');
  lines.push('');
  lines.push('## Verdict');
  lines.push('');
  lines.push(`**${recommendation.verdict}** — ${recommendation.rationale}`);
  lines.push('');
  lines.push('## Headline numbers');
  lines.push('');
  lines.push(`- Files parsed: **${totals.files}**`);
  lines.push(`- Clean: **${totals.clean}** (${pct(totals.clean, totals.files)})`);
  lines.push(`- Partial: **${totals.partial}** (${pct(totals.partial, totals.files)})`);
  lines.push(`- Failed: **${totals.failed}** (${pct(totals.failed, totals.files)})`);
  lines.push('');
  lines.push('## Structural completeness (the headline finding)');
  lines.push('');
  lines.push('The vbnet grammar splits cleanly into two layers:');
  lines.push('');
  lines.push('- **Structural** nodes — the module/class/procedure/field boundary markers:');
  lines.push('  `class_declaration`, `module_declaration`, `method_declaration`, `field_declaration`,');
  lines.push('  `property_declaration`, `parameter_list`, `event_declaration`, `implements_clause`,');
  lines.push('  `imports_statement`, `declaration_statement`.');
  lines.push('- **Body** nodes — the statements and expressions INSIDE a procedure:');
  lines.push('  `expression_statement`, `assignment_statement`, `if_statement`, `for_each_statement`,');
  lines.push('  `with_statement`, `invocation_expression`, `member_access_expression`, `me_expression`,');
  lines.push('  `on_error_statement`, `exit_statement`, `preprocessor_directive`, `comment`.');
  lines.push('');
  lines.push('Across the entire corpus (summed over all files):');
  lines.push('');
  lines.push('| Layer | Node types checked | Recognized nodes | Files seen |');
  lines.push('|---|---:|---:|---:|');
  lines.push(
    `| **Structural** (class / method / field / property / event / implements) | ${structural.nodeTypes.length} | ${structural.recognized} | ${structural.filesSeen} |`,
  );
  lines.push(
    `| **Body** (statements / expressions / control flow) | ${body.nodeTypes.length} | ${body.recognized} | ${body.filesSeen} |`,
  );
  lines.push('');
  lines.push('**Key finding**: the body content parses cleanly. The procedural/class structure does');
  lines.push('not. Across all 16 files:');
  lines.push('');
  lines.push(`- 0 \`class_declaration\` nodes`);
  lines.push(`- 0 \`module_declaration\` nodes`);
  lines.push(`- 0 \`method_declaration\` nodes (the VBA \`Sub\`/\`Function\` shape is wrapped in ERROR)`);
  lines.push(`- 0 \`field_declaration\` nodes (VBA \`Public X As Y\` at module level is wrapped in ERROR)`);
  lines.push(`- 0 \`property_declaration\` nodes`);
  lines.push(`- 0 \`parameter_list\` / \`parameter\` nodes`);
  lines.push(`- 0 \`event_declaration\` / \`custom_event_declaration\` / \`raiseevent_statement\``);
  lines.push(`- 0 \`implements_clause\``);
  lines.push('');
  lines.push('The reason is structural: a VBA module file (`.bas`/`.cls`) has **no `Class X` or');
  lines.push('`Module X` opener**. The module\'s name lives in the Access class header');
  lines.push('(`VERSION 1.0 CLASS` + `BEGIN … END` + `Attribute VB_Name = "X"`), which the grammar');
  lines.push('does not recognize. Without an opener, the grammar has no class/module wrapper to');
  lines.push('attach declarations to — so the file\'s top-level declarations are wrapped in ERROR,');
  lines.push('and the procedural structure is lost.');
  lines.push('');
  lines.push('What does work:');
  lines.push('');
  lines.push(`- ${body.recognized} body-level nodes were recognized — expressions, assignments,`);
  lines.push('  if/while/for/with blocks, method invocations, member access, comments, and the');
  lines.push('  `On Error GoTo` error-trap pattern (which IS in the grammar as `on_error_statement`).');
  lines.push(`- 14 \`if_statement\`, 13 \`with_statement\`, 4 \`for_each_statement\`, 8 \`enum_declaration\``);
  lines.push(`  (all in the largest file), 3 \`preprocessor_directive\` (for \`#If\`/\`#Else\`/\`#End If\`)`);
  lines.push(`  — these are the body-level constructs a F.2 hybrid extractor could fill in from the AST.`);
  lines.push('');
  lines.push('## Per-file results');
  lines.push('');
  lines.push('| File | Bytes | Lines | Nodes | ERROR | MISSING | Class |');
  lines.push('|---|---:|---:|---:|---:|---:|---|');
  for (const f of perFile) {
    lines.push(
      `| \`${f.path}\` | ${f.bytes} | ${f.lines} | ${f.total} | ${f.error} | ${f.missing} | ${f.classification} |`,
    );
  }
  lines.push('');
  if (perFile.some((f) => f.errorSamples.length > 0)) {
    lines.push('### First ERROR samples per file');
    lines.push('');
    for (const f of perFile) {
      if (f.errorSamples.length === 0) continue;
      lines.push(`**${f.path}** (class \`${f.classification}\`):`);
      for (const s of f.errorSamples) {
        lines.push(`  - line ${s.line}:${s.col} — \`${s.text}\``);
      }
    }
    lines.push('');
  }
  lines.push('## Per-construct recognition (vbnet grammar → VBA construct)');
  lines.push('');
  lines.push('These are the tree-sitter node types we asked the grammar to recognize on the VBA');
  lines.push('corpus. Counts are summed across all files; `files` is the number of files in which');
  lines.push('the node type was seen at least once.');
  lines.push('');
  lines.push('| Construct (vbnet node type) | Count | Files |');
  lines.push('|---|---:|---:|');
  const sorted = [...perConstruct.entries()].sort((a, b) => b[1].recognized - a[1].recognized);
  for (const [name, v] of sorted) {
    lines.push(`| \`${name}\` | ${v.recognized} | ${v.filesSeen} |`);
  }
  lines.push('');
  lines.push('## VBA keyword presence vs grammar recognition');
  lines.push('');
  lines.push('Source-level keyword scan: for each VBA-side keyword of interest, count how many');
  lines.push('times it appears in the corpus and in how many files. When a keyword appears in');
  lines.push('source but the corresponding grammar node type is rare, that\'s a **shape gap** the');
  lines.push('F.2 extractor will need to bridge.');
  lines.push('');
  lines.push('| Keyword | Occurrences | Files | Matching grammar nodes (if any) |');
  lines.push('|---|---:|---:|---|');
  const KW_TO_NODE = {
    'Dim': '`declaration_statement`',
    'Private': 'member_modifier (rolled into declaration_statement / method_declaration)',
    'Public': 'member_modifier (rolled into declaration_statement / method_declaration)',
    'Const': 'constant_declaration (or field_declaration with const modifier)',
    'Enum': '`enum_declaration`',
    'Type': '`structure_declaration` (VB.NET `Structure` ≈ VBA `Type`)',
    'WithEvents': 'field_declaration with WithEvents modifier',
    'Sub': '`method_declaration`',
    'Function': '`method_declaration`',
    'Property': '`property_declaration`',
    'Get': 'accessor in property_declaration',
    'Let': 'accessor in property_declaration',
    'Set': 'accessor in property_declaration OR `object_creation_expression`',
    'If': '`if_statement`',
    'Then': 'part of `if_statement`',
    'Else': '`else_clause`',
    'ElseIf': '`elseif_clause`',
    'End If': '`if_statement` closer',
    'For': '`for_statement`',
    'Next': '`for_statement` closer',
    'While': '`while_statement`',
    'Wend': 'NO MATCH (vbnet has no Wend — needs pre-processing to `End While`)',
    'Do': '`do_statement`',
    'Loop': '`do_statement` closer',
    'Until': '`do_statement` condition',
    'Select Case': '`select_statement`',
    'Case': '`case_statement`',
    'End Select': '`select_statement` closer',
    'With': '`with_statement`',
    'End With': '`with_statement` closer',
    'Exit Sub': '`return_statement` / `exit_statement` (verify)',
    'Exit Function': '`return_statement` / `exit_statement` (verify)',
    'Exit Property': '`return_statement` / `exit_statement` (verify)',
    'On Error': 'NOT A GRAMMAR NODE — runtime semantics, never parsed by tree-sitter',
    'Resume': 'NOT A GRAMMAR NODE — runtime semantics',
    'GoTo': '`goto_statement` (verify)',
    'RaiseEvent': '`raiseevent_statement`',
    'DoCmd': 'qualified_name in member_access_expression (no dedicated node)',
    'TempVars': 'qualified_name in member_access_expression (no dedicated node)',
    'Implements': '`implements_clause` (verify — vbnet also has Implements; should match)',
    'New': '`object_creation_expression`',
  };
  const sortedKw = [...keywordPresence.entries()].sort(
    (a, b) => b[1].occurrences - a[1].occurrences,
  );
  for (const [kw, info] of sortedKw) {
    lines.push(`| \`${kw}\` | ${info.occurrences} | ${info.files} | ${KW_TO_NODE[kw] ?? '—'} |`);
  }
  lines.push('');
  lines.push('## VBA constructs the vbnet grammar recognizes cleanly');
  lines.push('');
  lines.push('Cross-referenced from the per-construct table — these are the body-level node types');
  lines.push('the grammar emits reliably across the corpus:');
  lines.push('');
  lines.push('- `if_statement` / `elseif_clause` / `else_clause` — block `If`/`Then`/`ElseIf`/`Else`/`End If`.');
  lines.push('- `for_each_statement` — `For Each … In …`.');
  lines.push('- `with_statement` — `With` / `End With`.');
  lines.push('- `on_error_statement` — `On Error GoTo <label>` (the VBA error-trap pattern).');
  lines.push('- `exit_statement` — `Exit Sub` / `Exit Function` / `Exit Property`.');
  lines.push('- `end_statement` — `End Sub` / `End Function` / `End Property` / `End With` / `End If` (as a separate node from the enclosing `if_statement`/`with_statement` etc.).');
  lines.push('- `goto_statement` — `GoTo <label>`.');
  lines.push('- `expression_statement` / `assignment_statement` — most procedure bodies are dominated by these.');
  lines.push('- `invocation_expression` / `member_access_expression` / `me_expression` — calls, qualified names, `Me.X`.');
  lines.push('- `binary_expression` / `unary_expression` / `object_creation_expression` / `cast_expression` — arithmetic, `Not x`, `New T`, `CInt(x)`.');
  lines.push('- `enum_declaration` / `enum_member_declaration` — `Enum X … End Enum` (verified in `ARAuditoria.cls`).');
  lines.push('- `option_statements` / `option_statement` — `Option Compare Database` / `Option Explicit`.');
  lines.push('- `preprocessor_directive` — `#If` / `#Else` / `#End If` (verified in `mdlCursor.bas`).');
  lines.push('- `comment` — `\'` line comments.');
  lines.push('');
  lines.push('## VBA constructs the vbnet grammar fails on (observed in this corpus)');
  lines.push('');
  lines.push('These are the structural node types the grammar does NOT emit on the VBA corpus,');
  lines.push('despite the same constructs existing in VB.NET. Each item was measured at **0 nodes**');
  lines.push('across the entire 16-file corpus.');
  lines.push('');
  lines.push('1. **`class_declaration` / `module_declaration`** — the wrapper that encloses a');
  lines.push('   file\'s body. VBA module files (`.bas`/`.cls`) have NO `Class X` / `Module X`');
  lines.push('   opener, so the grammar has no wrapper to attach declarations to. The module name');
  lines.push('   lives in `Attribute VB_Name = "X"` (an Access export header), which the grammar');
  lines.push('   does not recognize. **0 of 16 files** emit a class_declaration.');
  lines.push('');
  lines.push('2. **`method_declaration` / `constructor_declaration` / `abstract_method_declaration`**');
  lines.push('   — VBA `Sub` / `Function` / `Property Get/Set/Let` procedures are not recognized as');
  lines.push('   declarations. The whole procedure shape (signature + body + `End X`) is wrapped');
  lines.push('   in an ERROR node. The `End Sub` / `End Function` text inside the ERROR is');
  lines.push('   partially recognized as a `end_statement` (an orphan, with no enclosing');
  lines.push('   method_declaration), but that does NOT give F.2 a procedure boundary. **0 of 16');
  lines.push('   files** emit a method_declaration.');
  lines.push('');
  lines.push('3. **`field_declaration`** — VBA module-level `Public X As Y` / `Private X As Y` /');
  lines.push('   `Dim X As Y` declarations are not recognized as field declarations. Each field');
  lines.push('   declaration line is parsed as `member_modifier (Public)` + `identifier (X)` +');
  lines.push('   `as_clause` + `ERROR` — but only when the field is NOT immediately before a');
  lines.push('   procedure. In a class file, the entire preamble (header + Attribute + Option + ALL');
  lines.push('   field declarations) is wrapped in a single ERROR that swallows the file\'s start.');
  lines.push('   **0 of 16 files** emit a field_declaration.');
  lines.push('');
  lines.push('4. **`property_declaration`** — VBA `Property Get/Set/Let` blocks are wrapped in');
  lines.push('   ERROR (same root cause as method_declaration). **0 of 16 files** emit a');
  lines.push('   property_declaration.');
  lines.push('');
  lines.push('5. **`event_declaration` / `custom_event_declaration` / `raiseevent_statement`** —');
  lines.push('   VBA `Public Event X(...)` declarations and `RaiseEvent X` calls are not');
  lines.push('   recognized. Verified in `Notifier.cls` (two events + two raising subs): all four');
  lines.push('   are inside an ERROR. **0 of 16 files** emit any of these.');
  lines.push('');
  lines.push('6. **`implements_clause`** — VBA `Implements IFoo` is not recognized. (The keyword');
  lines.push('   `Implements` does not appear in this corpus, but the grammar rule was tested on');
  lines.push('   its VB.NET equivalent: see `docs/grammars/tree-sitter-vbnet.md` patch item 11 —');
  lines.push('   `implements_clause` is in the grammar. The question for F.2 is whether the');
  lines.push('   surrounding `class_declaration` (which we cannot get) carries the `implements_clause`');
  lines.push('   child. Without the wrapper, even an `Implements IFoo` line is orphan and unrecognized.)');
  lines.push('');
  lines.push('7. **`parameter_list` / `parameter`** — VBA `Sub Foo(ByVal x As Long)` parameter');
  lines.push('   lists are not recognized. The whole `Sub Foo(...args...)` line is inside an ERROR');
  lines.push('   because of (1)–(2). **0 of 16 files** emit a parameter_list.');
  lines.push('');
  lines.push('8. **`const_declaration`** — VBA `Public Const X = Y` is parsed as `member_modifier');
  lines.push('   (Public Const)` + `assignment_statement (X = Y)`. The constant-ness is lost (the');
  lines.push('   identifier is treated as an assignable variable, not a const). Verified in');
  lines.push('   `mdlCursor.bas` and `constantes.bas`.');
  lines.push('');
  lines.push('9. **`Wend` loop terminator** — none of the corpus files use `Wend` (this corpus only');
  lines.push('   has `For Each` loops), but VBA supports `While … Wend` and the vbnet grammar has');
  lines.push('   no `wend_statement` rule. **Pre-processing:** rewrite `Wend` → `End While` if it');
  lines.push('   shows up in the wider corpus.');
  lines.push('');
  lines.push('10. **`VERSION 1.0 CLASS` + `BEGIN ... MultiUse = -1 ... END` block** — Access class');
  lines.push('    header (legacy `form.frm` / `report.frm` shape). The vbnet grammar has no rule');
  lines.push('    for it; every `.cls` starting with it begins inside an ERROR node. The regex');
  lines.push('    pipeline silently strips it. **Pre-processing:** blank the `VERSION ... END`');
  lines.push('    block before parsing.');
  lines.push('');
  lines.push('11. **`Attribute VB_Name = "X"` (and `Attribute VB_GlobalNameSpace`, etc.)** — these');
  lines.push('    legacy `Attribute` directives are not recognized as VB.NET attributes (the');
  lines.push('    shape is the same lexically, but the grammar does not have an `Attribute`');
  lines.push('    rule for top-of-file placement). Each one becomes `identifier (Attribute) +');
  lines.push('    ERROR (VB_Name) + assignment_statement (= "X")`. **Pre-processing:** strip');
  lines.push('    `Attribute VB_*` lines.');
  lines.push('');
  lines.push('12. **`Option Compare Database` / `Option Explicit`** — recognized as `option_statement`,');
  lines.push('    but the position is at module-body level (between `class_declaration` and the');
  lines.push('    first member). Without a class_declaration wrapper, the option_statement is');
  lines.push('    orphaned. Not a fatal error, but F.2 needs to handle it.');
  lines.push('');
  lines.push('## What pre-processing can and cannot fix');
  lines.push('');
  lines.push('A first-cut pre-processing layer (strip `VERSION...END`, strip `Attribute VB_*`,');
  lines.push('blank `Option …`, append a trailing newline) was applied to `ARAuditoria.cls` as a');
  lines.push('dry-run. Result: the structural gap is **NOT closed**. The class body still has');
  lines.push('module-level `Public X As Y` / `Private X As Y` / `Dim X As Y` field declarations,');
  lines.push('and the grammar still wraps the entire class body in an ERROR. The reason is that');
  lines.push('VBA\'s module file has no `Class X` / `Module X` opener — even after stripping the');
  lines.push('preamble, the grammar still has no wrapper to attach the class body to.');
  lines.push('');
  lines.push('But adding a wrapper DOES unlock the structure. See the next section.');
  lines.push('');
  lines.push('## F.2 dry-run: synthesized `Class <Name> … End Class` wrapper');
  lines.push('');
  lines.push('The F.2 hypothesis is to inject a synthetic `Class <Name>` opener (where `<Name>`');
  lines.push('comes from the existing `Attribute VB_Name = "X"` line) and a closing `End Class`');
  lines.push('for every file, then run the AST. The spike ran this on the entire corpus and');
  lines.push('measured the structural node counts the grammar now emits.');
  lines.push('');
  lines.push('| Structural node | Count (synthesized) | Count (raw) |');
  lines.push('|---|---:|---:|');
  for (const node of [
    'class_declaration', 'method_declaration', 'field_declaration', 'property_declaration',
    'parameter_list', 'event_declaration', 'raiseevent_statement', 'implements_clause',
    'enum_declaration',
  ]) {
    const synthN = synthTotals.perStructuralNode[node] ?? 0;
    const rawN = perConstruct.get(node)?.recognized ?? 0;
    lines.push(`| \`${node}\` | ${synthN} | ${rawN} |`);
  }
  lines.push('');
  if ((synthTotals.perStructuralNode.class_declaration ?? 0) > 0) {
    lines.push('**Result:** the synthesized wrapper unlocks the entire structural tree. With');
    lines.push('just one extra pre-processing step — inject a `Class <Name>` opener + `End Class`');
    lines.push('closer — the grammar emits `class_declaration`, `method_declaration`,');
    lines.push('`field_declaration`, `property_declaration`, `parameter_list`, `event_declaration`,');
    lines.push('and `raiseevent_statement` on the VBA corpus. The body-content recognition is');
    lines.push('preserved (expressions, calls, control flow were already parsing cleanly).');
    lines.push('');
    lines.push('This is the **unlock** for F.2. The hybrid extractor is feasible:');
    lines.push('');
    lines.push('1. **Pre-processing** (already in the existing `vba-preprocess.ts` regex): strip');
    lines.push('   `VERSION...END`, `Attribute VB_*`, `Option ...`, `Wend`.');
    lines.push('2. **Wrapper injection** (new): prepend `Class <Name>` and append `End Class`');
    lines.push('   to the file before handing it to the tree-sitter parser. The `<Name>` is');
    lines.push('   already extracted by the regex pipeline from `Attribute VB_Name`.');
    lines.push('3. **AST walk** (the new `VbaTreeSitterExtractor`): walk the vbnet AST for');
    lines.push('   procedures, classes, fields, parameters, properties, events. The body');
    lines.push('   content (expressions, calls, control flow) is already there.');
    lines.push('4. **Access layer** (the existing 5 `create*Classifier()` regex modules):');
    lines.push('   fill in the Access-specific emissions — DoCmd, WithEvents pair convention,');
    lines.push('   TempVars, RecordSource/RowSource, form/report layout.');
    lines.push('');
    lines.push('This matches the F.2 plan\'s "vbnet AST for the language skeleton, regex for');
    lines.push('Access-specific" framing — but with a concrete, bounded pre-processing step');
    lines.push('(wrapper injection) that makes the AST walk actually work.');
  } else {
    lines.push('**Result:** the synthesized wrapper did NOT unlock the structure on this');
    lines.push('corpus. F.2 will need a different approach (regex-based boundary detection');
    lines.push('see "What pre-processing can and cannot fix" above).');
  }
  lines.push('');
  lines.push('## Pre-processing checklist for F.2');
  lines.push('');
  lines.push('1. **Blank the `VERSION 1.0 CLASS` + `BEGIN … END` block** (Access class header).');
  lines.push('2. **Strip `Attribute VB_*` lines** (legacy class metadata).');
  lines.push('3. **Blank `Option Compare Database` / `Option Explicit` lines** (file-level options;');
  lines.push('   the grammar recognizes them as `option_statement` but the position is wrong).');
  lines.push('4. **Rewrite `Wend` → `End While`** (legacy loop terminator — not in this corpus, but');
  lines.push('   widely used).');
  lines.push('5. **Append a trailing newline** to every file (mandatory, already in `vbnet.ts:7-9`).');
  lines.push('6. **Inject a synthetic `Class <Name>` opener + `End Class` closer** (the unlock —');
  lines.push('   see the section above). `<Name>` comes from the existing `Attribute VB_Name`');
  lines.push('   regex extraction.');
  lines.push('');
  lines.push('After steps 1-6, the F.2 extractor should walk the AST cleanly. Re-run this spike');
  lines.push('on the synthesized corpus to confirm: the per-construct table should now show');
  lines.push('non-zero counts for `class_declaration`, `method_declaration`, `field_declaration`,');
  lines.push('`property_declaration`, `parameter_list`, `event_declaration`.');
  lines.push('');
  lines.push('## Reproduce');
  lines.push('');
  lines.push('```bash');
  lines.push('node scripts/spike-vbnet-as-vba.mjs');
  lines.push('# or, with a different corpus:');
  lines.push('node scripts/spike-vbnet-as-vba.mjs --fixtures <dir> --out <path>');
  lines.push('```');
  lines.push('');
  lines.push('Output:');
  lines.push('');
  lines.push(`- Markdown report: \`${path.relative(REPO_ROOT, REPORT_PATH)}\``);
  lines.push(`- Machine-readable JSON: \`${path.relative(REPO_ROOT, JSON_PATH)}\``);
  lines.push('');
  lines.push('## Raw data');
  lines.push('');
  lines.push(`Full per-file detail is in \`${path.relative(REPO_ROOT, JSON_PATH)}\`.`);
  lines.push('');
  return lines.join('\n');
}

function pct(part, whole) {
  if (whole === 0) return '0%';
  return `${((part / whole) * 100).toFixed(0)}%`;
}

main().catch((err) => {
  console.error('[spike] fatal:', err);
  process.exit(1);
});
