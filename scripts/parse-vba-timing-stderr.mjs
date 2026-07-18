#!/usr/bin/env node
// scripts/parse-vba-timing-stderr.mjs
//
// Parse per-file `[vba-timing]` blocks from `CODEGRAPH_VBA_TIMING=2 codegraph`
// stderr logs (issue #156 / #166) and emit a median-across-runs table for
// the four required measurement targets plus a corpus-wide aggregate.
//
// Usage:
//   node scripts/parse-vba-timing-stderr.mjs <log1> [log2] [log3] ...
//
// Each log is a stderr capture from a single `codegraph init` or
// `codegraph index` run. Output is plain markdown tables written to
// stdout — pipe to a file with `> summary.md`.
//
// Scope (intentionally narrow — issue #166):
//   - Parses the seven documented stages:
//       preprocess.stripVbaComments, preprocess.joinLineContinuations,
//       preprocess.conditionalCompilation, preprocess.cc.lexer+parser,
//       walk.procedures, walk.main, plus the 6 classifiers (any name).
//   - Aggregates per file across runs (median).
//   - Emits per-stage + corpus-wide tables.
//   - Knows nothing about specific files — caller filters with `--files`
//     or the REQUIRED list below.
//
// The format `[vba-timing] <basename>` followed by indented `bucket: stage Nms`
// lines is the contract documented at src/extraction/vba-timing.ts:20.

import * as fs from 'node:fs';
import * as path from 'node:path';

const REQUIRED_TARGETS = [
  'ACAuditoriaOperaciones.cls',
  'ARAuditoria.cls',
  'Form_FormGestionRiesgos.cls',
  'mdlCursor.bas',
];

function parseLog(logPath) {
  const body = fs.readFileSync(logPath, 'utf8');
  const lines = body.split(/\r?\n/);
  const result = new Map();

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const headerMatch = /^\[vba-timing\] (.+\.(?:cls|bas|frm|dsr))$/.exec(line);
    if (headerMatch) {
      const file = headerMatch[1];
      const t = {
        preprocessStripMs: 0,
        preprocessJoinMs: 0,
        preprocessCcMs: 0,
        preprocessCcInnerMs: 0,
        walkProceduresMs: 0,
        walkMainMs: 0,
        classifiersMs: 0,
        classifierBreakdown: {},
        // The extractor's post-preprocessing `source.split('\n').length`
        // count, surfaced via the `(n=N)` annotation on classifier lines.
        // NOT the file's physical line count — see
        // docs/vba-extraction-perf.md for the distinction.
        splitSlots: 0,
      };
      i++;
      while (i < lines.length) {
        const sub = lines[i].trim();
        if (sub === '' || sub.startsWith('[')) break;
        const stageMatch =
          /^(preprocess|classifiers|walk):\s+(\S+)\s+([\d.]+)ms(?:\s*\(n=(\d+)\))?/.exec(sub);
        if (stageMatch) {
          const bucket = stageMatch[1];
          const name = stageMatch[2];
          const ms = parseFloat(stageMatch[3]);
          const n = stageMatch[4] ? parseInt(stageMatch[4], 10) : 0;
          if (bucket === 'preprocess' && name === 'preprocess.stripVbaComments') {
            t.preprocessStripMs = ms;
          } else if (bucket === 'preprocess' && name === 'preprocess.joinLineContinuations') {
            t.preprocessJoinMs = ms;
          } else if (bucket === 'preprocess' && name === 'preprocess.conditionalCompilation') {
            t.preprocessCcMs = ms;
          } else if (bucket === 'preprocess' && name === 'preprocess.cc.lexer+parser') {
            t.preprocessCcInnerMs = ms;
          } else if (bucket === 'walk' && name === 'walk.procedures') {
            t.walkProceduresMs = ms;
          } else if (bucket === 'walk' && name === 'walk.main') {
            t.walkMainMs = ms;
          } else if (bucket === 'classifiers') {
            t.classifierBreakdown[name] = ms;
            t.classifiersMs += ms;
            if (n > 0) t.splitSlots = n;
          }
        }
        i++;
      }
      result.set(file, t);
    } else {
      i++;
    }
  }
  return result;
}

function totalMs(t) {
  return (
    t.preprocessStripMs +
    t.preprocessJoinMs +
    t.preprocessCcMs +
    t.preprocessCcInnerMs +
    t.walkProceduresMs +
    t.walkMainMs +
    t.classifiersMs
  );
}

function median(values) {
  if (values.length === 0) return 0;
  const s = [...values].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 === 1 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

function fmt(n, digits = 3) {
  return Number.isFinite(n) ? n.toFixed(digits) : '—';
}

function renderPerFile(file, rows) {
  const out = [];
  out.push(`### \`${file}\`\n`);
  out.push('| run | totalMs | preprocess.strip | preprocess.join | preprocess.cc | preprocess.cc.inner | walk.procedures | walk.main | classifiers.sum | splitSlots |');
  out.push('|---|---|---|---|---|---|---|---|---|---|');
  for (const r of rows) {
    if (!r.t) {
      out.push(`| ${r.run} | (missing) | | | | | | | | |`);
    } else {
      const t = r.t;
      out.push(
        `| ${r.run} | ${fmt(t.totalMs ?? totalMs(t))} | ${fmt(t.preprocessStripMs)} | ${fmt(t.preprocessJoinMs)} | ${fmt(t.preprocessCcMs)} | ${fmt(t.preprocessCcInnerMs)} | ${fmt(t.walkProceduresMs)} | ${fmt(t.walkMainMs)} | ${fmt(t.classifiersMs)} | ${t.splitSlots || '—'} |`,
      );
    }
  }
  const totals = rows.map((r) => (r.t ? totalMs(r.t) : 0)).filter((v) => v > 0);
  const pre = rows.map((r) => r.t?.preprocessStripMs ?? 0).filter((v) => v >= 0);
  const join = rows.map((r) => r.t?.preprocessJoinMs ?? 0).filter((v) => v >= 0);
  const cc = rows.map((r) => r.t?.preprocessCcMs ?? 0).filter((v) => v >= 0);
  const ccIn = rows.map((r) => r.t?.preprocessCcInnerMs ?? 0).filter((v) => v >= 0);
  const walkP = rows.map((r) => r.t?.walkProceduresMs ?? 0).filter((v) => v >= 0);
  const walkM = rows.map((r) => r.t?.walkMainMs ?? 0).filter((v) => v >= 0);
  const clf = rows.map((r) => r.t?.classifiersMs ?? 0).filter((v) => v >= 0);
  const slots = rows.map((r) => r.t?.splitSlots ?? 0).filter((v) => v > 0);
  out.push(
    `| **MEDIAN** | ${fmt(median(totals))} | ${fmt(median(pre))} | ${fmt(median(join))} | ${fmt(median(cc))} | ${fmt(median(ccIn))} | ${fmt(median(walkP))} | ${fmt(median(walkM))} | ${fmt(median(clf))} | ${slots.length ? median(slots).toFixed(0) : '—'} |`,
  );
  out.push('');
  out.push('Classifier breakdown (per run, ms):');
  out.push(
    '| classifier | ' +
      rows.map((r) => r.run).join(' | ') +
      ' | MEDIAN |',
  );
  out.push('|---|' + rows.map(() => '---').join('|') + '|---|');
  const clfNames = new Set();
  for (const r of rows) {
    if (r.t) {
      for (const n of Object.keys(r.t.classifierBreakdown)) clfNames.add(n);
    }
  }
  for (const clf of [...clfNames].sort()) {
    const vals = rows
      .map((r) => r.t?.classifierBreakdown[clf] ?? 0)
      .filter((v) => v >= 0);
    out.push(
      `| ${clf} | ` +
        rows
          .map((r) =>
            r.t && r.t.classifierBreakdown[clf] !== undefined
              ? fmt(r.t.classifierBreakdown[clf])
              : '—',
          )
          .join(' | ') +
        ` | ${fmt(median(vals))} |`,
    );
  }
  out.push('');
  return out.join('\n');
}

function renderCorpusTotals(perRun) {
  const out = [];
  out.push('## Corpus-wide totals (sum across all `.cls`/`.bas` files, ms)\n');
  out.push('| run | files | totalMs | preprocess | walk | classifiers |');
  out.push('|---|---|---|---|---|---|');
  for (const r of perRun) {
    let totalAll = 0,
      preAll = 0,
      walkAll = 0,
      clfAll = 0;
    for (const t of r.data.values()) {
      totalAll += totalMs(t);
      preAll += t.preprocessStripMs + t.preprocessJoinMs + t.preprocessCcMs;
      walkAll += t.walkProceduresMs + t.walkMainMs;
      clfAll += t.classifiersMs;
    }
    out.push(
      `| ${r.name} | ${r.data.size} | ${totalAll.toFixed(0)} | ${preAll.toFixed(0)} | ${walkAll.toFixed(0)} | ${clfAll.toFixed(0)} |`,
    );
  }
  out.push('');
  return out.join('\n');
}

function main(argv) {
  const args = argv.slice(2);
  if (args.length === 0) {
    console.error(
      'Usage: node scripts/parse-vba-timing-stderr.mjs <log1> [log2] [log3] ...',
    );
    console.error('');
    console.error(
      'Each <log> is a stderr capture from a `codegraph init` or `codegraph index`',
    );
    console.error(
      'run with CODEGRAPH_VBA_TIMING=1 or =2. Emits a markdown summary on stdout.',
    );
    process.exit(1);
  }
  for (const a of args) {
    if (!fs.existsSync(a)) {
      console.error(`error: ${a}: not found`);
      process.exit(1);
    }
  }
  const perRun = args.map((p) => ({ name: path.basename(p), data: parseLog(p) }));

  console.log(
    `# VBA timing parse — ${perRun.length} run(s): ${perRun.map((r) => r.name).join(', ')}\n`,
  );
  console.log('## Required measurement targets\n');
  for (const file of REQUIRED_TARGETS) {
    const rows = perRun.map((r) => ({ run: r.name, t: r.data.get(file) }));
    console.log(renderPerFile(file, rows));
  }
  console.log(renderCorpusTotals(perRun));
}

main(process.argv);
