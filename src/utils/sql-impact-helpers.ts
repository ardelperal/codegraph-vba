import * as fs from 'fs';
import * as path from 'path';

/**
 * Interface representing the extracted bindings of a VBA form or report layout.
 */
export interface FormBindings {
  recordSource?: string;
  rowSources: Array<{ control: string; target: string }>;
}

/**
 * Interface representing the SQL lineage analysis result.
 */
export interface SqlLineage {
  tables: string[];
  lineage: Array<{ source: string; resolved: string }>;
}

/**
 * Traces references to the target query name inside a VBA module's content.
 * Returns 1-based line numbers where references are found.
 */
export function traceVbaCallers(content: string, queryName: string): number[] {
  const lines = content.split(/\r?\n/);
  const matchedLines: number[] = [];

  // Escape query name for regex literal construction
  const escapedQuery = queryName.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');

  // Match OpenRecordset("queryName") or OpenRecordset "queryName"
  const openRecordsetRegex = new RegExp(`OpenRecordset\\s*(?:\\(\\s*["']${escapedQuery}["']|\\s+["']${escapedQuery}["'])`, 'i');
  // Match QueryDefs("queryName")
  const queryDefsRegex = new RegExp(`QueryDefs\\s*\\(\\s*["']${escapedQuery}["']`, 'i');

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line !== undefined && (openRecordsetRegex.test(line) || queryDefsRegex.test(line))) {
      matchedLines.push(i + 1);
    }
  }

  return matchedLines;
}

/**
 * Parses a VBA form definition (.form.txt/.report.txt) and extracts RecordSource and RowSource bindings.
 */
export function extractFormBindings(content: string): FormBindings {
  const lines = content.split(/\r?\n/);
  let recordSource: string | undefined;
  const rowSources: Array<{ control: string; target: string }> = [];

  interface Container {
    type: string;
    name?: string;
  }
  const stack: Container[] = [];

  for (const line of lines) {
    const trimmed = line.trim();

    // Check for start of form controls (e.g., Begin ComboBox, Begin Form)
    const beginMatch = trimmed.match(/^Begin\s+(\w+)/i);
    if (beginMatch && beginMatch[1] !== undefined) {
      stack.push({ type: beginMatch[1] });
      continue;
    }

    // Check for end of block
    if (trimmed === 'End') {
      stack.pop();
      continue;
    }

    // Keep track of current control Name property
    const nameMatch = trimmed.match(/Name\s*=\s*"([^"]+)"/i);
    if (nameMatch && nameMatch[1] !== undefined && stack.length > 0) {
      const current = stack[stack.length - 1];
      if (current !== undefined) {
        current.name = nameMatch[1];
      }
    }

    // Extract RecordSource at Form level
    const recordSourceMatch = trimmed.match(/RecordSource\s*=\s*"([^"]+)"/i);
    if (recordSourceMatch && recordSourceMatch[1] !== undefined) {
      recordSource = recordSourceMatch[1];
    }

    // Extract RowSource inside ComboBox/ListBox
    const rowSourceMatch = trimmed.match(/RowSource\s*=\s*"([^"]+)"/i);
    if (rowSourceMatch && rowSourceMatch[1] !== undefined && stack.length > 0) {
      const rawVal = rowSourceMatch[1];
      let target = rawVal;

      // If it is a query (starts with SELECT), extract table/query name from the FROM clause
      if (/select\s+/i.test(rawVal)) {
        const fromMatch = rawVal.match(/from\s+([a-zA-Z0-9_]+)/i);
        if (fromMatch && fromMatch[1] !== undefined) {
          target = fromMatch[1];
        }
      }

      const currentControl = stack[stack.length - 1];
      if (currentControl !== undefined) {
        const controlName = currentControl.name || currentControl.type;
        rowSources.push({ control: controlName, target });
      }
    }
  }

  return { recordSource, rowSources };
}

/**
 * Resolves table and column aliases from a SQL string to build a column-level lineage.
 */
export function resolveSqlLineage(sql: string): SqlLineage {
  const normalized = sql.replace(/\s+/g, ' ');
  const SQL_KEYWORDS = new Set([
    'as', 'on', 'inner', 'left', 'right', 'join', 'where', 'order', 'group', 'by',
    'and', 'or', 'select', 'from', 'using', 'cross', 'outer'
  ]);

  const aliasMap = new Map<string, string>(); // alias lowercased -> table
  const tables = new Set<string>();

  // Extract tables and aliases from FROM and JOIN clauses
  const fromJoinPattern = /(?:from|join)\s+([a-zA-Z0-9_]+)(?:\s+as\s+([a-zA-Z0-9_]+)|\s+([a-zA-Z0-9_]+))?/gi;
  let match;
  while ((match = fromJoinPattern.exec(normalized)) !== null) {
    const table = match[1];
    if (table !== undefined) {
      tables.add(table);

      const alias = match[2] || match[3];
      if (alias && !SQL_KEYWORDS.has(alias.toLowerCase())) {
        aliasMap.set(alias.toLowerCase(), table);
      }
      // A table name can also qualify itself, e.g., tblTable.col
      aliasMap.set(table.toLowerCase(), table);
    }
  }

  // Extract column references (e.g., alias.col) and map to resolved tables
  const colPattern = /\b([a-zA-Z0-9_]+)\.([a-zA-Z0-9_*]+)\b/g;
  const lineageMap = new Map<string, string>();

  while ((match = colPattern.exec(normalized)) !== null) {
    const prefix = match[1];
    const col = match[2];
    if (prefix !== undefined && col !== undefined) {
      const source = `${prefix}.${col}`;
      const resolvedTable = aliasMap.get(prefix.toLowerCase());

      if (resolvedTable) {
        lineageMap.set(source, `${resolvedTable}.${col}`);
      }
    }
  }

  const lineage = Array.from(lineageMap.entries()).map(([source, resolved]) => ({
    source,
    resolved
  }));

  return {
    tables: Array.from(tables),
    lineage
  };
}

/**
 * Internal recursive file finder.
 */
function getAllFiles(dir: string): string[] {
  let results: string[] = [];
  if (!fs.existsSync(dir)) return [];
  const list = fs.readdirSync(dir);
  for (const file of list) {
    const filePath = path.join(dir, file);
    const stat = fs.statSync(filePath);
    if (stat && stat.isDirectory()) {
      results = results.concat(getAllFiles(filePath));
    } else {
      results.push(filePath);
    }
  }
  return results;
}

/**
 * Node kinds that can own a call chain — a VBA procedure, under any of the
 * kind names the graph has used.
 */
const CALLER_NODE_KINDS = ['function', 'method', 'event', 'sub'];

/**
 * Node kinds that represent a UI object an event handler can be wired to.
 * The canonical Access kinds first, then the legacy ones older graphs used.
 */
const UI_NODE_KINDS = new Set([
  'form-instance-control',
  'form-layout',
  'report-layout',
  'control',
  'form',
  'report',
]);

const LAYOUT_NODE_KINDS = new Set(['form-layout', 'report-layout']);

interface GraphNodeRow {
  id: string;
  name: string;
  kind: string;
  file_path: string;
}

/**
 * Walks the graph backwards from a procedure to every procedure that can
 * reach it.
 *
 * Only caller relationships are followed. `calls` is the explicit one and
 * `defines-event` the legacy one; a `references` edge counts only when it
 * comes FROM a procedure, which is how the resolver stores VBA's
 * statement-form Sub call (`SaveOrderTotals` on its own line resolves as an
 * ambiguous bare identifier — see issue #265 — so the edge is `references`,
 * not `calls`). Containment is deliberately not a caller relationship: a
 * module contains a procedure, it does not call it.
 */
function findCallChain(db: any, startNodeId: string): string[] {
  const visited = new Set<string>([startNodeId]);
  const chain: string[] = [startNodeId];
  const stmt = db.prepare(
    `SELECT e.source AS id
       FROM edges e
       JOIN nodes n ON n.id = e.source
      WHERE e.target = ?
        AND (e.kind IN ('calls', 'defines-event')
             OR (e.kind = 'references'
                 AND n.kind IN (${CALLER_NODE_KINDS.map(() => '?').join(', ')})))
      ORDER BY e.source`
  );

  const queue = [startNodeId];
  while (queue.length > 0) {
    const nodeId = queue.shift()!;
    const rows = stmt.all(nodeId, ...CALLER_NODE_KINDS) as Array<{ id: string }>;
    for (const row of rows) {
      if (row.id === undefined || visited.has(row.id)) continue;
      visited.add(row.id);
      chain.push(row.id);
      queue.push(row.id);
    }
  }

  return chain;
}

/**
 * Names the form or report that owns a UI node.
 *
 * A layout node names itself. A control is resolved through the `contains`
 * edge its layout emits — ownership is a graph relationship, not a guess
 * made from the file name. The basename fallback only fires for a control
 * whose layout file was never indexed (a cross-file stub), and for the
 * legacy `control`/`form` kinds that predate layout containment.
 */
function resolveOwningLayout(db: any, uiNodeId: string): string | null {
  const node = db
    .prepare('SELECT id, name, kind, file_path FROM nodes WHERE id = ?')
    .get(uiNodeId) as GraphNodeRow | undefined;
  if (!node) return null;

  if (LAYOUT_NODE_KINDS.has(node.kind)) return node.name;

  const owner = db
    .prepare(
      `SELECT n.name AS name
         FROM edges e
         JOIN nodes n ON n.id = e.source
        WHERE e.target = ?
          AND e.kind = 'contains'
          AND n.kind IN ('form-layout', 'report-layout')
        ORDER BY e.source
        LIMIT 1`
    )
    .get(uiNodeId) as { name: string } | undefined;
  if (owner) return owner.name;

  if (!node.file_path) return null;
  const base = path.basename(node.file_path);
  const stripped = base.replace(/\.(form|report)\.txt$/i, '');
  return (stripped === base ? base.split('.')[0] : stripped) || null;
}

/**
 * Runs a combined database and file-level downstream impact analysis.
 */
export function runImpactAnalysis(
  db: any,
  workspaceDir: string,
  queryName: string
): any {
  const output: any = {
    query_name: queryName,
    callers: [],
    form_bindings: [],
    tables_touched: [],
    lineage: [],
    downstream_impact: {
      queries: [queryName],
      forms: [],
      vba_callers: []
    },
    warnings: []
  };

  // 1. Resolve SQL lineage
  const sqlPath = path.join(workspaceDir, 'queries', `${queryName}.sql`);
  if (fs.existsSync(sqlPath)) {
    const sql = fs.readFileSync(sqlPath, 'utf-8');
    const sqlLin = resolveSqlLineage(sql);
    output.tables_touched = sqlLin.tables;
    output.lineage = sqlLin.lineage;
  }

  // 2. Scan and parse files
  const allFiles = getAllFiles(workspaceDir);
  const matchedForms = new Set<string>();

  for (const file of allFiles) {
    const relativePath = path.relative(workspaceDir, file).replace(/\\/g, '/');
    const ext = path.extname(file).toLowerCase();

    // VBA files
    if (['.bas', '.cls', '.frm'].includes(ext)) {
      const content = fs.readFileSync(file, 'utf-8');
      const lines = traceVbaCallers(content, queryName);
      if (lines.length > 0) {
        const baseName = path.basename(file);
        if (!output.downstream_impact.vba_callers.includes(baseName)) {
          output.downstream_impact.vba_callers.push(baseName);
        }
        const linesSplit = content.split(/\r?\n/);
        for (const lineNum of lines) {
          const contextLine = linesSplit[lineNum - 1];
          output.callers.push({
            file: relativePath,
            line: lineNum,
            context: contextLine !== undefined ? contextLine : ''
          });
        }
      }
    }

    // Form/Report layout files
    const lowerFile = file.toLowerCase();
    const isFormOrReport = lowerFile.endsWith('.form.txt') || lowerFile.endsWith('.report.txt');
    if (isFormOrReport) {
      const content = fs.readFileSync(file, 'utf-8');
      const bindings = extractFormBindings(content);
      const formName = path.basename(file).replace(/\.(form|report)\.txt$/i, '');

      let isFormBound = false;
      if (bindings.recordSource && (bindings.recordSource === queryName || output.tables_touched.includes(bindings.recordSource))) {
        output.form_bindings.push({
          file: relativePath,
          control: 'Form',
          property: 'RecordSource',
          target: bindings.recordSource
        });
        isFormBound = true;
      }

      for (const rowSrc of bindings.rowSources) {
        if (rowSrc.target === queryName || output.tables_touched.includes(rowSrc.target)) {
          output.form_bindings.push({
            file: relativePath,
            control: rowSrc.control,
            property: 'RowSource',
            target: rowSrc.target
          });
          isFormBound = true;
        }
      }

      if (isFormBound) {
        matchedForms.add(formName);
        if (!output.downstream_impact.forms.includes(formName)) {
          output.downstream_impact.forms.push(formName);
        }
      }
    }
  }

  // 3. Database traversal: from the procedures that touch the query, back up
  //    the call chain and across the handler -> control/layout binding.
  //
  //    The binding is stored HANDLER -> control (and HANDLER -> layout for a
  //    form-level event), so the owning form is found by following the
  //    handler's OUTGOING `event-handler` edge — not by walking further up
  //    the ancestor chain, which never reaches the UI at all. That inverted
  //    assumption is why a query reached only through code, with no
  //    RecordSource/RowSource binding, used to report no affected form.
  const stmtUiTargets = db.prepare(
    "SELECT target AS id FROM edges WHERE source = ? AND kind = 'event-handler' ORDER BY target"
  );

  const addForm = (formName: string | null) => {
    if (formName && !output.downstream_impact.forms.includes(formName)) {
      output.downstream_impact.forms.push(formName);
    }
  };

  for (const caller of output.callers) {
    // Find the procedure node whose line range contains this reference.
    const matchingNodes = db.prepare(`
      SELECT id, name, kind, file_path
      FROM nodes
      WHERE (file_path = ? OR file_path LIKE ?)
        AND start_line <= ?
        AND end_line >= ?
        AND kind IN (${CALLER_NODE_KINDS.map(() => '?').join(', ')})
    `).all(
      caller.file,
      `%/${path.basename(caller.file)}`,
      caller.line,
      caller.line,
      ...CALLER_NODE_KINDS,
    ) as GraphNodeRow[];

    for (const node of matchingNodes) {
      const chain = findCallChain(db, node.id);
      for (const chainNodeId of chain) {
        const chainNode = db
          .prepare('SELECT id, name, kind, file_path FROM nodes WHERE id = ?')
          .get(chainNodeId) as GraphNodeRow | undefined;
        if (!chainNode) continue;

        if (chainNodeId !== node.id) {
          const fileBase = path.basename(chainNode.file_path ?? '');
          if (['.bas', '.cls', '.frm'].includes(path.extname(fileBase).toLowerCase())) {
            if (!output.downstream_impact.vba_callers.includes(fileBase)) {
              output.downstream_impact.vba_callers.push(fileBase);
            }
          }
        }

        // A legacy graph put the control itself in the caller chain.
        if (UI_NODE_KINDS.has(chainNode.kind)) {
          addForm(resolveOwningLayout(db, chainNodeId));
          continue;
        }

        for (const ui of stmtUiTargets.all(chainNodeId) as Array<{ id: string }>) {
          addForm(resolveOwningLayout(db, ui.id));
        }
      }
    }
  }

  return output;
}
