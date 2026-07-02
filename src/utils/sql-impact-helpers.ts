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
 * Recursively queries the SQLite graph database backwards to find caller chain ancestors.
 */
function findGraphAncestors(db: any, startNodeId: string): string[] {
  const visited = new Set<string>();
  const ancestors: string[] = [];

  function traverse(nodeId: string) {
    if (visited.has(nodeId)) return;
    visited.add(nodeId);

    const rows = db.prepare('SELECT source FROM edges WHERE target = ?').all(nodeId) as Array<{ source: string }>;
    for (const row of rows) {
      if (row.source !== undefined) {
        ancestors.push(row.source);
        traverse(row.source);
      }
    }
  }

  traverse(startNodeId);
  return ancestors;
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

  // 3. Database traversal for event handlers & controls
  for (const caller of output.callers) {
    // Find matching function node containing caller.line in caller.file
    const matchingNodes = db.prepare(`
      SELECT id, name, kind, file_path
      FROM nodes
      WHERE (file_path = ? OR file_path LIKE ?)
        AND start_line <= ?
        AND end_line >= ?
        AND kind IN ('function', 'event', 'sub')
    `).all(caller.file, `%/${path.basename(caller.file)}`, caller.line, caller.line) as Array<{ id: string; name: string; kind: string; file_path: string }>;

    for (const node of matchingNodes) {
      const ancestors = findGraphAncestors(db, node.id);
      for (const ancestorId of ancestors) {
        const ancestorNode = db.prepare('SELECT name, kind, file_path FROM nodes WHERE id = ?').get(ancestorId) as { name: string; kind: string; file_path: string } | undefined;
        if (ancestorNode) {
          const fileBase = path.basename(ancestorNode.file_path);
          if (['.bas', '.cls', '.frm'].includes(path.extname(fileBase).toLowerCase())) {
            if (!output.downstream_impact.vba_callers.includes(fileBase)) {
              output.downstream_impact.vba_callers.push(fileBase);
            }
          }
          if (ancestorNode.kind === 'control' || ancestorNode.kind === 'form') {
            const formBase = path.basename(ancestorNode.file_path).split('.')[0] || '';
            if (formBase && !output.downstream_impact.forms.includes(formBase)) {
              output.downstream_impact.forms.push(formBase);
            }
          }
        }
      }
    }
  }

  return output;
}
