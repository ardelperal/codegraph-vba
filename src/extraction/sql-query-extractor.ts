/**
 * SqlQueryExtractor — regex extractor for Dysflow-exported saved Access
 * queries (`queries/<Name>.sql`).
 *
 * Dysflow exports each saved Access QueryDef as a raw `.sql` file alongside a
 * `queries.json` manifest. A `.sql` file is only routed here when that sibling
 * manifest is present (the directory-discovery gate in
 * `src/extraction/index.ts`), so non-Access repos with ordinary `.sql` files
 * are never affected.
 *
 * Emits:
 *  - one `file` node (so the watcher tracks the file);
 *  - one `query` node named after the file basename (e.g. `Consulta3`);
 *  - one `references` edge from the `query` node to each table named after a
 *    `FROM` / `JOIN` / `INTO` / `UPDATE` keyword, tagged
 *    `metadata.synthesizedBy = 'sql-query-table'`. Table targets are synthetic
 *    `class`-placeholder nodes (the same shape `VbaExtractor` uses for SQL
 *    table references), so the data layer is queryable and table usage is
 *    traceable. Bracketed names (`[Order Details]`) are unwrapped; a table
 *    referenced multiple times in one query produces exactly one node + edge.
 *
 * Issue #203: the table-name regex and reserved-word reject list live in
 * the shared leaf module `src/extraction/sql-table-scan.ts` — the same
 * scanner every other call site (`vba/sql-wrapper.ts`,
 * `vba-form-extractor.ts`) imports, so a SQL reserved word can never
 * be emitted as a table name across the project.
 */
import * as path from 'path';
import {
  Node,
  Edge,
  ExtractionResult,
  ExtractionError,
} from '../types';
import { generateNodeId } from './tree-sitter-helpers';
import { scanSqlTables } from './sql-table-scan';

export class SqlQueryExtractor {
  private filePath: string;
  private source: string;
  private nodes: Node[] = [];
  private edges: Edge[] = [];
  private errors: ExtractionError[] = [];

  constructor(filePath: string, source: string) {
    this.filePath = filePath;
    this.source = source;
  }

  extract(): ExtractionResult {
    const startTime = Date.now();
    try {
      const queryName = this.basenameWithoutExt();
      this.nodes.push(this.createFileNode());

      const queryId = generateNodeId(this.filePath, 'query', queryName, 1);
      this.nodes.push({
        id: queryId,
        kind: 'query',
        name: queryName,
        qualifiedName: queryName,
        filePath: this.filePath,
        language: 'sql',
        startLine: 1,
        endLine: this.source.split('\n').length,
        startColumn: 0,
        endColumn: 0,
        updatedAt: Date.now(),
      });

      this.sweepTables(queryId);
    } catch (error) {
      this.errors.push({
        message: `SQL query extraction error: ${error instanceof Error ? error.message : String(error)}`,
        filePath: this.filePath,
        severity: 'error',
        code: 'parse_error',
      });
    }

    return {
      nodes: this.nodes,
      edges: this.edges,
      unresolvedReferences: [],
      errors: this.errors,
      durationMs: Date.now() - startTime,
    };
  }

  private basenameWithoutExt(): string {
    return path.basename(this.filePath).replace(/\.[^.]+$/, '');
  }

  private createFileNode(): Node {
    const lines = this.source.split('\n');
    return {
      id: generateNodeId(this.filePath, 'file', this.filePath, 1),
      kind: 'file',
      name: path.basename(this.filePath),
      qualifiedName: this.filePath,
      filePath: this.filePath,
      language: 'sql',
      startLine: 1,
      endLine: lines.length,
      startColumn: 0,
      endColumn: lines[lines.length - 1]?.length ?? 0,
      updatedAt: Date.now(),
    };
  }

  /**
   * Scan the SQL for table names and emit one synthetic `class` node + one
   * `references` edge per distinct table. The table node id is line-independent
   * (line 0) so the same table referenced N times collapses to one node.
   *
   * Issue #203: delegates to the shared `scanSqlTables` scanner in
   * `src/extraction/sql-table-scan.ts` so a SQL reserved word can never
   * be emitted as a table name — `WHERE x=1`, `ORDER BY a`, `SET a=1`
   * never become phantom `class` nodes here.
   */
  private sweepTables(queryId: string): void {
    const seen = new Set<string>();
    for (const row of scanSqlTables(this.source)) {
      const key = row.table.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);

      const tableId = generateNodeId(this.filePath, 'class', row.table, 0);
      this.nodes.push({
        id: tableId,
        kind: 'class', // placeholder kind; cross-file resolution re-types at lookup
        name: row.table,
        qualifiedName: row.table,
        filePath: this.filePath,
        language: 'sql',
        startLine: 1,
        endLine: 1,
        startColumn: 0,
        endColumn: 0,
        updatedAt: Date.now(),
      });
      this.edges.push({
        source: queryId,
        target: tableId,
        kind: 'references',
        provenance: 'heuristic',
        metadata: { synthesizedBy: 'sql-query-table' },
        line: 1,
        column: 0,
      });
    }
  }
}
