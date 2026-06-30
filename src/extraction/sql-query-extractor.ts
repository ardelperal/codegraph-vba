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
 */
import * as path from 'path';
import {
  Node,
  Edge,
  ExtractionResult,
  ExtractionError,
} from '../types';
import { generateNodeId } from './tree-sitter-helpers';

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

  /**
   * Table name following `FROM` / `JOIN` / `INTO` / `UPDATE`. Captures either a
   * bracketed name (which may contain spaces, e.g. `[Order Details]`) or a bare
   * identifier. `\p{L}` covers accented identifiers common in localized schemas.
   */
  private static readonly TABLE_RE =
    /\b(?:FROM|JOIN|INTO|UPDATE)\s+(\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*)/giu;

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
   */
  private sweepTables(queryId: string): void {
    const seen = new Set<string>();
    const re = new RegExp(
      SqlQueryExtractor.TABLE_RE.source,
      SqlQueryExtractor.TABLE_RE.flags,
    );
    let m: RegExpExecArray | null;
    while ((m = re.exec(this.source)) !== null) {
      const raw = (m[1] ?? '').trim();
      const table = raw.replace(/^\[|\]$/g, '').trim();
      if (!table) continue;
      const key = table.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);

      const tableId = generateNodeId(this.filePath, 'class', table, 0);
      this.nodes.push({
        id: tableId,
        kind: 'class', // placeholder kind; cross-file resolution re-types at lookup
        name: table,
        qualifiedName: table,
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
