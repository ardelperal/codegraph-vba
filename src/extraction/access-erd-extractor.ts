/**
 * AccessErdExtractor — table structure from the Access ERD export
 * (`ERD/Estructura_Datos.md`). Issue #257, half B.
 *
 * Table references in the graph are otherwise synthetic `class` placeholders
 * recovered from SQL text: a name and nothing else. Columns cannot be inferred
 * from a `SELECT` list without being wrong by construction — but an Access
 * application can ship a GENERATED structure dump of its backend, and that dump
 * is a declaration rather than an inference:
 *
 * ```markdown
 * # Estructura de Datos: Gestion_Riesgos_Datos.accdb
 *
 * ## Tabla: TbAnexos
 * | Campo | Tipo | Longitud |
 * | :--- | :--- | :--- |
 * | IDAnexo | 4 | 4 |
 *
 * ## Tabla: TbAplicaciones
 * > (LINKED) **Tabla Vinculada**
 * > *Origen:* TbAplicaciones
 * > *Conexión:* MS Access;PWD=***;DATABASE=\\srv\...\Lanzadera_Datos.accdb
 * ```
 *
 * Emits:
 *  - one `file` node (so the watcher tracks the document);
 *  - one `class` node per `## Tabla:` section — the SAME kind the SQL
 *    placeholders already use, so a post-extraction resolution pass can promote
 *    the ERD declaration to canonical and repoint the placeholders onto it
 *    (`resolveAccessErdTableNodes` in `src/resolution/index.ts`);
 *  - one `type_member` node per field row, `contains`-ed by its table and
 *    carrying `metadata: { accessType, length, position }`;
 *  - for a `(LINKED)` table, a `references` edge to the external backend file
 *    it really lives in, tagged `synthesizedBy: 'vba-linked-table'`.
 *
 * ## Two gates, both required
 *
 * `ERD/` is a very ordinary folder name, so a path match alone would drag
 * hand-written diagrams in every repo into the graph. Routing therefore gates
 * on the path (`isAccessErdFile` in `grammars.ts`) AND on the document's own
 * header line, exactly the way a `.sql` is only an Access saved query when a
 * sibling `queries.json` sits beside it. A markdown file that fails the header
 * gate produces no nodes, no edges and no errors.
 *
 * The header gate has one trap worth stating out loud: the real generated file
 * starts with a **UTF-8 BOM**, so its first line is
 * `\uFEFF# Estructura de Datos: …`. A gate that tests the raw first non-empty
 * line rejects the one input this whole feature exists for while happily
 * accepting every hand-written fixture. The BOM is stripped before the test.
 *
 * ## Optional by design
 *
 * Only some Access projects ship this export. There is no "missing ERD"
 * warning, no empty placeholder node, and no error: a project without one
 * indexes byte-for-byte as it did before.
 */
import * as path from 'path';
import { Node, Edge, ExtractionResult, ExtractionError } from '../types';
import { generateNodeId } from './tree-sitter-helpers';
import { buildExternalBackendNode, normalizeBackendPath } from './sql-table-scan';

/**
 * `metadata.synthesizedBy` stamped on every table node this extractor emits.
 * The resolution pass keys on it to tell an ERD DECLARATION apart from a SQL
 * placeholder that merely mentions the same name.
 */
export const ERD_TABLE_SYNTHESIZED_BY = 'access-erd-table';

/**
 * `metadata.synthesizedBy` stamped on the `references` edge from a linked table
 * to the external backend file it lives in.
 *
 * Deliberately NOT `vba-external-backend` (#256's tag): the NODE shape is
 * shared with #256 — a `file` node keyed on the normalized path with
 * `metadata.external` — but the mechanism that discovered it differs (a linked
 * table's connection string here, a query's `IN "<path>"` clause there), and
 * the tag is what tells the two apart in the graph.
 */
export const ERD_LINKED_TABLE_SYNTHESIZED_BY = 'vba-linked-table';

/** The header every generated Access structure export opens with. */
const ERD_HEADER_PREFIX = '# Estructura de Datos:';

/** `## Tabla: <Name>` — the section header that opens one table. */
const TABLE_HEADER_RE = /^##\s+Tabla:\s*(.+?)\s*$/;

/** A `(LINKED)` blockquote marker introducing a linked-table origin block. */
const LINKED_MARKER_RE = /^>\s*\(LINKED\)/i;

/** `> *Origen:* <name>` — the table's name in the backend it is linked from. */
const ORIGIN_RE = /^>\s*\*\s*Origen\s*:?\s*\*\s*(.+?)\s*$/i;

/**
 * `DATABASE=<path>` inside an Access connection string. An ODBC link has no
 * such operand (it names a DSN instead), and a link we cannot key on a file
 * emits nothing rather than a guessed node — the same "silent beats wrong"
 * doctrine the SQL table scanner follows.
 */
const CONNECT_DATABASE_RE = /DATABASE\s*=\s*([^;]+)/i;

/** A markdown separator cell (`:---`, `---`, `---:`, `:---:`). */
const SEPARATOR_CELL_RE = /^:?-{3,}:?$/;

/**
 * Strip a leading UTF-8 BOM. The generated export always carries one; every
 * hand-written fixture does not.
 */
function stripBom(source: string): string {
  return source.charCodeAt(0) === 0xfeff ? source.slice(1) : source;
}

/**
 * Content-shape gate: is `source` a generated Access structure export — i.e.
 * does its FIRST non-empty line open with `# Estructura de Datos:`?
 *
 * Pure and exported so the gate is testable on its own; the path match is
 * applied separately by `isAccessErdFile` in `grammars.ts`.
 */
export function isAccessErdShape(source: string): boolean {
  if (!source) return false;
  for (const line of stripBom(source).split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    return trimmed.startsWith(ERD_HEADER_PREFIX);
  }
  return false;
}

/** One parsed `| a | b | c |` markdown row, cells trimmed. */
function parseRowCells(line: string): string[] | null {
  const trimmed = line.trim();
  if (!trimmed.startsWith('|')) return null;
  const inner = trimmed.replace(/^\|/, '').replace(/\|$/, '');
  return inner.split('|').map((c) => c.trim());
}

/**
 * Access reports a field's type and length as integers (`10`, `255`), but a
 * differently-generated export may spell them out. Keep an integer as a number
 * and anything else as its raw text rather than coercing to `NaN`.
 */
function coerceCell(raw: string): number | string {
  if (/^-?\d+$/.test(raw)) return Number(raw);
  return raw;
}

interface PendingTable {
  node: Node;
  fieldCount: number;
}

export class AccessErdExtractor {
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
      if (isAccessErdShape(this.source)) {
        this.parse();
      }
    } catch (error) {
      this.errors.push({
        message: `Access ERD extraction error: ${
          error instanceof Error ? error.message : String(error)
        }`,
        filePath: this.filePath,
        severity: 'warning',
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

  private parse(): void {
    const lines = stripBom(this.source).split(/\r?\n/);
    this.nodes.push(this.createFileNode(lines));

    /** Table nodes already emitted, keyed by lower-cased name. */
    const tables = new Map<string, PendingTable>();
    /** External backend node ids already emitted, so one file is one node. */
    const backends = new Set<string>();
    /** Field ids already emitted, so a repeated column collapses. */
    const fieldIds = new Set<string>();

    let current: PendingTable | null = null;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i]!;
      const lineNumber = i + 1;

      const header = TABLE_HEADER_RE.exec(line);
      if (header) {
        current = this.openTable(tables, header[1]!, lineNumber);
        continue;
      }
      if (!current) continue;

      if (LINKED_MARKER_RE.test(line)) {
        this.markLinked(current, lines, i, lineNumber, backends);
        continue;
      }

      this.maybeAddField(current, line, lineNumber, fieldIds);
    }
  }

  /**
   * Open (or reuse) the table node for a `## Tabla:` header. The node id is
   * line-independent so a name declared twice in one document collapses onto
   * one node instead of two half-populated ones.
   */
  private openTable(
    tables: Map<string, PendingTable>,
    rawName: string,
    lineNumber: number,
  ): PendingTable {
    const name = rawName.trim();
    const key = name.toLowerCase();
    const existing = tables.get(key);
    if (existing) return existing;

    const node: Node = {
      id: generateNodeId(this.filePath, 'class', name, 0),
      // `class` is the placeholder kind the SQL table scanners already use, so
      // the ERD declaration and a `SELECT`'s placeholder are the same kind and
      // the resolution pass can promote one onto the other.
      kind: 'class',
      name,
      qualifiedName: name,
      filePath: this.filePath,
      language: 'vba',
      startLine: lineNumber,
      endLine: lineNumber,
      startColumn: 0,
      endColumn: name.length,
      metadata: { synthesizedBy: ERD_TABLE_SYNTHESIZED_BY, linked: false },
      updatedAt: Date.now(),
    };
    this.nodes.push(node);

    const pending: PendingTable = { node, fieldCount: 0 };
    tables.set(key, pending);
    return pending;
  }

  /**
   * A `(LINKED)` marker opens a short blockquote block naming the origin table
   * and the connection string. Read forward over the contiguous `>` lines,
   * flag the table, and — when the connection string names a database FILE —
   * emit the external-backend node plus the `references` edge to it.
   */
  private markLinked(
    table: PendingTable,
    lines: string[],
    markerIndex: number,
    lineNumber: number,
    backends: Set<string>,
  ): void {
    let originTable: string | undefined;
    let backendPath = '';

    for (let j = markerIndex; j < lines.length; j++) {
      const line = lines[j]!;
      if (!line.trim().startsWith('>')) break;
      const origin = ORIGIN_RE.exec(line);
      if (origin) {
        originTable = origin[1]!.trim();
        continue;
      }
      const db = CONNECT_DATABASE_RE.exec(line);
      if (db) backendPath = normalizeBackendPath(db[1]!);
    }

    table.node.metadata = {
      ...(table.node.metadata ?? {}),
      linked: true,
      ...(originTable ? { originTable } : {}),
      ...(backendPath ? { backendPath } : {}),
    };

    // No file operand (an ODBC/DSN link, or a malformed line) → the table is
    // still flagged linked, but no backend node is guessed.
    if (!backendPath) return;

    const backend = buildExternalBackendNode(backendPath);
    if (!backends.has(backend.id)) {
      backends.add(backend.id);
      this.nodes.push(backend);
    }
    this.edges.push({
      source: table.node.id,
      target: backend.id,
      kind: 'references',
      provenance: 'heuristic',
      metadata: {
        synthesizedBy: ERD_LINKED_TABLE_SYNTHESIZED_BY,
        ...(originTable ? { originTable } : {}),
        backendPath,
      },
      line: lineNumber,
      column: 0,
    });
  }

  /**
   * A `| Campo | Tipo | Longitud |` data row becomes one `type_member` node
   * `contains`-ed by its table. The markdown header row and the `:---`
   * separator row are not fields.
   */
  private maybeAddField(
    table: PendingTable,
    line: string,
    lineNumber: number,
    fieldIds: Set<string>,
  ): void {
    const cells = parseRowCells(line);
    if (!cells || cells.length < 3) return;

    const [name, type, length] = cells as [string, string, string];
    if (!name) return;
    if (SEPARATOR_CELL_RE.test(name)) return;
    // The generated header row, in the export's own language.
    if (name.toLowerCase() === 'campo' && type.toLowerCase() === 'tipo') return;

    const qualifiedName = `${table.node.name}.${name}`;
    const id = generateNodeId(this.filePath, 'type_member', qualifiedName, 0);
    if (fieldIds.has(id)) return;
    fieldIds.add(id);

    table.fieldCount += 1;
    this.nodes.push({
      id,
      kind: 'type_member',
      name,
      qualifiedName,
      filePath: this.filePath,
      language: 'vba',
      startLine: lineNumber,
      endLine: lineNumber,
      startColumn: 0,
      endColumn: line.length,
      metadata: {
        accessType: coerceCell(type),
        length: coerceCell(length),
        position: table.fieldCount,
      },
      updatedAt: Date.now(),
    });
    this.edges.push({
      source: table.node.id,
      target: id,
      kind: 'contains',
      provenance: 'parser',
      line: lineNumber,
      column: 0,
    });

    // The table's span grows to cover its last field, so `codegraph_node` on a
    // table shows the whole section rather than just its header line.
    if (lineNumber > table.node.endLine) table.node.endLine = lineNumber;
  }

  private createFileNode(lines: string[]): Node {
    return {
      id: generateNodeId(this.filePath, 'file', this.filePath, 1),
      kind: 'file',
      name: path.basename(this.filePath),
      qualifiedName: this.filePath,
      filePath: this.filePath,
      language: 'vba',
      startLine: 1,
      endLine: lines.length,
      startColumn: 0,
      endColumn: lines[lines.length - 1]?.length ?? 0,
      updatedAt: Date.now(),
    };
  }
}
