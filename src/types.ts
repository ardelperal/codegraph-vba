/**
 * CodeGraph Type Definitions
 *
 * Core types for the semantic knowledge graph system.
 */

// =============================================================================
// Union Types
// =============================================================================

/**
 * Types of nodes in the knowledge graph.
 *
 * Defined as a runtime-iterable `as const` array so the same source
 * of truth backs both the TS type and any runtime validation
 * (e.g. the search query parser).
 */
export const NODE_KINDS = [
  'file',
  'module',
  'class',
  'struct',
  'interface',
  'trait',
  'protocol',
  'function',
  'method',
  'property',
  'field',
  'variable',
  'constant',
  'enum',
  'enum_member',
  'event',
  'type',
  'type_member',
  'declare',
  'type_alias',
  'namespace',
  'parameter',
  'import',
  'export',
  'route',
  'component',
  // 'query' — a saved Access/SQL query (Dysflow `queries/*.sql`). The data
  // layer of an Access app; emits `references` edges to the tables it names.
  'query',
  // --- VBA form-control modeling (added 2026-06-29, Phase B1) ---
  // 'form-layout' — the form-level container node emitted from a
  // `.form.txt` / `.report.txt` (formerly mis-emitted as `module`).
  // 'form-instance-control' — a single Access control instance declared
  // in a `.form.txt`, identified by its `Name = "..."` attribute. Carries
  // `metadata.controlType` for the Access control type (Label, TextBox,
  // CommandButton, etc.) and is the bridge target for `event-handler`
  // edges synthesized from `<Control>_<Event>` handler Subs in the
  // sibling `.cls`.
  // 'report-layout' — Issue #48: the report-level stub node synthesized
  // for `DoCmd.OpenReport "<Name>"`. Mirrors `form-layout` but is keyed
  // to the report naming convention (`Report_<Name>` qualifiedName).
  // Distinct kind so downstream tooling can filter form-vs-report stubs
  // without inspecting qualifiedName.
  'form-layout',
  'form-instance-control',
  'report-layout',
] as const;

export type NodeKind = (typeof NODE_KINDS)[number];

/**
 * Types of edges (relationships) between nodes
 */
export type EdgeKind =
  | 'contains'        // Parent contains child (file→class, class→method)
  | 'calls'           // Function/method calls another
  | 'imports'         // File imports from another
  | 'exports'         // File exports a symbol
  | 'extends'         // Class/interface extends another
  | 'implements'      // Class implements interface
  | 'references'      // Generic reference to another symbol
  | 'type_of'         // Variable/parameter has type
  | 'returns'         // Function returns type
  | 'instantiates'    // Creates instance of class
  | 'overrides'       // Method overrides parent method
  | 'decorates'       // Decorator applied to symbol
  // --- VBA form-control modeling (added 2026-06-29, Phase B1) ---
  // 'event-handler' — synthesized from a `<Control>_<Event>` handler Sub
  // in a `.cls` to the matching `form-instance-control` node in the
  // sibling `.form.txt`. Carries `metadata.eventName` (e.g. 'Click').
  // Provenance: 'heuristic' (a naming convention, not a static parse fact).
  // 'opens-form' — `DoCmd.OpenForm "<FormName>"` modeled as an edge from
  // the calling function to a target form module. Carries
  // `metadata.targetFormName` until the target `.cls`/`.form.txt` is
  // indexed and resolved.
  // 'opens-report' — Issue #48: `DoCmd.OpenReport "<ReportName>"`
  // modeled as an edge from the calling function to a target report
  // module. Carries `metadata.targetReportName`. Symmetric to
  // `opens-form` (different edge kind, different stub kind, different
  // qualifiedName prefix `Report_<Name>` vs `Form_<Name>`).
  | 'event-handler'
  | 'opens-form'
  | 'opens-report'
  | 'raises-event'
  | 'subscribes-event'
  | 'type-member';

/**
 * Supported programming languages. See NODE_KINDS for why this is a
 * runtime-iterable const array.
 */
export const LANGUAGES = [
  'typescript',
  'javascript',
  'tsx',
  'jsx',
  'arkts',
  'python',
  'go',
  'rust',
  'java',
  'c',
  'cpp',
  'csharp',
  'razor',
  'php',
  'ruby',
  'swift',
  'kotlin',
  'dart',
  'svelte',
  'vue',
  'astro',
  'liquid',
  'pascal',
  'scala',
  'lua',
  'luau',
  'objc',
  'r',
  'vba',
  // SQL — Dysflow-exported saved Access queries (`queries/*.sql`). No grammar;
  // the `SqlQueryExtractor` regex-models each query file as a `query` node with
  // `references` edges to the tables it names. Only treated as a query when a
  // sibling `queries.json` manifest is present (see directory discovery).
  'sql',
  'solidity',
  'nix',
  'yaml',
  'twig',
  'xml',
  'properties',
  'cfml',
  'cfscript',
  'cfquery',
  'cobol',
  'vbnet',
  'erlang',
  'terraform',
  'unknown',
] as const;

export type Language = (typeof LANGUAGES)[number];

// =============================================================================
// Core Graph Types
// =============================================================================

/**
 * A node in the knowledge graph representing a code symbol
 */
export interface Node {
  /** Unique identifier (hash of file path + qualified name) */
  id: string;

  /** Type of code element */
  kind: NodeKind;

  /** Simple name (e.g., "calculateTotal") */
  name: string;

  /** Fully qualified name (e.g., "src/utils.ts::MathHelper.calculateTotal") */
  qualifiedName: string;

  /** File path relative to project root */
  filePath: string;

  /** Programming language */
  language: Language;

  /** Starting line number (1-indexed) */
  startLine: number;

  /** Ending line number (1-indexed) */
  endLine: number;

  /** Starting column (0-indexed) */
  startColumn: number;

  /** Ending column (0-indexed) */
  endColumn: number;

  /** Documentation string if present */
  docstring?: string;

  /** Function/method signature */
  signature?: string;

  /** Visibility modifier */
  visibility?: 'public' | 'private' | 'protected' | 'internal';

  /** Whether symbol is exported */
  isExported?: boolean;

  /** Whether symbol is async */
  isAsync?: boolean;

  /** Whether symbol is static */
  isStatic?: boolean;

  /** Whether symbol is abstract */
  isAbstract?: boolean;

  /** Decorators/annotations applied */
  decorators?: string[];

  /** Generic type parameters */
  typeParameters?: string[];

  /**
   * Normalized return/result type name for a function/method (the bare class
   * name, smart-pointer pointee unwrapped). Captured for C/C++ so resolution
   * can infer a chained receiver's type from what the inner call returns —
   * `Foo::instance().bar()` resolves `bar` on `Foo` (issue #645). Undefined for
   * languages/symbols where it isn't captured.
   */
  returnType?: string;

  /** When the node was last updated */
  updatedAt: number;

  /**
   * Optional metadata bag for language-/extractor-specific annotations.
   * Used by VBA to carry `hasClassInitializer` / `initializerName` on class
   * nodes and `controlType` on form property nodes; reserved by other
   * extractors for similar per-language extras. Mirrors the
   * `Edge.metadata` shape (`Record<string, unknown>`).
   */
  metadata?: Record<string, unknown>;
}

/**
 * An edge representing a relationship between two nodes
 */
export interface Edge {
  /**
   * Database row id (AUTOINCREMENT PK on the `edges` table). Only populated
   * when the edge was read back from the DB (e.g. via `getIncomingEdges`);
   * undefined for edges constructed in-memory by an extractor/resolver
   * before insertion. Needed by `resolveVbaCallStubs` to target a specific
   * row for `repointEdgeTarget`/`deleteEdgeById` (vba-graph-connectivity-
   * fixes, #12) since `edges` has no natural unique key.
   */
  id?: number;

  /** Source node ID */
  source: string;

  /** Target node ID */
  target: string;

  /** Type of relationship */
  kind: EdgeKind;

  /** Additional context about the relationship */
  metadata?: Record<string, unknown>;

  /** Line number where relationship occurs (e.g., call site) */
  line?: number;

  /** Column number where relationship occurs */
  column?: number;

  /** How this edge was created */
  provenance?: 'tree-sitter' | 'scip' | 'parser' | 'heuristic';
}

/**
 * Metadata about a tracked file
 */
export interface FileRecord {
  /** File path relative to project root */
  path: string;

  /** Content hash for change detection */
  contentHash: string;

  /** Detected language */
  language: Language;

  /** File size in bytes */
  size: number;

  /** Last modification timestamp */
  modifiedAt: number;

  /** When last indexed */
  indexedAt: number;

  /** Number of nodes extracted */
  nodeCount: number;

  /** Any extraction errors */
  errors?: ExtractionError[];
}

// =============================================================================
// Extraction Types
// =============================================================================

/**
 * Result from parsing a source file
 */
export interface ExtractionResult {
  /** Extracted nodes */
  nodes: Node[];

  /** Extracted edges */
  edges: Edge[];

  /** References that couldn't be resolved yet */
  unresolvedReferences: UnresolvedReference[];

  /** Any errors during extraction */
  errors: ExtractionError[];

  /** Extraction duration in milliseconds */
  durationMs: number;
}

/**
 * Error during code extraction
 */
export interface ExtractionError {
  /** Error message */
  message: string;

  /** File path where the error occurred */
  filePath?: string;

  /** Line number if available */
  line?: number;

  /** Column number if available */
  column?: number;

  /** Error severity */
  severity: 'error' | 'warning';

  /** Error code for categorization */
  code?: string;
}

/**
 * Kinds an unresolved reference can carry. `function_ref` is internal-only —
 * a function name used as a VALUE (callback registration, #756). It never
 * becomes an edge kind: resolution maps it to a `references` edge targeting
 * function/method nodes only (see `matchFunctionRef`).
 *
 * Round-3 (issue #108) extends the union with shape-based classifiers so
 * the consumer can filter `unresolved_refs` by syntactic shape at the SQL
 * layer. The legacy `EdgeKind` literal `'references'` is preserved as
 * `references` here via `EdgeKind` — back-compat for any push site this
 * round does not reclassify (e.g. the Implements emitter). New literals:
 *   - `calls`            paren-form `Name(...)` or statement-form Sub call
 *   - `qualified-call`   `Receiver.Member(...)` (qualified-paren) or
 *                        `Receiver.Member args` (qualified-statement)
 *   - `property-get`     `Me.Name` (dot access, read)
 *   - `property-set`     `Me.Name = value` (dot access, assignment)
 *   - `bang-get`         `Me!SubCtl` (bang read) or
 *                        `Forms!FormX!Ctl` / `Forms("FormX")!Ctl` (cross-form)
 *   - `bang-set`         `Me!SubCtl = value` (bang assignment)
 *   - `unqualified-ident` bare identifier without `(` after; default for
 *                        `HayErrorEnRiesgo`-style hits — Const-read takes
 *                        priority via FR-3.1 disambiguation
 *   - `member-with`      `.Member` inside a `With <receiver>` block
 *   - `dao-query`        `DoCmd.OpenQuery "X"` argument
 * `dao-field-get` / `dao-field-set` are deliberately deferred to round-4.
 */
/** `calls` is the canonical kind for unresolved procedure and function calls. */
export type ReferenceKind = EdgeKind | 'function_ref' | 'qualified-call' | 'property-get' | 'property-set' | 'bang-get' | 'bang-set' | 'unqualified-ident' | 'member-with' | 'dao-query';

/**
 * Runtime guard for whether a `ReferenceKind` literal is also an `EdgeKind`
 * literal (i.e. the value flows from an unresolved-ref row straight onto the
 * kind column of a resolved edge without remapping). Round-3 (issue #108)
 * introduced shape-based classifier literals (`qualified-call`,
 * `property-get`, …) that are valid as a `ReferenceKind` but are NOT valid
 * edge kinds — edges still use the `calls`/`references` family of literals.
 * The resolver uses this guard to fall back to `'references'` for the edge
 * kind while keeping the shape preserved on the unresolved-ref row, so the
 * consumer's SQL filter still works after resolution.
 */
export function isEdgeKindLiteral(value: ReferenceKind): value is EdgeKind {
  switch (value) {
    case 'contains':
    case 'calls':
    case 'imports':
    case 'exports':
    case 'extends':
    case 'implements':
    case 'references':
    case 'type_of':
    case 'returns':
    case 'instantiates':
    case 'overrides':
    case 'decorates':
    case 'event-handler':
    case 'opens-form':
    case 'opens-report':
    case 'raises-event':
    case 'subscribes-event':
    case 'type-member':
      return true;
    default:
      return false;
  }
}

/**
 * A reference that couldn't be resolved during extraction
 */
export interface UnresolvedReference {
  /** ID of the node containing the reference */
  fromNodeId: string;

  /** Name being referenced */
  referenceName: string;

  /** Type of reference (call, type, import, etc.) */
  referenceKind: ReferenceKind;

  /** Location of the reference */
  line: number;
  column: number;

  /** File path where reference occurs (denormalized for performance) */
  filePath?: string;

  /** Language of the source file (denormalized for performance) */
  language?: Language;

  /** Possible qualified names it might resolve to */
  candidates?: string[];

  /**
   * Optional metadata bag for language-/extractor-specific annotations on
   * unresolved references. Used by VBA to carry
   * `synthesizedBy: 'vba-form-binding'` on the form → sibling-`.cls`
   * reference so downstream resolvers can render the provenance inline.
   * Mirrors `Edge.metadata` and `Node.metadata` shapes
   * (`Record<string, unknown>`).
   */
  metadata?: Record<string, unknown>;
}

// =============================================================================
// Query Types
// =============================================================================

/**
 * A subgraph containing a subset of the knowledge graph
 */
export interface Subgraph {
  /** Nodes in this subgraph */
  nodes: Map<string, Node>;

  /** Edges in this subgraph */
  edges: Edge[];

  /** Root node IDs (entry points) */
  roots: string[];

  /**
   * Retrieval confidence for context-style queries. `'low'` means the query
   * resolved only to isolated common-word matches (no entry point corroborated
   * by 2+ distinct query terms) — callers should surface an honest handoff to
   * explore/trace rather than present the results as comprehensive. Undefined
   * for graph traversals that don't run the search-ranking path.
   */
  confidence?: 'high' | 'low';
}

/**
 * Options for graph traversal
 */
export interface TraversalOptions {
  /** Maximum depth to traverse (default: Infinity) */
  maxDepth?: number;

  /** Edge types to follow (default: all) */
  edgeKinds?: EdgeKind[];

  /** Node types to include (default: all) */
  nodeKinds?: NodeKind[];

  /** Direction of traversal */
  direction?: 'outgoing' | 'incoming' | 'both';

  /** Maximum nodes to return */
  limit?: number;

  /** Whether to include the starting node */
  includeStart?: boolean;
}

/**
 * Options for searching the graph
 */
export interface SearchOptions {
  /** Node types to search */
  kinds?: NodeKind[];

  /** Languages to include */
  languages?: Language[];

  /** File path patterns to include */
  includePatterns?: string[];

  /** File path patterns to exclude */
  excludePatterns?: string[];

  /** Maximum results to return */
  limit?: number;

  /** Offset for pagination */
  offset?: number;

  /** Whether search is case-sensitive */
  caseSensitive?: boolean;
}

/**
 * A search result with relevance scoring
 */
export interface SearchResult {
  /** Matching node */
  node: Node;

  /**
   * Relevance score for relative ranking only — higher is more relevant.
   * NOT normalized and NOT a 0-1 fraction: the FTS path returns an unbounded
   * BM25 magnitude (often in the tens or hundreds), while the fuzzy/exact
   * paths return ~0-1. Use it to order results, not as an absolute percentage.
   */
  score: number;

  /** Matched text snippets for highlighting */
  highlights?: string[];
}

/**
 * A symbol whose name-segments match prose words from a prompt — the
 * graph-derived signal behind the front-load hook's medium tier
 * (CodeGraph.getSegmentMatches). Always verified to exist in `nodes` at the
 * time it is returned.
 */
export interface SegmentMatch {
  /** Symbol name as indexed (e.g. `OrderStateMachine`). */
  name: string;
  /** Kind of the representative definition. */
  kind: NodeKind;
  /** File of the representative definition. */
  filePath: string;
  /** 1-based start line of the representative definition. */
  startLine: number;
  /** The prompt words (normalized) that matched this name's segments. */
  matchedWords: string[];
}

// =============================================================================
// Context Types
// =============================================================================

/**
 * Context information for code understanding
 */
export interface Context {
  /** Primary node being examined */
  focal: Node;

  /** Nodes containing the focal node (file, class, etc.) */
  ancestors: Node[];

  /** Nodes directly contained by focal node */
  children: Node[];

  /** Incoming references (who calls/uses this) */
  incomingRefs: Array<{ node: Node; edge: Edge }>;

  /** Outgoing references (what this calls/uses) */
  outgoingRefs: Array<{ node: Node; edge: Edge }>;

  /** Related type information */
  types: Node[];

  /** Relevant imports */
  imports: Node[];
}

/**
 * A block of code with context
 */
export interface CodeBlock {
  /** The code content */
  content: string;

  /** File path */
  filePath: string;

  /** Starting line */
  startLine: number;

  /** Ending line */
  endLine: number;

  /** Language for syntax highlighting */
  language: Language;

  /** Associated node if extracted */
  node?: Node;
}

// =============================================================================
// Database Types
// =============================================================================

/**
 * Database schema version info
 */
export interface SchemaVersion {
  /** Current schema version */
  version: number;

  /** When schema was created/updated */
  appliedAt: number;

  /** Description of this version */
  description?: string;
}

/**
 * Statistics about the knowledge graph
 */
export interface GraphStats {
  /** Total number of nodes */
  nodeCount: number;

  /** Total number of edges */
  edgeCount: number;

  /** Number of tracked files */
  fileCount: number;

  /** Node counts by kind */
  nodesByKind: Record<NodeKind, number>;

  /** Edge counts by kind */
  edgesByKind: Record<EdgeKind, number>;

  /** File counts by language */
  filesByLanguage: Record<Language, number>;

  /** Database size in bytes */
  dbSizeBytes: number;

  /** Last update timestamp */
  lastUpdated: number;
}

// =============================================================================
// Task Context Types (for buildContext)
// =============================================================================

/**
 * Input for building task context
 */
export type TaskInput = string | { title: string; description?: string };

/**
 * Options for building task context
 */
export interface BuildContextOptions {
  /** Maximum number of nodes to include (default: 50) */
  maxNodes?: number;

  /** Maximum number of code blocks to include (default: 10) */
  maxCodeBlocks?: number;

  /** Maximum characters per code block (default: 2000) */
  maxCodeBlockSize?: number;

  /** Whether to include code blocks (default: true) */
  includeCode?: boolean;

  /** Output format (default: 'markdown') */
  format?: 'markdown' | 'json';

  /** Number of semantic search results (default: 5) */
  searchLimit?: number;

  /** Graph traversal depth from entry points (default: 2) */
  traversalDepth?: number;

  /** Minimum semantic similarity score (default: 0.3) */
  minScore?: number;
}

/**
 * Full context for a task, ready for Claude
 */
export interface TaskContext {
  /** The original query/task */
  query: string;

  /** Subgraph of relevant nodes and edges */
  subgraph: Subgraph;

  /** Entry point nodes (from semantic search) */
  entryPoints: Node[];

  /** Code blocks extracted from key nodes */
  codeBlocks: CodeBlock[];

  /** Files involved in this context */
  relatedFiles: string[];

  /** Brief summary of the context */
  summary: string;

  /** Statistics about the context */
  stats: {
    /** Number of nodes included */
    nodeCount: number;
    /** Number of edges included */
    edgeCount: number;
    /** Number of files touched */
    fileCount: number;
    /** Number of code blocks included */
    codeBlockCount: number;
    /** Total characters in code blocks */
    totalCodeSize: number;
  };
}

/**
 * Options for finding relevant context
 */
export interface FindRelevantContextOptions {
  /** Number of semantic search results (default: 5) */
  searchLimit?: number;

  /** Graph traversal depth (default: 2) */
  traversalDepth?: number;

  /** Maximum nodes in result (default: 50) */
  maxNodes?: number;

  /** Minimum semantic similarity score (default: 0.3) */
  minScore?: number;

  /** Edge types to follow in traversal */
  edgeKinds?: EdgeKind[];

  /** Node types to include */
  nodeKinds?: NodeKind[];
}
