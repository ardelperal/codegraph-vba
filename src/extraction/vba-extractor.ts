/**
 * VbaExtractor — regex extractor for `.bas` / `.cls` / `.frm` / `.dsr` source.
 *
 * Emits the symbol + edge shape expected by `codegraph_explore`:
 *  - one `module` node per `.bas` / one `class` node per `.cls` / one `file`
 *    stub for `.frm` / `.dsr` legacy
 *  - one `function` node per `Sub` / `Function` / `Property Get|Let|Set`
 *    declaration, with `node.visibility` set to the canonical lowercase enum
 *    (`'public'` / `'private'`; `Friend`, `Static`, and missing visibility fold
 *    to `'public'`).
 *  - `contains` edges from the module/class node to each function node.
 *  - `calls` edges from a procedure to a same-file procedure (`Sub Outer()` →
 *    `Inner()`); qualified cross-module calls carry `provenance: 'heuristic'`
 *    and `metadata.synthesizedBy: 'vba-name-resolution'`.
 *  - `implements` edges for `Implements IFoo` declarations.
 *  - `references` edges for `Dim x As Foo.Bar` (qualified Dim → outer type)
 *    tagged `synthesizedBy: 'vba-name-resolution'`.
 *  - `references` edges for `WithEvents m_X As Form_Foo` tagged
 *    `synthesizedBy: 'vba-withevents'`.
 *  - `references` edges for SQL table names found inside string literals passed
 *    directly to SQL wrappers, or assigned to variables before
 *    `getdb().OpenRecordset(...)` / `getdb().Execute ...`, tagged
 *    `synthesizedBy: 'vba-sql-table'`.
 *
 * Rejects `.form.txt` / `.report.txt` (REQ-CODE-9 — the form UI extractor
 * owns those). `Option ...` directives are inert — `stripVbaComments()`
 * drops them.
 */
import * as path from 'path';
import {
  Node,
  Edge,
  ExtractionResult,
  ExtractionError,
  UnresolvedReference,
} from '../types';
import { generateNodeId } from './tree-sitter-helpers';
import {
  joinLineContinuations,
  stripVbaComments,
  extractStringLiterals,
} from './vba-preprocess';

interface ProcInfo {
  name: string;
  qualifiedName: string;
  kind: 'sub' | 'function' | 'property';
  visibility: 'public' | 'private' | 'protected' | 'internal';
  startLine: number;
}

export class VbaExtractor {
  private filePath: string;
  private source: string;
  private nodes: Node[] = [];
  private edges: Edge[] = [];
  private errors: ExtractionError[] = [];
  private unresolvedReferences: UnresolvedReference[] = [];
  private moduleOrClassNode: Node | null = null;
  /**
   * Map of procedure name (within this file) → list of ProcInfo. The list
   * (rather than a single value) is needed because VBA allows multiple
   * Property accessors with the same name: `Property Get Foo`, `Property
   * Let Foo`, and `Property Set Foo` all share the name `Foo`. Audit W6
   * (June 2026): the previous Map<string, ProcInfo> kept only the last
   * accessor, breaking same-file call resolution and emitting the wrong
   * `findFunctionNodeByName` match.
   */
  private localProcs = new Map<string, ProcInfo[]>();

  /**
   * Cache: procedure name → first matching function node emitted for that
   * name. Audit S2 (June 2026): the previous implementation did an O(n)
   * linear scan of all nodes for every call site; with W6's multimap
   * the same-name collision can return multiple matches but bare-name
   * call resolution only needs the first.
   */
  private functionNodeByName = new Map<string, Node>();

  /**
   * Cache: startLine → function node emitted for that line. Used by Fix 1
   * (June 2026): when Property Get/Let/Set share a name, we need to find
   * the SPECIFIC accessor's node by its declaration line rather than relying
   * on `functionNodeByName` which only stores the first accessor.
   */
  private functionNodeByStartLine = new Map<number, Node>();

  constructor(filePath: string, source: string) {
    this.filePath = filePath;
    this.source = source;
  }

  extract(): ExtractionResult {
    const startTime = Date.now();

    try {
      // REQ-CODE-9: hand form/report files off to VbaFormExtractor. If
      // somehow routed here, emit just a file node and return.
      if (this.isFormOrReportFile()) {
        this.nodes.push(this.createFileNode());
        return this.result(startTime);
      }

      // Pre-process pipeline (per design): join continuations, strip comments,
      // then we sweep the joined-but-uncommented source for call sites and
      // SQL strings. Comments are gone before regex runs, so no commented
      // SQL can match the SQL regex.
      const joined = joinLineContinuations(this.source);
      const uncommented = stripVbaComments(joined);

      // Create the file node.
      this.nodes.push(this.createFileNode());

      // Resolve the file's "kind" — .cls → class, else module (including .bas
      // and the legacy .frm/.dsr stubs).
      const isCls = this.filePath.toLowerCase().endsWith('.cls');

      // Detect VB_Name attribute (first non-empty line).
      const vbName = this.detectVbName(this.source);

      // Module/class node — created lazily so a file containing only
      // Option directives (REQ-CODE-10: "emits nothing") doesn't carry a
      // module node with no symbols attached.
      let hasAnySymbols = false;

      // Sweep: procedures → function nodes + contains edges.
      const procs = this.sweepProcedures(uncommented);
      if (procs.length > 0) hasAnySymbols = true;

      // Sweep: Implements (REQ-CODE-5).
      const implCount = this.sweepImplements(uncommented);
      if (implCount > 0) hasAnySymbols = true;

      // Sweep: qualified Dim (REQ-CODE-6) and WithEvents (REQ-CODE-7).
      const dimCount = this.sweepDimsAndWithEvents(uncommented);
      if (dimCount > 0) hasAnySymbols = true;

      // Sweep: call sites (REQ-CODE-4) and SQL-in-strings (REQ-CODE-8).
      // Both walk the same per-line view; combining them in one pass is
      // simpler and keeps line tracking consistent.
      this.sweepCallsAndSql(uncommented);

      // Create the module/class node ONLY when the file has actual symbols
      // (REQ-CODE-10 — a file with only Option directives emits zero symbol
      // nodes).
      if (hasAnySymbols) {
        this.moduleOrClassNode = this.createModuleOrClassNode(isCls, vbName);
        this.nodes.push(this.moduleOrClassNode);

        // Re-attribute contains/implements/reference edges whose source was
        // held in `pendingModuleOrClassSource` — see below. Since we now
        // know the module id, redirect those edges to it.
        for (const edge of this.pendingModuleOrClassSource) {
          edge.source = this.moduleOrClassNode.id;
        }
        this.pendingModuleOrClassSource.length = 0;

        // Sub New marker on the class node (REQ-CODE-3).
        if (isCls) {
          const hasNew = procs.some((p) => p.name === 'New');
          if (hasNew) {
            const md = (this.moduleOrClassNode.metadata ?? {}) as Record<string, unknown>;
            md.hasClassInitializer = true;
            md.initializerName = 'New';
            this.moduleOrClassNode.metadata = md;
          }
        }
      }
    } catch (error) {
      this.errors.push({
        message: `VBA extraction error: ${error instanceof Error ? error.message : String(error)}`,
        filePath: this.filePath,
        severity: 'error',
        code: 'parse_error',
      });
    }

    return this.result(startTime);
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  private result(startTime: number): ExtractionResult {
    return {
      nodes: this.nodes,
      edges: this.edges,
      unresolvedReferences: this.unresolvedReferences,
      errors: this.errors,
      durationMs: Date.now() - startTime,
    };
  }

  private isFormOrReportFile(): boolean {
    const lower = this.filePath.toLowerCase();
    return lower.endsWith('.form.txt') || lower.endsWith('.report.txt');
  }

  private createFileNode(): Node {
    const lines = this.source.split('\n');
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

  private detectVbName(src: string): string | null {
    // Walk past the Access class metadata header block (VERSION / BEGIN /
    // MultiUse / END / Attribute …) so VB_Name on a later line is found.
    // The previous implementation returned null at the first non-Attribute
    // line, which meant real Access .cls files — which start with
    // `VERSION 1.0 CLASS / BEGIN / MultiUse = ... / END` before
    // `Attribute VB_Name = "..."` — always fell through to the basename
    // fallback. Audit W2 (June 2026).
    for (const line of src.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      const m = /^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/i.exec(trimmed);
      if (m) return m[1] ?? null;
      // Skip known header keywords so we keep scanning.
      if (
        /^\s*VERSION\b/i.test(trimmed) ||
        /^\s*BEGIN\b/i.test(trimmed) ||
        /^\s*END\b/i.test(trimmed) ||
        /^\s*(?:MultiUse|Persistable|DataBindingBehavior|DataSourceBehavior)\s*=/i.test(trimmed) ||
        /^\s*Attribute\s+/i.test(trimmed)
      ) {
        continue;
      }
      // Any other non-empty line that isn't a known header — stop.
      return null;
    }
    return null;
  }

  private createModuleOrClassNode(isCls: boolean, vbName: string | null): Node {
    const fallbackName = path.basename(this.filePath).replace(/\.[^.]+$/, '');
    const name = vbName ?? fallbackName;
    return {
      id: generateNodeId(this.filePath, isCls ? 'class' : 'module', name, 1),
      kind: isCls ? 'class' : 'module',
      name,
      qualifiedName: name,
      filePath: this.filePath,
      language: 'vba',
      startLine: 1,
      endLine: this.source.split('\n').length,
      startColumn: 0,
      endColumn: 0,
      metadata: {},
      updatedAt: Date.now(),
    };
  }

  /** Sub/Function/Property regex — captures visibility prefix, kind, and name. */
  private static readonly PROC_RE =
    /^\s*((?:Public|Private|Friend|Static)\s+)?(?:Static\s+)?(Sub|Function|Property(?:\s+(?:Get|Let|Set))?)\s+(\p{L}[\p{L}\p{N}_]*)/iu;

  /**
   * Walk the (uncommented, line-joined) source and emit one `function` node
   * per `Sub` / `Function` / `Property` declaration. Also records the proc
   * in `localProcs` so call-site resolution can distinguish same-file from
   * cross-module calls.
   */
  private sweepProcedures(src: string): ProcInfo[] {
    const procs: ProcInfo[] = [];
    const lines = src.split('\n');
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const m = VbaExtractor.PROC_RE.exec(line);
      if (!m) continue;
      const visibilityRaw = (m[1] ?? '').trim();
      const kindRaw = (m[2] ?? '').trim().toLowerCase();
      const name = m[3] ?? '';
      if (!name) continue;
      const lineNum = i + 1;

      // Normalize the visibility. VBA's "Static" keyword is a storage
      // specifier, not visibility, so we treat bare-static declarations as
      // 'Public' (the default) per spec R1. "Friend" is not in the Node
      // visibility enum; treat it as 'Public' since it's the closest
      // broader-than-private modifier.
      let visibility: ProcInfo['visibility'];
      switch (visibilityRaw.toLowerCase()) {
        case 'private':
          visibility = 'private';
          break;
        case 'public':
        case 'static':
        case 'friend':
        case '':
        default:
          visibility = 'public';
          break;
      }

      const kind: ProcInfo['kind'] = kindRaw.startsWith('sub')
        ? 'sub'
        : kindRaw.startsWith('function')
          ? 'function'
          : 'property';

      const proc: ProcInfo = {
        name,
        qualifiedName: name,
        kind,
        visibility,
        startLine: lineNum,
      };
      procs.push(proc);
      const bucket = this.localProcs.get(name);
      if (bucket) bucket.push(proc);
      else this.localProcs.set(name, [proc]);

      const nodeId = generateNodeId(this.filePath, 'function', name, lineNum);
      const fnNode: Node = {
        id: nodeId,
        kind: 'function',
        name,
        qualifiedName: name,
        filePath: this.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: 0,
        endColumn: line.length,
        visibility,
        updatedAt: Date.now(),
      };
      this.nodes.push(fnNode);
      // Cache the first node emitted for this name — `findFunctionNodeByName`
      // (audit S2) becomes O(1) instead of O(n) per call site.
      if (!this.functionNodeByName.has(name)) {
        this.functionNodeByName.set(name, fnNode);
      }
      // Fix 1: also index by startLine so Property Get/Let/Set with the same
      // name can each be found by their exact declaration line.
      this.functionNodeByStartLine.set(lineNum, fnNode);

      // Hueco 3: synthesize the event-handler edge for Access naming
      // convention `<ControlName>_<EventName>` (e.g. `ComandoAltaPM_Click`,
      // `MotivoBorrado_AfterUpdate`). The `<X>_<Y>` shape is split on the
      // LAST underscore so multi-word events like `BeforeDelConfirm` parse
      // correctly. Form-level events (`Form_Load`, `Form_Open`,
      // `Form_Unload`, …) are NOT control handlers — they fire on the
      // form object itself, not on a control — so they are skipped here
      // (the form module node is the conceptual source for those).
      //
      // Scope guard: only emit when this .cls looks like a form code-
      // behind (`Form_*.cls`). Without this guard, regular service classes
      // whose methods happen to have underscores in their names
      // (e.g. `InformeRiesgoPDFServicio.cls` declares
      // `GenerarHTML_Principal`, `GetEstilosCSS_PDF`,
      // `Class_Initialize`) would synthesize ~550 spurious
      // `form-instance-control` stubs in real Dysflow projects. The
      // `Form_` prefix is the canonical Access code-behind naming
      // convention and matches the .form.txt siblings' basename.
      //
      // Cross-file synthesis caveat: the edge's source is a
      // `form-instance-control` node that lives in the sibling .form.txt,
      // not this .cls file. Two consequences:
      //   1. We also emit a STUB form-instance-control node locally with
      //      the deterministic id so the per-file edge filter
      //      (`insertedIds.has(source)`) accepts the edge. When the
      //      sibling .form.txt is later indexed, VbaFormExtractor emits
      //      the real form-instance-control node with the same id; the
      //      `INSERT OR REPLACE` semantics in queries.ts:insertNode
      //      overwrite the stub with the real one (preserving the
      //      metadata.controlType, filePath, line range, etc.).
      //   2. When the .form.txt is processed FIRST (alphabetically
      //      unlikely but possible), the real node exists in the DB
      //      before this edge is committed; insertEdges's DB-level
      //      endpoint check passes the edge naturally without the stub.
      //      Either order converges on the same final state.
      // See vba-form-extractor.ts:findControlName for the matching real
      // form-instance-control node emission.
      const handler = parseEventHandlerName(name);
      const isFormCodeBehind = /Form_[^/\\]*\.cls$/i.test(this.filePath);
      if (handler && isFormCodeBehind) {
        const formFilePath = this.filePath.replace(/\.cls$/i, '.form.txt');
        const controlNodeId = generateNodeId(
          formFilePath,
          'form-instance-control',
          handler.controlName,
          0,
        );
        // Stub form-instance-control: local so the per-file edge filter
        // passes the event-handler edge. Overwritten by the real node
        // emitted from the sibling .form.txt at index time (same id, same
        // schema, INSERT OR REPLACE). No metadata.controlType here — the
        // .form.txt side carries the real control type.
        this.nodes.push({
          id: controlNodeId,
          kind: 'form-instance-control',
          name: handler.controlName,
          qualifiedName: `${formFilePath}::${handler.controlName}`,
          filePath: formFilePath,
          language: 'vba',
          startLine: 0,
          endLine: 0,
          startColumn: 0,
          endColumn: 0,
          updatedAt: Date.now(),
        });
        this.edges.push({
          source: controlNodeId,
          target: nodeId,
          kind: 'event-handler',
          provenance: 'heuristic',
          metadata: { eventName: handler.eventName },
          line: lineNum,
          column: 0,
        });
      }

      if (this.moduleOrClassNode) {
        this.edges.push({
          source: this.moduleOrClassNode.id,
          target: nodeId,
          kind: 'contains',
        });
      } else {
        // Module/class node is created lazily (see extract). Hold the edge in
        // pending so its source can be rewritten once the module exists.
        const edge: Edge = {
          source: '',
          target: nodeId,
          kind: 'contains',
        };
        this.edges.push(edge);
        this.pendingModuleOrClassSource.push(edge);
      }
    }
    return procs;
  }

  /** Implements regex. */
  private static readonly IMPLEMENTS_RE =
    /^\s*Implements\s+(\p{L}[\p{L}\p{N}_]*)/iu;

  /** Edges whose source needs to be set to the module/class id once it exists. */
  private pendingModuleOrClassSource: Edge[] = [];

  private sweepImplements(src: string): number {
    const lines = src.split('\n');
    let count = 0;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const m = VbaExtractor.IMPLEMENTS_RE.exec(line);
      if (!m) continue;
      const name = m[1] ?? '';
      if (!name) continue;
      const lineNum = i + 1;
      const targetId = generateNodeId(
        this.filePath,
        'interface',
        name,
        lineNum,
      );
      this.nodes.push({
        id: targetId,
        kind: 'interface',
        name,
        qualifiedName: name,
        filePath: this.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: 0,
        endColumn: line.length,
        updatedAt: Date.now(),
      });
      const edge: Edge = {
        source: '', // placeholder; rewritten after module node exists
        target: targetId,
        kind: 'implements',
        // S4 fix: `Implements IFoo` is a static, source-declared fact —
        // not a guess. Use the `parser` provenance (generalizes
        // `tree-sitter` for non-tree-sitter extractors like our regex
        // sweepers) instead of `heuristic`, which is reserved for
        // guessed/inferred edges.
        provenance: 'parser',
        line: lineNum,
        column: 0,
      };
      this.edges.push(edge);
      this.pendingModuleOrClassSource.push(edge);
      count++;
    }
    return count;
  }

  /**
   * Check that a line is a variable declaration and NOT a Sub/Function/
   * Property/Const/WithEvents header (those have their own sweeps).
   * Fix 1+3 (Issues #1,#3): replaced the old single-match DIM_QUAL_RE /
   * DIM_UNQUAL_RE pair with a prefix-check + global scan that handles
   * `As New <Type>`, multi-variable `Dim a As Foo, b As Bar`, and all
   * visibility keywords in one pass.
   */
  private static readonly DIM_DECL_PREFIX_RE =
    /^\s*(?:Dim|Private|Public)\s+(?!(?:Function|Sub|Property|Const|WithEvents)\b)/i;

  /**
   * Globally scan all `identifier As [New] TypePart1[.TypePart2]` on a
   * variable declaration line. Run with /g after confirming DIM_DECL_PREFIX_RE.
   *
   * Groups: (1) variable name, (2) type outer part, (3) type inner part (if qualified).
   * `(?:New\s+)?` consumes the VBA auto-instantiation keyword so it is
   * never captured as the type name (Fix 1).
   */
  private static readonly DIM_ALL_VARS_RE =
    /\b(\p{L}[\p{L}\p{N}_]*)\s+As\s+(?:New\s+)?(\p{L}[\p{L}\p{N}_]*)(?:\.(\p{L}[\p{L}\p{N}_]*))?/giu;

  /**
   * VBA primitive type names — skipped when emitted as Dim targets so
   * we don't pollute the graph with `As Long` / `As String` references.
   * Fix 4: all entries are LOWERCASE and the lookup lowercases the
   * captured type name so `As long`, `As LONG`, `As Long` all match.
   * Fix 1 (Issue #1): added `'new'` as a backstop so that if the `As New`
   * pattern is ever captured as a type name it is silently skipped.
   */
  private static readonly PRIMITIVE_TYPES = new Set([
    'long', 'integer', 'short', 'byte', 'single', 'double', 'currency',
    'string', 'boolean', 'date', 'variant', 'object', 'error',
    'empty', 'null', 'longptr', 'longlong', 'new',
  ]);

  /** `WithEvents m_X As Form_Foo` — Dim/Private/Public prefix is optional. */
  private static readonly WITHEVENTS_RE =
    /^\s*(?:(?:Dim|Private|Public)\s+)?WithEvents\s+\p{L}[\p{L}\p{N}_]*\s+As\s+(\p{L}[\p{L}\p{N}_]*)/iu;

  private sweepDimsAndWithEvents(src: string): number {
    const lines = src.split('\n');
    let count = 0;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const lineNum = i + 1;

      // Fix 1 + Fix 3 (Issues #1, #3): replace the old single-match
      // DIM_QUAL_RE / DIM_UNQUAL_RE pair with a global scan that handles
      // `As New <Type>`, multi-variable `Dim a As Foo, b As Bar`, and
      // qualified `Dim x As Foo.Bar` in a single pass.
      if (VbaExtractor.DIM_DECL_PREFIX_RE.test(line)) {
        VbaExtractor.DIM_ALL_VARS_RE.lastIndex = 0;
        let m: RegExpExecArray | null;
        while ((m = VbaExtractor.DIM_ALL_VARS_RE.exec(line)) !== null) {
          const varName = m[1] ?? '';
          const outerType = m[2] ?? '';
          const innerType = m[3] ?? '';

          // Fix 2 (Issue #2): populate the local var type map so that
          // `sweepCallsAndSql` can gate qualified statement-form calls.
          if (varName && outerType) {
            this.localVarTypeMap.set(varName.toLowerCase(), {
              outer: outerType,
              qualified: !!innerType,
            });
          }

          if (innerType) {
            // Qualified form (`Dim x As Foo.Bar`) — emit reference to the
            // outer type `Foo` (same behaviour as the old DIM_QUAL_RE path).
            if (outerType) {
              this.emitReference(outerType, lineNum, 0, 'vba-name-resolution');
              count++;
            }
          } else {
            // Unqualified form (`Dim x As SomeType`, `Dim x As New SomeType`)
            // — emit reference only when the type is not a primitive or keyword.
            if (outerType && !VbaExtractor.PRIMITIVE_TYPES.has(outerType.toLowerCase())) {
              this.emitReference(outerType, lineNum, 0, 'vba-name-resolution');
              count++;
            }
          }
        }
      }

      // WithEvents declarations — handled by their own regex; also populate
      // the local var type map for completeness.
      const weMatch = VbaExtractor.WITHEVENTS_RE.exec(line);
      if (weMatch) {
        const formType = weMatch[1] ?? '';
        if (formType) {
          // Extract the variable name from the WithEvents line for the map.
          const weVarM = /^\s*(?:(?:Dim|Private|Public)\s+)?WithEvents\s+(\p{L}[\p{L}\p{N}_]*)/iu.exec(line);
          const weVarName = weVarM?.[1] ?? '';
          if (weVarName) {
            this.localVarTypeMap.set(weVarName.toLowerCase(), {
              outer: formType,
              qualified: false,
            });
          }
          this.emitReference(formType, lineNum, 0, 'vba-withevents');
          count++;
        }
      }
    }
    return count;
  }

  /**
   * Walk the (uncommented) source per line. For each line:
   *  - If it starts a `Sub`/`Function`/`Property`, push the proc onto a stack.
   *  - If it ends one (`End Sub` / `End Function` / `End Property`), pop.
   *  - While inside a procedure, scan the line for call-site patterns and
   *    SQL-wrapper patterns.
   *
   * Call-site regex: `(?<!\w)([A-Za-z_]\w*)(?:\.([A-Za-z_]\w*))?\s*\(` —
   * captures either `Name(...)` (same-file candidate) or `Receiver.Member(...)`
   * (qualified — emit a synthetic node + heuristic edge).
   */
  private static readonly CALL_RE =
    /(?<![\w.])(\p{L}[\p{L}\p{N}_]*)(?:\.(\p{L}[\p{L}\p{N}_]*))?\s*\(/gu;

  /** SQL wrapper helpers — order matters because `db.Execute` is a suffix of others. */
  private static readonly SQL_WRAPPERS: ReadonlyArray<{ name: string; re: RegExp }> = [
    { name: 'DoCmd.RunSQL', re: /\bDoCmd\.RunSQL\s+"((?:[^"]|"")*)"/g },
    { name: 'CurrentDb.OpenRecordset', re: /\bCurrentDb\.OpenRecordset\s+"((?:[^"]|"")*)"/g },
    { name: 'CurrentDb.Execute', re: /\bCurrentDb\.Execute\s+"((?:[^"]|"")*)"/g },
    { name: 'db.Execute', re: /\bdb\.Execute\s+"((?:[^"]|"")*)"/g },
    // Fix 4 (Issue #4): inline-literal forms `getdb().Execute "..."` and
    // `getdb().OpenRecordset "..."` — the variable form is covered by
    // SQL_VAR_EXEC_RE but the direct-literal form was missing.
    { name: 'getdb().Execute', re: /\bgetdb\(\)\.Execute\s+"((?:[^"]|"")*)"/g },
    { name: 'getdb().OpenRecordset', re: /\bgetdb\(\)\.OpenRecordset\s+"((?:[^"]|"")*)"/g },
  ];

  /** SQL assigned to a local variable, e.g. `m_SQL = "SELECT ..." & ...`. */
  private static readonly SQL_VAR_ASSIGN_RE =
    /^\s*(\p{L}[\p{L}\p{N}_]*)\s*=\s*(.*)$/iu;

  /** SQL wrapper called with a variable, e.g. `getdb().Execute m_SQL`. */
  private static readonly SQL_VAR_EXEC_RE =
    /\b(?:getdb\(\)|CurrentDb|db)\.(?:OpenRecordset|Execute)\s*\(?\s*(\p{L}[\p{L}\p{N}_]*)\s*\)?/giu;

  /** SQL table-name regex scoped to FROM / INTO / UPDATE. */
  private static readonly SQL_TABLE_RE =
    /\b(?:FROM|INTO|UPDATE)\s+(\[?\p{L}[\p{L}\p{N}_]*\]?)/giu;

  /**
   * `Me.<ControlName>` reference capture — hole 1 of VBA control-modeling.
   *
   * Real VBA idiom:
   *   `Me.lblTitulo.Caption = "Hello"`            ← property assignment
   *   `Me.txtDescripcion.Value = "World"`          ← property assignment
   *   `Me.ComandoGrabar.Enabled = True`           ← property assignment
   *   `If Nz(Me.MotivoBorrado, "") = "" Then`      ← read in expression
   *
   * The existing call-site scanner (CALL_RE) only fires on `Name(`
   * (paren form) and `Me` is in its keyword blacklist anyway, so
   * `Me.<Control>` references are silently invisible. This regex matches
   * the FIRST identifier after `Me.` regardless of what follows (a
   * property, an index, an assignment, a call argument, etc.) so the
   * form → control binding is surfaced as an UnresolvedReference for the
   * resolver to pick up later. Subsequent segments (`.Caption`, `.Value`,
   * `.Enabled`) are intentionally NOT captured — they are properties of
   * the control, not new symbols.
   *
   * Provenance: `metadata.synthesizedBy = 'vba-me-control'`. Mirrors the
   * `vba-form-binding` (form→sibling-`.cls`) and `vba-name-resolution`
   * (qualified Dim) patterns already documented on `UnresolvedReference.metadata`
   * in `src/types.ts`.
   */
  private static readonly ME_CONTROL_RE = /\bMe\.(\p{L}[\p{L}\p{N}_]*)/gu;

  /** Keywords we never want to match as call receivers. */
  private static readonly CALL_KEYWORD_BLACKLIST = new Set([
    'If',
    'For',
    'While',
    'Do',
    'Select',
    'Case',
    'Then',
    'Else',
    'ElseIf',
    'With',
    'Loop',
    'Wend',
    'End',
    'Return',
    'Dim',
    'Set',
    'Let',
    'Const',
    'ReDim',
    'Static',
    'Public',
    'Private',
    'Friend',
    'Sub',
    'Function',
    'Property',
    'Class',
    'Module',
    'Option',
    'On',
    'Error',
    'Resume',
    'Exit',
    'New',
    'Call',
    'Rem',
    'LBound',
    'UBound',
    'Me',
    'Nothing',
  ]);

  /**
   * Access runtime objects and singletons. Calls on these receivers are
   * real VBA calls but the targets are NOT user-defined modules or
   * classes — they're Access/DAO/ADO runtime types. Synthesizing a
   * `function` node for each would pollute the graph with ~20+ junk
   * nodes per real-world file (audit W4, June 2026). Skip synthesis
   * for any receiver or member in this set.
   *
   * Note: `DoCmd.RunSQL`, `DoCmd.OpenForm`, etc. still get SQL/edge
   * tracking via the dedicated `SQL_WRAPPERS` regex path (REQ-CODE-8),
   * which fires BEFORE this scan and uses its own dispatch — so
   * blacklisting DoCmd here doesn't lose the SQL-flow edges.
   */
  private static readonly RUNTIME_RECEIVER_BLACKLIST = new Set([
    // Form / page references
    'Screen',
    // Access application singletons
    'Application',
    'DoCmd',
    'SysCmd',
    // VBA debugging intrinsic — Debug.Print / Debug.Assert
    'Debug',
    // Access object collections
    'Forms',
    'Reports',
    'Modules',
    'References',
    'CommandBars',
    // Error-handling intrinsic
    'Err',
    // Late-binding factories (return IDispatch — not user code)
    'CreateObject',
    'GetObject',
    // DAO/ADO recordset field collection access (e.g. rcdDatos.Fields("ID"))
    'Fields',
  ]);

  /**
   * Fix 2 (Issue #2): replace string-literal content with spaces so that
   * call-site patterns inside `"..."` spans are invisible to CALL_RE and
   * the statement-form detectors.  Column positions are preserved (each
   * character is replaced 1-for-1) so any col-based metadata stays correct.
   */
  private static maskStringContent(line: string): string {
    let result = '';
    let i = 0;
    while (i < line.length) {
      const ch = line[i]!;
      if (ch === '"') {
        result += ' '; // opening quote masked
        i++;
        while (i < line.length) {
          const c = line[i]!;
          if (c === '"' && line[i + 1] === '"') {
            result += '  '; // doubled-quote escape masked (2 chars)
            i += 2;
          } else if (c === '"') {
            result += ' '; // closing quote masked
            i++;
            break;
          } else {
            result += ' '; // string content → space
            i++;
          }
        }
      } else {
        result += ch;
        i++;
      }
    }
    return result;
  }

  /**
   * Fix 2 (Issue #2): return true iff `receiverName` is a file-local variable
   * (appears in `localVarTypeMap`) whose declared type is a SIMPLE (non-
   * qualified, non-primitive) identifier — a candidate project-defined class.
   * Qualified types (e.g. `DAO.Recordset`) and primitives (`String`, `Long`)
   * return false so runtime/DAO calls are suppressed.
   */
  private isLocalProjectClassVar(receiverName: string): boolean {
    const entry = this.localVarTypeMap.get(receiverName.toLowerCase());
    if (!entry) return false; // not declared in this file → silent
    if (entry.qualified) return false; // DAO.Recordset etc. → silent
    if (VbaExtractor.PRIMITIVE_TYPES.has(entry.outer.toLowerCase())) return false;
    return true;
  }

  private sweepCallsAndSql(src: string): void {
    const lines = src.split('\n');
    const procedureStartLines = new Set<number>();
    // S5 fix: also match `End Sub`/`End Function`/`End Property` after a
    // colon (`:`), so single-line `Public Sub X(): End Sub` is recognized
    // as ending the procedure. The previous `/^\s*End...` only matched at
    // line start, so the proc stack never popped for colon-separated
    // single-line declarations.
    const procedureEndRe = /(?:^|:\s*)End\s+(?:Sub|Function|Property)\b/i;
    const sqlTargetsThisFile = new Set<string>();

    // Walk the source once, emitting call edges and SQL edges per line and
    // tracking the current procedure stack. The previous implementation
    // did this in two passes; audit S1 (June 2026) flagged the first pass
    // as dead code (its `procStack` was never read after the loop). One
    // pass suffices.
    const stack: ProcInfo[] = [];
    const sqlVariables = new Map<string, string>();
    // C2 fix: track each procedure's `endLine` (the line containing the
    // matching `End Sub`/`End Function`/`End Property`) keyed by its
    // `startLine`. After the loop, we update every function node's
    // `endLine` so `codegraph_explore` returns the full body span —
    // not just the signature line.
    const procEndLines = new Map<number, number>();
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const lineNum = i + 1;

      const procStart = VbaExtractor.PROC_RE.exec(line);
      if (procStart) {
        const name = procStart[3] ?? '';
        const bucket = this.localProcs.get(name);
        if (bucket) {
          // Fix 1: select the ProcInfo whose startLine matches the current
          // line, not always bucket[0]. When Property Get/Let/Set share a
          // name, all three exist in the bucket; pushing bucket[0] every time
          // meant Let/Set bodies erroneously attributed to the Get's ProcInfo
          // (wrong endLine and wrong caller source on call edges).
          const proc = bucket.find((p) => p.startLine === lineNum) ?? bucket[0];
          if (proc) stack.push(proc);
        }
        procedureStartLines.add(lineNum);
      } else if (procedureEndRe.test(line) && stack.length > 0) {
        const ending = stack.pop()!;
        procEndLines.set(ending.startLine, lineNum);
        continue;
      }

      // Fix 2 (Issue #2): mask string-literal content before call scanning so
      // patterns like `modHelper.BuildQuery(` inside a string argument are not
      // mistakenly treated as call sites.  SQL scanning still uses the original
      // line because SQL lives INSIDE string literals.
      const callScanLine = VbaExtractor.maskStringContent(line);

      // Don't scan call sites on the line that declares the procedure — it
      // would match the proc name itself in `Sub Outer()`.
      if (!procedureStartLines.has(lineNum) && stack.length > 0) {
        this.scanCallSites(callScanLine, stack[stack.length - 1]!, lineNum);
      }

      // Hueco 1: capture `Me.<Control>` references (property assignments,
      // read expressions, method calls, anything after `Me.`). Only inside
      // procedures because `Me` is only meaningful inside a form's class
      // module. Skipped on the proc-declaration line for the same reason as
      // the call-site scan above.
      if (!procedureStartLines.has(lineNum) && stack.length > 0) {
        this.scanMeControlReferences(callScanLine, stack[stack.length - 1]!, lineNum);
      }

      // SQL wrappers — only inside a procedure (don't pollute module scope
      // with a stray string literal).  Use the ORIGINAL line — SQL is inside
      // string literals, so the masked line would strip the SQL content.
      if (stack.length > 0) {
        this.trackSqlVariableAssignment(lines, i, sqlVariables);
        this.scanSqlInLine(line, lineNum, sqlTargetsThisFile, sqlVariables);
      }

      // H1 fix: detect statement-form Sub calls (no parens, no `Call`
      // keyword). Audit H1 (June 2026): the parens-only CALL_RE made the
      // dominant VBA idiom invisible — `EstablecerDatos m_Error` on the
      // Form_FormNCAuditoriaMotivoEliminado.cls fixture contributed
      // nothing to the call graph. Walked here so we share the proc stack
      // already maintained by this loop.
      if (stack.length > 0 && !procedureStartLines.has(lineNum)) {
        const stmtCall = this.detectStatementCall(callScanLine);
        if (stmtCall) {
          const caller = stack[stack.length - 1]!;
          this.emitStatementCallEdge(caller, stmtCall, lineNum);
        }

        // Fix 7 + Fix 2: qualified statement-form calls (`Receiver.Member args`) —
        // the dominant cross-object call shape in real Dysflow fixtures.
        // `CALL_RE` only matches the paren form; this path covers the no-paren
        // statement form and emits a heuristic `calls` edge ONLY when the
        // receiver is a file-local variable typed as a candidate project class
        // (Fix 2: REQ-CODE-4 "unresolvable call is silent").
        const qualStmt = this.detectQualifiedStatementCall(callScanLine);
        if (qualStmt) {
          const caller = stack[stack.length - 1]!;
          if (this.isLocalProjectClassVar(qualStmt.receiver)) {
            this.emitQualifiedStatementCallEdge(caller, qualStmt.receiver, qualStmt.member, lineNum);
          }
        }
      }
    }

    // Apply endLine to every emitted function node keyed by its startLine.
    // Functions without a recorded endLine (e.g. malformed VBA without an
    // `End`) keep their `endLine = startLine` from sweepProcedures —
    // which is the correct "single line" representation.
    for (const n of this.nodes) {
      if (n.kind !== 'function') continue;
      const end = procEndLines.get(n.startLine);
      if (end !== undefined) n.endLine = end;
    }
  }

  private scanCallSites(line: string, from: ProcInfo, lineNum: number): void {
    VbaExtractor.CALL_RE.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = VbaExtractor.CALL_RE.exec(line)) !== null) {
      const receiver = m[1] ?? '';
      const member = m[2] ?? '';
      if (!receiver) continue;
      // Skip VBA control-flow keywords.
      if (VbaExtractor.CALL_KEYWORD_BLACKLIST.has(receiver)) continue;
      if (member && VbaExtractor.CALL_KEYWORD_BLACKLIST.has(member)) continue;
      // Skip Access runtime objects — `Me`, `DoCmd`, `Application`, etc.
      // These calls are real but the targets are NOT user code; emitting
      // synthetic function nodes for them pollutes the graph (audit W4).
      if (VbaExtractor.RUNTIME_RECEIVER_BLACKLIST.has(receiver)) continue;
      if (member && VbaExtractor.RUNTIME_RECEIVER_BLACKLIST.has(member)) continue;
      // Skip the receiver when it equals the containing procedure (self-call).
      if (receiver === from.name && !member) continue;

      const col = m.index;

      if (!member) {
        // Bare `Name(...)` — same-file resolution.
        const bucket = this.localProcs.get(receiver);
        const local = bucket?.[0];
        if (!local) continue; // unresolvable — silent per spec R4.
        const localFuncNode = this.findFunctionNodeByName(receiver);
        if (!localFuncNode) continue;
        this.edges.push({
          source: this.findOrCreateFunctionNodeId(from),
          target: localFuncNode.id,
          kind: 'calls',
          line: lineNum,
          column: col,
        });
      } else {
        // Qualified `Receiver.Member(...)` — synthesize the call target.
        const qualified = `${receiver}.${member}`;
        // Avoid emitting duplicate edges for the same call (within a line).
        const dedupeKey = `${from.name}->${qualified}@${lineNum}`;
        if (this.callDedupe.has(dedupeKey)) continue;
        this.callDedupe.add(dedupeKey);

        const synthId = generateNodeId(
          this.filePath,
          'function',
          qualified,
          lineNum,
        );
        // Only add the synthetic function node once per (file, qualified, line).
        if (!this.synthFunctionNodeIds.has(synthId)) {
          this.synthFunctionNodeIds.add(synthId);
          this.nodes.push({
            id: synthId,
            kind: 'function',
            name: qualified,
            qualifiedName: qualified,
            filePath: this.filePath,
            language: 'vba',
            startLine: lineNum,
            endLine: lineNum,
            startColumn: col,
            endColumn: col + qualified.length,
            visibility: 'public',
            updatedAt: Date.now(),
          });
        }
        this.edges.push({
          source: this.findOrCreateFunctionNodeId(from),
          target: synthId,
          kind: 'calls',
          provenance: 'heuristic',
          metadata: { synthesizedBy: 'vba-name-resolution' },
          line: lineNum,
          column: col,
        });
      }
    }
  }

  /** Cache so we don't re-emit the same proc function node per call site. */
  private procNodeIdCache = new Map<string, string>();
  private callDedupe = new Set<string>();
  private synthFunctionNodeIds = new Set<string>();

  /**
   * Hueco 1: scan a line for `Me.<ControlName>` patterns and emit one
   * UnresolvedReference per occurrence, tagged
   * `metadata.synthesizedBy: 'vba-me-control'`.
   *
   * Operates on the masked `callScanLine` (string-literal content already
   * replaced with spaces) so `Me.X` inside a string literal is not falsely
   * captured. Per-site emission — every `Me.lblTitulo` reference site
   * produces its own UnresolvedReference carrying the line/column; the
   * resolver fans them out into multiple `references` edges at index time
   * once the matching `form-instance-control` node exists.
   *
   * `fromNodeId` is the current procedure's function node — that's the
   * "owner" of the reference (the Sub body that wrote `Me.lblTitulo = …`).
   */
  private scanMeControlReferences(line: string, from: ProcInfo, lineNum: number): void {
    VbaExtractor.ME_CONTROL_RE.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = VbaExtractor.ME_CONTROL_RE.exec(line)) !== null) {
      const controlName = m[1] ?? '';
      if (!controlName) continue;
      this.unresolvedReferences.push({
        fromNodeId: this.findOrCreateFunctionNodeId(from),
        referenceName: controlName,
        referenceKind: 'references',
        line: lineNum,
        column: m.index + 3, // +3 to skip the `Me.` prefix
        filePath: this.filePath,
        language: 'vba',
        metadata: { synthesizedBy: 'vba-me-control' },
      });
    }
  }

  private findOrCreateFunctionNodeId(proc: ProcInfo): string {
    // Fix 1: key the cache by `name:startLine` so Property Get/Let/Set
    // accessors with the same name each resolve to their own node.
    const cacheKey = `${proc.name}:${proc.startLine}`;
    const cached = this.procNodeIdCache.get(cacheKey);
    if (cached) return cached;
    // Prefer the startLine-indexed lookup (O(1), correct for all three
    // Property accessors). Fall back to the name-indexed cache for procs
    // whose startLine wasn't recorded (should not happen in practice).
    const fn =
      this.functionNodeByStartLine.get(proc.startLine) ??
      this.findFunctionNodeByName(proc.name);
    if (fn) {
      this.procNodeIdCache.set(cacheKey, fn.id);
      return fn.id;
    }
    // Fallback: synthesize a node id matching generateNodeId's input shape.
    const id = generateNodeId(this.filePath, 'function', proc.name, proc.startLine);
    this.procNodeIdCache.set(cacheKey, id);
    return id;
  }

  private findFunctionNodeByName(name: string): Node | undefined {
    // O(1) via the cache populated as function nodes are emitted in
    // `sweepProcedures`. Audit S2 (June 2026): the previous
    // `this.nodes.find(...)` was an O(n) linear scan per call site —
    // meaningful on real .cls files with hundreds of procedures.
    return this.functionNodeByName.get(name);
  }

  /**
   * H1: detect a statement-form Sub call.
   *
   * Real VBA idioms:
   *   `MySub`                           — bare no-argument statement call
   *   `MySub arg1, Nz(x, 0)`            — bare statement call
   *   `Call MySub arg1, Nz(x, 0)`       — Call keyword, no parens around call
   *   `MySub(arg1, arg2)`               — parens call (handled by CALL_RE)
   *   `Call MySub(arg1, arg2)`          — Call keyword + parens (also CALL_RE)
   *
   * Returns the called proc name if the line matches the statement form
   * AND the proc name is in localProcs (same-file resolution). Returns
   * null for declarations, assignments, comments, keyword lines, etc.
   *
   * The `Call` keyword form is handled by stripping `Call ` and
   * running the same logic on the remainder — they're structurally the
   * same call after the keyword.
   */
  private detectStatementCall(line: string): string | null {
    let trimmed = line.trimStart();
    if (!trimmed) return null;
    // Strip `Call ` keyword if present — same call shape after.
    if (/^Call\s/i.test(trimmed)) {
      trimmed = trimmed.replace(/^Call\s+/i, '');
    }
    // Skip comment lines.
    if (trimmed.startsWith("'") || trimmed.startsWith('Rem ')) return null;
    // Skip declarations: Dim/Private/Public/Static/Global/Const/ReDim.
    if (/^(Dim|Private|Public|Static|Global|Const|ReDim)\s/i.test(trimmed)) return null;
    // Extract the leading identifier.
    const m = /^(\p{L}[\p{L}\p{N}_]*)/u.exec(trimmed);
    if (!m) return null;
    const procName = m[1] ?? '';
    const rest = trimmed.slice(procName.length);
    // `MySub(...)` is parens-form and already handled by CALL_RE. Parentheses
    // are valid later in statement-form argument expressions (`MySub Nz(x, 0)`),
    // so only the character immediately after the proc name distinguishes the
    // call form.
    if (rest.startsWith('(')) return null;
    // Bare `MySub` is a valid no-argument statement-form Sub call.
    if (rest.length === 0) return procName;
    const nextCh = trimmed.charAt(procName.length);
    if (nextCh !== ' ' && nextCh !== '\t') return null;
    const args = rest.trimStart();
    // Skip leading-identifier assignments (`X = ...`). Do not reject `=` inside
    // argument expressions because named arguments use `:=` and comparisons can
    // appear in expressions.
    if (args.startsWith('=')) return null;
    return procName;
  }

  /**
   * H1: emit a same-file `calls` edge for a statement-form Sub call.
   * `caller` is the current procedure (from the stack); `procName` is
   * the called procedure name. Emits a plain `calls` edge to the
   * already-emitted function node.
   */
  private emitStatementCallEdge(
    caller: ProcInfo,
    procName: string,
    lineNum: number,
  ): void {
    if (procName === caller.name) return; // skip self-call
    if (VbaExtractor.CALL_KEYWORD_BLACKLIST.has(procName)) return;
    if (VbaExtractor.RUNTIME_RECEIVER_BLACKLIST.has(procName)) return;
    const bucket = this.localProcs.get(procName);
    if (!bucket || !bucket[0]) return;
    const target = this.findFunctionNodeByName(procName);
    if (!target) return;
    this.edges.push({
      source: this.findOrCreateFunctionNodeId(caller),
      target: target.id,
      kind: 'calls',
      line: lineNum,
      column: 0,
    });
  }

  /**
   * Fix 7: detect a qualified statement-form call — `Receiver.Member <args>`
   * where `Receiver.Member` is NOT followed by `(`.
   *
   * Real VBA idioms this covers:
   *   `m_NCOp.Registrar m_ARAlInicio, p_Error`
   *   `m_Obj.Init`
   *
   * These are distinct from the paren form (`m_NCOp.Registrar(...)`) which is
   * already handled by `CALL_RE`. Returns `{receiver, member}` or `null`.
   *
   * Property assignments (`Receiver.Prop = value`) are excluded via the `=`
   * guard — consistent with `detectStatementCall`'s same-file assignment skip.
   * Blacklisted receivers (runtime objects, keywords) are excluded too.
   */
  private detectQualifiedStatementCall(
    line: string,
  ): { receiver: string; member: string } | null {
    let trimmed = line.trimStart();
    if (!trimmed) return null;
    // Strip `Call` keyword — same call shape after it.
    if (/^Call\s/i.test(trimmed)) trimmed = trimmed.replace(/^Call\s+/i, '');
    // Skip comment lines.
    if (trimmed.startsWith("'") || /^Rem(\s|$)/i.test(trimmed)) return null;
    // Skip declarations.
    if (/^(Dim|Private|Public|Static|Global|Const|ReDim)\s/i.test(trimmed)) return null;
    // Extract receiver identifier.
    const receiverM = /^(\p{L}[\p{L}\p{N}_]*)/u.exec(trimmed);
    if (!receiverM) return null;
    const receiver = receiverM[1] ?? '';
    const rest = trimmed.slice(receiver.length);
    // Must have a dot separator.
    if (!rest.startsWith('.')) return null;
    // Extract member identifier.
    const memberRest = rest.slice(1); // skip the dot
    const memberM = /^(\p{L}[\p{L}\p{N}_]*)/u.exec(memberRest);
    if (!memberM) return null;
    const member = memberM[1] ?? '';
    const afterMember = memberRest.slice(member.length);
    // Must NOT be followed by `(` — the paren form is handled by CALL_RE.
    if (afterMember.startsWith('(')) return null;
    // Must be followed by space/tab (args present) OR end of line (no args).
    if (afterMember.length > 0) {
      const ch = afterMember.charAt(0);
      if (ch !== ' ' && ch !== '\t') return null;
      // Skip property assignments: `Receiver.Prop = value`.
      const argsText = afterMember.trimStart();
      if (argsText.startsWith('=')) return null;
    }
    // Respect the keyword and runtime blacklists.
    if (VbaExtractor.CALL_KEYWORD_BLACKLIST.has(receiver)) return null;
    if (VbaExtractor.RUNTIME_RECEIVER_BLACKLIST.has(receiver)) return null;
    if (VbaExtractor.CALL_KEYWORD_BLACKLIST.has(member)) return null;
    if (VbaExtractor.RUNTIME_RECEIVER_BLACKLIST.has(member)) return null;
    return { receiver, member };
  }

  /**
   * Fix 7: emit a heuristic `calls` edge for a qualified statement-form call.
   * Same shape as the qualified-paren path in `scanCallSites` — reuses the
   * same `callDedupe` / `synthFunctionNodeIds` sets for de-duplication so a
   * paren and non-paren form on the same line don't create duplicate edges.
   */
  private emitQualifiedStatementCallEdge(
    caller: ProcInfo,
    receiver: string,
    member: string,
    lineNum: number,
  ): void {
    const qualified = `${receiver}.${member}`;
    const dedupeKey = `${caller.name}->${qualified}@${lineNum}`;
    if (this.callDedupe.has(dedupeKey)) return;
    this.callDedupe.add(dedupeKey);

    const synthId = generateNodeId(this.filePath, 'function', qualified, lineNum);
    if (!this.synthFunctionNodeIds.has(synthId)) {
      this.synthFunctionNodeIds.add(synthId);
      this.nodes.push({
        id: synthId,
        kind: 'function',
        name: qualified,
        qualifiedName: qualified,
        filePath: this.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: 0,
        endColumn: qualified.length,
        visibility: 'public',
        updatedAt: Date.now(),
      });
    }
    this.edges.push({
      source: this.findOrCreateFunctionNodeId(caller),
      target: synthId,
      kind: 'calls',
      provenance: 'heuristic',
      metadata: { synthesizedBy: 'vba-name-resolution' },
      line: lineNum,
      column: 0,
    });
  }

  private scanSqlInLine(
    line: string,
    lineNum: number,
    dedupe: Set<string>,
    sqlVariables: Map<string, string>,
  ): void {
    for (const { re } of VbaExtractor.SQL_WRAPPERS) {
      // Each wrapper regex is stateful (has /g); reset before use.
      const localRe = new RegExp(re.source, re.flags);
      let m: RegExpExecArray | null;
      while ((m = localRe.exec(line)) !== null) {
        this.emitSqlTableReferences(m[1] ?? '', lineNum, dedupe);
      }
    }

    const localRe = new RegExp(
      VbaExtractor.SQL_VAR_EXEC_RE.source,
      VbaExtractor.SQL_VAR_EXEC_RE.flags,
    );
    let vm: RegExpExecArray | null;
    while ((vm = localRe.exec(line)) !== null) {
      const varName = (vm[1] ?? '').toLowerCase();
      const sqlString = sqlVariables.get(varName);
      if (!sqlString) continue;
      this.emitSqlTableReferences(sqlString, lineNum, dedupe);
    }
  }

  private trackSqlVariableAssignment(
    lines: string[],
    lineIndex: number,
    sqlVariables: Map<string, string>,
  ): void {
    const line = lines[lineIndex] ?? '';
    const m = VbaExtractor.SQL_VAR_ASSIGN_RE.exec(line);
    if (!m) return;
    const varName = (m[1] ?? '').toLowerCase();
    const sqlText = this.collectStringLiteralText(lines, lineIndex);
    if (!sqlText) return;
    sqlVariables.set(varName, sqlText);
  }

  private collectStringLiteralText(lines: string[], startIndex: number): string {
    const fragments: string[] = [];
    for (let i = startIndex; i < lines.length; i++) {
      const line = lines[i] ?? '';
      for (const lit of extractStringLiterals(line)) {
        fragments.push(lit.text);
      }
      if (!line.trimEnd().endsWith('&')) break;
    }
    return fragments.join(' ');
  }

  private emitSqlTableReferences(
    sqlString: string,
    lineNum: number,
    dedupe: Set<string>,
  ): void {
    // Scan the SQL string for FROM/INTO/UPDATE <table>.
    // Preserve the source regex's `/u` flag (Unicode property classes)
    // — hardcoding `'gi'` here would silently break non-ASCII identifiers.
    const tableRe = new RegExp(
      VbaExtractor.SQL_TABLE_RE.source,
      VbaExtractor.SQL_TABLE_RE.flags,
    );
    let tm: RegExpExecArray | null;
    while ((tm = tableRe.exec(sqlString)) !== null) {
      let table = (tm[1] ?? '').replace(/[\[\]]/g, '');
      if (!table) continue;
      const key = `${lineNum}:${table}`;
      if (dedupe.has(key)) continue;
      dedupe.add(key);
      this.emitReference(table, lineNum, 0, 'vba-sql-table');
    }
  }

  /**
   * Emit a `references` edge from the file's module/class node to a synthetic
   * node named `targetName`. Used by Dim, WithEvents, Implements, and SQL.
   */
  private emitReference(
    targetName: string,
    lineNum: number,
    column: number,
    synthesizedBy: string,
  ): void {
    if (!targetName) return;
    // Fix 5: key the synthetic node id on (filePath, kind, name) WITHOUT
    // lineNum so the same type/table referenced on N different lines produces
    // exactly ONE node. The edge carries the per-site `line`/`column`.
    const targetId = generateNodeId(
      this.filePath,
      'class', // placeholder kind; cross-file resolution will re-type at lookup
      targetName,
      0,        // stable — line-independent
    );
    if (this.synthClassNodeIds.has(targetId)) {
      // Node already emitted for this name — only add the edge below.
    } else {
      this.synthClassNodeIds.add(targetId);
      this.nodes.push({
        id: targetId,
        kind: 'class',
        name: targetName,
        qualifiedName: targetName,
        filePath: this.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: column,
        endColumn: column + targetName.length,
        updatedAt: Date.now(),
      });
    }
    const edge: Edge = {
      source: '', // rewritten after module node exists
      target: targetId,
      kind: 'references',
      provenance: 'heuristic',
      metadata: { synthesizedBy },
      line: lineNum,
      column,
    };
    this.edges.push(edge);
    this.pendingModuleOrClassSource.push(edge);
  }

  private synthClassNodeIds = new Set<string>();

  /**
   * Fix 2 (Issue #2): maps `variableName.toLowerCase()` → declared type info.
   * Built by `sweepDimsAndWithEvents`; consulted by `sweepCallsAndSql` to gate
   * qualified statement-form calls — only receivers that are file-local variables
   * typed as a SIMPLE (non-qualified, non-primitive) identifier emit edges.
   */
  private localVarTypeMap = new Map<string, { outer: string; qualified: boolean }>();
}

/**
 * Hueco 3 helper: parse an Access event-handler Sub name into its
 * `<ControlName>_<EventName>` components. Returns null when the name does
 * not match the convention, when the split yields an empty segment, or
 * when the control name is `Form` (form-level events are NOT control
 * handlers — `Form_Load` is the form's own lifecycle event, not a
 * command-button click).
 *
 * Splitting on the LAST underscore (rather than the first) lets
 * multi-word event names parse correctly: `ComandoAltaPM_BeforeDelConfirm`
 * yields control=`ComandoAltaPM`, event=`BeforeDelConfirm`, not
 * control=`ComandoAlta`, event=`PM_BeforeDelConfirm`.
 */
function parseEventHandlerName(
  name: string,
): { controlName: string; eventName: string } | null {
  if (!name) return null;
  const lastUnderscore = name.lastIndexOf('_');
  if (lastUnderscore <= 0) return null; // no underscore OR starts with underscore
  const controlName = name.slice(0, lastUnderscore);
  const eventName = name.slice(lastUnderscore + 1);
  if (!controlName || !eventName) return null;
  // Form-level events live on the form, not on a control.
  if (controlName.toLowerCase() === 'form') return null;
  return { controlName, eventName };
}
