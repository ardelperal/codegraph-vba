/**
 * VbaExtractor — regex extractor for `.bas` / `.cls` / `.frm` / `.dsr` source.
 *
 * Emits the symbol + edge shape expected by `codegraph_explore`:
 *  - one `module` node per `.bas` / one `class` node per `.cls` / one `file`
 *    stub for `.frm` / `.dsr` legacy
 *  - one `function` node per `Sub` / `Function` / `Property Get|Let|Set`
 *    declaration, with `metadata.visibility` set to `'Public'`, `'Private'`,
 *    `'Friend'`, `'Static'`, or `'Public'` (default when no keyword).
 *  - `contains` edges from the module/class node to each function node.
 *  - `calls` edges from a procedure to a same-file procedure (`Sub Outer()` →
 *    `Inner()`); qualified cross-module calls carry `provenance: 'heuristic'`
 *    and `metadata.synthesizedBy: 'vba-name-resolution'`.
 *  - `implements` edges for `Implements IFoo` declarations.
 *  - `references` edges for `Dim x As Foo.Bar` (qualified Dim → outer type)
 *    tagged `synthesizedBy: 'vba-name-resolution'`.
 *  - `references` edges for `WithEvents m_X As Form_Foo` tagged
 *    `synthesizedBy: 'vba-withevents'`.
 *  - `references` edges for SQL table names found inside string literals
 *    passed to `DoCmd.RunSQL`, `CurrentDb.OpenRecordset`, `CurrentDb.Execute`,
 *    or `db.Execute`, tagged `synthesizedBy: 'vba-sql-table'`.
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
        provenance: 'heuristic',
        metadata: { synthesizedBy: 'vba-name-resolution' },
        line: lineNum,
        column: 0,
      };
      this.edges.push(edge);
      this.pendingModuleOrClassSource.push(edge);
      count++;
    }
    return count;
  }

  /** Qualified Dim `Dim x As Foo.Bar` — emits `references` to `Foo`. */
  private static readonly DIM_QUAL_RE =
    /^\s*(?:Dim|Private|Public)\s+\p{L}[\p{L}\p{N}_]*\s+As\s+(\p{L}[\p{L}\p{N}_]*)\.(\p{L}[\p{L}\p{N}_]*)/iu;

  /** `WithEvents m_X As Form_Foo` — Dim/Private/Public prefix is optional. */
  private static readonly WITHEVENTS_RE =
    /^\s*(?:(?:Dim|Private|Public)\s+)?WithEvents\s+\p{L}[\p{L}\p{N}_]*\s+As\s+(\p{L}[\p{L}\p{N}_]*)/iu;

  private sweepDimsAndWithEvents(src: string): number {
    const lines = src.split('\n');
    let count = 0;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const lineNum = i + 1;

      const dimMatch = VbaExtractor.DIM_QUAL_RE.exec(line);
      if (dimMatch) {
        const outerType = dimMatch[1] ?? '';
        if (outerType) {
          this.emitReference(outerType, lineNum, 0, 'vba-name-resolution');
          count++;
        }
      }

      const weMatch = VbaExtractor.WITHEVENTS_RE.exec(line);
      if (weMatch) {
        const formType = weMatch[1] ?? '';
        if (formType) {
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
  ];

  /** SQL table-name regex scoped to FROM / INTO / UPDATE. */
  private static readonly SQL_TABLE_RE =
    /\b(?:FROM|INTO|UPDATE)\s+(\[?\p{L}[\p{L}\p{N}_]*\]?)/giu;

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
  ]);

  private sweepCallsAndSql(src: string): void {
    const lines = src.split('\n');
    const procedureStartLines = new Set<number>();
    const procedureEndRe = /^\s*End\s+(?:Sub|Function|Property)\b/i;
    const sqlTargetsThisFile = new Set<string>();

    // Walk the source once, emitting call edges and SQL edges per line and
    // tracking the current procedure stack. The previous implementation
    // did this in two passes; audit S1 (June 2026) flagged the first pass
    // as dead code (its `procStack` was never read after the loop). One
    // pass suffices.
    const stack: ProcInfo[] = [];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const lineNum = i + 1;

      const procStart = VbaExtractor.PROC_RE.exec(line);
      if (procStart) {
        const name = procStart[3] ?? '';
        const bucket = this.localProcs.get(name);
        if (bucket && bucket[0]) stack.push(bucket[0]);
        procedureStartLines.add(lineNum);
      } else if (procedureEndRe.test(line) && stack.length > 0) {
        stack.pop();
        continue;
      }

      // Don't scan call sites on the line that declares the procedure — it
      // would match the proc name itself in `Sub Outer()`.
      if (!procedureStartLines.has(lineNum) && stack.length > 0) {
        this.scanCallSites(line, stack[stack.length - 1]!, lineNum);
      }

      // SQL wrappers — only inside a procedure (don't pollute module scope
      // with a stray string literal).
      if (stack.length > 0) {
        this.scanSqlInLine(line, lineNum, sqlTargetsThisFile);
      }
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

  private findOrCreateFunctionNodeId(proc: ProcInfo): string {
    const cached = this.procNodeIdCache.get(proc.name);
    if (cached) return cached;
    const fn = this.findFunctionNodeByName(proc.name);
    if (fn) {
      this.procNodeIdCache.set(proc.name, fn.id);
      return fn.id;
    }
    // Fallback: synthesize a node id matching generateNodeId's input shape.
    const id = generateNodeId(this.filePath, 'function', proc.name, proc.startLine);
    this.procNodeIdCache.set(proc.name, id);
    return id;
  }

  private findFunctionNodeByName(name: string): Node | undefined {
    // O(1) via the cache populated as function nodes are emitted in
    // `sweepProcedures`. Audit S2 (June 2026): the previous
    // `this.nodes.find(...)` was an O(n) linear scan per call site —
    // meaningful on real .cls files with hundreds of procedures.
    return this.functionNodeByName.get(name);
  }

  private scanSqlInLine(
    line: string,
    lineNum: number,
    dedupe: Set<string>,
  ): void {
    for (const { re } of VbaExtractor.SQL_WRAPPERS) {
      // Each wrapper regex is stateful (has /g); reset before use.
      const localRe = new RegExp(re.source, re.flags);
      let m: RegExpExecArray | null;
      while ((m = localRe.exec(line)) !== null) {
        const sqlString = m[1] ?? '';
        // Scan the captured SQL string for FROM/INTO/UPDATE <table>.
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
    const targetId = generateNodeId(
      this.filePath,
      'class', // placeholder kind; cross-file resolution will re-type at lookup
      targetName,
      lineNum,
    );
    if (this.synthClassNodeIds.has(targetId)) {
      // Edge already exists? Emit anyway with a fresh line attribution.
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
}