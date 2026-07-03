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
  preprocessConditionalCompilation,
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
   * Class-name prefix for `qualifiedName` composition.
   *
   * B3 (hueco 5): for `.cls` files every function node's `qualifiedName`
   * is composed as `${className}.${procName}` (e.g. `Form_TestForm.Form_Load`)
   * so cross-class callers can disambiguate which form each `Form_Load`
   * belongs to. For `.bas` files this is `null` — module-scoped procs
   * keep their bare-name `qualifiedName` (e.g. `DoThing`), preserving
   * the pre-hueco-5 behavior and every existing qualifiedName-based
   * assertion in `extraction-vba.test.ts`.
   *
   * Resolved once per `extract()` from `isCls` + `vbName` (both already
   * computed for `createModuleOrClassNode`); `sweepProcedures` reads it
   * to set `ProcInfo.qualifiedName` and the function node's qualifiedName.
   */
  private classNamePrefix: string | null = null;
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

      // Pre-process pipeline (per design): join continuations, blank inactive
      // conditional-compilation branches, strip comments, then sweep the
      // joined-but-uncommented source. The conditional preprocessor preserves
      // line count by replacing directives/inactive lines with empty strings.
      const joined = joinLineContinuations(this.source);
      const preprocessed = preprocessConditionalCompilation(joined);
      const uncommented = stripVbaComments(preprocessed);

      // Create the file node.
      this.nodes.push(this.createFileNode());

      // Resolve the file's "kind" — .cls → class, else module (including .bas
      // and the legacy .frm/.dsr stubs).
      const isCls = this.filePath.toLowerCase().endsWith('.cls');

      // Detect VB_Name attribute (first non-empty line).
      const vbName = this.detectVbName(this.source);

      // B3 (hueco 5): compute the class-name prefix used to compose every
      // `function` node's `qualifiedName` in this file. For `.cls` files
      // the prefix is the resolved VB_Name (or the file basename when
      // VB_Name is absent), producing qualifiedNames like
      // `Form_TestForm.Form_Load`. For `.bas` (and `.frm`/`.dsr`) files
      // the prefix is null, preserving the bare-name qualifiedName that
      // every existing test asserts.
      const fallbackName = path
        .basename(this.filePath)
        .replace(/\.[^.]+$/, '');
      const className = isCls ? (vbName ?? fallbackName) : null;
      this.classNamePrefix = className;

      // Module/class node — created lazily so a file containing only
      // Option directives (REQ-CODE-10: "emits nothing") doesn't carry a
      // module node with no symbols attached.
      let hasAnySymbols = false;

      // Sweep: procedures → function nodes + contains edges.
      const procs = this.sweepProcedures(uncommented);
      if (procs.length > 0) hasAnySymbols = true;

      // Sweep: first-class Event / Type / Declare declarations (roadmap #26).
      // This runs before call-site scanning so `RaiseEvent Foo` and calls to a
      // `Declare` can resolve to real local nodes rather than fall through.
      const declarationCount = this.sweepEventsTypesAndDeclares(uncommented);
      if (declarationCount > 0) hasAnySymbols = true;

      // Sweep: Implements (REQ-CODE-5).
      const implCount = this.sweepImplements(uncommented);
      if (implCount > 0) hasAnySymbols = true;

      // Sweep: qualified Dim (REQ-CODE-6) and WithEvents (REQ-CODE-7).
      const dimCount = this.sweepDimsAndWithEvents(uncommented);
      if (dimCount > 0) hasAnySymbols = true;

      // Sweep: Enum/Const declarations (REQ-CODE-12, REQ-CODE-13). Dysflow
      // exports the full module text, so a constants module carries its
      // `Const` lines and `Enum ... End Enum` blocks verbatim — these are the
      // project's domain dictionary and must be in the graph.
      const enumConstCount = this.sweepEnumsAndConsts(uncommented);
      if (enumConstCount > 0) hasAnySymbols = true;

      // Sweep: call sites (REQ-CODE-4) and SQL-in-strings (REQ-CODE-8).
      // Both walk the same per-line view; combining them in one pass is
      // simpler and keeps line tracking consistent.
      this.sweepCallsAndSql(uncommented);

      // Create the module/class node ONLY when the file has actual symbols
      // (REQ-CODE-10 — a file with ONLY Option directives emits zero symbol
      // nodes. A file with Enum/Const but no procedures DOES emit a module
      // node, since those are real symbols — see REQ-CODE-12/13).
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
        // B3 (hueco 5): when this file is a `.cls`, prefix the
        // qualifiedName with the resolved class name so cross-class
        // queries (e.g. `Form_Load`) match only the owning class.
        // `.bas` files leave `qualifiedName === name`.
        qualifiedName: this.classNamePrefix
          ? `${this.classNamePrefix}.${name}`
          : name,
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
        // B3 (hueco 5): same prefix rule as the ProcInfo above — class
        // methods get `${className}.${name}`, module-level Subs keep
        // their bare name.
        qualifiedName: this.classNamePrefix
          ? `${this.classNamePrefix}.${name}`
          : name,
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
          source: nodeId,
          target: controlNodeId,
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

  /** `[visibility] Event <Name>(...)` custom event declaration. */
  private static readonly EVENT_DECL_RE =
    /^\s*((?:Public|Private|Friend)\s+)?Event\s+(\p{L}[\p{L}\p{N}_]*)\b/iu;

  /** `[visibility] Type <Name>` user-defined type block start. */
  private static readonly TYPE_START_RE =
    /^\s*((?:Public|Private|Friend)\s+)?Type\s+(\p{L}[\p{L}\p{N}_]*)\b/iu;

  /** `End Type` user-defined type block end. */
  private static readonly TYPE_END_RE = /^\s*End\s+Type\b/iu;

  /** `<MemberName> As <Type>` inside a user-defined type block. */
  private static readonly TYPE_MEMBER_RE =
    /^\s*(\p{L}[\p{L}\p{N}_]*)\s*(?:\([^)]*\))?\s+As\s+(.+?)\s*$/iu;

  /** `[visibility] Declare [PtrSafe] Sub|Function <Name> Lib "dll" [Alias "x"] ...` */
  private static readonly DLL_DECLARE_RE =
    /^\s*((?:Public|Private)\s+)?Declare\s+(PtrSafe\s+)?(Sub|Function)\s+(\p{L}[\p{L}\p{N}_]*)\s+Lib\s+"([^"]+)"(?:\s+Alias\s+"([^"]+)")?/iu;

  /**
   * Roadmap #26 declaration sweep:
   * - Event declarations become `event` nodes and `RaiseEvent` can point to them.
   * - Type...End Type blocks become `type` + `type_member` nodes.
   * - Win32 API Declare statements become `declare` nodes, while still being
   *   cached by name so normal call-site scanning can emit `calls` edges.
   */
  private sweepEventsTypesAndDeclares(src: string): number {
    const lines = src.split('\n');
    let count = 0;
    let currentType: { id: string; name: string } | null = null;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const lineNum = i + 1;

      if (currentType) {
        if (VbaExtractor.TYPE_END_RE.test(line)) {
          currentType = null;
          continue;
        }
        const member = VbaExtractor.TYPE_MEMBER_RE.exec(line);
        if (member) {
          const memberName = member[1] ?? '';
          const memberType = (member[2] ?? '').trim();
          if (!memberName) continue;
          const memberId = generateNodeId(this.filePath, 'type_member', memberName, lineNum);
          this.nodes.push({
            id: memberId,
            kind: 'type_member',
            name: memberName,
            qualifiedName: `${currentType.name}.${memberName}`,
            filePath: this.filePath,
            language: 'vba',
            startLine: lineNum,
            endLine: lineNum,
            startColumn: 0,
            endColumn: line.length,
            metadata: { memberType },
            updatedAt: Date.now(),
          });
          this.edges.push({
            source: currentType.id,
            target: memberId,
            kind: 'type-member',
            provenance: 'parser',
          });
        }
        continue;
      }

      const eventDecl = VbaExtractor.EVENT_DECL_RE.exec(line);
      if (eventDecl) {
        const visibility = VbaExtractor.foldVisibility(eventDecl[1] ?? '');
        const name = eventDecl[2] ?? '';
        if (!name) continue;
        const eventId = generateNodeId(this.filePath, 'event', name, lineNum);
        const eventNode: Node = {
          id: eventId,
          kind: 'event',
          name,
          signature: line.trim(),
          qualifiedName: this.classNamePrefix ? `${this.classNamePrefix}.${name}` : name,
          filePath: this.filePath,
          language: 'vba',
          startLine: lineNum,
          endLine: lineNum,
          startColumn: 0,
          endColumn: line.length,
          visibility,
          updatedAt: Date.now(),
        };
        this.nodes.push(eventNode);
        this.localEvents.set(name.toLowerCase(), eventNode);
        this.pushContainsFromModule(eventId);
        count++;
        continue;
      }

      const typeStart = VbaExtractor.TYPE_START_RE.exec(line);
      if (typeStart) {
        const visibility = VbaExtractor.foldVisibility(typeStart[1] ?? '');
        const name = typeStart[2] ?? '';
        if (!name) continue;
        const typeId = generateNodeId(this.filePath, 'type', name, lineNum);
        this.nodes.push({
          id: typeId,
          kind: 'type',
          name,
          qualifiedName: this.classNamePrefix ? `${this.classNamePrefix}.${name}` : name,
          filePath: this.filePath,
          language: 'vba',
          startLine: lineNum,
          endLine: lineNum,
          startColumn: 0,
          endColumn: line.length,
          visibility,
          updatedAt: Date.now(),
        });
        this.pushContainsFromModule(typeId);
        currentType = { id: typeId, name };
        count++;
        continue;
      }

      const declaration = VbaExtractor.DLL_DECLARE_RE.exec(line);
      if (declaration) {
        const visibility = VbaExtractor.foldVisibility(declaration[1] ?? '');
        const ptrSafe = !!declaration[2];
        const declareKind = (declaration[3] ?? '').toLowerCase();
        const name = declaration[4] ?? '';
        const dll = declaration[5] ?? '';
        const aliasName = declaration[6] ?? undefined;
        if (!name) continue;
        const declareId = generateNodeId(this.filePath, 'declare', name, lineNum);
        const declareNode: Node = {
          id: declareId,
          kind: 'declare',
          name,
          qualifiedName: this.classNamePrefix ? `${this.classNamePrefix}.${name}` : name,
          filePath: this.filePath,
          language: 'vba',
          startLine: lineNum,
          endLine: lineNum,
          startColumn: 0,
          endColumn: line.length,
          visibility,
          metadata: {
            isDeclare: true,
            dll,
            declareKind,
            ptrSafe,
            ...(aliasName ? { aliasName } : {}),
          },
          updatedAt: Date.now(),
        };
        this.nodes.push(declareNode);
        if (!this.functionNodeByName.has(name)) {
          this.functionNodeByName.set(name, declareNode);
        }
        this.pushContainsFromModule(declareId);
        count++;
      }
    }

    return count;
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
   * Issue #47: now also accepts `Global` (module-level typed instance) and
   * `Static` (procedure-local retention modifier) so they emit the same
   * `references` edge and `localVarTypeMap` registration as their `Dim`
   * siblings today. The negative lookahead is unchanged: `Const` is still
   * routed to `sweepEnumsAndConsts`.
   */
  private static readonly DIM_DECL_PREFIX_RE =
    /^\s*(?:Dim|Private|Public|Global|Static)\s+(?!(?:Function|Sub|Property|Const|WithEvents)\b)/i;

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
   * Bare-declared variable capture for the `Dim|Private|Public|Global|Static`
   * prefix. Captures (1) the variable name. Used to register bare `Dim x`
   * (no `As` clause) and explicit-primitive `Dim x As Long|String|...`
   * declarations into `localVarTypeMap` so the type tracking is consistent
   * across all three Dim shapes:
   *
   *   `Dim x`                → outer = 'variant' (VBA default)
   *   `Dim x As Variant`     → outer = 'variant'  (PRIMITIVE_TYPES member)
   *   `Dim x As Long`        → outer = 'long'     (PRIMITIVE_TYPES member)
   *   `Dim x As Foo`         → outer = 'foo'      (project class — non-primitive)
   *
   * Antigravity audit Task 3: the previous `DIM_ALL_VARS_RE` only matched
   * the `... As <Type>` form, so a bare `Dim x` was invisible to
   * `isLocalProjectClassVar` / `scanCallSites` and `x.Method(1)` produced
   * a dead-end `calls` edge to a stub named `x.Method` that no resolver
   * could repoint. Registering bare Dim with `outer = 'variant'` closes
   * the gate, so `scanCallSites` skips ONLY when the receiver is mapped
   * as a primitive — leaving the "undeclared receiver → stub → resolver
   * repoints" path intact for cross-module qualified calls like
   * `modUtils.Foo(1)` (`modUtils` is not in `localVarTypeMap`).
   */
  private static readonly BARE_DIM_VAR_RE =
    /^\s*(?:Dim|Private|Public|Global|Static)\s+(\p{L}[\p{L}\p{N}_]*)\s*(?:,|$|\b)/iu;

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

  /** `WithEvents m_X As Form_Foo` — Dim/Private/Public/Global/Static prefix is optional. */
  private static readonly WITHEVENTS_RE =
    /^\s*(?:(?:Dim|Private|Public|Global|Static)\s+)?WithEvents\s+\p{L}[\p{L}\p{N}_]*\s+As\s+(\p{L}[\p{L}\p{N}_]*)/iu;

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

        // Antigravity audit Task 3: bare `Dim x` (no `As` clause) and
        // explicit-primitive `Dim x As <primitive>` declarations must
        // still register `x` in `localVarTypeMap` so the qualified-call
        // site scan in `scanCallSites` can gate the dead-end stub that
        // no resolver could ever repoint.
        //
        // Captures the FIRST variable name only. Multi-variable bare
        // Dim (e.g. `Dim a, b, c` without an `As` clause) is rare in
        // real Dysflow fixtures and is intentionally NOT tracked here —
        // those bare variables fall back to the undeclared-receiver
        // path, which is the conservative choice. Skip if the typed-form
        // loop already populated the entry.
        const bm = VbaExtractor.BARE_DIM_VAR_RE.exec(line);
        if (bm) {
          const varName = bm[1] ?? '';
          if (varName) {
            const key = varName.toLowerCase();
            if (!this.localVarTypeMap.has(key)) {
              // Look for an `As <Type>` continuation on the same line so the
              // outer type matches the existing typed-form behaviour. If
              // absent, the variable is implicit `Variant` per VBA semantics.
              const asRe = /\bAs\s+(\p{L}[\p{L}\p{N}_]*)/iu;
              const asMatch = asRe.exec(line);
              const outer = asMatch ? (asMatch[1] ?? '').toLowerCase() : 'variant';
              this.localVarTypeMap.set(key, {
                outer,
                qualified: false,
              });
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
          const weVarM = /^\s*(?:(?:Dim|Private|Public|Global|Static)\s+)?WithEvents\s+(\p{L}[\p{L}\p{N}_]*)/iu.exec(line);
          const weVarName = weVarM?.[1] ?? '';
          if (weVarName) {
            this.localVarTypeMap.set(weVarName.toLowerCase(), {
              outer: formType,
              qualified: false,
              withEvents: true,
              variableName: weVarName,
            });
          }
          this.emitReference(formType, lineNum, 0, 'vba-withevents');
          const targetId = generateNodeId(this.filePath, 'class', formType, 0);
          const subscriberEdge: Edge = {
            source: this.moduleOrClassNode?.id ?? '',
            target: targetId,
            kind: 'subscribes-event',
            provenance: 'heuristic',
            metadata: {
              synthesizedBy: 'vba-withevents',
              variableName: weVarName || undefined,
            },
            line: lineNum,
            column: 0,
          };
          this.edges.push(subscriberEdge);
          if (!this.moduleOrClassNode) {
            this.pendingModuleOrClassSource.push(subscriberEdge);
          }
          count++;
        }
      }
    }
    return count;
  }

  /** `[visibility] Enum <Name>` — opens an enum block. */
  private static readonly ENUM_START_RE =
    /^\s*(?:(Public|Private|Friend|Global)\s+)?Enum\s+(\p{L}[\p{L}\p{N}_]*)/iu;

  /** `End Enum` — closes the current enum block. */
  private static readonly ENUM_END_RE = /^\s*End\s+Enum\b/i;

  /**
   * An enum member line: a leading identifier optionally followed by `=
   * <value>`. Runs only inside an open Enum block, on the (already
   * comment-stripped) source, so a trailing `'comment` never reaches here.
   * `\p{L}` covers accented member names (e.g. `Sí`).
   */
  private static readonly ENUM_MEMBER_RE = /^\s*(\p{L}[\p{L}\p{N}_]*)\s*(?:=|$)/u;

  /** `[visibility] Const <decls>` — captures visibility (1) and the rest (2). */
  private static readonly CONST_DECL_RE =
    /^\s*(?:(Public|Private|Friend|Global)\s+)?Const\s+(.+)$/i;

  /**
   * Fold a VBA visibility keyword to the canonical lowercase enum, matching
   * the procedure convention: `Private` → 'private'; `Public`, `Global`,
   * `Friend`, or none → 'public' (VBA's default module-level `Const`/`Enum`
   * is Private, but we follow the same broader-than-private fold the proc
   * sweep uses so visibility is consistent across symbol kinds).
   */
  private static foldVisibility(raw: string): 'public' | 'private' {
    return raw.trim().toLowerCase() === 'private' ? 'private' : 'public';
  }

  /**
   * Walk the (uncommented, line-joined) source and emit:
   *  - one `enum` node per `Enum <Name>` block, with one `enum_member` node
   *    per member and a `contains` edge enum→member;
   *  - one `constant` node per name declared on a `Const` line (multi-name
   *    lines emit one node per name);
   *  - a `contains` edge from the module/class node to each enum and constant
   *    (held in `pendingModuleOrClassSource` until the module node exists).
   *
   * Returns the number of top-level symbols (enums + constants) emitted so
   * the caller can flip `hasAnySymbols`.
   */
  private sweepEnumsAndConsts(src: string): number {
    const lines = src.split('\n');
    let count = 0;
    let currentEnum: { id: string; name: string } | null = null;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const lineNum = i + 1;

      if (currentEnum) {
        if (VbaExtractor.ENUM_END_RE.test(line)) {
          currentEnum = null;
          continue;
        }
        const mm = VbaExtractor.ENUM_MEMBER_RE.exec(line);
        if (mm) {
          const memberName = mm[1] ?? '';
          if (memberName) {
            const memberId = generateNodeId(
              this.filePath,
              'enum_member',
              memberName,
              lineNum,
            );
            this.nodes.push({
              id: memberId,
              kind: 'enum_member',
              name: memberName,
              qualifiedName: `${currentEnum.name}.${memberName}`,
              filePath: this.filePath,
              language: 'vba',
              startLine: lineNum,
              endLine: lineNum,
              startColumn: 0,
              endColumn: line.length,
              updatedAt: Date.now(),
            });
            // enum → member: source is known, emit directly (not pending).
            this.edges.push({
              source: currentEnum.id,
              target: memberId,
              kind: 'contains',
            });
          }
        }
        continue;
      }

      const enumStart = VbaExtractor.ENUM_START_RE.exec(line);
      if (enumStart) {
        const visibility = VbaExtractor.foldVisibility(enumStart[1] ?? '');
        const name = enumStart[2] ?? '';
        if (!name) continue;
        const enumId = generateNodeId(this.filePath, 'enum', name, lineNum);
        this.nodes.push({
          id: enumId,
          kind: 'enum',
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
        });
        this.pushContainsFromModule(enumId);
        currentEnum = { id: enumId, name };
        count++;
        continue;
      }

      const constDecl = VbaExtractor.CONST_DECL_RE.exec(line);
      if (constDecl) {
        const visibility = VbaExtractor.foldVisibility(constDecl[1] ?? '');
        const body = constDecl[2] ?? '';
        const declarations = parseConstDeclarations(body);
        for (const declaration of declarations) {
          const constName = declaration.name;
          if (!constName) continue;
          if (declaration.value !== null) {
            this.localConstants.set(constName.toLowerCase(), declaration.value);
          }
          const constId = generateNodeId(
            this.filePath,
            'constant',
            constName,
            lineNum,
          );
          this.nodes.push({
            id: constId,
            kind: 'constant',
            name: constName,
            qualifiedName: constName,
            filePath: this.filePath,
            language: 'vba',
            startLine: lineNum,
            endLine: lineNum,
            startColumn: 0,
            endColumn: line.length,
            visibility,
            metadata: declaration.value !== null ? { value: declaration.value } : undefined,
            updatedAt: Date.now(),
          });
          this.pushContainsFromModule(constId);
          count++;
        }
      }
    }
    return count;
  }

  /**
   * Emit a `contains` edge from the (lazily-created) module/class node to
   * `targetId`. Mirrors the pending-source pattern the procedure and
   * implements sweeps use: the module node doesn't exist yet, so the edge's
   * source is rewritten once `extract()` creates it.
   */
  private pushContainsFromModule(targetId: string): void {
    if (this.moduleOrClassNode) {
      this.edges.push({
        source: this.moduleOrClassNode.id,
        target: targetId,
        kind: 'contains',
      });
      return;
    }
    const edge: Edge = { source: '', target: targetId, kind: 'contains' };
    this.edges.push(edge);
    this.pendingModuleOrClassSource.push(edge);
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
    { name: '*db.OpenRecordset', re: /\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.OpenRecordset\s+"((?:[^"]|"")*)"/giu },
    { name: '*db.Execute', re: /\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.Execute\s+"((?:[^"]|"")*)"/giu },
    // Fix 4 (Issue #4): inline-literal forms `getdb().Execute "..."` and
    // `getdb().OpenRecordset "..."` — the variable form is covered by
    // SQL_VAR_EXEC_RE but the direct-literal form was missing.
    { name: 'getdb().Execute', re: /\bgetdb\(\)\.Execute\s+"((?:[^"]|"")*)"/g },
    { name: 'getdb().OpenRecordset', re: /\bgetdb\(\)\.OpenRecordset\s+"((?:[^"]|"")*)"/g },
  ];

  /**
   * `DoCmd.OpenForm "<FormName>"` modelling regex — B4 (hueco 6).
   *
   * Real VBA idiom (matches literal and bare-identifier forms):
   *   `DoCmd.OpenForm "MyForm"`
   *   `DoCmd.OpenForm FORM_MY_FORM`
   *   `DoCmd.OpenForm "MyForm", acNormal, , , acFormEdit`
   *
   * Captures the first argument (group 1). String literals are unwrapped;
   * bare identifiers resolve against local Const declarations, falling back
   * to the identifier name when unknown. The trailing positional args
   * (`acNormal`, `acFormEdit`, etc.) are intentionally NOT captured.
   *
   * Why a separate dispatch: `DoCmd` is in `RUNTIME_RECEIVER_BLACKLIST`
   * (R4 invariant), so `DoCmd.OpenForm` is intentionally SKIPPED by the
   * generic `CALL_RE` path that would otherwise emit a junk `calls` edge
   * to a synthetic `function` node for `DoCmd.OpenForm`. This regex
   * matches BEFORE the call-site scan and uses its own dispatch to emit
   * the `opens-form` edge instead — sharing no logic with CALL_RE.
   */
  private static readonly OPEN_FORM_ARG_RE =
    /\bDoCmd\.OpenForm\s+("(?:(?:[^"]|"")*)"|\p{L}[\p{L}\p{N}_]*)/gu;

  /** SQL assigned to a local variable, e.g. `m_SQL = "SELECT ..." & ...`. */
  private static readonly SQL_VAR_ASSIGN_RE =
    /^\s*(\p{L}[\p{L}\p{N}_]*)\s*=\s*(.*)$/iu;

  /** SQL wrapper called with a variable, e.g. `getdb().Execute m_SQL`. */
  private static readonly SQL_VAR_EXEC_RE =
    /\b(?:\p{L}[\p{L}\p{N}_]*)?db\b(?:\(\))?\.(?:OpenRecordset|Execute)\s*\(?\s*(\p{L}[\p{L}\p{N}_]*)\s*\)?/giu;

  /**
   * SQL table-name regex scoped to the clauses that introduce a table
   * reference: `FROM <t>`, `JOIN <t>`, `INTO <t>`, `UPDATE <t>`. Adding
   * `JOIN` lets the scanner pick up tables from joined fragments that
   * arrive via `&`-concatenated wrapper literals (e.g.
   * `db.Execute "FROM A" & " JOIN B"`); without it the second literal's
   * table was silently dropped even though the wrapper regex now matches
   * the chain.
   *
   * The captured table name is an optional bracketed/unbracketed schema
   * prefix followed by a `.`, then a bracketed-or-bare identifier — so
   * `FROM dbo.tblCustomers` and `FROM [My Schema].[My Table]` come
   * through as one composite reference. Without the prefix the regex
   * still matches a single identifier byte-identical to the old shape.
   * Brackets in the captured composite are stripped by
   * `emitSqlTableReferences` (`replace(/[\[\]]/g, '')`), so the public
   * node name is the unwrapped form `dbo.tblCustomers` /
   * `My Schema.My Table` — matching how plain `[Order Details]` is also
   * unwrapped to `Order Details`. The identifier class
   * `\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*` (same as the saved-queries
   * `TABLE_RE` in `sql-query-extractor.ts`) ensures bracketed names
   * with spaces — `[Order Details]`, `[My Schema]`, `[My Table]` —
   * are captured whole.
   */
  private static readonly SQL_TABLE_RE =
    /\b(?:FROM|JOIN|INTO|UPDATE)\s+((?:(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*)\.)?(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*))/giu;

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

  /**
   * Issue #46: scan `Set <var> = New <Type>[.<Inner>]` lines — the dominant
   * VBA late-instantiation idiom. Run inside `sweepCallsAndSql`'s proc-stack
   * loop so the surrounding procedure is known. For each match:
   *   - register `<var>` in `localVarTypeMap` with `outer=<Type>`,
   *     `qualified=<hasInner>`, `assignedWithSet=true` so the PR #61 refined
   *     gate lets subsequent `<var>.Member ...` qualified calls resolve via
   *     the resolved class name;
   *   - emit a `references` edge from the module/class node to a synthetic
   *     node named `<Type>`, tagged `synthesizedBy: 'vba-set-new'`.
   *
   * Groups: (1) variable name, (2) outer type, (3) optional inner type.
   * Operates on the MASKED line (string-literal content already replaced
   * with spaces) so `Set x = New Foo` inside a string literal never matches.
   */
  private static readonly SET_NEW_RE =
    /\bSet\s+(\p{L}[\p{L}\p{N}_]*)\s*=\s*New\s+(\p{L}[\p{L}\p{N}_]*)(?:\.(\p{L}[\p{L}\p{N}_]*))?/iu;

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

  /**
   * #12a: resolve the "receiver type" used to build a qualified call-stub's
   * name/qualifiedName. When `receiverName` is a file-local variable typed
   * as a candidate project class (`isLocalProjectClassVar`), returns the
   * RESOLVED CLASS NAME from `localVarTypeMap` (e.g. `m_NCOp` typed
   * `As NCOperaciones` → `'NCOperaciones'`) so the stub's qualifiedName
   * matches the real `.cls` method's `${className}.${proc}` shape and the
   * post-extraction resolver (#12b) can find it via an exact qualifiedName
   * match. Otherwise returns `receiverName` unchanged — this is the case
   * for `.bas`-qualified module calls (`modUtils.Foo`), where the receiver
   * IS already the target module's name and no resolution is needed.
   */
  private resolveReceiverType(receiverName: string): string {
    if (this.isLocalProjectClassVar(receiverName)) {
      const entry = this.localVarTypeMap.get(receiverName.toLowerCase());
      if (entry) return entry.outer;
    }
    return receiverName;
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
        const currentProc = stack[stack.length - 1]!;
        this.scanRaiseEvents(callScanLine, currentProc, lineNum);
        this.scanCallSites(callScanLine, currentProc, lineNum);
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
        // Issue #46: `Set x = New <Type>[.<Inner>]` late-instantiation.
        // Run BEFORE the call-site scan so a later `<x>.Member ...` line
        // finds `x` already registered in `localVarTypeMap` and the PR #61
        // refined gate lets the qualified call resolve to `<Type>.Member`.
        const setNew = VbaExtractor.SET_NEW_RE.exec(callScanLine);
        if (setNew) {
          const varName = setNew[1] ?? '';
          const outerType = setNew[2] ?? '';
          const innerType = setNew[3] ?? '';
          if (varName && outerType) {
            // Skip primitives defensively — `Set x = New Long` is nonsense
            // in practice but the gate is cheap and consistent with the
            // Dim sweep's PRIMITIVE_TYPES guard.
            if (!VbaExtractor.PRIMITIVE_TYPES.has(outerType.toLowerCase())) {
              this.localVarTypeMap.set(varName.toLowerCase(), {
                outer: outerType,
                // Mirror `Dim x As Foo.Bar`: qualified `Set rs = New
                // DAO.Recordset` registers `qualified: true` so the PR #61
                // gate keeps downstream `rs.Method` calls silent (DAO is
                // a runtime / external library, not a project class).
                qualified: !!innerType,
                assignedWithSet: true,
                variableName: varName,
              });
              this.emitReference(outerType, lineNum, 0, 'vba-set-new');
            }
          }
        }

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

        // B4 (hueco 6): `DoCmd.OpenForm "FormName"` modelling.
        //
        // The literal form name lives INSIDE a string literal, so we
        // scan the ORIGINAL (unmasked) line — `callScanLine` has its
        // string content replaced with spaces. The receiver is the same
        // proc-stack frame so we attribute the edge to the calling Sub,
        // not the file-level module.
        //
        // Cross-file edge filter note: the synthesized stub node is
        // emitted locally (same file's extraction result), so the
        // per-file filter at `index.ts:insertedIds.has(source/target)`
        // passes the edge naturally — no exemption to the filter is
        // required. The real form-layout node, when VbaFormExtractor
        // later processes the matching .form.txt file, gets a DIFFERENT
        // node id (it uses the real .form.txt path); the stub and the
        // real node coexist harmlessly — the edge from THIS file still
        // references the stub. Future work can collapse the stub once
        // the indexer learns to re-resolve cross-file edges to the real
        // node id (the same pattern the cross-file incoming-edges
        // snapshot already uses at `index.ts:getCrossFileIncomingEdges`).
        const caller2 = stack[stack.length - 1]!;
        this.scanOpenFormCalls(line, caller2, lineNum);
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

  private static readonly RAISE_EVENT_RE =
    /\bRaiseEvent\s+(\p{L}[\p{L}\p{N}_]*)\b/giu;

  private scanRaiseEvents(line: string, from: ProcInfo, lineNum: number): void {
    VbaExtractor.RAISE_EVENT_RE.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = VbaExtractor.RAISE_EVENT_RE.exec(line)) !== null) {
      const eventName = m[1] ?? '';
      const eventNode = this.localEvents.get(eventName.toLowerCase());
      if (!eventNode) continue;
      this.edges.push({
        source: this.findOrCreateFunctionNodeId(from),
        target: eventNode.id,
        kind: 'raises-event',
        provenance: 'parser',
        metadata: { eventName },
        line: lineNum,
        column: m.index,
      });
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
        // #12a: `receiverType` resolves to the real class name when
        // `receiver` is a declared project-class local var (matching a real
        // `.cls` method's `${className}.${proc}` qualifiedName shape so the
        // #12b resolver can find it by exact match); otherwise it's the raw
        // `receiver` text unchanged (e.g. `.bas`-qualified module calls).
        //
        // Antigravity audit Task 3 (refined gate): if `receiver` is a
        // file-local variable declared as a PRIMITIVE (Variant, Object,
        // Empty, Null, LongPtr, LongLong, New, Long, String, ...), skip
        // emission. The previous behaviour emitted a heuristic `calls`
        // edge to a stub named `<receiver>.<member>` that no resolver
        // could ever repoint (Variant can hold anything, including
        // runtime singletons; the stub is dead-end graph pollution).
        //
        // This refined gate does NOT regress cross-module qualified
        // calls like `modUtils.Foo(1)` because `modUtils` is never
        // declared as a file-local variable and therefore is NOT in
        // `localVarTypeMap` — `localVarTypeMap.has(receiver.toLowerCase())`
        // returns false and the gate is skipped. The "stub emitted for
        // undeclared receivers → resolver repoints if a real module
        // exists" behaviour is preserved.
        const recvEntry = this.localVarTypeMap.get(receiver.toLowerCase());
        if (recvEntry && VbaExtractor.PRIMITIVE_TYPES.has(recvEntry.outer.toLowerCase())) {
          continue;
        }
        const receiverType = this.resolveReceiverType(receiver);
        const qualified = `${receiverType}.${member}`;
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
            // #12a: tag the stub so the post-extraction resolver (#12b)
            // can find and repoint it. Mirrors the DoCmd.OpenForm stub
            // precedent (`emitOpensFormEdge`).
            metadata: { stub: true },
            updatedAt: Date.now(),
          });
        }
        this.edges.push({
          source: this.findOrCreateFunctionNodeId(from),
          target: synthId,
          kind: 'calls',
          provenance: 'heuristic',
          metadata: {
            synthesizedBy: 'vba-name-resolution',
            stub: true,
            receiverType,
            member,
          },
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
   * B4 (hueco 6): cache of stub `form-layout` node ids we've already emitted
   * for a given target form name in this file. Avoids emitting duplicate
   * stubs when `DoCmd.OpenForm "FormTest"` shows up N times across N calls.
   * Keyed by the lowercased form name so `FormTest` / `formtest` collapse.
   */
  private opensFormStubIdsByName = new Map<string, string>();

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
    // #12a: the caller already checked `isLocalProjectClassVar(receiver)`
    // before calling this method, so `resolveReceiverType` always returns
    // the RESOLVED CLASS NAME here — the stub's name/qualifiedName matches
    // the real `.cls` method's `${className}.${proc}` shape (e.g.
    // `m_NCOp` typed `As NCOperaciones` → `NCOperaciones.Registrar`, not
    // `m_NCOp.Registrar`) so the #12b resolver can find it by exact
    // qualifiedName match.
    const receiverType = this.resolveReceiverType(receiver);
    const qualified = `${receiverType}.${member}`;
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
        metadata: { stub: true },
        updatedAt: Date.now(),
      });
    }
    this.edges.push({
      source: this.findOrCreateFunctionNodeId(caller),
      target: synthId,
      kind: 'calls',
      provenance: 'heuristic',
      metadata: {
        synthesizedBy: 'vba-name-resolution',
        stub: true,
        receiverType,
        member,
      },
      line: lineNum,
      column: 0,
    });
  }

  /**
   * B4 (hueco 6): scan one line of VBA source for `DoCmd.OpenForm "X"`
   * calls. For each match, emit:
   *  - a stub `form-layout` node for the target form (cached by name so
   *    the same form referenced from N sites emits exactly ONE stub),
   *  - an `opens-form` heuristic edge from the calling Sub to that stub.
   *
   * Both endpoints are pushed into `this.nodes` / `this.edges`, so the
   * per-file edge filter at `index.ts:insertedIds.has(source) &&
   * insertedIds.has(target)` passes the edge naturally without any
   * exemption to the filter.
   *
   * Why a stub and not a direct lookup: the target form lives in a
   * DIFFERENT file (its own `.form.txt`), and the extractor doesn't have
   * DB access at parse time. The stub's synthetic file path
   * (`synthetic:opensFormStub/<FormName>.form.txt`) guarantees a
   * deterministic node id so re-indexes collapse to the same stub.
   * When the consumer's `.form.txt` is later indexed, the real
   * `form-layout` node carries a different id (it uses the real file
   * path); the stub and the real coexist harmlessly. The orchestrator
   * flagged this as acceptable for B4 — only `OpenForm` is in scope.
   * `OpenReport`, `OpenQuery`, `OpenTable`, … are follow-up work.
   *
   * Scope note: this regex matches literal-string and bare-identifier forms.
   * Bare identifiers are resolved only through local `Const` declarations;
   * arbitrary variable data-flow remains intentionally out of scope.
   */
  private scanOpenFormCalls(
    line: string,
    caller: ProcInfo,
    lineNum: number,
  ): void {
    // Each regex has /g so we MUST reset `lastIndex` before use; cloning
    // the regex is the simplest way to avoid leaking state across lines.
    const localRe = new RegExp(
      VbaExtractor.OPEN_FORM_ARG_RE.source,
      VbaExtractor.OPEN_FORM_ARG_RE.flags,
    );
    let m: RegExpExecArray | null;
    while ((m = localRe.exec(line)) !== null) {
      const rawArg = (m[1] ?? '').trim();
      const targetFormName = rawArg.startsWith('"')
        ? unwrapVbaStringLiteral(rawArg)
        : (this.localConstants.get(rawArg.toLowerCase()) ?? rawArg);
      if (!targetFormName) continue;
      this.emitOpensFormEdge(caller, targetFormName, lineNum, m.index);
    }
  }

  /**
   * B4 (hueco 6): emit a stub `form-layout` node for `targetFormName`
   * (cached so duplicates collapse) and an `opens-form` heuristic edge
   * from `caller` to that stub.
   *
   * The edge carries:
   *  - `kind: 'opens-form'`                — new cross-file edge kind
   *  - `provenance: 'heuristic'`           — synthesized, not parsed
   *  - `metadata.targetFormName`           — the captured literal
   *  - `metadata.synthesizedBy: 'vba-opens-form'` — distinguishes this
   *    synthesis from the dim/sql/event-handler families
   *
   * The stub's `metadata.stub: true` flag lets downstream UI render
   * unresolved references distinctly (e.g. with a dashed border) and
   * gives later re-resolution pass a hook for collapse. The stub is
   * line-independent (`line = 0`) so re-indexes produce identical ids.
   */
  private emitOpensFormEdge(
    caller: ProcInfo,
    targetFormName: string,
    lineNum: number,
    column: number,
  ): void {
    const key = targetFormName.toLowerCase();
    let stubId = this.opensFormStubIdsByName.get(key);
    if (!stubId) {
      // Synthetic file path keeps the stub's id deterministic AND
      // disambiguates it from any real `.form.txt` indexed later.
      // The directory prefix (`synthetic:opensFormStub/`) is intentionally
      // not a real filesystem path — it just namespaces the id space.
      const syntheticFilePath = `synthetic:opensFormStub/${targetFormName}.form.txt`;
      stubId = generateNodeId(
        syntheticFilePath,
        'form-layout',
        targetFormName,
        0,
      );
      this.opensFormStubIdsByName.set(key, stubId);
      this.nodes.push({
        id: stubId,
        kind: 'form-layout',
        name: targetFormName,
        // Convention: form module names in Access are `Form_<Name>`.
        // We follow the same convention in the synthetic stub's
        // qualifiedName so cross-file lookups can find it consistently.
        qualifiedName: `Form_${targetFormName}`,
        filePath: syntheticFilePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: 0,
        endColumn: 0,
        metadata: { stub: true },
        updatedAt: Date.now(),
      });
    }
    this.edges.push({
      source: this.findOrCreateFunctionNodeId(caller),
      target: stubId,
      kind: 'opens-form',
      provenance: 'heuristic',
      metadata: {
        synthesizedBy: 'vba-opens-form',
        targetFormName,
      },
      line: lineNum,
      column,
    });
  }

  /**
   * Regex matching the chained `& "..."` literals that may follow a
   * wrapper's first literal on the same physical line. Captures the
   * literal CONTENT (group 1); the surrounding `&` and quotes are
   * structural, not data. VBA allows whitespace around `&` and around
   * the inner quotes — handled with `\s*`. The `((?:[^"]|"")*)` body
   * mirrors the wrapper regex so a `""` inside a chained literal still
   * decodes to a single `"`.
   *
   * Cross-physical-line concat via `_` continuation is OUT OF SCOPE for
   * v1 (deferred; see commit message).
   */
  private static readonly SQL_WRAPPER_CHAIN_RE = /&\s*"((?:[^"]|"")*)"/g;

  /**
   * Given the text that follows a SQL wrapper's first literal on the same
   * physical line, return the contents of every `& "..."` chained literal
   * in source order. Operates per-physical-line only — VBA `_` line
   * continuation across physical lines is handled separately by
   * `collectStringLiteralText` for the variable-assignment path.
   */
  private collectSqlWrapperChain(rest: string): string[] {
    const out: string[] = [];
    const re = new RegExp(
      VbaExtractor.SQL_WRAPPER_CHAIN_RE.source,
      VbaExtractor.SQL_WRAPPER_CHAIN_RE.flags,
    );
    let m: RegExpExecArray | null;
    while ((m = re.exec(rest)) !== null) {
      out.push(m[1] ?? '');
    }
    return out;
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
        const firstLiteral = m[1] ?? '';
        // After the wrapper regex consumes up to and including the closing
        // `"` of the first literal, walk the rest of the line for any
        // `& "..."` chains and concatenate every literal's content. Joining
        // with a space (mirrors `collectStringLiteralText`) keeps adjacent
        // `FROM tblA` & `FROM tblB` separated so `SQL_TABLE_RE` finds both.
        const rest = line.slice(m.index + m[0].length);
        const chain = this.collectSqlWrapperChain(rest);
        const joined = [firstLiteral, ...chain].join(' ');
        this.emitSqlTableReferences(joined, lineNum, dedupe);
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

  /**
   * #13 fix: `sql = sql & "..."` (self-referential concatenation) must
   * ACCUMULATE the new fragment onto whatever was already tracked for
   * `varName`, not overwrite it. Overwriting silently dropped earlier
   * fragments' tables — typically the initial `FROM <table>` in
   * `sql = "SELECT * FROM tblA"` followed by `sql = sql & " WHERE x=1"`.
   *
   * Detection: the RHS (`m[2]`, trimmed) starts with `<varName> &`,
   * case-insensitively — matching VBA's case-insensitive identifiers (`Sql`
   * and `sql` are the same variable). A genuine fresh assignment (RHS does
   * NOT start with the self-reference) still RESETS tracking — that
   * behavior is unchanged.
   */
  private trackSqlVariableAssignment(
    lines: string[],
    lineIndex: number,
    sqlVariables: Map<string, string>,
  ): void {
    const line = lines[lineIndex] ?? '';
    const m = VbaExtractor.SQL_VAR_ASSIGN_RE.exec(line);
    if (!m) return;
    const rawVarName = m[1] ?? '';
    const varName = rawVarName.toLowerCase();
    const rhs = (m[2] ?? '').trim();
    const newFragment = this.collectStringLiteralText(lines, lineIndex);
    if (!newFragment) return;

    const selfRefRe = new RegExp(`^${escapeRegExpLiteral(rawVarName)}\\s*&`, 'i');
    const existing = sqlVariables.get(varName);
    if (existing !== undefined && selfRefRe.test(rhs)) {
      sqlVariables.set(varName, `${existing} ${newFragment}`);
    } else {
      sqlVariables.set(varName, newFragment);
    }
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
  private localVarTypeMap = new Map<string, {
    outer: string;
    qualified: boolean;
    withEvents?: boolean;
    variableName?: string;
    assignedWithSet?: boolean;
  }>();

  /** Local constant name (lowercase) → simple literal value for OpenForm resolution. */
  private localConstants = new Map<string, string>();

  /** Local event name (lowercase) → event node for `RaiseEvent` edge emission. */
  private localEvents = new Map<string, Node>();
}

/**
 * #13 helper: escape a variable name for safe interpolation into the
 * self-reference RegExp built by `trackSqlVariableAssignment`. VBA
 * identifiers are alphanumeric+underscore only, so in practice nothing here
 * ever needs escaping — this guards against regex metacharacters anyway
 * rather than assume the input is always well-formed.
 */
function escapeRegExpLiteral(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function parseConstDeclarations(
  body: string,
): Array<{ name: string; value: string | null }> {
  const declarations: Array<{ name: string; value: string | null }> = [];
  for (const part of splitOutsideVbaStrings(body, ',')) {
    const m =
      /^\s*(\p{L}[\p{L}\p{N}_]*)\s*(?:As\s+[^=]+?)?\s*=\s*(.+?)\s*$/iu.exec(part);
    if (!m) continue;
    const name = m[1] ?? '';
    const rawValue = (m[2] ?? '').trim();
    declarations.push({
      name,
      value: rawValue.startsWith('"') ? unwrapVbaStringLiteral(rawValue) : rawValue || null,
    });
  }
  return declarations;
}

function splitOutsideVbaStrings(value: string, separator: string): string[] {
  const parts: string[] = [];
  let current = '';
  let inString = false;
  for (let i = 0; i < value.length; i++) {
    const ch = value[i];
    const next = value[i + 1];
    if (ch === '"') {
      current += ch;
      if (inString && next === '"') {
        current += next;
        i++;
        continue;
      }
      inString = !inString;
      continue;
    }
    if (!inString && ch === separator) {
      parts.push(current);
      current = '';
      continue;
    }
    current += ch;
  }
  parts.push(current);
  return parts;
}

function unwrapVbaStringLiteral(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed.startsWith('"')) return trimmed;
  let text = '';
  for (let i = 1; i < trimmed.length; i++) {
    const ch = trimmed[i];
    const next = trimmed[i + 1];
    if (ch === '"' && next === '"') {
      text += '"';
      i++;
      continue;
    }
    if (ch === '"') break;
    text += ch;
  }
  return text;
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
