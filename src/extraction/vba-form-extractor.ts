/**
 * VbaFormExtractor — regex extractor for Dysflow's `.form.txt` / `.report.txt`
 * form/report UI files.
 *
 * This extractor is the strict counterpart of `VbaExtractor`:
 *  - It NEVER emits `function`, `sub`, `class`, or non-form `module` nodes.
 *    The form CODE lives in the sibling `.cls` (parsed by `VbaExtractor`),
 *    not here. Dysflow overwrites `.form.txt`'s embedded `CodeBehindForm`
 *    block on the next import, so any code emitted from this file would be
 *    wrong AND ephemeral.
 *  - It emits one `form-layout` node per file (named per `Attribute VB_Name`
 *    when present, otherwise the file basename). The `form-layout` kind
 *    (B2 hueco 4) replaces what was historically emitted as `kind: 'module'`
 *    so consumers can dispatch on a UI-specific kind and avoid confusing
 *    form/report UI files with `.bas` modules. `module` remains the kind
 *    for `.bas` standard modules emitted by `VbaExtractor`.
 *  - It emits one `property` node per Access control declaration with
 *    `metadata.controlType` set to the control type (e.g. `'TextBox'`,
 *    `'CommandButton'`).
 *  - It emits one unresolved reference (synthesizedBy=vba-form-binding) from
 *    the form-layout node to the sibling `.cls` basename, so
 *    `codegraph_explore` can resolve form → class at lookup time.
 *  - It emits one `references` edge per `RecordSource` binding (Issue #49),
 *    tagged `metadata.synthesizedBy: 'vba-record-source'`. Bare table names
 *    and `SELECT/FROM/...` SQL are both supported; SQL goes through the
 *    shared `SQL_TABLE_RE` regex and emits one edge per table. The same
 *    happens for `RowSource` bindings on individual controls, but the
 *    source is the enclosing `form-instance-control` node and the tag is
 *    `'vba-row-source'`. Value-list controls (`RowSourceType = "Value List"`)
 *    are skipped — their RowSource is a literal list, not a data binding.
 *
 * Hard invariant (REQ-FORM-4): ZERO `function`/`sub`/`module`/`event`
 * /`declare`/`type` nodes from `.form.txt` / `.report.txt`. The form-side
 * behavior lives in `vba-form-ui-extraction`; the code-side lives in
 * `vba-code-extraction`. Note: synthetic `class` PLACEHOLDER nodes ARE
 * emitted for external data bindings (table/query references from
 * RecordSource/RowSource) — they are not the form's own class binding
 * (that one stays an UnresolvedReference) and are deduplicated by name
 * so the same table referenced from N sites collapses to one node.
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
import { stripVbaComments } from './vba-preprocess';

export class VbaFormExtractor {
  private filePath: string;
  private source: string;
  private nodes: Node[] = [];
  private edges: Edge[] = [];
  private errors: ExtractionError[] = [];
  private unresolvedReferences: UnresolvedReference[] = [];
  /**
   * Issue #49: synthetic `class` placeholder node ids for table/query
   * references emitted by RecordSource/RowSource sweeps. The same table
   * referenced from N different sites collapses to a single placeholder
   * node (id keyed on `(filePath, 'class', tableName, 0)`), but every
   * call site still emits its own `references` edge with a per-site
   * `line` value. Mirrors `VbaExtractor.synthClassNodeIds` (private to
   * VbaExtractor — we keep a separate cache here so the form extractor
   * stays self-contained per the project's per-extractor state rule).
   */
  private synthClassNodeIds = new Set<string>();

  constructor(filePath: string, source: string) {
    this.filePath = filePath;
    this.source = source;
  }

  extract(): ExtractionResult {
    const startTime = Date.now();

    try {
      // File node (always present so the watcher tracks the file).
      this.nodes.push(this.createFileNode());

      // Strip VBA comments — comments can contain literal `Sub`/`Function`
      // text that we don't want to misinterpret. Note: we deliberately do
      // NOT run joinLineContinuations here — `.form.txt` files don't have
      // VBA-style line continuations.
      const cleaned = stripVbaComments(this.source);

      // Resolve VB_Name attribute (first non-empty line).
      const vbName = this.detectVbName(this.source);
      const basename = this.basenameWithoutExtension();
      const moduleName = vbName ?? basename;

      // Form/report file-level node — `form-layout` (NOT `module`). B2
      // hueco 4 promoted this from `module` to `form-layout` so the form
      // UI is dispatched on its own kind and never confused with a `.bas`
      // standard module. The `metadata.containerKind` preserves the
      // historical `module` label so downstream tooling that still keys on
      // it has a back-compat path; new code should test `kind === 'form-layout'`.
      const formLayoutNode = this.createFormLayoutNode(moduleName);
      this.nodes.push(formLayoutNode);

      // Sibling `.cls` binding (REQ-FORM-1, REQ-FORM-3). Emit an
      // UnresolvedReference rather than a hardcoded `references` edge with
      // a synthetic target node: REQ-FORM-4 forbids ANY `class` node from
      // form files, and a hardcoded target would have to be a `class` kind
      // for downstream resolution to match it. The unresolved reference is
      // resolved at index time when the actual sibling `.cls` is processed.
      this.unresolvedReferences.push({
        fromNodeId: formLayoutNode.id,
        referenceName: basename,
        referenceKind: 'references',
        line: 1,
        column: 0,
        filePath: this.filePath,
        language: 'vba',
        metadata: { synthesizedBy: 'vba-form-binding' },
      });

      // Control declarations → property nodes (REQ-FORM-2).
      this.sweepControls(cleaned);

      // Issue #49: RecordSource/RowSource bindings → references edges to
      // placeholder class nodes (one per table/query). Swept AFTER
      // sweepControls so the `form-instance-control` nodes for any
      // enclosing controls are already in `this.nodes` and
      // `sweepRowSources` can attribute each edge to its source node.
      this.sweepRecordSources(cleaned, formLayoutNode.id);
      this.sweepRowSources(cleaned, formLayoutNode.id);
    } catch (error) {
      this.errors.push({
        message: `VBA form extraction error: ${error instanceof Error ? error.message : String(error)}`,
        filePath: this.filePath,
        severity: 'error',
        code: 'parse_error',
      });
    }

    return {
      nodes: this.nodes,
      edges: this.edges,
      unresolvedReferences: this.unresolvedReferences,
      errors: this.errors,
      durationMs: Date.now() - startTime,
    };
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

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

  /**
   * Strip the two-segment extension (`.form.txt`, `.report.txt`) or the
   * fallback single-segment extension. Returns the file basename without
   * its extension — used as the sibling `.cls` name when `VB_Name` is
   * absent and as the default module name.
   */
  private basenameWithoutExtension(): string {
    const basename = path.basename(this.filePath);
    return basename.replace(/\.(form|report)\.txt$|\.[^.]+$/i, '');
  }

  private detectVbName(src: string): string | null {
    // Same as VbaExtractor's detectVbName — see audit W2 (June 2026).
    // Walk past Access form metadata headers (VERSION / BEGIN / END /
    // Attribute …) so VB_Name on a later line is found.
    for (const line of src.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      const m = /^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/i.exec(trimmed);
      if (m) return m[1] ?? null;
      if (
        /^\s*VERSION\b/i.test(trimmed) ||
        /^\s*BEGIN\b/i.test(trimmed) ||
        /^\s*END\b/i.test(trimmed) ||
        /^\s*(?:MultiUse|Persistable|DataBindingBehavior|DataSourceBehavior)\s*=/i.test(trimmed) ||
        /^\s*Attribute\s+/i.test(trimmed)
      ) {
        continue;
      }
      return null;
    }
    return null;
  }

  /**
   * Build the per-file node for a `.form.txt` / `.report.txt` source.
   * Emits `kind: 'form-layout'` (B2 hueco 4), replacing the historical
   * `kind: 'module'` so consumers can dispatch on a UI-specific kind.
   *
   * The deterministic id formula `generateNodeId(filePath, 'form-layout',
   * name, 1)` is preserved so cross-extractor stubs (e.g. event-handler
   * synthesis on the `.cls` side) can produce a matching id when needed
   * — though the immediate use case is the `module` → `form-layout`
   * rename. `metadata.containerKind` keeps the historical `'module'`
   * label as a back-compat marker; consumers should prefer `node.kind`.
   */
  private createFormLayoutNode(name: string): Node {
    const lines = this.source.split('\n');
    return {
      id: generateNodeId(this.filePath, 'form-layout', name, 1),
      kind: 'form-layout',
      name,
      qualifiedName: name,
      filePath: this.filePath,
      language: 'vba',
      startLine: 1,
      endLine: lines.length,
      startColumn: 0,
      endColumn: 0,
      metadata: { containerKind: 'module' },
      updatedAt: Date.now(),
    };
  }

  /**
   * Match each `Begin <ControlType>` declaration line and emit one `property`
   * node per match. The control type is the second token after `Begin`.
   *
   * The Dysflow SaveAsText format uses:
   *   Begin TextBox
   *       Name = "txtFoo"
   *   End
   *
   * We do not try to balance `Begin`/`End` blocks — each `Begin <Type>` on
   * its own line is a control. Form/report files are shallow enough that
   * the extra `Begin`/`End` for the form's root block are matched but
   * filtered by the control-type blacklist (see below).
   */
  private static readonly BEGIN_RE = /^\s*Begin\s+(\p{L}[\p{L}\p{N}_]*)\s*$/u;

  /**
   * `Name = "..."` attribute line — emits the Access control instance name
   * (e.g. `lblTitulo`, `ComandoAltaPM`). Capture group 1 is the name.
   * The Dysflow SaveAsText format always wraps the value in double quotes
   * — even when the name is a simple identifier — so we anchor on `"…"`
   * without trying to handle unquoted forms.
   */
  private static readonly NAME_RE = /^\s*Name\s*=\s*"([^"]+)"\s*$/u;

  /**
   * Control type tokens that are NOT user-visible Access controls and must be
   * filtered out so they don't appear as `property` nodes.
   *
   * - `Form`    — the form's own root `Begin Form` / `End` container.
   * - `Section` — Access section containers (Header / Detail / Footer).
   *
   * `Rectangle` and `Image` are real Access controls and must NOT appear here.
   */
  private static readonly NON_CONTROL_TYPES = new Set<string>([
    'Form',
    'Section',
  ]);

  /**
   * Maximum scan window for the `Name = "..."` attribute after a
   * `Begin <Type>` line. Real Dysflow exports have at most a handful of
   * whitespace-only lines and the `Name` line within the first 3–6 lines
   * of the block. 16 is a generous bound; if `Name` is missing within
   * that window, the control is treated as a nameless container and only
   * the legacy `property` node is emitted (preserves REQ-FORM-2 for
   * pre-Name `.form.txt` files exported by older Dysflow versions).
   */
  private static readonly NAME_SCAN_WINDOW = 16;

  private sweepControls(src: string): void {
    const lines = src.split('\n');
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const m = VbaFormExtractor.BEGIN_RE.exec(line);
      if (!m) continue;
      const controlType = m[1] ?? '';
      if (!controlType) continue;
      if (VbaFormExtractor.NON_CONTROL_TYPES.has(controlType)) continue;
      // Skip lines whose captured token is the GUID-prefix form of the root
      // Begin (some Dysflow exports write `Begin {XXXXXXXX-XXXX-...}` with
      // a CLSID, not a control type). The GUID pattern won't match
      // [A-Za-z_]\w* — it starts with `{`, so the regex naturally rejects
      // it.
      const lineNum = i + 1;

      // ---- Legacy `property` node (REQ-FORM-2, unchanged). ---------------
      // This node's `name` is the control TYPE (e.g. "CommandButton"). Kept
      // intact for the 11 existing extraction-vba-form.test.ts tests and
      // for the 4 realfixture tests that assert on property-kind counts.
      const nodeId = generateNodeId(
        this.filePath,
        'property',
        controlType,
        lineNum,
      );
      this.nodes.push({
        id: nodeId,
        kind: 'property',
        name: controlType,
        qualifiedName: `${this.filePath}::${controlType}`,
        filePath: this.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: 0,
        endColumn: line.length,
        metadata: { controlType },
        updatedAt: Date.now(),
      });

      // ---- Hueco 2: emit a `form-instance-control` node per NAME. -------
      // Scan ahead up to NAME_SCAN_WINDOW lines for the first
      // `Name = "..."` attribute. The control's `Name` (e.g. "lblTitulo",
      // "ComandoAltaPM") is what the .cls side references via
      // `Me.<ControlName>` (hueco 1) and what event handlers are wired to
      // via the `<ControlName>_<Event>` naming convention (hueco 3).
      // line=0 in the generated id keeps the id STABLE across re-indexes
      // of the same control — the VbaExtractor side synthesizes the
      // matching event-handler edge using the same id formula (see
      // vba-extractor.ts: synthesizeEventHandlerEdge).
      const { name: controlName, nameLine } = this.findControlName(
        lines,
        i,
        lineNum,
      );
      if (!controlName) continue;

      const controlNodeId = generateNodeId(
        this.filePath,
        'form-instance-control',
        controlName,
        0,
      );
      this.nodes.push({
        id: controlNodeId,
        kind: 'form-instance-control',
        name: controlName,
        qualifiedName: `${this.filePath}::${controlName}`,
        filePath: this.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: nameLine, // spans from Begin to the Name attribute line
        startColumn: 0,
        endColumn: 0,
        metadata: { controlType },
        updatedAt: Date.now(),
      });
    }
  }

  /**
   * Scan ahead from a `Begin <Type>` line for the first `Name = "…"`
   * attribute. Returns the captured name and the line number where it
   * was found, or `{ name: '', nameLine: 0 }` when no Name is present
   * within the scan window (e.g. the `Begin Form` root block which has
   * `Caption = "..."` but no `Name`, or pre-Name legacy exports).
   *
   * Stops at the next `Begin` or `End` boundary so a misaligned scan
   * never crosses into a sibling control's attribute block.
   */
  private findControlName(
    lines: string[],
    beginLineIndex: number,
    beginLineNum: number,
  ): { name: string; nameLine: number } {
    const end = Math.min(
      lines.length,
      beginLineIndex + 1 + VbaFormExtractor.NAME_SCAN_WINDOW,
    );
    for (let j = beginLineIndex + 1; j < end; j++) {
      const line = lines[j] ?? '';
      // Boundary check: a sibling Begin or End closes this block. Don't
      // look past it (a missing Name line is the common case for the
      // root `Begin Form` block — its `Caption` is the visible label,
      // not a `Name`).
      if (/^\s*(Begin|End)\b/i.test(line)) break;
      const m = VbaFormExtractor.NAME_RE.exec(line);
      if (m) {
        const name = m[1] ?? '';
        if (name) {
          return { name, nameLine: j + 1 };
        }
      }
    }
    // No Name within the window — that's the case for the `Begin Form`
    // root block (which has `Caption`, not `Name`) and for the
    // `Begin Section` Access section blocks (which group controls but
    // carry no Name of their own). The legacy `property` node was
    // already emitted above; we simply skip the form-instance-control
    // emission so hueco-4 stays RED for the .form.txt module node
    // transition (a separate B2 task).
    return { name: '', nameLine: beginLineNum };
  }

  // ---------------------------------------------------------------------------
  // RecordSource / RowSource edge emission (Issue #49)
  // ---------------------------------------------------------------------------

  /**
   * Issue #49 — `RecordSource = "..."` line regex. The Dysflow SaveAsText
   * format always wraps the value in double quotes; the inner `(?:[^"]|"")*`
   * body tolerates the doubled-quote escape (`""` → `"`) so the captured
   * group carries the literal text exactly as Access stores it.
   *
   * Anchored at the start of the line — properties in the SaveAsText
   * format are always indented but the attribute name is unambiguous, so
   * a leading whitespace-tolerant anchor is enough.
   */
  private static readonly RECORD_SOURCE_RE =
    /^\s*RecordSource\s*=\s*"((?:[^"]|"")*)"/iu;

  /** Issue #49 — `RowSource = "..."` line regex. Mirrors `RECORD_SOURCE_RE`. */
  private static readonly ROW_SOURCE_RE =
    /^\s*RowSource\s*=\s*"((?:[^"]|"")*)"/iu;

  /** Issue #49 — `RowSourceType = "..."` line regex. Used by the control-block
   * scan to detect value-list controls whose RowSource is a literal list. */
  private static readonly ROW_SOURCE_TYPE_RE =
    /^\s*RowSourceType\s*=\s*"([^"]*)"/iu;

  /**
   * Issue #49 — copy of `VbaExtractor.SQL_TABLE_RE`. Same source / flags:
   * captures the table name that follows `FROM`/`JOIN`/`INTO`/`UPDATE`,
   * tolerates bracketed `[Order Details]` identifiers and `\p{L}` Unicode
   * identifiers, and allows an optional schema prefix (`[dbo].[tblA]`).
   *
   * We duplicate the regex (rather than exporting it from `VbaExtractor`)
   * because `VbaExtractor.SQL_TABLE_RE` is `private static` and the
   * project's per-extractor state rule keeps each extractor's helpers
   * self-contained — see the file-level JSDoc on `VbaExtractor`.
   */
  private static readonly SQL_TABLE_RE =
    /\b(?:FROM|JOIN|INTO|UPDATE)\s+((?:(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*)\.)?(?:\[[^\]]+\]|\p{L}[\p{L}\p{N}_]*))/giu;

  /**
   * Issue #49 — classify a RecordSource/RowSource value as SQL.
   * Anything starting with a SQL keyword (`SELECT`, `PARAMETERS`, `WITH`,
   * `UPDATE`, `INSERT`, `DELETE`) — case-insensitive — is treated as a
   * SQL statement and run through `SQL_TABLE_RE`. Anything else is a
   * bare table-or-query name and emitted as a single reference.
   *
   * `WITH` is included because Access/JET supports CTE-style `WITH` queries
   * as record sources. `PARAMETERS` is included because Access uses
   * `PARAMETERS [foo] Text; SELECT … FROM …` to declare parameter types
   * before the body.
   */
  private static readonly SQL_PREFIX_RE =
    /^\s*(?:SELECT|PARAMETERS|WITH|UPDATE|INSERT|DELETE)\b/i;

  private isLikelySql(value: string): boolean {
    return VbaFormExtractor.SQL_PREFIX_RE.test(value);
  }

  /**
   * Issue #49 — emit one `references` edge from `sourceNodeId` to a
   * synthetic `class` placeholder node named `tableName`. The placeholder
   * node id is `(filePath, 'class', tableName, 0)` (line-independent, so
   * the same table referenced from N sites collapses to a single node);
   * the edge carries a per-site `line` so consumers can navigate to the
   * call site.
   *
   * The placeholder `class` kind mirrors `VbaExtractor.emitReference`'s
   * convention. The cross-file resolver may re-type these at lookup time
   * (a real `query` node takes precedence when both exist for the same
   * name — same dual-match `vba-sql-impact`'s `extractFormBindings` does).
   */
  private emitTableReference(
    sourceNodeId: string,
    tableName: string,
    lineNum: number,
    synthesizedBy: string,
  ): void {
    if (!tableName) return;
    const targetId = generateNodeId(this.filePath, 'class', tableName, 0);
    if (!this.synthClassNodeIds.has(targetId)) {
      this.synthClassNodeIds.add(targetId);
      this.nodes.push({
        id: targetId,
        kind: 'class',
        name: tableName,
        qualifiedName: tableName,
        filePath: this.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: 0,
        endColumn: tableName.length,
        updatedAt: Date.now(),
      });
    }
    this.edges.push({
      source: sourceNodeId,
      target: targetId,
      kind: 'references',
      provenance: 'heuristic',
      metadata: { synthesizedBy },
      line: lineNum,
      column: 0,
    });
  }

  /**
   * Issue #49 — dispatch the value of a RecordSource/RowSource binding.
   * If SQL, run `SQL_TABLE_RE` over the value and emit one edge per
   * distinct table (within-value dedup so `FROM tblA JOIN tblA` emits a
   * single edge). If a bare name, emit a single edge with the name as-is
   * — the resolver handles the dual-match against `query` and `class`
   * nodes at lookup time.
   *
   * Doubled-quote escapes (`""`) are unwarapped before classification;
   * the regex capture already includes them as part of the value, so a
   * single `.replace(/""/g, '"')` collapses them — same technique the
   * `extractStringLiterals` helper uses for its emitted `text` field.
   */
  private emitBinding(
    sourceNodeId: string,
    rawValue: string,
    lineNum: number,
    synthesizedBy: string,
  ): void {
    if (!rawValue) return;
    // Unwrap VBA doubled-quote escapes so a literal `""` inside the
    // value becomes a single `"` before we classify and parse.
    const value = rawValue.replace(/""/g, '"');
    if (this.isLikelySql(value)) {
      const seen = new Set<string>();
      for (const m of value.matchAll(VbaFormExtractor.SQL_TABLE_RE)) {
        const table = (m[1] ?? '').replace(/[\[\]]/g, '');
        if (!table) continue;
        if (seen.has(table)) continue;
        seen.add(table);
        this.emitTableReference(sourceNodeId, table, lineNum, synthesizedBy);
      }
      return;
    }
    // Bare name → single reference. Trim defensively so trailing
    // whitespace from the SaveAsText export doesn't pollute the target.
    const name = value.trim();
    if (!name) return;
    this.emitTableReference(sourceNodeId, name, lineNum, synthesizedBy);
  }

  /**
   * Issue #49 — sweep for the form-level `RecordSource` line. RecordSource
   * always attributes to the form-layout node (even if the line is
   * written inside a control's Begin block — defensive: in real
   * SaveAsText exports the form's RecordSource sits at the root, but
   * the agent's spec says "emit from the form-layout node" regardless).
   *
   * The sweep intentionally matches ONLY `RecordSource` here;
   * `RowSource` is handled by `sweepRowSources` so the per-binding
   * attribution logic (control vs. form-layout) stays in one place.
   */
  private sweepRecordSources(src: string, formLayoutNodeId: string): void {
    const lines = src.split('\n');
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const m = VbaFormExtractor.RECORD_SOURCE_RE.exec(line);
      if (!m) continue;
      const rawValue = m[1] ?? '';
      const lineNum = i + 1;
      this.emitBinding(formLayoutNodeId, rawValue, lineNum, 'vba-record-source');
    }
  }

  /**
   * Issue #49 — sweep for per-control `RowSource` lines. RowSource
   * attributes to the enclosing control's `form-instance-control` node
   * when one is in scope; if the line is found outside any control
   * Begin block (defensive — unusual in real exports), it falls back to
   * the form-layout node. The tag is always `'vba-row-source'` so
   * consumers can distinguish a control-level data binding from the
   * form-level RecordSource even when both happen to share the same
   * source node.
   *
   * Control-block tracking: a stack of `{ controlName, rowSourceType }`
   * entries mirrors the current scope as the sweep walks lines in order.
   * `Begin <Type>` pushes; `End` pops. `Begin Form` / `Begin Section`
   * push a non-control entry so a RowSource written at the form root
   * correctly falls through to the form-layout fallback.
   *
   * Value-list skip: when the CURRENT TOP scope's `RowSourceType` is
   * `"Value List"` (captured by the same scan-window as the control's
   * `Name`), the control's RowSource is a literal list and we skip the
   * emission — the data is in code, not a table.
   */
  private sweepRowSources(src: string, formLayoutNodeId: string): void {
    const lines = src.split('\n');
    type Scope = { controlName: string; rowSourceType: string };
    const stack: Scope[] = [];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const lineNum = i + 1;

      const beginM = VbaFormExtractor.BEGIN_RE.exec(line);
      if (beginM) {
        const controlType = beginM[1] ?? '';
        if (VbaFormExtractor.NON_CONTROL_TYPES.has(controlType)) {
          stack.push({ controlName: '', rowSourceType: '' });
        } else {
          // Reuse findControlName so the control-name attribution stays
          // in lockstep with the form-instance-control node id we
          // already produced in sweepControls. Issue #41's event-handler
          // edges and Issue #49's references edges then point at the
          // same id consistently.
          const { name: controlName } = this.findControlName(
            lines,
            i,
            lineNum,
          );
          const rowSourceType = this.findRowSourceType(lines, i);
          stack.push({ controlName, rowSourceType });
        }
        continue;
      }

      if (/^\s*End\s*$/i.test(line)) {
        if (stack.length > 0) stack.pop();
        continue;
      }

      const rowM = VbaFormExtractor.ROW_SOURCE_RE.exec(line);
      if (!rowM) continue;

      // Find the nearest enclosing control scope. We don't just look at
      // the stack top because a row could conceivably be written inside
      // a `Begin Section` (controlName='' entry on top) — in that case
      // the section entry hides the actual control below it. Walk the
      // stack from top to bottom and pick the first controlName we see.
      let currentControl = '';
      for (let j = stack.length - 1; j >= 0; j--) {
        if (stack[j]?.controlName) {
          currentControl = stack[j]!.controlName;
          break;
        }
      }

      // Skip value-list controls. The check is on the TOP scope's
      // rowSourceType, which is the most-recent control's type — even
      // if a nested Begin/End hid it from the stack-walk above, the
      // RowSource line we're matching was written inside that scope so
      // the top of stack carries the right value.
      const top = stack[stack.length - 1];
      if (top && top.rowSourceType === 'Value List') continue;

      const rawValue = rowM[1] ?? '';
      const sourceId = currentControl
        ? generateNodeId(
            this.filePath,
            'form-instance-control',
            currentControl,
            0,
          )
        : formLayoutNodeId;
      this.emitBinding(sourceId, rawValue, lineNum, 'vba-row-source');
    }
  }

  /**
   * Issue #49 — companion helper to `findControlName`: scan ahead from a
   * `Begin <Type>` line for the first `RowSourceType = "..."` attribute
   * within the same `NAME_SCAN_WINDOW` (so a missing attribute gracefully
   * no-ops instead of mis-attributing). Returns the captured value (e.g.
   * `"Value List"`, `"Table/Query"`) or `''` when absent.
   *
   * Stops at the next `Begin` or `End` boundary — same scan-window
   * discipline as `findControlName`, so a misaligned scan never crosses
   * into a sibling control's attribute block.
   */
  private findRowSourceType(
    lines: string[],
    beginLineIndex: number,
  ): string {
    const end = Math.min(
      lines.length,
      beginLineIndex + 1 + VbaFormExtractor.NAME_SCAN_WINDOW,
    );
    for (let j = beginLineIndex + 1; j < end; j++) {
      const line = lines[j] ?? '';
      if (/^\s*(Begin|End)\b/i.test(line)) break;
      const m = VbaFormExtractor.ROW_SOURCE_TYPE_RE.exec(line);
      if (m) {
        return m[1] ?? '';
      }
    }
    return '';
  }
}