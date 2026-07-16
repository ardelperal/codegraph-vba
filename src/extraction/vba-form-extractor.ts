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
 *  - It emits one layout node per file (named per `Attribute VB_Name` when
 *    present, otherwise the file basename): `form-layout` for forms and
 *    `report-layout` for reports. `module` remains the kind for `.bas`
 *    standard modules emitted by `VbaExtractor`.
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
import { ACCESS_EVENT_PROPERTIES } from './vba/events';

interface FormBlockFrame {
  controlType: string;
  beginLine: number;
  beginColumnLength: number;
  name: string;
  nameLine: number;
  section?: string;
  properties: Map<string, { value: string; line: number }>;
}

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

      // A SaveAsText document is a recursive Begin/End block tree. Walk it
      // once so names, bindings, and section membership share one scope model.
      this.walkBlocks(cleaned, formLayoutNode.id);
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
   * Emits `form-layout` for forms and `report-layout` for reports, replacing
   * the historical `module` kind with a UI-specific layout kind.
   *
   * The deterministic id formula remains
   * `generateNodeId(filePath, layoutKind, name, 1)`. Report ids intentionally
   * use `report-layout` so real reports agree with DoCmd.OpenReport stubs.
   * `metadata.containerKind` keeps the historical `module` label as a
   * back-compat marker; consumers should prefer `node.kind`.
   */
  private createFormLayoutNode(name: string): Node {
    const lines = this.source.split('\n');
    const kind = /\.report\.txt$/i.test(this.filePath)
      ? 'report-layout'
      : 'form-layout';
    return {
      id: generateNodeId(this.filePath, kind, name, 1),
      kind,
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
   * The block walker balances every `Begin`/`End` pair, including untyped
   * root blocks and non-control Form/Report/Section containers.
   */
  private static readonly BEGIN_RE =
    /^\s*Begin(?:\s+(.+?))?\s*$/i;

  /** Access also serializes binary/GUID values as `Property = Begin ... End`. */
  private static readonly PROPERTY_BLOCK_BEGIN_RE =
    /^\s*\p{L}[\p{L}\p{N}_]*\s*=\s*Begin\s*$/iu;

  /** Quoted SaveAsText property captured inside the current block frame. */
  private static readonly QUOTED_PROPERTY_RE =
    /^\s*(\p{L}[\p{L}\p{N}_]*)\s*=\s*"((?:[^"]|"")*)"/u;

  private walkBlocks(src: string, formLayoutNodeId: string): void {
    const lines = src.split('\n');
    const stack: FormBlockFrame[] = [];
    let recordSource: string | undefined;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      const lineNum = i + 1;
      const begin = VbaFormExtractor.BEGIN_RE.exec(line);
      if (begin) {
        const rawType = (begin[1] ?? '').trim();
        const controlType = /^\p{L}[\p{L}\p{N}_]*$/u.test(rawType)
          ? rawType
          : '';
        stack.push({
          controlType,
          beginLine: lineNum,
          beginColumnLength: line.length,
          name: '',
          nameLine: lineNum,
          section: this.enclosingSection(stack),
          properties: new Map(),
        });
        continue;
      }

      if (VbaFormExtractor.PROPERTY_BLOCK_BEGIN_RE.test(line)) {
        stack.push({
          controlType: '',
          beginLine: lineNum,
          beginColumnLength: line.length,
          name: '',
          nameLine: lineNum,
          section: this.enclosingSection(stack),
          properties: new Map(),
        });
        continue;
      }

      if (/^\s*End\s*$/i.test(line)) {
        const frame = stack.pop();
        if (frame) this.emitBlock(frame, formLayoutNodeId, recordSource);
        continue;
      }

      const property = VbaFormExtractor.QUOTED_PROPERTY_RE.exec(line);
      if (!property) continue;
      const key = (property[1] ?? '').toLowerCase();
      const value = (property[2] ?? '').replace(/""/g, '"');
      const frame = stack[stack.length - 1];
      if (!frame) {
        if (key === 'recordsource') {
          this.emitBinding(
            formLayoutNodeId,
            value,
            lineNum,
            'vba-record-source',
          );
        } else if (key === 'rowsource') {
          this.emitBinding(
            formLayoutNodeId,
            value,
            lineNum,
            'vba-row-source',
          );
        }
        this.emitExpressionHandler(formLayoutNodeId, key, value, lineNum);
        continue;
      }

      frame.properties.set(key, { value, line: lineNum });
      if (key === 'name') {
        frame.name = value;
        frame.nameLine = lineNum;
      }

      // RecordSource is a layout-level binding even though SaveAsText places
      // it inside the root Form/Report block.
      if (key === 'recordsource') {
        recordSource = value;
        this.emitBinding(
          formLayoutNodeId,
          value,
          lineNum,
          'vba-record-source',
        );
      }
    }

    // Malformed/truncated exports still yield the facts from every complete
    // frame accumulated so far, matching the previous tolerant extractor.
    while (stack.length > 0) {
      this.emitBlock(stack.pop()!, formLayoutNodeId, recordSource);
    }
  }

  private enclosingSection(stack: readonly FormBlockFrame[]): string | undefined {
    for (let i = stack.length - 1; i >= 0; i--) {
      const frame = stack[i];
      if (frame?.controlType.toLowerCase() === 'section' && frame.name) {
        return frame.name;
      }
    }
    return undefined;
  }

  private emitBlock(
    frame: FormBlockFrame,
    formLayoutNodeId: string,
    recordSource: string | undefined,
  ): void {
    const { controlType, beginLine: lineNum } = frame;
    if (!controlType) {
      const rowSource = frame.properties.get('rowsource');
      if (rowSource) {
        this.emitBinding(
          formLayoutNodeId,
          rowSource.value,
          rowSource.line,
          'vba-row-source',
        );
      }
      this.emitExpressionHandlers(formLayoutNodeId, frame.properties);
      return;
    }

    if (controlType.toLowerCase() === 'section') return;
    if (/^(form|report)$/i.test(controlType)) {
      const rowSource = frame.properties.get('rowsource');
      if (rowSource) {
        this.emitBinding(
          formLayoutNodeId,
          rowSource.value,
          rowSource.line,
          'vba-row-source',
        );
      }
      this.emitExpressionHandlers(formLayoutNodeId, frame.properties);
      return;
    }

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
        endColumn: frame.beginColumnLength,
        metadata: { controlType },
        updatedAt: Date.now(),
      });

      // ---- Hueco 2: emit a `form-instance-control` node per NAME. -------
      // The block-scoped walk records the Name property at any distance.
      // The control's `Name` (e.g. "lblTitulo",
      // "ComandoAltaPM") is what the .cls side references via
      // `Me.<ControlName>` (hueco 1) and what event handlers are wired to
      // via the `<ControlName>_<Event>` naming convention (hueco 3).
      // line=0 in the generated id keeps the id STABLE across re-indexes
      // of the same control — the VbaExtractor side synthesizes the
      // matching event-handler edge using the same id formula (see
      // vba-extractor.ts: synthesizeEventHandlerEdge).
      const controlName = frame.name;
      if (!controlName) return;

      const controlNodeId = generateNodeId(
        this.filePath,
        'form-instance-control',
        controlName,
        0,
      );
      const controlSource = frame.properties.get('controlsource');
      const sourceObject = frame.properties.get('sourceobject');
      this.nodes.push({
        id: controlNodeId,
        kind: 'form-instance-control',
        name: controlName,
        qualifiedName: `${this.filePath}::${controlName}`,
        filePath: this.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: frame.nameLine, // spans from Begin to the Name attribute line
        startColumn: 0,
        endColumn: 0,
        metadata: {
          controlType,
          ...(frame.section ? { section: frame.section } : {}),
          ...(controlSource ? { controlSource: controlSource.value } : {}),
          ...(sourceObject ? { sourceObject: sourceObject.value } : {}),
        },
        updatedAt: Date.now(),
      });

      this.edges.push({
        source: formLayoutNodeId,
        target: controlNodeId,
        kind: 'contains',
        provenance: 'parser',
      });

      this.emitExpressionHandlers(controlNodeId, frame.properties);

      if (controlSource) {
        this.emitControlSourceReference(
          controlNodeId,
          controlSource.value,
          controlSource.line,
          recordSource,
        );
      }

      if (sourceObject) {
        this.emitSourceObjectReference(
          controlNodeId,
          sourceObject.value,
          sourceObject.line,
        );
      }

      const rowSource = frame.properties.get('rowsource');
      const rowSourceType = frame.properties.get('rowsourcetype');
      if (
        rowSource &&
        rowSourceType?.value.toLowerCase() !== 'value list'
      ) {
        this.emitBinding(
          controlNodeId,
          rowSource.value,
          rowSource.line,
          'vba-row-source',
        );
      }
  }

  private emitExpressionHandlers(
    wiringSiteNodeId: string,
    properties: ReadonlyMap<string, { value: string; line: number }>,
  ): void {
    for (const [propertyName, property] of properties) {
      this.emitExpressionHandler(
        wiringSiteNodeId,
        propertyName,
        property.value,
        property.line,
      );
    }
  }

  private emitExpressionHandler(
    wiringSiteNodeId: string,
    propertyName: string,
    rawValue: string,
    lineNum: number,
  ): void {
    const eventName = ACCESS_EVENT_PROPERTIES.get(propertyName.toLowerCase());
    if (!eventName) return;

    // `[Event Procedure]` is handled by the existing code-behind naming path.
    // Bare values name Access macros, which are not graph nodes; silent beats
    // inventing a function edge for either form.
    const expression = rawValue.trim();
    const match = /^=\s*([\p{L}_][\p{L}\p{N}_]*)\s*\(/u.exec(expression);
    if (!match) return;
    let depth = 0;
    let quoted = false;
    let completeAt = -1;
    for (let i = expression.indexOf('(', match.index); i < expression.length; i++) {
      const char = expression[i];
      if (char === '"') quoted = !quoted;
      if (quoted) continue;
      if (char === '(') depth++;
      if (char === ')' && --depth === 0) { completeAt = i; break; }
    }
    if (completeAt < 0 || expression.slice(completeAt + 1).trim() !== '') return;

    this.unresolvedReferences.push({
      fromNodeId: wiringSiteNodeId,
      referenceName: match[1]!,
      referenceKind: 'event-handler',
      line: lineNum,
      column: 0,
      filePath: this.filePath,
      language: 'vba',
      metadata: {
        eventName,
        synthesizedBy: 'vba-expression-handler',
      },
    });
  }

  private emitSourceObjectReference(
    controlNodeId: string,
    rawValue: string,
    lineNum: number,
  ): void {
    const match = /^(?:(Form|Report|Table|Query)\.)?(.*)$/i.exec(rawValue.trim());
    if (!match) return;
    const prefix = match[1]?.toLowerCase();
    const target = match[2]?.trim() ?? '';
    if (!target) return;

    if (prefix === 'table' || prefix === 'query') {
      this.emitTableReference(
        controlNodeId,
        target,
        lineNum,
        'vba-source-object',
        { sourceObjectType: prefix },
      );
      return;
    }

    this.unresolvedReferences.push({
      fromNodeId: controlNodeId,
      referenceName: target,
      referenceKind: 'references',
      line: lineNum,
      column: 0,
      filePath: this.filePath,
      language: 'vba',
      metadata: {
        synthesizedBy: 'vba-source-object',
        embeds: true,
        accessObjectKind: prefix === 'report' ? 'report' : 'form',
      },
    });
  }

  // ---------------------------------------------------------------------------
  // RecordSource / RowSource edge emission (Issue #49)
  // ---------------------------------------------------------------------------

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
    extraMetadata: Record<string, unknown> = {},
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
      // RecordSource/RowSource are data-display bindings — a form/report or a
      // list/combo control READS its source. Tagging `access: 'read'` keeps
      // the metadata.access field uniformly present across every SQL-derived
      // table reference (the in-code SQL sweep classifies read vs write from
      // the statement verb; a binding is always a read).
      metadata: { synthesizedBy, access: 'read', ...extraMetadata },
      line: lineNum,
      column: 0,
    });
  }

  /**
   * Link a bound control to its enclosing form/report's single bare table.
   * Expressions and SQL/absent RecordSource values stay metadata-only: column
   * lineage through expressions or SELECT projections cannot be inferred here
   * without risking false graph edges.
   */
  private emitControlSourceReference(
    controlNodeId: string,
    controlSource: string,
    lineNum: number,
    recordSource: string | undefined,
  ): void {
    const rawField = controlSource.trim();
    if (!rawField || rawField.startsWith('=') || !recordSource) return;

    const bracketedField = /^\[([^\]]+)\]$/.exec(rawField);
    const field = bracketedField?.[1] ?? rawField;
    if (!bracketedField && !/^\p{L}[\p{L}\p{N}_]*$/u.test(rawField)) return;

    const source = recordSource.trim();
    if (!source || this.isLikelySql(source)) return;
    const bracketedSource = /^\[([^\]]+)\]$/.exec(source);
    const tableName = bracketedSource?.[1] ?? source;
    if (!bracketedSource && !/^\p{L}[\p{L}\p{N}_]*$/u.test(source)) return;

    this.emitTableReference(
      controlNodeId,
      tableName,
      lineNum,
      'vba-control-source',
      { column: field },
    );
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
}
