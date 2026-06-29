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
 *  - It emits one `module` node per file (named per `Attribute VB_Name`
 *    when present, otherwise the file basename).
 *  - It emits one `property` node per Access control declaration with
 *    `metadata.controlType` set to the control type (e.g. `'TextBox'`,
 *    `'CommandButton'`).
 *  - It emits one `references` edge from the form module to a node whose
 *    name matches the sibling `.cls` basename, so `codegraph_explore` can
 *    resolve form → class at lookup time. Tagged
 *    `metadata.synthesizedBy: 'vba-form-binding'`.
 *
 * Hard invariant (REQ-FORM-4): ZERO `function`/`sub`/`class` nodes from
 * `.form.txt` / `.report.txt`. The form-side behavior lives in
 * `vba-form-ui-extraction`; the code-side lives in `vba-code-extraction`.
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

      // Form/report module node — NEVER a `class` node.
      const moduleNode = this.createModuleNode(moduleName);
      this.nodes.push(moduleNode);

      // Sibling `.cls` binding (REQ-FORM-1, REQ-FORM-3). Emit an
      // UnresolvedReference rather than a hardcoded `references` edge with
      // a synthetic target node: REQ-FORM-4 forbids ANY `class` node from
      // form files, and a hardcoded target would have to be a `class` kind
      // for downstream resolution to match it. The unresolved reference is
      // resolved at index time when the actual sibling `.cls` is processed.
      this.unresolvedReferences.push({
        fromNodeId: moduleNode.id,
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

  private createModuleNode(name: string): Node {
    const lines = this.source.split('\n');
    return {
      id: generateNodeId(this.filePath, 'module', name, 1),
      kind: 'module',
      name,
      qualifiedName: name,
      filePath: this.filePath,
      language: 'vba',
      startLine: 1,
      endLine: lines.length,
      startColumn: 0,
      endColumn: 0,
      metadata: {},
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
    }
  }
}