/**
 * Procedure sweep: emit one `function` node per `Sub` / `Function` /
 * `Property Get|Let|Set` declaration, record each proc in `localProcs` for
 * same-file call resolution, and synthesize the Access `<Control>_<Event>`
 * event-handler edge for form/report code-behind classes (hueco 3 / issue #41).
 */
import * as path from 'path';
import { Node, Edge } from '../../types';
import { generateNodeId } from '../tree-sitter-helpers';
import { PROC_RE, PRIMITIVE_TYPES } from './constants';
import { parseEventHandlerName } from './text-utils';
import { ProcInfo, VbaClassifier } from './context';
import { defineRule, matchRule, VbaExtractionRule } from './rules';

/**
 * Parse a `Function`/`Property Get` declaration's return type — the `As
 * <Type>` that follows the parameter list (or the name, for a paren-less
 * declaration). Parameters can carry their own `As <Type>`, so we anchor on
 * the text AFTER the last `)` when parens are present; the bracketed form
 * `As [Type With Spaces]` is unwrapped. Returns null when there is no return
 * type (a `Sub`, or a function with an implicit `Variant` return).
 */
function parseReturnType(line: string): string | null {
  const afterParams = line.includes(')')
    ? line.slice(line.lastIndexOf(')') + 1)
    : line;
  const m = /\bAs\s+(?:New\s+)?(?:\[([^\]]+)\]|(\p{L}[\p{L}\p{N}_]*))/iu.exec(afterParams);
  if (!m) return null;
  return m[1] ?? m[2] ?? null;
}

function parseArrayParameters(line: string): string[] {
  const openIdx = line.indexOf('(');
  if (openIdx < 0) return [];
  const closeIdx = line.lastIndexOf(')');
  if (closeIdx < openIdx) return [];
  const body = line.slice(openIdx + 1, closeIdx);
  const params: string[] = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '(') depth++;
    else if (ch === ')') depth--;
    else if (ch === ',' && depth === 0) {
      params.push(body.slice(start, i));
      start = i + 1;
    }
  }
  params.push(body.slice(start));
  const arrayParamRe =
    /^\s*(?:ByRef|ByVal)\s+(\p{L}[\p{L}\p{N}_]*)\s*\(\s*\)/iu;
  const out: string[] = [];
  for (const raw of params) {
    const match = arrayParamRe.exec(raw);
    const name = match?.[1];
    if (name) out.push(name.toLowerCase());
  }
  return out;
}

/**
 * Issue #153: the declarative rule table for the procedures concern.
 * One rule — `procedure` — matches a `Sub` / `Function` /
 * `Property Get|Let|Set` declaration and emits the function node,
 * the `localProcs` / `functionNodeByName` / `functionNodeByStartLine`
 * / `functionReturnTypes` registration, and (for form code-behind)
 * the synthesized `event-handler` edge.
 *
 * The emit body is intentionally not split across multiple rules —
 * these are 5+ small operations that all fire on the same declaration
 * line. Splitting them would force the orchestrator to know they
 * always co-occur, which is more coupling than the original cascade
 * had. The pattern is the discriminating surface; everything past
 * `match` is one emit.
 */
export const RULES: readonly VbaExtractionRule<unknown>[] = [
  defineRule({
    id: 'procedure',
    description:
      'Match a `Sub` / `Function` / `Property Get|Let|Set <Name>(...) [As <Type>]` declaration; emit a function node, register the proc in `localProcs` / `functionNodeByName` / `functionNodeByStartLine` / `functionReturnTypes`, and (for `Form_*.cls` / `Report_*.cls`) synthesize a `form-instance-control` stub + `event-handler` edge.',
    pattern: PROC_RE,
    emit: (m, ctx, line, lineNum) => {
      const visibilityRaw = (m[1] ?? '').trim();
      const kindRaw = (m[2] ?? '').trim().toLowerCase();
      const name = m[3] ?? '';
      if (!name) return null;

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

      // Factory-return inference: record a function's project-class return type
      // so the call sweep can type `Set x = <name>(...)`. Restricted to `Sub`'s
      // sibling `Function` (a `Property Let/Set` has no return type and a
      // `Property Get`'s `As <Type>` is rarely a factory target). Primitives are
      // skipped — `x.Method` on a primitive is never a project call.
      if (kind === 'function') {
        const retType = parseReturnType(line);
        if (retType && !PRIMITIVE_TYPES.has(retType.toLowerCase())) {
          const key = name.toLowerCase();
          if (!ctx.functionReturnTypes.has(key)) {
            ctx.functionReturnTypes.set(key, retType);
          }
        }
      }

      const proc: ProcInfo = {
        name,
        // B3 (hueco 5): when this file is a `.cls`, prefix the
        // qualifiedName with the resolved class name so cross-class
        // queries (e.g. `Form_Load`) match only the owning class.
        // `.bas` files leave `qualifiedName === name`.
        qualifiedName: ctx.classNamePrefix
          ? `${ctx.classNamePrefix}.${name}`
          : name,
        kind,
        visibility,
        startLine: lineNum,
        // Issue #190: capture parameter-array names so the call sweep can
        // suppress `name(index)` accesses inside this procedure body
        // without polluting the file-global `localVarTypeMap`.
        arrayParameters: parseArrayParameters(line),
      };
      ctx.procedures.push(proc);
      const bucket = ctx.localProcs.get(name);
      if (bucket) bucket.push(proc);
      else ctx.localProcs.set(name, [proc]);

      const nodeId = generateNodeId(ctx.filePath, 'function', name, lineNum);
      const fnNode: Node = {
        id: nodeId,
        kind: 'function',
        name,
        // B3 (hueco 5): same prefix rule as the ProcInfo above — class
        // methods get `${className}.${name}`, module-level Subs keep
        // their bare-name qualifiedName.
        qualifiedName: ctx.classNamePrefix
          ? `${ctx.classNamePrefix}.${name}`
          : name,
        filePath: ctx.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: 0,
        endColumn: line.length,
        visibility,
        updatedAt: Date.now(),
      };
      ctx.nodes.push(fnNode);
      // Cache the first node emitted for this name — `findFunctionNodeByName`
      // (audit S2) becomes O(1) instead of O(n) per call site.
      if (!ctx.functionNodeByName.has(name)) {
        ctx.functionNodeByName.set(name, fnNode);
      }
      // Fix 1: also index by startLine so Property Get/Let/Set with the same
      // name can each be found by their exact declaration line.
      ctx.functionNodeByStartLine.set(lineNum, fnNode);

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
      // Prefix-driven sibling binding (issue #41). Both `Form_*.cls` and
      // `Report_*.cls` Dysflow code-behind files share the same code path;
      // only the sibling extension differs (`.form.txt` vs `.report.txt`).
      // The check is on the BASENAME prefix so a class called
      // `FormularioVentas.cls` or `ReportingHelper.cls` (no trailing
      // underscore) does not match — the trailing `_` is the discriminator.
      // Any other `.cls` (e.g. `InformeRiesgoPDFServicio.cls` with methods
      // like `GenerarHTML_Principal`) gets `codeBehindExt === null` and is
      // skipped, preserving the original Form_-only guard's behaviour for
      // non-form classes.
      const basename = path.basename(ctx.filePath).toLowerCase();
      const codeBehindExt = basename.startsWith('report_')
        ? '.report.txt'
        : basename.startsWith('form_')
          ? '.form.txt'
          : null;
      const isFormCodeBehind = codeBehindExt !== null;
      if (handler && isFormCodeBehind) {
        const siblingPath = ctx.filePath.replace(/\.cls$/i, codeBehindExt!);
        const controlNodeId = generateNodeId(
          siblingPath,
          'form-instance-control',
          handler.controlName,
          0,
        );
        // Stub form-instance-control: local so the per-file edge filter
        // passes the event-handler edge. Overwritten by the real node
        // emitted from the sibling .form.txt (or .report.txt) at index time
        // (same id, same schema, INSERT OR REPLACE). No metadata.controlType
        // here — the sibling side carries the real control type.
        ctx.nodes.push({
          id: controlNodeId,
          kind: 'form-instance-control',
          name: handler.controlName,
          qualifiedName: `${siblingPath}::${handler.controlName}`,
          filePath: siblingPath,
          language: 'vba',
          startLine: 0,
          endLine: 0,
          startColumn: 0,
          endColumn: 0,
          updatedAt: Date.now(),
        });
        ctx.edges.push({
          source: nodeId,
          target: controlNodeId,
          kind: 'event-handler',
          provenance: 'heuristic',
          metadata: { eventName: handler.eventName },
          line: lineNum,
          column: 0,
        });
      }

      if (ctx.moduleOrClassNode) {
        ctx.edges.push({
          source: ctx.moduleOrClassNode.id,
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
        ctx.edges.push(edge);
        ctx.pendingModuleOrClassSource.push(edge);
      }
      return { name };
    },
  }),
];

export function createProceduresClassifier(): VbaClassifier {
  return {
    name: 'procedures',
    count: 0,
    classifyLine(line, i, ctx) {
      const lineNum = i + 1;
      for (const rule of RULES) {
        const m = matchRule(rule.pattern, line);
        if (!m) continue;
        const result = rule.emit(m, ctx, line, lineNum);
        if (result !== null && result !== undefined) {
          this.count += rule.count ? rule.count(result as never) : 1;
        }
      }
    },
  };
}
