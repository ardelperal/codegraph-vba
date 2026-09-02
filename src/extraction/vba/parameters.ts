/**
 * Issue #257 (task T16, half A): turn a procedure's already-parsed signature
 * into `parameter` nodes.
 *
 * This module is EMISSION ONLY. The signature was parsed once by
 * `signature.ts` — after the preprocessor joined `_` line continuations — and
 * stamped onto the `function` node's `metadata.params` by `procedures.ts`.
 * The same `VbaParameterInfo[]` is handed here, so a continued signature and
 * its one-line equivalent produce byte-identical nodes by construction. Never
 * re-parse the declaration line here; that is how the two views drift.
 */
import { Node } from '../../types';
import { generateNodeId } from '../tree-sitter-helpers';
import { PRIMITIVE_TYPES } from './constants';
import { VbaExtractorContext } from './context';
import type { VbaParameterInfo } from './signature';

/**
 * The type name a `type_of` edge should point at, or `null` when the
 * declaration deserves no edge at all.
 *
 * This mirrors what `dims.ts` does for `Dim x As Foo`, deliberately:
 *
 *  - an untyped parameter (implicit `Variant`) yields nothing;
 *  - a primitive (`Long`, `String`, `Variant`, …) yields nothing —
 *    `PRIMITIVE_TYPES` is the same gate the Dim sweep uses;
 *  - a qualified type (`As DAO.Recordset`) yields its OUTER segment, because
 *    that is the name the Dim sweep resolves and the only one a project
 *    declaration could ever match.
 *
 * The gate is intentionally coarse: `type_of` is emitted for anything that is
 * not a primitive, and `resolveVbaReferenceStubs` decides afterwards whether a
 * real project class of that name exists. A library type nobody declares keeps
 * its stub, exactly as a `Dim r As DAO.Recordset` does today.
 */
function typeEdgeTarget(declaredType: string | null): string | null {
  if (!declaredType) return null;
  const outer = declaredType.split('.')[0] ?? '';
  if (!outer) return null;
  if (PRIMITIVE_TYPES.has(outer.toLowerCase())) return null;
  return outer;
}

/**
 * Emit one `parameter` node per declared parameter of `procedureNode`, the
 * `contains` edge that binds it to its procedure, and — for a non-primitive
 * declared type — a `type_of` edge onto the type's node.
 *
 * Parameters have no line of their own: VBA writes them inside the procedure
 * header, and after continuation joining even the physical line is a fiction.
 * So `startLine` / `endLine` are the procedure's declaration line, and the id
 * is disambiguated by `<Procedure>.<param>` rather than by position — two
 * accessors of the same property (`Property Get X` / `Property Let X`) sit on
 * different lines and therefore still get distinct ids.
 *
 * `qualifiedName` is `<ModuleOrClass>.<Procedure>.<param>`, using
 * `ctx.moduleName` (the resolved `Attribute VB_Name`, or the basename) rather
 * than `ctx.classNamePrefix` — the latter is `null` for a `.bas`, and a
 * parameter of a module-level Sub wants the module prefix just as much as a
 * class method's does. Same choice the module-level `variable` nodes made in
 * issue #251.
 */
export function emitParameterNodes(
  ctx: VbaExtractorContext,
  procedureNode: Node,
  params: readonly VbaParameterInfo[],
  lineNum: number,
  line: string,
): void {
  if (params.length === 0) return;
  const procedureName = procedureNode.name;
  const emitted = new Set<string>();

  for (let position = 0; position < params.length; position++) {
    const param = params[position]!;
    const name = param.name.trim();
    // `parseSignature` never drops a fragment it cannot understand — a wrong
    // arity is worse than a coarse one — so a nameless entry can reach here.
    // It is not a symbol, so it gets no node.
    if (!name) continue;

    const nodeId = generateNodeId(
      ctx.filePath,
      'parameter',
      `${procedureName}.${name}`,
      lineNum,
    );
    // VBA forbids two parameters of the same name on one signature; this only
    // fires if a malformed declaration parsed into a repeat. Keep the first.
    if (emitted.has(nodeId)) continue;
    emitted.add(nodeId);

    const qualifiedPrefix = ctx.moduleName
      ? `${ctx.moduleName}.${procedureName}`
      : procedureName;

    ctx.nodes.push({
      id: nodeId,
      kind: 'parameter',
      name,
      qualifiedName: `${qualifiedPrefix}.${name}`,
      filePath: ctx.filePath,
      language: 'vba',
      startLine: lineNum,
      endLine: lineNum,
      startColumn: 0,
      endColumn: line.length,
      metadata: {
        position,
        byRef: param.byRef,
        optional: param.optional,
        isArray: param.isArray,
        hasDefault: param.hasDefault,
        // An `As`-less parameter is `Variant` per VBA semantics. Recorded as
        // the language's own default rather than left null, matching how
        // issue #251 stamps `declaredType` on module-level variables.
        declaredType: param.type ?? 'Variant',
      },
      updatedAt: Date.now(),
    });

    ctx.edges.push({
      source: procedureNode.id,
      target: nodeId,
      kind: 'contains',
    });

    const typeName = typeEdgeTarget(param.type);
    if (typeName) {
      ctx.emitTypeOf(nodeId, typeName, lineNum, 0, 'vba-parameter-type');
    }
  }
}
