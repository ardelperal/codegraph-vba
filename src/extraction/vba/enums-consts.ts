/**
 * Enum / Const sweep (REQ-CODE-12, REQ-CODE-13). Emits `enum` + `enum_member`
 * nodes and module-level `constant` nodes, while tracking proc scope so
 * proc-local Consts populate the per-scope resolution bucket (issue #52)
 * without emitting a spurious module-level `constant` node.
 */
import { generateNodeId } from '../tree-sitter-helpers';
import { PROC_RE, PROCEDURE_END_RE } from './constants';
import { foldVisibility, parseConstDeclarations } from './text-utils';
import { VbaExtractorContext } from './context';

/** `[visibility] Enum <Name>` — opens an enum block. */
const ENUM_START_RE =
  /^\s*(?:(Public|Private|Friend|Global)\s+)?Enum\s+(\p{L}[\p{L}\p{N}_]*)/iu;

/** `End Enum` — closes the current enum block. */
const ENUM_END_RE = /^\s*End\s+Enum\b/i;

/**
 * An enum member line: a leading identifier optionally followed by `=
 * <value>`. Runs only inside an open Enum block, on the (already
 * comment-stripped) source, so a trailing `'comment` never reaches here.
 * `\p{L}` covers accented member names (e.g. `Sí`).
 */
const ENUM_MEMBER_RE = /^\s*(\p{L}[\p{L}\p{N}_]*)\s*(?:=|$)/u;

/** `[visibility] Const <decls>` — captures visibility (1) and the rest (2). */
const CONST_DECL_RE =
  /^\s*(?:(Public|Private|Friend|Global)\s+)?Const\s+(.+)$/i;

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
export function sweepEnumsAndConsts(ctx: VbaExtractorContext, src: string): number {
  const lines = src.split('\n');
  let count = 0;
  let currentEnum: { id: string; name: string } | null = null;

  // Issue #52: reset the shared scope stack + lookup key so leftover
  // state from a previous extract() (impossible in production but
  // possible in unit tests that construct a fresh extractor and run
  // twice) never leaks across sweeps. The walk below updates both
  // every iteration; `sweepCallsAndSql` resets again at its own
  // start, before any OpenForm/OpenQuery reader consults them.
  ctx.procStack.length = 0;
  ctx.currentProcKey = 'module';

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? '';
    const lineNum = i + 1;

    // Issue #52: track proc scope so Const declarations on this line
    // can decide whether they belong to the module (currentProcKey
    // === 'module') or to the top-most procedure (write the per-proc
    // bucket, skip the module-level `constant` node emission).
    //
    // PROC_RE cannot overlap with CONST_DECL_RE on the same physical
    // line (different leading keywords), so it is safe to advance the
    // stack here and then fall through to the rest of the body.
    const procStart = PROC_RE.exec(line);
    if (procStart) {
      ctx.procStack.push(lineNum);
      ctx.currentProcKey = String(lineNum);
    } else if (PROCEDURE_END_RE.test(line) && ctx.procStack.length > 0) {
      ctx.procStack.pop();
      ctx.currentProcKey =
        ctx.procStack.length > 0
          ? String(ctx.procStack[ctx.procStack.length - 1])
          : 'module';
    }

    if (currentEnum) {
      if (ENUM_END_RE.test(line)) {
        currentEnum = null;
        continue;
      }
      const mm = ENUM_MEMBER_RE.exec(line);
      if (mm) {
        const memberName = mm[1] ?? '';
        if (memberName) {
          const memberId = generateNodeId(
            ctx.filePath,
            'enum_member',
            memberName,
            lineNum,
          );
          ctx.nodes.push({
            id: memberId,
            kind: 'enum_member',
            name: memberName,
            qualifiedName: `${currentEnum.name}.${memberName}`,
            filePath: ctx.filePath,
            language: 'vba',
            startLine: lineNum,
            endLine: lineNum,
            startColumn: 0,
            endColumn: line.length,
            updatedAt: Date.now(),
          });
          // enum → member: source is known, emit directly (not pending).
          ctx.edges.push({
            source: currentEnum.id,
            target: memberId,
            kind: 'contains',
          });
        }
      }
      continue;
    }

    const enumStart = ENUM_START_RE.exec(line);
    if (enumStart) {
      const visibility = foldVisibility(enumStart[1] ?? '');
      const name = enumStart[2] ?? '';
      if (!name) continue;
      const enumId = generateNodeId(ctx.filePath, 'enum', name, lineNum);
      ctx.nodes.push({
        id: enumId,
        kind: 'enum',
        name,
        qualifiedName: name,
        filePath: ctx.filePath,
        language: 'vba',
        startLine: lineNum,
        endLine: lineNum,
        startColumn: 0,
        endColumn: line.length,
        visibility,
        updatedAt: Date.now(),
      });
      ctx.pushContainsFromModule(enumId);
      currentEnum = { id: enumId, name };
      count++;
      continue;
    }

    const constDecl = CONST_DECL_RE.exec(line);
    if (constDecl) {
      const visibility = foldVisibility(constDecl[1] ?? '');
      const body = constDecl[2] ?? '';
      const declarations = parseConstDeclarations(body);
      for (const declaration of declarations) {
        const constName = declaration.name;
        if (!constName) continue;
        // Issue #52: every Const line (module-level or proc-local)
        // writes into a per-scope resolution bucket so
        // `DoCmd.OpenForm FORM_X` later resolves correctly; module-level
        // Consts additionally emit a `constant` graph node + the
        // module→constant `contains` edge. Proc-local Consts skip both
        // (the const is not a module symbol, so the wrong-containment
        // node + edge the pre-fix code emitted are gone), but the
        // per-proc bucket keeps OpenForm/OpenQuery argument
        // resolution working exactly as before.
        if (declaration.value !== null) {
          ctx.setLocalConstInScope(ctx.currentProcKey, constName, declaration.value);
        }
        if (ctx.procStack.length > 0) continue;
        const constId = generateNodeId(
          ctx.filePath,
          'constant',
          constName,
          lineNum,
        );
        ctx.nodes.push({
          id: constId,
          kind: 'constant',
          name: constName,
          qualifiedName: constName,
          filePath: ctx.filePath,
          language: 'vba',
          startLine: lineNum,
          endLine: lineNum,
          startColumn: 0,
          endColumn: line.length,
          visibility,
          metadata: declaration.value !== null ? { value: declaration.value } : undefined,
          updatedAt: Date.now(),
        });
        ctx.pushContainsFromModule(constId);
        count++;
      }
    }
  }
  return count;
}
