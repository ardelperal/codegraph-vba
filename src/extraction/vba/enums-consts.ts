/**
 * Enum / Const sweep (REQ-CODE-12, REQ-CODE-13). Emits `enum` + `enum_member`
 * nodes and module-level `constant` nodes, while tracking proc scope so
 * proc-local Consts populate the per-scope resolution bucket (issue #52)
 * without emitting a spurious module-level `constant` node.
 */
import { generateNodeId } from '../tree-sitter-helpers';
import { PROC_RE, PROCEDURE_END_RE } from './constants';
import { foldVisibility, parseConstDeclarations } from './text-utils';
import { VbaExtractorContext, VbaClassifier } from './context';
import { defineRule, matchRule, VbaExtractionRule } from './rules';

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
 * Issue #153: the declarative rule table for the enum / const concern.
 * Six rules, partitioned by where they fire:
 *
 *  - `proc-start`      (outside-enum-block) — match a `Sub` /
 *                      `Function` / `Property` header; push onto
 *                      `ctx.procStack` and set `ctx.currentProcKey`
 *                      so subsequent Const writes land in the
 *                      per-proc resolution bucket.
 *  - `proc-end`        (outside-enum-block) — match `End Sub` /
 *                      `End Function` / `End Property`; pop the
 *                      matching `procStack` entry.
 *  - `enum-start`      (outside-enum-block) — match `[visibility]
 *                      Enum <Name>`; emit an `enum` node + `contains`
 *                      edge and mark `ctx.vbaEnumBlock`.
 *  - `enum-end`        (inside-enum-block)  — match `End Enum`; clear
 *                      `ctx.vbaEnumBlock`.
 *  - `enum-member`     (inside-enum-block)  — match `<MemberName>
 *                      [= <value>]`; emit an `enum_member` node + a
 *                      `contains` edge from the open enum.
 *  - `const-decl`      (outside-enum-block) — match `[visibility]
 *                      Const <decls>`; always write the per-scope
 *                      resolution bucket, and emit a `constant` node
 *                      + `contains` edge ONLY when not inside a proc
 *                      (proc-local Consts are not module symbols).
 *
 * The inter-line `ctx.procStack` / `ctx.currentProcKey` state is
 * SHARED with the calls/SQL classifier (issue #52 protocol). The
 * `ctx.vbaEnumBlock` state is local to this concern and lives on
 * `ctx` for the same RULES-table-friendliness reason as
 * `vbaDeclTypeBlock`.
 */
export const RULES: readonly VbaExtractionRule<unknown>[] = [
  defineRule({
    id: 'proc-start',
    description:
      'Match a `Sub` / `Function` / `Property` header; push `lineNum` onto `ctx.procStack` and set `ctx.currentProcKey` so subsequent Const writes land in the per-proc resolution bucket.',
    pattern: PROC_RE,
    requires: 'outside-enum-block',
    emit: (_m, ctx, _line, lineNum) => {
      ctx.procStack.push(lineNum);
      ctx.currentProcKey = String(lineNum);
      return { startLine: lineNum };
    },
  }),
  defineRule({
    id: 'proc-end',
    description:
      'Match an `End Sub` / `End Function` / `End Property`; pop the matching `procStack` entry and rewind `ctx.currentProcKey` to the new top (or `module`).',
    pattern: PROCEDURE_END_RE,
    requires: 'outside-enum-block',
    emit: (_m, ctx) => {
      if (ctx.procStack.length > 0) ctx.procStack.pop();
      ctx.currentProcKey =
        ctx.procStack.length > 0
          ? String(ctx.procStack[ctx.procStack.length - 1])
          : 'module';
      return { kind: 'proc-end' as const };
    },
  }),
  defineRule({
    id: 'enum-start',
    description:
      'Match a `[Public|Private|Friend|Global] Enum <Name>` block header; emit an `enum` node + `contains` edge and mark `ctx.vbaEnumBlock` so subsequent member lines route to `enum-member` / `End Enum` lines route to `enum-end`.',
    pattern: ENUM_START_RE,
    requires: 'outside-enum-block',
    emit: (m, ctx, line, lineNum) => {
      const visibility = foldVisibility(m[1] ?? '');
      const name = m[2] ?? '';
      if (!name) return null;
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
      ctx.vbaEnumBlock = { id: enumId, name };
      return { name, enumId };
    },
  }),
  defineRule({
    id: 'enum-end',
    description:
      'Match `End Enum` while inside an enum block; emit nothing and clear `ctx.vbaEnumBlock` so subsequent lines route back to the outside-enum-block rules.',
    pattern: ENUM_END_RE,
    requires: 'inside-enum-block',
    emit: (_m, ctx) => {
      ctx.vbaEnumBlock = null;
      return { kind: 'end-enum-block' as const };
    },
  }),
  defineRule({
    id: 'enum-member',
    description:
      'Match `<MemberName> [= <value>]` inside an open enum block; emit an `enum_member` node + a `contains` edge from the open enum.',
    pattern: ENUM_MEMBER_RE,
    requires: 'inside-enum-block',
    emit: (m, ctx, line, lineNum) => {
      const memberName = m[1] ?? '';
      if (!memberName || !ctx.vbaEnumBlock) return null;
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
        qualifiedName: `${ctx.vbaEnumBlock.name}.${memberName}`,
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
        source: ctx.vbaEnumBlock.id,
        target: memberId,
        kind: 'contains',
      });
      return { name: memberName };
    },
  }),
  defineRule({
    id: 'const-decl',
    description:
      'Match a `[Public|Private|Friend|Global] Const <decls>` declaration; always write the per-scope resolution bucket (issue #52) and additionally emit a `constant` node + `contains` edge ONLY when not inside a proc (proc-local Consts are not module symbols).',
    pattern: CONST_DECL_RE,
    requires: 'outside-enum-block',
    emit: (m, ctx, line, lineNum) => {
      const visibility = foldVisibility(m[1] ?? '');
      const body = m[2] ?? '';
      const declarations = parseConstDeclarations(body);
      let count = 0;
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
      return { count };
    },
  }),
];

/**
 * Issue #83: factory for the enum / const classifier.
 *
 * The body walks the declarative `RULES` table (Issue #153) and
 * honours each rule's `requires` precondition. The inter-line
 * state — `ctx.procStack`, `ctx.currentProcKey`, and
 * `ctx.vbaEnumBlock` — lives on `ctx` so the RULES table's `emit`
 * functions can read/write it without taking a closure reference.
 *
 * Issue #52: the first invocation also resets the shared proc-stack
 * + lookup key so leftover state from a previous `extract()` (only
 * possible in tests that construct a fresh extractor and run twice)
 * never leaks across sweeps. The walk below updates both every
 * iteration; `sweepCallsAndSql` resets again at its own start, so
 * the protocol stays consistent across both classifiers.
 */
export function createEnumsConstsClassifier(): VbaClassifier {
  let initialized = false;
  const cls: VbaClassifier = {
    name: 'enumsConsts',
    count: 0,
    classifyLine(line, i, ctx) {
      if (!initialized) {
        ctx.procStack.length = 0;
        ctx.currentProcKey = 'module';
        initialized = true;
      }
      const lineNum = i + 1;
      for (const rule of RULES) {
        if (rule.requires === 'inside-enum-block' && !ctx.vbaEnumBlock) continue;
        if (rule.requires === 'outside-enum-block' && ctx.vbaEnumBlock) continue;
        const m = matchRule(rule.pattern, line);
        if (!m) continue;
        const result = rule.emit(m, ctx, line, lineNum);
        if (result !== null && result !== undefined) {
          this.count += rule.count ? rule.count(result as never) : 1;
        }
        // `enum-end` cleared the enum block — nothing else on this
        // line can apply. The legacy cascade did the same short-circuit
        // via `return`.
        if (rule.id === 'enum-end') break;
      }
    },
  };
  return cls;
}

/**
 * Backward-compat wrapper (see procedures.ts). Returns the classifier's
 * `count` so the orchestrator can decide `hasAnySymbols`.
 */
export function sweepEnumsAndConsts(ctx: VbaExtractorContext, src: string): number {
  const cls = createEnumsConstsClassifier();
  const lines = src.split('\n');
  for (let i = 0; i < lines.length; i++) {
    cls.classifyLine(lines[i] ?? '', i, ctx);
  }
  return cls.count;
}
