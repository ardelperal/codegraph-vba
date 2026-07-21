/**
 * Call/SQL sweep orchestrator (REQ-CODE-4, REQ-CODE-8). Walks the
 * (uncommented) source once, maintaining the procedure stack and `With`
 * receiver stack, and dispatches every per-line scanner: raise-events,
 * call-sites, Me/Forms control refs, SQL wrappers, statement/qualified calls,
 * `Set x = New …` late-instantiation, `DoCmd.Open*`, and TempVars.
 */
import { PROC_RE, PROCEDURE_END_RE, PRIMITIVE_TYPES, isVbaKeyword } from './constants';
import { maskStringContent } from './text-utils';
import { VbaExtractorContext, ProcInfo, VbaClassifier } from './context';
import { defineRule, matchRuleForScan, VbaExtractionRule } from './rules';
import {
  scanRaiseEvents,
  scanCallSites,
  splitSingleLineIfClauses,
  detectStatementCall,
  emitStatementCallEdge,
  detectQualifiedStatementCall,
  emitQualifiedStatementCallEdge,
  normalizeWithReceiver,
  detectWithMemberCall,
} from './calls';
import { scanMeControlReferences, scanFormsBang } from './controls';
import { scanDoCmdOpenCalls, scanDoCmdOpenQuery } from './docmd';
import { sweepTempVars } from './tempvars';
import { scanSqlInLine, trackSqlVariableAssignment } from './sql-wrapper';

/**
 * Issue #46: `Set <var> = New <Type>[.<Inner>]` late-instantiation.
 * Groups: (1) variable name, (2) outer type, (3) optional inner type.
 * Operates on the MASKED line so `Set x = New Foo` inside a string literal
 * never matches.
 */
const SET_NEW_RE =
  /\bSet\s+(\p{L}[\p{L}\p{N}_]*)\s*=\s*New\s+(\p{L}[\p{L}\p{N}_]*)(?:\.(\p{L}[\p{L}\p{N}_]*))?/iu;

/**
 * Factory-return inference: `Set <var> = <Factory>(...)` or `Set <var> =
 * <Factory>` where <Factory> is a bare same-file function. Groups: (1)
 * variable, (2) factory name. The trailing `(?:\(|$)` requires the factory
 * name to be followed by `(` or end-of-expression, so a qualified
 * `Set x = obj.Member` (the `.` breaks the match) and the `New` form
 * (`New Foo` — `New` is followed by a space + type, not `(`/end) never match.
 * Runs on the MASKED line so a string literal never triggers it.
 */
const SET_CALL_RE =
  /\bSet\s+(\p{L}[\p{L}\p{N}_]*)\s*=\s*(\p{L}[\p{L}\p{N}_]*)\s*(?:\(|$)/iu;

/** Issue #43: track the receiver for `With <expr>` / `End With` blocks. */
const WITH_START_RE = /^\s*With\b\s+(.+?)\s*$/iu;
const WITH_END_RE = /^\s*End\s+With\b/iu;

/**
 * Issue #153: the declarative rule table for the calls/SQL concern.
 *
 * Of the call-sweep's full machinery, only the per-line *patterns*
 * fit the declarative shape. The procedural scanners
 * (`scanRaiseEvents`, `scanCallSites`, `scanMeControlReferences`,
 * `scanSqlInLine`, `scanDoCmdOpenCalls`, `scanDoCmdOpenQuery`,
 * `scanFormsBang`, `sweepTempVars`, statement/qualified call
 * detection) walk the masked line scanning for call shapes that
 * don't reduce to a single regex — they stay inside the factory's
 * `classifyLine` (acknowledged in the Issue #153 spec: "The
 * inter-line state machines (procedure stack, with stack,
 * sqlVariables) NEED a class, not a pure rule.").
 *
 * Four rules, all driven by the masked line (the call site / Set
 * patterns must not match inside string literals):
 *
 *  - `set-new`    — `Set <var> = New <Type>[.<Inner>]`; registers the
 *                   receiver in `localVarTypeMap` and emits a
 *                   `vba-set-new` `references` edge.
 *  - `set-call`   — `Set <var> = <Factory>(...)`; types the receiver
 *                   from a same-file function's project-class return
 *                   type so a later `x.Method` resolves to the
 *                   factory's class. Emits a `vba-factory-return`
 *                   `references` edge.
 *  - `with-start` — `With <receiver>`; normalizes the receiver and
 *                   pushes it onto `ctx.vbaWithStack` (replaces the
 *                   pre-#153 per-factory closure variable).
 *  - `with-end`   — `End With`; pops the matching `ctx.vbaWithStack`
 *                   entry.
 *
 * The `set-new` rule's emit function is a "pure" (match, ctx, line)
 * consumer — it reads/writes only `ctx` and the unmasked/ masked
 * line, with no closure references. Same for the other three.
 */
export const RULES: readonly VbaExtractionRule<unknown>[] = [
  defineRule({
    id: 'set-new',
    description:
      'Match `Set <var> = New <Type>[.<Inner>]` (masked line, inside-procedure); register the receiver in `localVarTypeMap` and emit a `vba-set-new` `references` edge to the outer type.',
    pattern: SET_NEW_RE,
    scan: 'masked',
    requires: 'inside-procedure',
    emit: (m, ctx, _line, lineNum) => {
      const varName = m[1] ?? '';
      const outerType = m[2] ?? '';
      const innerType = m[3] ?? '';
      if (!varName || !outerType) return null;
      // Skip primitives defensively — consistent with the Dim sweep guard.
      if (PRIMITIVE_TYPES.has(outerType.toLowerCase())) return null;
      // Issue #205: write to the current procedure's bucket so a
      // `Set x = New Foo` inside `Sub Bar` does not silently
      // overwrite a `Dim x As Whatever` declaration in `Sub Baz`
      // (or a module-level `Dim x As ModuleThing`). `currentProcKey`
      // is maintained by the enum/const classifier, the sole writer of
      // the shared scope stack during the main walk.
      ctx.setLocalVarTypeInScope(ctx.currentProcKey, varName, {
        outer: outerType,
        // Mirror `Dim x As Foo.Bar`: qualified `Set rs = New
        // DAO.Recordset` registers `qualified: true` so the PR #61
        // gate keeps downstream `rs.Method` calls silent.
        qualified: !!innerType,
        assignedWithSet: true,
        variableName: varName,
      });
      ctx.emitReference(outerType, lineNum, 0, 'vba-set-new');
      return { varName, outerType };
    },
  }),
  defineRule({
    id: 'set-call',
    description:
      'Match `Set <var> = <Factory>(...)` (masked line, inside-procedure); type the receiver from a same-file function\'s project-class return type (looked up in `ctx.functionReturnTypes`) so a later `x.Method` resolves to the factory\'s class. Yields to an explicit `Dim x As <ProjectClass>` — the declaration is the authoritative type.',
    pattern: SET_CALL_RE,
    scan: 'masked',
    requires: 'inside-procedure',
    emit: (m, ctx, _line, lineNum) => {
      const varName = m[1] ?? '';
      const factory = (m[2] ?? '').toLowerCase();
      const retType = factory ? ctx.functionReturnTypes.get(factory) : undefined;
      if (!varName || !retType) return null;
      // Issue #205: the existing-key check uses the two-tier
      // lookup (current proc → module) so a `Set x = Factory()`
      // inside `Sub Bar` correctly yields to a `Dim x As Class`
      // declared in `Sub Bar` (proc bucket) AND to a module-level
      // `Dim x As ModuleClass` (module bucket). A proc-local
      // `Dim x As Class` declared in `Sub Baz` (a different
      // procedure) does NOT suppress this `set-call`.
      const existing = ctx.lookupLocalVarType(varName);
      const existingIsProjectClass =
        !!existing &&
        !existing.qualified &&
        !PRIMITIVE_TYPES.has(existing.outer.toLowerCase());
      if (existingIsProjectClass) return null;
      ctx.setLocalVarTypeInScope(ctx.currentProcKey, varName, {
        outer: retType,
        qualified: false,
        assignedWithSet: true,
        variableName: varName,
      });
      ctx.emitReference(retType, lineNum, 0, 'vba-factory-return');
      return { varName, retType };
    },
  }),
  defineRule({
    id: 'with-start',
    description:
      'Match `With <receiver>` (masked line, inside-procedure); normalize the receiver expression and push it onto `ctx.vbaWithStack` so a subsequent `.Member` statement-form call routes through the With-member path.',
    pattern: WITH_START_RE,
    scan: 'masked',
    requires: 'inside-procedure',
    emit: (m, ctx) => {
      const receiver = normalizeWithReceiver(m[1] ?? '');
      if (!receiver) return null;
      ctx.vbaWithStack.push(receiver);
      return { receiver };
    },
  }),
  defineRule({
    id: 'with-end',
    description:
      'Match `End With` (inside-procedure); pop the matching entry from `ctx.vbaWithStack`.',
    pattern: WITH_END_RE,
    scan: 'masked',
    requires: 'inside-procedure',
    emit: (_m, ctx) => {
      ctx.vbaWithStack.pop();
      return { kind: 'with-end' as const };
    },
  }),
];

/**
 * Issue #83: factory for the calls/SQL classifier. The factory takes the
 * pre-split `lines` array (so `trackSqlVariableAssignment` can do its
 * multi-line look-ahead for `&`-accumulate semantics) and closes over the
 * per-file state the legacy `sweepCallsAndSql` declared locally.
 *
 * Issue #153: the per-line pattern matching has been extracted into
 * the `RULES` table. The factory body is now a thin shell that
 * dispatches `RULES` and runs the procedural scanners (call sites,
 * raise events, SQL, DoCmd, TempVars, etc.) that don't reduce to a
 * single regex.
 */
export function createCallsAndSqlClassifier(
  lines: readonly string[],
): VbaClassifier {
  const procedureStartLines = new Set<number>();
  const sqlTargetsThisFile = new Set<string>();
  // Walk the source once, emitting call edges and SQL edges per line and
  // tracking the current procedure stack.
  const stack: ProcInfo[] = [];
  const moduleSqlVariables = new Map<string, string>();
  const sqlVariables = new Map<string, string>();
  // C2 fix: track each procedure's `endLine` (the line containing the
  // matching `End Sub`/`End Function`/`End Property`) keyed by its
  // `startLine`. After the loop, we update every function node's
  // `endLine` so `codegraph_explore` returns the full body span —
  // not just the signature line.
  const procEndLines = new Map<number, number>();

  const cls: VbaClassifier = {
    name: 'callsAndSql',
    count: 0,
    classifyLine(line, i, ctx) {
      const lineNum = i + 1;

      const procStart = PROC_RE.exec(line);
      if (procStart) {
        // SQL variables are procedure-scoped, with module assignments as a
        // fallback. Seed a fresh procedure bucket so common names such as
        // `strSQL` cannot leak from the preceding procedure.
        sqlVariables.clear();
        for (const [name, value] of moduleSqlVariables) {
          sqlVariables.set(name, value);
        }
        const name = procStart[3] ?? '';
        const bucket = ctx.localProcs.get(name);
        if (bucket) {
          // Fix 1: select the ProcInfo whose startLine matches the current
          // line, not always bucket[0]. When Property Get/Let/Set share a
          // name, all three exist in the bucket; pushing bucket[0] every time
          // meant Let/Set bodies erroneously attributed to the Get's ProcInfo.
          const proc = bucket.find((p) => p.startLine === lineNum) ?? bucket[0];
          if (proc) stack.push(proc);
        }
        procedureStartLines.add(lineNum);
        // FALL THROUGH — on a proc-start line the legacy sweep ALSO ran
        // the masked-line / with-receiver / call-scan / SQL-scan body.
        // A `return` here would silently drop the first body line of every
        // procedure (a real bug — `Sub Foo()\n  Bar 1\n` would never see
        // the `Bar 1` call).
      }
      // This is deliberately independent of the proc-start branch: a
      // colon-separated single-line procedure contains both markers.
      // Keep the frame alive until after scanners process its body.
      const endsProcedure = PROCEDURE_END_RE.test(line) && stack.length > 0;

      // Fix 2 (Issue #2): mask string-literal content before call scanning so
      // patterns like `modHelper.BuildQuery(` inside a string argument are not
      // mistakenly treated as call sites.  SQL scanning still uses the original
      // line because SQL lives INSIDE string literals.
      const callScanLine = maskStringContent(line);
      const procStartColon = procStart ? callScanLine.indexOf(':') : -1;
      const procStartBodyClauses = procStart
        ? procStartColon < 0
          ? []
          : callScanLine
            .slice(procStartColon + 1)
            .split(':')
            .map((clause) => clause.trim())
            .filter((clause) => clause.length > 0 && !/^End\s+(?:Sub|Function|Property)\b/i.test(clause))
        : null;

      // Issue #153: dispatch the declarative RULES table on the
      // masked line. The four per-line patterns (set-new, set-call,
      // with-start, with-end) all need the masked line because their
      // match shapes (`Set`, `With`) are noise inside string
      // literals. The dispatcher walks `RULES` and calls each rule's
      // `emit` when its `pattern.exec` matches. Each rule's
      // `requires` precondition is honoured here — `inside-procedure`
      // rules only fire when the call-sweep's per-instance proc stack
      // is non-empty (matching the legacy cascade's `if (stack.length
      // > 0)` gate).
      for (const rule of RULES) {
        if (rule.requires === 'inside-procedure' && stack.length === 0) continue;
        const matched = matchRuleForScan(rule, line, callScanLine);
        if (!matched) continue;
        const result = rule.emit(matched.match, ctx, matched.line, lineNum);
        if (result !== null && result !== undefined) {
          this.count += rule.count ? rule.count(result as never) : 1;
        }
        // `with-end` is short-circuited: a line that's `End With` is
        // not also a `Set x = New ...`, so break the loop.
        if (rule.id === 'with-end') break;
      }

      // Don't scan call sites on the line that declares the procedure — it
      // would match the proc name itself in `Sub Outer()`.
      if (stack.length > 0) {
        const currentProc = stack[stack.length - 1]!;
        const scanLines = procStartBodyClauses ?? [callScanLine];
        for (const scanLine of scanLines) {
          scanRaiseEvents(ctx, scanLine, currentProc, lineNum);
          scanCallSites(ctx, scanLine, currentProc, lineNum);
        }
      }

      // Hueco 1: capture `Me.<Control>` references. Only inside procedures
      // because `Me` is only meaningful inside a form's class module.
      if (stack.length > 0) {
        const scanLines = procStartBodyClauses ?? [callScanLine];
        for (const scanLine of scanLines) {
          scanMeControlReferences(ctx, scanLine, stack[stack.length - 1]!, lineNum);
        }
      }

      // SQL wrappers — only inside a procedure.  Use the ORIGINAL line — SQL is
      // inside string literals, so the masked line would strip the SQL content.
      if (stack.length > 0) {
        trackSqlVariableAssignment(lines as string[], i, sqlVariables);
        scanSqlInLine(ctx, line, lineNum, sqlTargetsThisFile, sqlVariables);
      } else {
        // Preserve file/module-level assignments for use as the fallback in
        // every procedure without mixing one procedure's locals into another.
        trackSqlVariableAssignment(lines as string[], i, moduleSqlVariables);
      }

      // H1 fix: detect statement-form Sub calls (no parens, no `Call` keyword).
      // Issue #45: split single-line `If … Then <body>` clauses first so the
      // actual call after `Then`/`Else`/`:` is not shadowed by the leading
      // keyword.
      if (stack.length > 0 && (!procedureStartLines.has(lineNum) || procStartBodyClauses?.length)) {
        // Note: `Set x = New ...` and `Set x = <Factory>(...)` are now
        // dispatched by the RULES table above. The factory body below
        // focuses on the procedural call-site scans that don't
        // reduce to a single regex.
        const clauseLines = procStartBodyClauses ?? splitSingleLineIfClauses(callScanLine);
        for (const clauseLine of clauseLines) {
          const stmtCall = detectStatementCall(clauseLine);
          if (stmtCall) {
            const caller = stack[stack.length - 1]!;
            // Returns true when a same-file calls edge was emitted; false
            // when the call was silenced (blacklist / runtime receiver /
            // unresolvable same-file target).
            const emitted = emitStatementCallEdge(ctx, caller, stmtCall, lineNum);
            // Round-3 (issue #108): if the statement-form Sub call did
            // NOT resolve, surface it as an `unqualified-ident` unresolved
            // reference. The const-first disambiguation rule (FR-3.1)
            // ensures Const reads do not pollute the call bucket — when
            // `stmtCall` resolves to a known Const, we still surface it as
            // `unqualified-ident` (NOT `call`) because that's the
            // unambiguous shape of a bare-ident read.
            if (!emitted && stmtCall !== caller.name && !isVbaKeyword(stmtCall)) {
              ctx.unresolvedReferences.push({
                fromNodeId: ctx.findOrCreateFunctionNodeId(caller),
                referenceName: stmtCall,
                referenceKind: 'unqualified-ident',
                line: lineNum,
                column: 0,
                filePath: ctx.filePath,
                language: 'vba',
                metadata: { synthesizedBy: 'vba-statement-call-unresolved' },
              });
            }
          }

          // Fix 7 + Fix 2 + Issue #40: qualified statement-form calls
          // (`Receiver.Member args`) — the dominant cross-object call shape.
          const qualStmt = detectQualifiedStatementCall(clauseLine);
          if (qualStmt) {
            const caller = stack[stack.length - 1]!;
            if (ctx.shouldProcessQualifiedCall(qualStmt.receiver)) {
              // Returns true when a `calls` edge was emitted to a synthetic
              // stub node; false when the receiver is not eligible
              // (primitive / DAO-qualified / runtime). Round-3 (issue
              // #108): when the synthetic-stub path skips, the parent
              // caller still needs to see an `unresolved_refs` row so the
              // SQL filter `WHERE reference_kind = 'qualified-call'`
              // surfaces these from `(caller, qualified, line)` tuples the
              // resolver couldn't bind.
              emitQualifiedStatementCallEdge(
                ctx,
                caller,
                qualStmt.receiver,
                qualStmt.member,
                lineNum,
              );
            }
          }

          const withReceiver = ctx.vbaWithStack[ctx.vbaWithStack.length - 1];
          if (withReceiver) {
            const withCall = detectWithMemberCall(clauseLine);
            if (withCall && ctx.isLocalProjectClassVar(withReceiver)) {
              const caller = stack[stack.length - 1]!;
              emitQualifiedStatementCallEdge(ctx, caller, withReceiver, withCall.member, lineNum);
            } else if (withCall) {
              // Round-3 (FR-2.5): `.Member` inside a `With` block where the
              // receiver is NOT a project-class local (e.g. a runtime /
              // DAO-qualified `With rs` where `rs` is `DAO.Recordset`).
              // surface the call as `member-with` so the SQL filter can
              // detect these by shape.
              const caller = stack[stack.length - 1]!;
              ctx.unresolvedReferences.push({
                fromNodeId: ctx.findOrCreateFunctionNodeId(caller),
                referenceName: `${withReceiver}.${withCall.member}`,
                referenceKind: 'member-with',
                line: lineNum,
                column: 0,
                filePath: ctx.filePath,
                language: 'vba',
                metadata: { synthesizedBy: 'vba-with-member-unresolved' },
              });
            }
          }
        }

        // B4 (hueco 6): `DoCmd.OpenForm "FormName"` modelling — the literal
        // form name lives INSIDE a string literal, so scan the ORIGINAL
        // (unmasked) line. The receiver is the same proc-stack frame.
        const caller2 = stack[stack.length - 1]!;
        // Issue #48: shared OpenForm/OpenReport dispatch; OpenQuery emits an
        // `UnresolvedReference` and stays separate.
        scanDoCmdOpenCalls(ctx, line, callScanLine, caller2, lineNum);
        scanDoCmdOpenQuery(ctx, line, callScanLine, caller2, lineNum);
        // Issue #44: cross-form bang references (`Forms!X` / `Forms("X")!Y`) —
        // scan the unmasked line (form name lives in a string literal in the
        // paren form).
        scanFormsBang(ctx, line, callScanLine, caller2, lineNum);

        // Issue #50: cross-form TempVars key accesses. Bang form scans the
        // masked line, paren + Add forms scan the original.
        sweepTempVars(ctx, callScanLine, line, lineNum, caller2);
      }

      if (endsProcedure) {
        const ending = stack.pop()!;
        procEndLines.set(ending.startLine, lineNum);
        sqlVariables.clear();
        for (const [name, value] of moduleSqlVariables) {
          sqlVariables.set(name, value);
        }
      }
    },
    finalize(ctx) {
      // Apply endLine to every emitted function node keyed by its startLine.
      // Functions without a recorded endLine (e.g. malformed VBA without an
      // `End`) keep their `endLine = startLine` from sweepProcedures —
      // which is the correct "single line" representation.
      for (const n of ctx.nodes) {
        if (n.kind !== 'function') continue;
        const end = procEndLines.get(n.startLine);
        if (end !== undefined) n.endLine = end;
      }
    },
  };

  return cls;
}

/**
 * Backward-compat wrapper (see procedures.ts). Returns void — the calls
 * sweep never contributed to `hasAnySymbols` directly (every other
 * concern's `count` is the signal the orchestrator reads).
 */
export function sweepCallsAndSql(ctx: VbaExtractorContext, src: string): void {
  const lines = src.split('\n');
  const cls = createCallsAndSqlClassifier(lines);
  for (let i = 0; i < lines.length; i++) {
    cls.classifyLine(lines[i] ?? '', i, ctx);
  }
  cls.finalize?.(ctx);
}
