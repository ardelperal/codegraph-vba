/**
 * Call/SQL sweep orchestrator (REQ-CODE-4, REQ-CODE-8). Walks the
 * (uncommented) source once, maintaining the procedure stack and `With`
 * receiver stack, and dispatches every per-line scanner: raise-events,
 * call-sites, Me/Forms control refs, SQL wrappers, statement/qualified calls,
 * `Set x = New …` late-instantiation, `DoCmd.Open*`, and TempVars.
 */
import { PROC_RE, PROCEDURE_END_RE, PRIMITIVE_TYPES } from './constants';
import { maskStringContent } from './text-utils';
import { VbaExtractorContext, ProcInfo } from './context';
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

export function sweepCallsAndSql(ctx: VbaExtractorContext, src: string): void {
  const lines = src.split('\n');
  const procedureStartLines = new Set<number>();
  const sqlTargetsThisFile = new Set<string>();

  // Issue #52: reset the shared scope state before the per-line walk
  // begins. `sweepEnumsAndConsts` already populated `procStack` /
  // `currentProcKey` during its own walk; clearing here guarantees the
  // `scanDoCmdOpenCalls` / `scanDoCmdOpenQuery` reads (which consult
  // `currentProcKey` per call-site) start in module scope and follow
  // the same push/pop discipline as the existing `stack` array below.
  ctx.procStack.length = 0;
  ctx.currentProcKey = 'module';

  // Walk the source once, emitting call edges and SQL edges per line and
  // tracking the current procedure stack.
  const stack: ProcInfo[] = [];
  const withReceiverStack: string[] = [];
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

    const procStart = PROC_RE.exec(line);
    if (procStart) {
      // Issue #52: mirror the proc push into the shared
      // `procStack` + `currentProcKey` so Const reads in this same
      // sweep see the same scope as the const writes did (during
      // `sweepEnumsAndConsts`).
      const procStartLine = lineNum;
      ctx.procStack.push(procStartLine);
      ctx.currentProcKey = String(procStartLine);

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
    } else if (PROCEDURE_END_RE.test(line) && stack.length > 0) {
      const ending = stack.pop()!;
      // Issue #52: mirror the pop into the shared scope state.
      ctx.procStack.pop();
      ctx.currentProcKey =
        ctx.procStack.length > 0
          ? String(ctx.procStack[ctx.procStack.length - 1])
          : 'module';
      procEndLines.set(ending.startLine, lineNum);
      continue;
    }

    // Fix 2 (Issue #2): mask string-literal content before call scanning so
    // patterns like `modHelper.BuildQuery(` inside a string argument are not
    // mistakenly treated as call sites.  SQL scanning still uses the original
    // line because SQL lives INSIDE string literals.
    const callScanLine = maskStringContent(line);

    if (stack.length > 0 && WITH_END_RE.test(callScanLine)) {
      withReceiverStack.pop();
      continue;
    }

    if (stack.length > 0 && !procedureStartLines.has(lineNum)) {
      const withStart = WITH_START_RE.exec(callScanLine);
      if (withStart) {
        const receiver = normalizeWithReceiver(withStart[1] ?? '');
        if (receiver) withReceiverStack.push(receiver);
        continue;
      }
    }

    // Don't scan call sites on the line that declares the procedure — it
    // would match the proc name itself in `Sub Outer()`.
    if (!procedureStartLines.has(lineNum) && stack.length > 0) {
      const currentProc = stack[stack.length - 1]!;
      scanRaiseEvents(ctx, callScanLine, currentProc, lineNum);
      scanCallSites(ctx, callScanLine, currentProc, lineNum);
    }

    // Hueco 1: capture `Me.<Control>` references. Only inside procedures
    // because `Me` is only meaningful inside a form's class module.
    if (!procedureStartLines.has(lineNum) && stack.length > 0) {
      scanMeControlReferences(ctx, callScanLine, stack[stack.length - 1]!, lineNum);
    }

    // SQL wrappers — only inside a procedure.  Use the ORIGINAL line — SQL is
    // inside string literals, so the masked line would strip the SQL content.
    if (stack.length > 0) {
      trackSqlVariableAssignment(lines, i, sqlVariables);
      scanSqlInLine(ctx, line, lineNum, sqlTargetsThisFile, sqlVariables);
    }

    // H1 fix: detect statement-form Sub calls (no parens, no `Call` keyword).
    // Issue #45: split single-line `If … Then <body>` clauses first so the
    // actual call after `Then`/`Else`/`:` is not shadowed by the leading
    // keyword.
    if (stack.length > 0 && !procedureStartLines.has(lineNum)) {
      // Issue #46: `Set x = New <Type>[.<Inner>]` late-instantiation.
      // Run BEFORE the call-site scan so a later `<x>.Member ...` line
      // finds `x` already registered in `localVarTypeMap`.
      const setNew = SET_NEW_RE.exec(callScanLine);
      if (setNew) {
        const varName = setNew[1] ?? '';
        const outerType = setNew[2] ?? '';
        const innerType = setNew[3] ?? '';
        if (varName && outerType) {
          // Skip primitives defensively — consistent with the Dim sweep guard.
          if (!PRIMITIVE_TYPES.has(outerType.toLowerCase())) {
            ctx.localVarTypeMap.set(varName.toLowerCase(), {
              outer: outerType,
              // Mirror `Dim x As Foo.Bar`: qualified `Set rs = New
              // DAO.Recordset` registers `qualified: true` so the PR #61
              // gate keeps downstream `rs.Method` calls silent.
              qualified: !!innerType,
              assignedWithSet: true,
              variableName: varName,
            });
            ctx.emitReference(outerType, lineNum, 0, 'vba-set-new');
          }
        }
      } else {
        // Factory-return inference: `Set x = <Factory>(...)`. Type x from a
        // same-file function's project-class return type so a later
        // `x.Method` resolves to the factory's class. Overrides a generic
        // `Dim x As Object/Variant` (or an untyped x), but yields to an
        // explicit `Dim x As <ProjectClass>` — the declaration is the
        // authoritative type. Runs after SET_NEW (which owns the `New` form).
        const setCall = SET_CALL_RE.exec(callScanLine);
        if (setCall) {
          const varName = setCall[1] ?? '';
          const factory = (setCall[2] ?? '').toLowerCase();
          const retType = factory ? ctx.functionReturnTypes.get(factory) : undefined;
          if (varName && retType) {
            const existing = ctx.localVarTypeMap.get(varName.toLowerCase());
            const existingIsProjectClass =
              !!existing &&
              !existing.qualified &&
              !PRIMITIVE_TYPES.has(existing.outer.toLowerCase());
            if (!existingIsProjectClass) {
              ctx.localVarTypeMap.set(varName.toLowerCase(), {
                outer: retType,
                qualified: false,
                assignedWithSet: true,
                variableName: varName,
              });
              ctx.emitReference(retType, lineNum, 0, 'vba-factory-return');
            }
          }
        }
      }

      const clauseLines = splitSingleLineIfClauses(callScanLine);
      for (const clauseLine of clauseLines) {
        const stmtCall = detectStatementCall(clauseLine);
        if (stmtCall) {
          const caller = stack[stack.length - 1]!;
          emitStatementCallEdge(ctx, caller, stmtCall, lineNum);
        }

        // Fix 7 + Fix 2 + Issue #40: qualified statement-form calls
        // (`Receiver.Member args`) — the dominant cross-object call shape.
        const qualStmt = detectQualifiedStatementCall(clauseLine);
        if (qualStmt) {
          const caller = stack[stack.length - 1]!;
          if (ctx.shouldProcessQualifiedCall(qualStmt.receiver)) {
            emitQualifiedStatementCallEdge(ctx, caller, qualStmt.receiver, qualStmt.member, lineNum);
          }
        }

        const withReceiver = withReceiverStack[withReceiverStack.length - 1];
        if (withReceiver) {
          const withCall = detectWithMemberCall(clauseLine);
          if (withCall && ctx.isLocalProjectClassVar(withReceiver)) {
            const caller = stack[stack.length - 1]!;
            emitQualifiedStatementCallEdge(ctx, caller, withReceiver, withCall.member, lineNum);
          }
        }
      }

      // B4 (hueco 6): `DoCmd.OpenForm "FormName"` modelling — the literal
      // form name lives INSIDE a string literal, so scan the ORIGINAL
      // (unmasked) line. The receiver is the same proc-stack frame.
      const caller2 = stack[stack.length - 1]!;
      // Issue #48: shared OpenForm/OpenReport dispatch; OpenQuery emits an
      // `UnresolvedReference` and stays separate.
      scanDoCmdOpenCalls(ctx, line, caller2, lineNum);
      scanDoCmdOpenQuery(ctx, line, caller2, lineNum);
      // Issue #44: cross-form bang references (`Forms!X` / `Forms("X")!Y`) —
      // scan the unmasked line (form name lives in a string literal in the
      // paren form).
      scanFormsBang(ctx, line, caller2, lineNum);

      // Issue #50: cross-form TempVars key accesses. Bang form scans the
      // masked line, paren + Add forms scan the original.
      sweepTempVars(ctx, callScanLine, line, lineNum, caller2);
    }

  }

  // Apply endLine to every emitted function node keyed by its startLine.
  // Functions without a recorded endLine (e.g. malformed VBA without an
  // `End`) keep their `endLine = startLine` from sweepProcedures —
  // which is the correct "single line" representation.
  for (const n of ctx.nodes) {
    if (n.kind !== 'function') continue;
    const end = procEndLines.get(n.startLine);
    if (end !== undefined) n.endLine = end;
  }
}
