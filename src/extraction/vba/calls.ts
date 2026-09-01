/**
 * Call-site detection (REQ-CODE-4): paren-form calls (`CALL_RE`),
 * statement-form calls (`Foo arg` / `Call Foo`), qualified statement calls
 * (`Receiver.Member args`), `RaiseEvent`, and the `With` receiver/member
 * helpers. Emits same-file `calls` edges and heuristic cross-module
 * `calls` edges to synthetic stubs (`vba-name-resolution`).
 */
import { generateNodeId } from '../tree-sitter-helpers';
import {
  PRIMITIVE_TYPES,
  CALL_KEYWORD_BLACKLIST,
  RUNTIME_RECEIVER_BLACKLIST,
} from './constants';
import { isRuntimeObject } from './runtime-objects';
import { VbaExtractorContext, ProcInfo } from './context';

/**
 * Call-site regex — captures either `Name(...)` (same-file candidate) or
 * `Receiver.Member(...)` (qualified). The receiver AND member alternatives
 * accept BOTH the bare form (`Foo`) and the VBA bracketed form (`[Foo Bar]`)
 * — bracketed captures win when present; brackets are stripped by the regex
 * itself. Issue #54 added the bracketed alternative so
 * `[FUNCIONES UTILES].FormatearFecha(fecha)` is no longer silently dropped.
 */
const CALL_RE =
  /(?<![\w.])(?:\[([^\]]+)\]|(\p{L}[\p{L}\p{N}_]*))(?:\.(?:\[([^\]]+)\]|(\p{L}[\p{L}\p{N}_]*)))?\s*\(/gu;

const RAISE_EVENT_RE = /\bRaiseEvent\s+(\p{L}[\p{L}\p{N}_]*)\b/giu;

export function scanRaiseEvents(
  ctx: VbaExtractorContext,
  line: string,
  from: ProcInfo,
  lineNum: number,
): void {
  RAISE_EVENT_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = RAISE_EVENT_RE.exec(line)) !== null) {
    const eventName = m[1] ?? '';
    const eventNode = ctx.localEvents.get(eventName.toLowerCase());
    if (!eventNode) continue;
    // Issue #152: bump the per-event fanout counter BEFORE pushing the
    // edge. The orchestrator's `applyRaiseFanoutGate` reads this map to
    // decide which event nodes to flag `metadata.highFanout: true` and
    // which `raises-event` edges to drop. Counting in the same walk is
    // free — one Map.get / Map.set per raise site.
    ctx.raiseEventCounts.set(
      eventNode.id,
      (ctx.raiseEventCounts.get(eventNode.id) ?? 0) + 1,
    );
    ctx.edges.push({
      source: ctx.findOrCreateFunctionNodeId(from),
      target: eventNode.id,
      kind: 'raises-event',
      provenance: 'parser',
      metadata: { eventName },
      line: lineNum,
      column: m.index,
    });
  }
}

export function scanCallSites(
  ctx: VbaExtractorContext,
  line: string,
  from: ProcInfo,
  lineNum: number,
): void {
  const arrayAssignment = /^\s*(\p{L}[\p{L}\p{N}_]*)\s*=\s*Array\s*\(/iu.exec(line);
  if (arrayAssignment?.[1]) {
    // Issue #205: scope the `isArray` mutation to the current proc
    // bucket so a `x = Array(...)` inside `Sub Foo` does not flip
    // the `isArray` flag on a `Dim x As String` declared in
    // `Sub Bar` (or at module level).
    const existing = ctx.lookupLocalVarType(arrayAssignment[1]);
    if (existing) existing.isArray = true;
  }
  CALL_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = CALL_RE.exec(line)) !== null) {
    // Issue #54: CALL_RE groups (1)/(2) are alternative captures for the
    // receiver position, (3)/(4) for the member position. The bracketed
    // alternative wins when present; the captured value is already
    // unwrapped by the regex.
    const receiver = m[1] ?? m[2] ?? '';
    const member = m[3] ?? m[4] ?? '';
    if (!receiver) continue;
    if (!member && ctx.lookupLocalVarType(receiver)?.isArray) continue;
    // Issue #190: procedure-scoped param-array recognition. A
    // `ByRef name() As Type` (or `ByVal name() As Type`) on this
    // procedure's signature makes `name` an array for the duration of
    // the body. The check is on the CURRENT procedure's
    // `arrayParameters` set, not on the file-global `localVarTypeMap`,
    // so a same-named parameter on a different procedure cannot suppress
    // a genuine missing-call elsewhere in the file.
    if (
      !member &&
      from.arrayParameters !== undefined &&
      from.arrayParameters.includes(receiver.toLowerCase())
    ) {
      continue;
    }
    // Skip VBA control-flow keywords.
    if (CALL_KEYWORD_BLACKLIST.has(receiver)) continue;
    if (member && CALL_KEYWORD_BLACKLIST.has(member)) continue;
    // Skip Access runtime objects — `Me`, `DoCmd`, `Application`, etc.
    // These calls are real but the targets are NOT user code; emitting
    // synthetic function nodes for them pollutes the graph (audit W4).
    if (RUNTIME_RECEIVER_BLACKLIST.has(receiver)) continue;
    if (member && RUNTIME_RECEIVER_BLACKLIST.has(member)) continue;
    // Skip the receiver when it equals the containing procedure (self-call).
    if (receiver === from.name && !member) continue;

    const col = m.index;

    if (!member) {
      // Bare `Name(...)` — same-file resolution.
      const localFuncNode = ctx.findFunctionNodeByName(receiver);
      if (!localFuncNode) {
        // Round-3 (FR-2.1, issue #108): the call-sweep path used to
        // silent-skip here, leaving the paren-form unresolvable call
        // invisible in `unresolved_refs`. Surface it as canonical `'calls'`
        // so consumers use the same kind as resolved call edges.
        ctx.unresolvedReferences.push({
          fromNodeId: ctx.findOrCreateFunctionNodeId(from),
          referenceName: receiver,
          referenceKind: 'calls',
          line: lineNum,
          column: col,
          filePath: ctx.filePath,
          language: 'vba',
          metadata: { synthesizedBy: 'vba-paren-call-unresolved' },
        });
        continue;
      }
      ctx.edges.push({
        source: ctx.findOrCreateFunctionNodeId(from),
        target: localFuncNode.id,
        kind: 'calls',
        line: lineNum,
        column: col,
      });
    } else {
      // Qualified `Receiver.Member(...)` — synthesize the call target only
      // for project-class local variables or undeclared module candidates.
      if (!ctx.shouldProcessQualifiedCall(receiver)) {
        // Round-3 (FR-2.2): receiver is a declared primitive or
        // runtime-blacklisted. Surface the qualified call as
        // `'qualified-call'` so the SQL filter still sees these from
        // `(caller, qualified, line)` tuples the resolver can't bind.
        ctx.unresolvedReferences.push({
          fromNodeId: ctx.findOrCreateFunctionNodeId(from),
          referenceName: `${receiver}.${member}`,
          referenceKind: 'qualified-call',
          line: lineNum,
          column: col,
          filePath: ctx.filePath,
          language: 'vba',
          metadata: { synthesizedBy: 'vba-qualified-call-unresolved' },
        });
        continue;
      }
      // #12a: `receiverType` resolves to the real class name when
      // `receiver` is a declared project-class local var; otherwise it's the
      // raw `receiver` text unchanged (e.g. `.bas`-qualified module calls).
      //
      // Antigravity audit Task 3 (refined gate): if `receiver` is a
      // file-local variable declared as a PRIMITIVE, skip emission — the
      // stub `<receiver>.<member>` would be dead-end graph pollution no
      // resolver could ever repoint. Cross-module qualified calls like
      // `modUtils.Foo(1)` are unaffected (`modUtils` is not a local var).
      // Issue #205: lookup is two-tier (current proc → module) so the
      // primitive-gate fires on the type declared IN this procedure,
      // not the type declared in whichever procedure happened to write
      // to `localVarTypeMap` last.
      const recvEntry = ctx.lookupLocalVarType(receiver);
      if (recvEntry && PRIMITIVE_TYPES.has(recvEntry.outer.toLowerCase())) {
        // Match the round-3 surfaced row so the SQL filter has the
        // receiver/member string even when the stub emission was
        // suppressed.
        ctx.unresolvedReferences.push({
          fromNodeId: ctx.findOrCreateFunctionNodeId(from),
          referenceName: `${receiver}.${member}`,
          referenceKind: 'qualified-call',
          line: lineNum,
          column: col,
          filePath: ctx.filePath,
          language: 'vba',
          metadata: { synthesizedBy: 'vba-qualified-call-unresolved' },
        });
        continue;
      }
      const receiverType = ctx.resolveReceiverType(receiver);
      // Issue #245: a call whose RESOLVED receiver type is a VBA/Access
      // runtime object (`Dim c As New Collection` → `Collection.Add`,
      // `VBA.DoEvents`, `fso.FileExists`) can never name user code. The
      // resolver already declines to repoint these edges
      // (`repointDecision: 'declined-runtime'`), but the stub NODE was
      // created anyway and leaked into symbol search and node counts.
      // Gate BEFORE `generateNodeId` so no node and no `calls` edge is
      // born. Nothing else changes: this branch never pushed an
      // `UnresolvedReference`, so `unresolvedByKind` is untouched.
      if (isRuntimeObject(receiverType)) continue;
      const qualified = `${receiverType}.${member}`;
      // Avoid emitting duplicate edges for the same call (within a line).
      const dedupeKey = `${from.name}->${qualified}@${lineNum}`;
      if (ctx.callDedupe.has(dedupeKey)) continue;
      ctx.callDedupe.add(dedupeKey);

      const synthId = generateNodeId(
        ctx.filePath,
        'function',
        qualified,
        lineNum,
      );
      // Only add the synthetic function node once per (file, qualified, line).
      if (!ctx.synthFunctionNodeIds.has(synthId)) {
        ctx.synthFunctionNodeIds.add(synthId);
        ctx.nodes.push({
          id: synthId,
          kind: 'function',
          name: qualified,
          qualifiedName: qualified,
          filePath: ctx.filePath,
          language: 'vba',
          startLine: lineNum,
          endLine: lineNum,
          startColumn: col,
          endColumn: col + qualified.length,
          visibility: 'public',
          // #12a: tag the stub so the post-extraction resolver (#12b)
          // can find and repoint it.
          metadata: { stub: true },
          updatedAt: Date.now(),
        });
      }
      ctx.edges.push({
        source: ctx.findOrCreateFunctionNodeId(from),
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

/**
 * Issue #45: split a single-line VBA `If <cond> Then <body>` into one or
 * more statement-clause fragments the statement-call detectors can process
 * (`If x Then Foo Else Bar`, colon-separated `If x Then DoA: DoB`). When the
 * line is not a single-line `If … Then` shape, returns `[<line>]` so
 * block-form `If` still works through the per-line scan. `GoTo`/`Exit`/
 * `Resume` clauses are filtered out. `line` is the string-literal-masked
 * scan line, which makes global `:`/`Else` splitting safe.
 */
export function splitSingleLineIfClauses(line: string): string[] {
  const trimmed = line.trimStart();
  if (!trimmed) return [];
  // Match `If <cond> Then <body>` with a non-greedy condition. Requiring
  // at least one whitespace character after `Then` ensures the block
  // form `If x Then` (with the body on subsequent lines) is left alone.
  const ifThenRe = /^If\s[\s\S]+?\bThen\b\s+/i;
  const m = ifThenRe.exec(trimmed);
  if (!m) {
    // Not a single-line `If … Then` — preserve the original line.
    return [line];
  }
  const body = trimmed.slice(m[0].length);
  // Split on top-level `Else` (case-insensitive; word-bounded).
  const elseClauses = body.split(/\s+Else\s+/i);
  const clauses: string[] = [];
  for (const elseClause of elseClauses) {
    // Split each Else-clause on `:` for multi-statement single-line
    // `If` bodies. VBA expressions never contain `:`, so a global
    // split is correct on a masked line.
    const subStatements = elseClause.split(':');
    for (const sub of subStatements) {
      const t = sub.trim();
      if (!t) continue;
      // Defense in depth: GoTo / Exit / Resume are VBA control-flow
      // statements, not Sub calls — drop them before they reach the
      // statement-call detectors.
      if (/^(?:GoTo|Exit|Resume)\b/i.test(t)) continue;
      clauses.push(t);
    }
  }
  return clauses;
}

/**
 * Issue #265: the outcome of `detectStatementCall`.
 *
 * `unambiguous` records whether the *syntax* of the statement can only be a
 * procedure call:
 *
 *  - `Call MySub`   — the `Call` keyword is only valid on a procedure.
 *  - `MySub 1, 2`   — a constant cannot take an argument list.
 *  - `MySub`        — AMBIGUOUS: a bare identifier is equally a `Const` read.
 *
 * The call sweep uses this to pick the unresolved-reference kind: `calls` for
 * the two unambiguous shapes, `unqualified-ident` for the bare read (the
 * bucket the const-first disambiguation rule of issue #108 / FR-3.1 needs).
 */
export interface StatementCall {
  /** The called procedure name (the leading identifier). */
  name: string;
  /** True when the `Call` keyword or an argument list rules out a Const read. */
  unambiguous: boolean;
}

/**
 * H1: detect a statement-form Sub call (`MySub`, `MySub arg1, x`,
 * `Call MySub arg1`). Returns the called proc name plus its syntactic
 * unambiguity (issue #265), or null for declarations, assignments,
 * comments, keyword lines, and the paren form (handled by CALL_RE).
 */
export function detectStatementCall(line: string): StatementCall | null {
  let trimmed = line.trimStart();
  if (!trimmed) return null;
  // Strip `Call ` keyword if present — same call shape after. Issue #265:
  // remember that it was there; `Call X` cannot be a Const read.
  const hasCallKeyword = /^Call\s/i.test(trimmed);
  if (hasCallKeyword) {
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
  // `MySub(...)` is parens-form and already handled by CALL_RE.
  if (rest.startsWith('(')) return null;
  // Bare `MySub` is a valid no-argument statement-form Sub call.
  if (rest.length === 0) return { name: procName, unambiguous: hasCallKeyword };
  const nextCh = trimmed.charAt(procName.length);
  if (nextCh !== ' ' && nextCh !== '\t') return null;
  const args = rest.trimStart();
  // Skip leading-identifier assignments (`X = ...`). Do not reject `=` inside
  // argument expressions because named arguments use `:=` and comparisons can
  // appear in expressions.
  if (args.startsWith('=')) return null;
  // Issue #265: `args` is measured AFTER trimming, so `MySub   ` (trailing
  // whitespace only) stays the ambiguous bare-identifier shape.
  return { name: procName, unambiguous: hasCallKeyword || args.length > 0 };
}

/**
 * H1: emit a same-file `calls` edge for a statement-form Sub call to the
 * already-emitted function node named `procName`. Returns `true` when a
 * `calls` edge was pushed, `false` when the call was silenced
 * (blacklist / runtime / self-call / unresolvable). Round-3 (issue
 * #108) needs that boolean so the call sweep can fall through to push
 * an `unqualified-ident` unresolved reference.
 */
export function emitStatementCallEdge(
  ctx: VbaExtractorContext,
  caller: ProcInfo,
  procName: string,
  lineNum: number,
): boolean {
  if (procName === caller.name) return false; // skip self-call
  if (CALL_KEYWORD_BLACKLIST.has(procName)) return false;
  if (RUNTIME_RECEIVER_BLACKLIST.has(procName)) return false;
  const target = ctx.findFunctionNodeByName(procName);
  if (!target) return false;
  ctx.edges.push({
    source: ctx.findOrCreateFunctionNodeId(caller),
    target: target.id,
    kind: 'calls',
    line: lineNum,
    column: 0,
  });
  return true;
}

/**
 * Fix 7: detect a qualified statement-form call — `Receiver.Member <args>`
 * where `Receiver.Member` is NOT followed by `(`. Distinct from the paren
 * form (handled by CALL_RE). Property assignments and blacklisted
 * receivers/members are excluded. Returns `{receiver, member}` or null.
 */
export function detectQualifiedStatementCall(
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
  // Issue #54: the receiver alternative accepts BOTH the bare form
  // (`Foo`) and the VBA bracketed form (`[Foo Bar]`). Same for the member.
  const receiverM = /^(?:\[([^\]]+)\]|(\p{L}[\p{L}\p{N}_]*))/u.exec(trimmed);
  if (!receiverM) return null;
  const receiver = receiverM[1] ?? receiverM[2] ?? '';
  const rest = trimmed.slice(receiverM[0].length);
  // Must have a dot separator.
  if (!rest.startsWith('.')) return null;
  // Extract member identifier.
  const memberRest = rest.slice(1); // skip the dot
  const memberM = /^(?:\[([^\]]+)\]|(\p{L}[\p{L}\p{N}_]*))/u.exec(memberRest);
  if (!memberM) return null;
  const member = memberM[1] ?? memberM[2] ?? '';
  const afterMember = memberRest.slice(memberM[0].length);
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
  if (CALL_KEYWORD_BLACKLIST.has(receiver)) return null;
  if (RUNTIME_RECEIVER_BLACKLIST.has(receiver)) return null;
  if (CALL_KEYWORD_BLACKLIST.has(member)) return null;
  if (RUNTIME_RECEIVER_BLACKLIST.has(member)) return null;
  return { receiver, member };
}

/**
 * Fix 7: emit a heuristic `calls` edge for a qualified statement-form call.
 * Same shape as the qualified-paren path in `scanCallSites` — reuses the
 * same `callDedupe` / `synthFunctionNodeIds` sets so a paren and non-paren
 * form on the same line don't create duplicate edges.
 */
export function emitQualifiedStatementCallEdge(
  ctx: VbaExtractorContext,
  caller: ProcInfo,
  receiver: string,
  member: string,
  lineNum: number,
): void {
  // Eligibility is checked before this call. Project-class local receivers
  // resolve to their class name; undeclared receivers stay as raw module-name
  // candidates.
  const receiverType = ctx.resolveReceiverType(receiver);
  // Issue #245 — same runtime-object node gate as the paren-form path in
  // `scanCallSites`. Statement-form `Call VBA.DoEvents` / `col.Add x` must
  // not synthesize a stub node either. No `UnresolvedReference` was emitted
  // on this path before, and none is emitted now.
  if (isRuntimeObject(receiverType)) return;
  const qualified = `${receiverType}.${member}`;
  const dedupeKey = `${caller.name}->${qualified}@${lineNum}`;
  if (ctx.callDedupe.has(dedupeKey)) return;
  ctx.callDedupe.add(dedupeKey);

  const synthId = generateNodeId(ctx.filePath, 'function', qualified, lineNum);
  if (!ctx.synthFunctionNodeIds.has(synthId)) {
    ctx.synthFunctionNodeIds.add(synthId);
    ctx.nodes.push({
      id: synthId,
      kind: 'function',
      name: qualified,
      qualifiedName: qualified,
      filePath: ctx.filePath,
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
  ctx.edges.push({
    source: ctx.findOrCreateFunctionNodeId(caller),
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
 * Issue #43: normalize the receiver of a `With <expr>` block to a bare
 * identifier, or null when it is a keyword / runtime object / unparseable.
 */
export function normalizeWithReceiver(expr: string): string | null {
  let receiver = expr.trim();
  if (!receiver) return null;
  if (/^Call\s/i.test(receiver)) receiver = receiver.replace(/^Call\s+/i, '').trimStart();
  if (receiver.startsWith('[')) {
    const m = /^\[([^\]]+)\]/u.exec(receiver);
    if (!m) return null;
    receiver = m[1] ?? '';
  } else {
    const m = /^(\p{L}[\p{L}\p{N}_]*)/u.exec(receiver);
    if (!m) return null;
    receiver = m[1] ?? '';
  }
  if (!receiver) return null;
  if (CALL_KEYWORD_BLACKLIST.has(receiver)) return null;
  if (RUNTIME_RECEIVER_BLACKLIST.has(receiver)) return null;
  return receiver;
}

/**
 * Issue #43: detect a `.Member` call inside a `With` block (leading-dot
 * member reference that is a call, not a property assignment).
 */
export function detectWithMemberCall(line: string): { member: string } | null {
  let trimmed = line.trimStart();
  if (!trimmed) return null;
  if (/^Call\s/i.test(trimmed)) trimmed = trimmed.replace(/^Call\s+/i, '').trimStart();
  if (trimmed.startsWith("'") || /^Rem(\s|$)/i.test(trimmed)) return null;
  if (!trimmed.startsWith('.')) return null;
  const memberRest = trimmed.slice(1);
  const memberM = /^(\p{L}[\p{L}\p{N}_]*)/u.exec(memberRest);
  if (!memberM) return null;
  const member = memberM[1] ?? '';
  const afterMember = memberRest.slice(member.length);
  if (afterMember.length > 0) {
    const ch = afterMember.charAt(0);
    if (ch !== '(' && ch !== ' ' && ch !== '\t') return null;
    const argsText = afterMember.trimStart();
    if (argsText.startsWith('=')) return null;
  }
  if (CALL_KEYWORD_BLACKLIST.has(member)) return null;
  if (RUNTIME_RECEIVER_BLACKLIST.has(member)) return null;
  return { member };
}
