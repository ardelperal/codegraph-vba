/**
 * Module-level variable read/write sweep (issue #251).
 *
 * The dims classifier emits one `variable` node per module-level
 * declaration and registers the name on `ctx.moduleVariables`. A node
 * nobody points at answers nothing, so this sweep supplies the other
 * half: for every procedure body line, an identifier that matches one of
 * THIS FILE's module-level variable names becomes an
 * `UnresolvedReference` tagged `synthesizedBy: 'vba-module-var'`, with
 * `referenceKind: 'property-get'` for a read and `'property-set'` for a
 * write.
 *
 * The gate is the whole design. The sweep never looks for "identifiers
 * that might be variables"; it only ever asks whether a name is already
 * registered as a module-level variable of the module being extracted.
 * Any looser rule turns a codebase-wide sweep into thousands of
 * references to names that merely collide.
 */
import { VbaExtractorContext, ProcInfo } from './context';
import { isDirectAssignment } from './controls';

/**
 * Every bare identifier on the line. The sweep is a membership test
 * against `ctx.moduleVariables`, so the pattern is deliberately generic —
 * the filtering happens on the name, not on the shape around it.
 */
const IDENTIFIER_RE = /\p{L}[\p{L}\p{N}_]*/gu;

/**
 * A `Set` that is the ENTIRE prefix of the matched identifier, i.e. the
 * line reads `Set <var> = …`. Stripping it lets the shared
 * `isDirectAssignment` predicate see the same "nothing before the name,
 * `=` after it" shape it was written for, so `Set gblConn = New …` is
 * classified as the write it is instead of a read.
 *
 * `Set` cannot be stripped unconditionally: in `x = Set` — or in any
 * expression where `Set` is not the leading statement keyword — the
 * prefix is not blank, and `isDirectAssignment` already answers "read"
 * correctly.
 */
const LEADING_SET_RE = /^\s*Set\s+$/i;

/**
 * Per-context de-dup of the emitted references, keyed by
 * `(procedure, variable, direction)`.
 *
 * A procedure that reads `gblConn` on twenty lines couples to it once;
 * twenty identical rows would not make that coupling any more visible,
 * and the volume scales with the size of the corpus rather than with the
 * number of real relationships. Reads and writes are keyed separately, so
 * a procedure that both reads and writes a variable still reports both —
 * which is the distinction the direction is there to make. The first
 * occurrence keeps the line number, mirroring the same trade-off the
 * `Me.<Control>` sweep already makes.
 */
const seenModuleVarRefs = new WeakMap<VbaExtractorContext, Set<string>>();

/**
 * Scan one procedure-body line for reads and writes of this module's
 * module-level variables.
 *
 * `line` must be the string-literal-masked line: a variable name that
 * appears inside `"…"` is prose, not an access.
 */
export function scanModuleVariableReferences(
  ctx: VbaExtractorContext,
  line: string,
  from: ProcInfo,
  lineNum: number,
): void {
  if (ctx.moduleVariables.size === 0) return;
  // Issue #205 scoping: names this procedure declares itself — its
  // parameters and its own `Dim` / `Static` declarations — are locals that
  // happen to share a name with module state. Reading one of those is not
  // a read of the module variable.
  const shadowed = ctx.procLocalNames.get(String(from.startLine));

  let seen = seenModuleVarRefs.get(ctx);
  if (!seen) {
    seen = new Set<string>();
    seenModuleVarRefs.set(ctx, seen);
  }

  IDENTIFIER_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = IDENTIFIER_RE.exec(line)) !== null) {
    const name = m[0];
    const key = name.toLowerCase();
    if (!ctx.moduleVariables.has(key)) continue;
    if (shadowed?.has(key)) continue;

    // `obj.gblConn` / `Me!gblConn` names a member of something else that
    // happens to be spelled like this module's variable. The module
    // variable itself is only ever reached by its bare name.
    const previous = m.index > 0 ? line.charAt(m.index - 1) : '';
    if (previous === '.' || previous === '!') continue;

    const before = line.slice(0, m.index);
    const after = line.slice(m.index + name.length);
    const isWrite = isDirectAssignment(before.replace(LEADING_SET_RE, ''), after);
    const referenceKind = isWrite ? 'property-set' : 'property-get';

    const fromNodeId = ctx.findOrCreateFunctionNodeId(from);
    const dedupeKey = `${fromNodeId}\0${key}\0${referenceKind}`;
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);

    ctx.unresolvedReferences.push({
      fromNodeId,
      referenceName: ctx.moduleVariables.get(key)!.name,
      referenceKind,
      line: lineNum,
      column: m.index,
      filePath: ctx.filePath,
      language: 'vba',
      metadata: {
        synthesizedBy: 'vba-module-var',
        access: isWrite ? 'write' : 'read',
      },
    });
  }
}
