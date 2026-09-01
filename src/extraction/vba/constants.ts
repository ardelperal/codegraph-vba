/**
 * Shared regexes and lookup sets used by more than one VBA extraction pass.
 * Kept in a leaf module so the per-concern pass modules can import them
 * without depending on each other or on the orchestrator.
 */

/** Sub/Function/Property regex — captures visibility prefix, kind, and name. */
export const PROC_RE =
  /^\s*((?:Public|Private|Friend|Static)\s+)?(?:Static\s+)?(Sub|Function|Property(?:\s+(?:Get|Let|Set))?)\s+(\p{L}[\p{L}\p{N}_]*)/iu;

/**
 * Issue #52: shared `End Sub` / `End Function` / `End Property` marker.
 * Promoted from a local regex in `sweepCallsAndSql` so `sweepEnumsAndConsts`
 * can walk the same proc boundaries and decide Const scope per line.
 * The `(?:^|:\s*)` prefix tolerates colon-separated single-line procs
 * (`Public Sub X(): ... : End Sub`). Classifiers must test procedure start
 * and end independently because both can occur on the same physical line.
 */
export const PROCEDURE_END_RE =
  /(?:^|:\s*)End\s+(?:Sub|Function|Property)\b/i;

/**
 * VBA primitive type names — skipped when emitted as Dim targets so
 * we don't pollute the graph with `As Long` / `As String` references.
 * Fix 4: all entries are LOWERCASE and the lookup lowercases the
 * captured type name so `As long`, `As LONG`, `As Long` all match.
 * Fix 1 (Issue #1): added `'new'` as a backstop so that if the `As New`
 * pattern is ever captured as a type name it is silently skipped.
 */
export const PRIMITIVE_TYPES = new Set([
  'long', 'integer', 'short', 'byte', 'single', 'double', 'currency',
  'string', 'boolean', 'date', 'variant', 'object', 'error',
  'empty', 'null', 'longptr', 'longlong', 'new',
]);

/** Keywords we never want to match as call receivers. */
export const CALL_KEYWORD_BLACKLIST = new Set([
  'If',
  'For',
  'While',
  'Do',
  'Select',
  'Case',
  'Then',
  'Else',
  'ElseIf',
  'With',
  'Loop',
  'Wend',
  'End',
  'Return',
  'Dim',
  'Set',
  'Let',
  'Const',
  'ReDim',
  'Static',
  'Public',
  'Private',
  'Friend',
  'Sub',
  'Function',
  'Property',
  'Class',
  'Module',
  'Option',
  'On',
  'Error',
  'Resume',
  'Exit',
  'New',
  'Call',
  'Rem',
  'LBound',
  'UBound',
  'Me',
  'Nothing',
]);

/**
 * VBA reserved words that cannot name an unqualified user-code target.
 * Stored lowercase so source casing cannot change classification.
 */
export const VBA_KEYWORDS = new Set([
  ...Array.from(CALL_KEYWORD_BLACKLIST, (keyword) => keyword.toLowerCase()),
  'and', 'append', 'as', 'base', 'binary', 'byref', 'byval', 'close',
  'compare', 'declare', 'empty', 'enum', 'eqv', 'erase', 'false', 'get',
  'global', 'gosub', 'goto', 'imp', 'input', 'kill', 'mod', 'mybase',
  'myclass', 'next', 'not', 'null', 'open', 'optional', 'or', 'output',
  'paramarray', 'preserve', 'print', 'put', 'seek', 'stop', 'true', 'type',
  'write', 'xor',
]);

export function isVbaKeyword(name: string): boolean {
  return VBA_KEYWORDS.has(name.toLowerCase());
}

/**
 * Access runtime objects and singletons whose calls the call-site scans drop
 * outright. Issue #245 collapsed this literal into the canonical
 * `RUNTIME_OBJECT_DEFS` in `./runtime-objects.ts`; it is re-exported here so
 * every existing `from './constants'` import keeps working unchanged.
 */
export { RUNTIME_RECEIVER_BLACKLIST } from './runtime-objects';
