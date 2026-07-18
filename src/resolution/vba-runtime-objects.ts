/**
 * Canonical list of VBA/Access runtime objects and singletons whose
 * `Receiver.Member` calls are NEVER user-defined code — DAO, FileSystemObject
 * (`fso`), intrinsic collections, error/debug intrinsics, and Access
 * application singletons.
 *
 * Consumed by the post-extraction call-stub resolver
 * (`ReferenceResolver.resolveVbaCallStubTarget`, see `./index.ts`) to DECLINE
 * repointing a synthetic `calls` stub whose receiver is a runtime object and
 * has no real user declaration in the project. Declining keeps the edge
 * `stub:true`, so a consumer's `WHERE stub=true` guardrail returns GENUINE
 * missing callees only, free of runtime-object noise (issue #110, supersedes
 * #109).
 *
 * A user class or module that happens to share a runtime-object name (a
 * "shadow" declaration, e.g. a user `.cls` literally named `DAO`) is still
 * linked: the resolver runs its normal two-step name resolution FIRST and
 * only falls back to this list when no real target exists, so a shadow
 * declaration is repointed exactly like any other real symbol (FR-2.1).
 *
 * NOTE — this is complementary to, NOT a duplicate of, the VBA extractor's
 * own `RUNTIME_RECEIVER_BLACKLIST` (`src/extraction/vba/constants.ts`). That
 * set is case-sensitive PascalCase and suppresses stub SYNTHESIS for the
 * common receivers at extraction time; this set is lowercased and catches the
 * runtime-object stubs that still reached the graph (lowercase receivers,
 * `Dim x As DAO.*` typed locals whose receiver survives as raw text, etc.).
 * The extractor blacklist deliberately stays untouched (AC-4: `src/extraction`
 * has 0 diff); the two layers are independent by design.
 *
 * Entries are lowercased so matching is case-insensitive against a stub's
 * receiver.
 */
export const RUNTIME_OBJECTS: ReadonlySet<string> = new Set<string>([
  'dao',
  'fso',
  'err',
  'listbox',
  'combobox',
  'textbox',
  'forms',
  'reports',
  'debug',
  'collection',
  'vba',
  'application',
  'screen',
  'docmd',
  'currentdb',
  'currentproject',
  'codedata',
  'codeproject',
]);

/**
 * True iff `receiver` (any case) names a known VBA/Access runtime object.
 * Leading/trailing brackets and surrounding whitespace are stripped
 * defensively so a bracketed receiver (`[DAO]`) still matches.
 */
export function isRuntimeObject(receiver: string | null | undefined): boolean {
  if (!receiver) return false;
  const key = receiver.replace(/^\[/, '').replace(/\]$/, '').trim().toLowerCase();
  return RUNTIME_OBJECTS.has(key);
}

/** Canonical VBA and Access built-in functions that never resolve to project code. */
export const VBA_STDLIB_FUNCTIONS: ReadonlySet<string> = new Set([
  'cstr', 'cint', 'clng', 'cdbl', 'csng', 'cbyte', 'cbool', 'cdate', 'cverr',
  'isnull', 'isempty', 'isnumeric', 'isdate', 'isarray', 'isobject', 'ismissing',
  'typename', 'vartype', 'len', 'lenb', 'instr', 'instrb', 'instrrev', 'lcase',
  'ucase', 'left', 'leftb', 'right', 'rightb', 'mid', 'midb', 'replace', 'trim',
  'ltrim', 'rtrim', 'space', 'string', 'strconv', 'asc', 'ascb', 'ascw', 'chr',
  'chrb', 'chrw', 'format', 'format$', 'array', 'lbound', 'ubound', 'split',
  'join', 'filter', 'abs', 'int', 'fix', 'round', 'sgn', 'sqr', 'val', 'now',
  'date', 'time', 'dateadd', 'datediff', 'datepart', 'dateserial', 'datevalue',
  'timeserial', 'timevalue', 'year', 'month', 'day', 'hour', 'minute', 'second',
  'weekday', 'weekdayname', 'monthname', 'timer', 'msgbox', 'inputbox', 'shell',
  'environ', 'command', 'createobject', 'getobject', 'rgb', 'qbcolor', 'doevents',
  'nz', 'iif', 'fv', 'pv', 'pmt', 'rate', 'npv', 'irr', 'sln', 'syd', 'ddb', 'mirr',
]);

export function isVbaStdlibFunction(name: string | null | undefined): boolean {
  return !!name && VBA_STDLIB_FUNCTIONS.has(name.trim().toLowerCase());
}
