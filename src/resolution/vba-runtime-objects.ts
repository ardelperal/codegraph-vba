import {
  RUNTIME_OBJECTS as CANONICAL_RUNTIME_OBJECTS,
  isRuntimeObject as canonicalIsRuntimeObject,
} from '../extraction/vba/runtime-objects';

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
 * Issue #245 — this used to be a SECOND literal, complementary to the VBA
 * extractor's `RUNTIME_RECEIVER_BLACKLIST`. The two drifted: the extractor
 * list omitted `VBA`, `fso`, `Collection` and `ListBox`, so those receivers
 * kept synthesizing stub function NODES even though the resolver already
 * declined their edges. Both sets now derive from one literal in
 * `src/extraction/vba/runtime-objects.ts` (a leaf module — `src/extraction`
 * must not import from `src/resolution`, so the canonical set lives on the
 * extraction side and is re-exported here). Every existing
 * `from '../src/resolution/vba-runtime-objects'` import keeps working.
 *
 * Entries are lowercased so matching is case-insensitive against a stub's
 * receiver.
 */
export const RUNTIME_OBJECTS: ReadonlySet<string> = CANONICAL_RUNTIME_OBJECTS;

/**
 * True iff `receiver` (any case) names a known VBA/Access runtime object.
 * Leading/trailing brackets and surrounding whitespace are stripped
 * defensively so a bracketed receiver (`[DAO]`) still matches.
 */
export const isRuntimeObject = canonicalIsRuntimeObject;

/** Canonical VBA and Access built-in functions that never resolve to project code. */
const VBA_STDLIB_FUNCTIONS: ReadonlySet<string> = new Set([
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

/**
 * Issue #188 — VBA intrinsic constants and DAO enum values reach
 * `unresolved_refs` as `reference_kind='unqualified-ident'` and used to leak
 * as `failed`. They are bare identifiers (NOT calls), so unlike the stdlib
 * set they need NO `metadata.synthesizedBy` gate — name membership alone is
 * the discriminator. The same case-insensitive / lowercased-compare
 * convention as the stdlib set applies. A user-defined
 * `Public Const dbFailOnError = 0` or `Public Function vbCrLf() As String`
 * shadow resolves normally first and never reaches the classifier, so the
 * name-only gate does not weaken user-symbol resolution.
 *
 * The bench corpus flagged ~104 confirmed rows (DAO enum literals in
 * `Execute` / `OpenArgs` / recordset options; `vbCrLf` / `vbLf` / `vbTab`
 * in `Debug.Print`; the full `vbMsgBoxStyle` + `vbMsgBoxHelpButton` +
 * `vbApplicationModal` / `vbSystemModal` matrix in `MsgBox` calls).
 */
export const DAO_ENUM_VALUES: ReadonlySet<string> = new Set<string>([
  // Recordset / Database option flags (DAO enum values used as bare args).
  'dbfailonerror',
  'dbseechanges',
  'dbdenywrite',
  'dbdenyread',
  'dbreadonly',
  'dbappendonly',
]);

export function isDaoEnumValue(name: string | null | undefined): boolean {
  return !!name && DAO_ENUM_VALUES.has(name.trim().toLowerCase());
}

/**
 * Issue #188 — VBA intrinsic constants. Lowercased so matching is
 * case-insensitive. Covers the constants listed in the issue's explicit SQL
 * set plus the rest of the `vbMsgBoxStyle` family and the
 * `vbApplicationModal` / `vbSystemModal` modal flags.
 */
export const VBA_INTRINSIC_CONSTANTS: ReadonlySet<string> = new Set<string>([
  // Newline / control-char constants.
  'vbcrlf', 'vblf', 'vbcr', 'vbnewline', 'vbtab', 'vbnullchar', 'vbnullstring',
  // MsgBox icon constants (`vbMsgBoxStyle`).
  'vbexclamation', 'vbquestion', 'vbinformation', 'vbcritical',
  // MsgBox button-set constants.
  'vbokonly', 'vbokcancel', 'vbabortretryignore',
  'vbyesnocancel', 'vbyesno', 'vbretrycancel',
  // MsgBox default-button constants.
  'vbdefaultbutton1', 'vbdefaultbutton2', 'vbdefaultbutton3', 'vbdefaultbutton4',
  // MsgBox modality constants.
  'vbapplicationmodal', 'vbsystemmodal',
]);

export function isVbaIntrinsicConstant(name: string | null | undefined): boolean {
  return !!name && VBA_INTRINSIC_CONSTANTS.has(name.trim().toLowerCase());
}

/**
 * Issue #192 — central decision: should this unresolved reference be parked
 * as `declined-runtime` instead of `failed`?
 *
 * VBA stdlib calls reach `unresolved_refs` under three shapes:
 *
 *   1. **paren-form call** — `MsgBox("hi")` is emitted by the call-site sweep
 *      as `reference_kind='calls'` with `metadata.synthesizedBy` either
 *      absent or set to `vba-paren-call-unresolved` (the post-sweep
 *      unifier). The classic PR #185 / issue #181 path.
 *
 *   2. **statement-form call** — `MsgBox "hi"`, `DoEvents`, `Shell "calc"`
 *      are emitted by the statement-call sweep as
 *      `reference_kind='unqualified-ident'` with
 *      `metadata.synthesizedBy='vba-statement-call-unresolved'`. Without
 *      this classifier, those 49+ rows from the bench corpus (MsgBox=44,
 *      DoEvents=4, Shell=1) leak as `failed`, polluting actionable
 *      failed-reference reports.
 *
 *   3. **NOT a call at all** — a bare `Public Const SomePublicConst = 1`
 *      read inside a procedure (e.g. `SomePublicConst` on its own line,
 *      same shape the const-first disambiguation rule checks) ALSO reaches
 *      `unresolved_refs` as `reference_kind='unqualified-ident'`. The
 *      resolver first attempts normal `matchFunctionRef` / matchConstRef
 *      resolution; for genuine misses this row will be deleted from
 *      unresolved_refs. When a Const cannot be resolved but still produces
 *      a row, classifying its status purely by name is wrong — a
 *      user-defined `Public Sub ProjectHelper s` and a public Const that
 *      elude resolution must NOT be marked `declined-runtime`.
 *
 * Issue #188 extends the contract with two new name-only branches for
 * bare DAO enum values and VBA intrinsic constants — these are emitted as
 * `reference_kind='unqualified-ident'` WITHOUT a `synthesizedBy` flag
 * because they are identifiers, not calls. Name membership alone is the
 * discriminator; the metadata gate does NOT apply. Normal user-symbol
 * resolution runs first, so a user-defined `Public Const dbFailOnError`
 * or `Public Function vbCrLf` shadow never reaches this classifier.
 *
 * The stdlib branch keeps its original tight gate: name match AND one of
 * the two real call shapes (paren-form `calls`, or statement-form
 * `unqualified-ident` stamped with `vba-statement-call-unresolved`).
 * Name-only matching for the stdlib set is still intentionally rejected —
 * a row like `MsgBox` on its own line without the statement-call stamp is
 * a user-defined reference, not a runtime call.
 */
export function classifyVbaReferenceAsRuntime(ref: {
  language?: string | null;
  referenceKind?: string | null;
  referenceName?: string | null;
  metadata?: { synthesizedBy?: unknown } | null;
}): boolean {
  if ((ref.language ?? '').toLowerCase() !== 'vba') return false;
  // Issue #188 — bare DAO enum values and VBA intrinsic constants are
  // emitted as `reference_kind='unqualified-ident'` with no
  // `synthesizedBy` stamp (they're identifiers, not calls). Name
  // membership alone is the discriminator; the stdlib metadata gate does
  // NOT apply to these sets.
  if (ref.referenceKind === 'unqualified-ident') {
    if (isDaoEnumValue(ref.referenceName)) return true;
    if (isVbaIntrinsicConstant(ref.referenceName)) return true;
  }
  if (!isVbaStdlibFunction(ref.referenceName)) return false;
  const kind = ref.referenceKind;
  if (kind === 'calls') return true;
  const synthesizedBy = ref.metadata?.synthesizedBy;
  if (kind === 'unqualified-ident' && synthesizedBy === 'vba-statement-call-unresolved') {
    return true;
  }
  return false;
}
