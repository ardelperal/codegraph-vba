/**
 * Canonical list of VBA/Access runtime objects and singletons whose
 * `Receiver.Member` calls are NEVER user-defined code — DAO, FileSystemObject
 * (`fso`), intrinsic collections, error/debug intrinsics, late-binding
 * factories, and Access application singletons.
 *
 * Issue #245: this used to be TWO literal sets written at different times.
 * `RUNTIME_OBJECTS` (formerly in `src/resolution/vba-runtime-objects.ts`)
 * decided whether a synthetic call EDGE was stamped `declined-runtime`;
 * `RUNTIME_RECEIVER_BLACKLIST` (formerly a second literal in
 * `./constants.ts`) decided whether the call-site scan looked at a receiver
 * at all. Because the second list omitted `VBA`, `fso`, `Collection`,
 * `ListBox` and friends, the extractor still synthesized a `function` NODE
 * for every one of those calls — 6,132 stub nodes against 3,840 real
 * procedures on the reference corpus. Edge consumers can filter on
 * `repointDecision`; symbol search and node counts cannot.
 *
 * The two sets are now one literal (`RUNTIME_OBJECT_DEFS`) with two derived
 * views, so they can never drift apart again:
 *
 *   - `RUNTIME_OBJECTS` — every canonical name, lowercased, for the
 *     case-insensitive `isRuntimeObject` predicate. Consumed by the
 *     post-extraction stub resolver (`ReferenceResolver.resolveVbaCallStubTarget`)
 *     to DECLINE repointing a stub whose receiver is a runtime object, and by
 *     the extractor's stub-NODE gate in `./calls.ts`.
 *   - `RUNTIME_RECEIVER_BLACKLIST` — the case-sensitive PascalCase subset
 *     flagged `blocksCallScan`. These receivers/members are dropped by the
 *     call-site scans BEFORE any reference is surfaced, so the set is
 *     deliberately narrow: widening it would silently delete
 *     `unresolved_refs` rows. Widening the canonical set is safe; promoting
 *     an entry to `blocksCallScan` is not.
 *
 * A user class or module that happens to share a runtime-object name (a
 * "shadow" declaration, e.g. a user `.cls` literally named `DAO`) is still
 * linked by the resolver: it runs its normal two-step name resolution FIRST
 * and only falls back to this list when no real target exists (FR-2.1).
 *
 * This module is a LEAF: it imports nothing. `src/extraction/` must not
 * depend on `src/resolution/`, so the canonical set lives here and
 * `src/resolution/vba-runtime-objects.ts` re-exports it — every existing
 * resolver import keeps working unchanged.
 */

interface RuntimeObjectDef {
  /** Canonical PascalCase (or conventional) spelling of the runtime object. */
  readonly name: string;
  /**
   * True when the call-site scans must drop this receiver/member outright
   * (`RUNTIME_RECEIVER_BLACKLIST`). These are the receivers that appear
   * literally in source text, so a case-sensitive match is sufficient and a
   * skip here loses no `unresolved_refs` row that a consumer relies on.
   *
   * Note: `DoCmd.RunSQL`, `DoCmd.OpenForm`, etc. still get SQL/edge tracking
   * via the dedicated `SQL_WRAPPERS` regex path (REQ-CODE-8), which fires
   * BEFORE this scan and uses its own dispatch — so blocking `DoCmd` here
   * doesn't lose the SQL-flow edges.
   */
  readonly blocksCallScan?: true;
}

/** THE single literal. Both exported sets below are views over it. */
const RUNTIME_OBJECT_DEFS: readonly RuntimeObjectDef[] = [
  // Access application singletons.
  { name: 'Application', blocksCallScan: true },
  { name: 'DoCmd', blocksCallScan: true },
  { name: 'SysCmd', blocksCallScan: true },
  { name: 'Screen', blocksCallScan: true },
  // Access object collections.
  { name: 'Forms', blocksCallScan: true },
  { name: 'Reports', blocksCallScan: true },
  { name: 'Modules', blocksCallScan: true },
  { name: 'References', blocksCallScan: true },
  { name: 'CommandBars', blocksCallScan: true },
  // VBA debugging / error-handling intrinsics.
  { name: 'Debug', blocksCallScan: true },
  { name: 'Err', blocksCallScan: true },
  // Late-binding factories (return IDispatch — not user code).
  { name: 'CreateObject', blocksCallScan: true },
  { name: 'GetObject', blocksCallScan: true },
  // DAO/ADO recordset field collection access (e.g. rcdDatos.Fields("ID")).
  { name: 'Fields', blocksCallScan: true },
  // Data-access libraries and the current-database singletons.
  { name: 'DAO' },
  { name: 'CurrentDb' },
  { name: 'CurrentProject' },
  { name: 'CodeData' },
  { name: 'CodeProject' },
  // Scripting runtime.
  { name: 'fso' },
  // The VBA library namespace itself — `VBA.DoEvents`, `VBA.Format`, ...
  { name: 'VBA' },
  // Intrinsic collection type.
  { name: 'Collection' },
  // Access control types used as declared local-variable types.
  { name: 'ListBox' },
  { name: 'ComboBox' },
  { name: 'TextBox' },
];

/**
 * Every canonical runtime-object name, lowercased so matching is
 * case-insensitive against a receiver as it appears in source.
 */
export const RUNTIME_OBJECTS: ReadonlySet<string> = new Set<string>(
  RUNTIME_OBJECT_DEFS.map((def) => def.name.toLowerCase()),
);

/**
 * Derived view: the case-sensitive receivers/members the call-site scans drop
 * before surfacing anything. NOT a second literal — see `blocksCallScan`.
 */
export const RUNTIME_RECEIVER_BLACKLIST: ReadonlySet<string> = new Set<string>(
  RUNTIME_OBJECT_DEFS.filter((def) => def.blocksCallScan).map((def) => def.name),
);

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
