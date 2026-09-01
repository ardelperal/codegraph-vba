/**
 * `VbaExtractionOptions` — the single object that carries every VBA
 * extraction knob from `codegraph.json` down to `VbaExtractor` (issue #243).
 *
 * ## Why this module exists
 *
 * The knobs used to travel as positional parameters through five files:
 *
 * ```
 * project-config.ts  loadVbaConfig()
 *   -> extraction/index.ts     reads vbaConfig
 *     -> extraction/parse-pool.ts  ParseTask fields
 *       -> extraction/parse-worker.ts  message shape
 *         -> extraction/tree-sitter.ts extractFromSource(..., vbaTargets, maxRaiseFanout, ...)
 *           -> new VbaExtractor(filePath, source, vbaTargets, maxRaiseFanout)
 * ```
 *
 * Every new knob meant editing all five signatures and every call site. One
 * object collapses that to "add a field here, read it in the extractor".
 *
 * ## Worker-boundary constraint — READ BEFORE ADDING A FIELD
 *
 * This object crosses the `parse-pool` -> `parse-worker` boundary through
 * `worker.postMessage`, which is `structuredClone`-based. Every field MUST be
 * plain, structured-cloneable data: primitives, plain objects, and arrays of
 * those. NO functions, NO `RegExp`, NO `Map`/`Set`, NO class instances — a
 * function throws `DataCloneError` at the boundary and a `RegExp` silently
 * loses its `lastIndex`/identity semantics.
 *
 * A derived value that is not cloneable (a compiled wrapper `RegExp`, a
 * lookup `Map`) belongs on the extractor context, built inside the extractor
 * from the plain data in this object.
 *
 * This module is a LEAF: it imports nothing, so the config layer, the pool,
 * the worker, and the extractor can all depend on it without a cycle.
 */

/**
 * Every VBA-specific extraction knob, in one structured-cloneable object.
 *
 * All fields are optional; `{}` is the zero-config default and must behave
 * exactly like the pre-#243 "caller passed nothing" path.
 */
export interface VbaExtractionOptions {
  /**
   * Conditional-compilation targets — the `#Const` name -> truth map used to
   * decide which `#If` branches stay active. Sourced from `codegraph.json` ->
   * `vba.targets`. Undefined means "no project overrides"; the preprocessor
   * falls back to its built-in defaults.
   */
  targets?: Record<string, boolean>;
  /**
   * Issue #152: per-file fanout cap for `RaiseEvent <EventName>` edges. An
   * event raised from more than this many sites in one file is flagged
   * `metadata.highFanout: true` and ALL its `raises-event` edges are dropped.
   * Undefined means the caller did not choose, so `VbaExtractor` applies
   * `DEFAULT_MAX_RAISE_FANOUT` (50). Sourced from `codegraph.json` ->
   * `vba.maxRaiseFanout`.
   */
  maxRaiseFanout?: number;
  /**
   * Names of project-specific procedures that wrap a SQL execution call
   * (e.g. a shared `EjecutarSQL(sql)` helper), so a call to one of them is
   * treated as a SQL execution site instead of an ordinary call.
   *
   * Declared here now and threaded end-to-end by #243; the SQL-wrapper task
   * is the consumer. Kept as a `readonly string[]` of plain strings — the
   * compiled matcher built from these names is NOT cloneable and therefore
   * belongs on the extractor context, never in this object.
   */
  sqlWrappers?: readonly string[];
}
