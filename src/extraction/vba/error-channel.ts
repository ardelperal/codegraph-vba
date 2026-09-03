/**
 * The error channel — the module-level variable this kind of Access codebase
 * actually propagates error messages through (issue #261, task E4 of
 * `docs/vba-error-handling-plan.md`).
 *
 * ## Why a whole module for four names
 *
 * §2.1/§2.3 of the plan measured it: **16 handlers out of 3,774 re-raise**.
 * `Err.Raise 1000` unwinds exactly one frame and the guard
 * `If Err.Number <> 1000` in the handler means *"an inner procedure already
 * wrote a human-readable message — do not overwrite it"*. The MESSAGE never
 * travels through VBA's error mechanism; it travels through a field:
 *
 * ```vba
 * errores:                              ' inner procedure
 *     If Err.Number <> 1000 Then
 *         m_Error = "Riesgo.Guardar ha devuelto el error: " & Err.Description
 *     End If
 *     Me.Error = m_Error
 * ```
 * ```vba
 * p_Error = m_ObjRiesgoActivo.Error      ' the form, later
 * If p_Error <> "" Then MsgBox p_Error
 * ```
 *
 * That is module-variable data flow, which issue #251 already models as
 * `property-set` / `property-get` references onto a `variable` node. This
 * module supplies the one thing #251 cannot know: **which** of those variables
 * is the error channel. Nothing here emits a node, an edge or a reference — it
 * only decides whether an existing row earns `metadata.errorChannel: true`.
 *
 * ## Why it is a LEAF, separate from `errors.ts`
 *
 * Two consumers need the same list and must never fork:
 *
 *   - `errors.ts` derives `errorPolicy.behavior` from a channel-WRITE matcher
 *     (is this statement `<channel> = …`?), and
 *   - `module-vars.ts` flags a read or a write by NAME (is this variable the
 *     channel?).
 *
 * `errors.ts` owned both the names and the write matcher before this task and
 * its own comment deferred the config knob here. Keeping the names in
 * `errors.ts` and importing them from `module-vars.ts` would have closed a
 * cycle the moment `errors.ts` needed the compiled, config-aware form back. So
 * the names and both matchers live here, in a module that imports only
 * `text-utils`, and `errors.ts` re-exports the constant for its existing
 * consumers.
 *
 * ## Names, never substrings, never user regex
 *
 * A channel entry is a bare VBA identifier and is matched as a WHOLE name, so
 * `ErrorCount` cannot match `Error`. `codegraph.json` → `vba.errorChannel`
 * therefore takes identifiers, not patterns: this runs on every identifier of
 * every line of every VBA file, which is precisely where a user-supplied regex
 * is a catastrophic-backtracking hazard (the same reasoning `vba.sqlWrappers`
 * was built on in #244).
 */
import { escapeRegExpLiteral } from './text-utils';

/**
 * The channel names used when `codegraph.json` → `vba.errorChannel` is absent.
 *
 * `m_Error` / `p_Error` are the private module-variable spellings, `g_Error`
 * the global one, and `Error` the public field a caller reads off the object
 * (`m_ObjRiesgoActivo.Error`). Identical to the probe's
 * `DEFAULT_ERROR_CHANNEL_NAMES` — `scripts/vba-coverage-probe.mjs` measured
 * the census this task is accepted against, so the two lists must agree.
 */
export const DEFAULT_ERROR_CHANNEL_NAMES: readonly string[] = [
  'm_Error',
  'p_Error',
  'g_Error',
  'Error',
];

/** A channel entry is a bare VBA identifier — nothing qualified, no regex. */
export const ERROR_CHANNEL_ENTRY_RE = /^\p{L}[\p{L}\p{N}_]*$/u;

/**
 * The per-extractor compiled channel. Built ONCE per `VbaExtractor` and parked
 * on `VbaExtractorContext`, for the same reason `CompiledSqlWrappers` is: the
 * name lookup runs per identifier per line, and the write matcher per statement
 * per line, so neither may be rebuilt inside the line loop.
 *
 * NOT structured-cloneable (it holds a `Set` and `RegExp`s) — which is exactly
 * why `VbaExtractionOptions.errorChannel` carries plain strings across the
 * worker boundary and the compilation happens on the far side.
 */
export interface CompiledErrorChannel {
  /** The channel names, defaults first, project entries appended. Display form. */
  readonly names: readonly string[];
  /** Lowercased `names`, for the per-identifier membership test. VBA is case-insensitive. */
  readonly lookup: ReadonlySet<string>;
  /**
   * One anchored regex per name: the channel in the assignment-TARGET position
   * of a statement — bare (`p_Error = …`), `Me.`-qualified (`Me.Error = …`) or
   * object-qualified (`obj.Error = …`).
   *
   * Anchoring at the start of the statement is the whole precision of this
   * matcher: it is what separates a WRITE from a READ, so the house guard
   * `If m_Error <> "" Then` and the copy `x = m_Error` both correctly miss.
   * Kept identical to the probe's `buildChannelWriteMatcher`.
   */
  readonly writeRes: readonly RegExp[];
}

/**
 * Compile the channel from the plain-string config entries.
 *
 * `configured` entries are APPENDED to {@link DEFAULT_ERROR_CHANNEL_NAMES},
 * never substituted for them — same contract as `vba.sqlWrappers`: a project
 * that names its own `lastFailure` field must not lose `m_Error` in the trade,
 * because in a codebase this size the two conventions coexist. Blank, dotted
 * and otherwise non-identifier entries are dropped here; `project-config.ts`
 * has already warned about them at load time.
 */
export function compileErrorChannel(
  configured?: readonly string[],
): CompiledErrorChannel {
  const names: string[] = [];
  const lookup = new Set<string>();
  for (const raw of [...DEFAULT_ERROR_CHANNEL_NAMES, ...(configured ?? [])]) {
    const entry = (raw ?? '').trim();
    if (!entry || !ERROR_CHANNEL_ENTRY_RE.test(entry)) continue;
    const key = entry.toLowerCase();
    if (lookup.has(key)) continue;
    lookup.add(key);
    names.push(entry);
  }
  const writeRes = names.map(
    (name) =>
      new RegExp(
        `^\\s*(?:Set\\s+)?(?:(?:Me|\\p{L}[\\p{L}\\p{N}_]*)\\s*\\.\\s*)?${escapeRegExpLiteral(name)}\\s*=(?!=)`,
        'iu',
      ),
  );
  return { names, lookup, writeRes };
}

/**
 * The zero-config channel, compiled once.
 *
 * The fallback for every context that was never handed options — tests and
 * out-of-repo callers of `VbaExtractor`. Memoised because the write regexes
 * are the hot path's and there is exactly one correct instance of them.
 */
let defaultChannel: CompiledErrorChannel | null = null;
export function defaultErrorChannel(): CompiledErrorChannel {
  if (!defaultChannel) defaultChannel = compileErrorChannel();
  return defaultChannel;
}

/**
 * Is `name` the error channel? A WHOLE-name, case-insensitive test — the
 * single reason `ErrorCount` does not match `Error`, pinned by a test.
 */
export function isErrorChannelName(
  channel: CompiledErrorChannel,
  name: string,
): boolean {
  return channel.lookup.has(name.toLowerCase());
}
