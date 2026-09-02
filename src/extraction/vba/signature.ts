/**
 * Issue #250 (task T8): parse a VBA procedure's parameter list off its
 * declaration line.
 *
 * The line handed in here has already been through the preprocessor, so a
 * signature split with `_` continuations arrives as ONE logical line and a
 * trailing `'` comment is already gone. This module must never re-implement
 * continuation joining — see `vba-preprocess.ts` and issue #202.
 */

/** One declared parameter of a `Sub` / `Function` / `Property` signature. */
export interface VbaParameterInfo {
  /** Declared name, original casing, brackets unwrapped. */
  name: string;
  /** The `As <Type>` type, or `null` for an implicitly-`Variant` parameter. */
  type: string | null;
  /**
   * `true` unless the declaration says `ByVal`. VBA passes by reference by
   * default, so an unqualified parameter really is a `ByRef` one.
   */
  byRef: boolean;
  /** The literal `Optional` keyword was present. */
  optional: boolean;
  /** Declared as an array — `name()`. */
  isArray: boolean;
  /** An `= <default>` clause was present. */
  hasDefault: boolean;
}

/**
 * Declared arity. `total` is `null` — not `Infinity` — when the signature ends
 * in a `ParamArray`, because this value is serialised into node metadata and
 * crosses the parse-worker boundary, and JSON has no `Infinity`.
 */
export interface VbaArity {
  required: number;
  total: number | null;
}

/** The parsed shape of one procedure signature. */
export interface VbaSignatureInfo {
  params: VbaParameterInfo[];
  arity: VbaArity;
}

/**
 * `[Optional] [ByVal|ByRef|ParamArray] name[()] [As Type] [= default]`.
 *
 * The mode alternatives all demand trailing whitespace, so a parameter simply
 * named `x` cannot be mistaken for a mode keyword. The type allows dots so
 * `As DAO.Recordset` survives, and the bracketed `As [Type With Spaces]` form
 * is unwrapped.
 */
const PARAM_RE =
  /^\s*(Optional\s+)?(ByVal\s+|ByRef\s+|ParamArray\s+)?(?:\[([^\]]+)\]|(\p{L}[\p{L}\p{N}_]*))\s*(\(\s*\))?\s*(?:As\s+(?:\[([^\]]+)\]|(\p{L}[\p{L}\p{N}_.]*)))?\s*(=[\s\S]*)?$/iu;

/**
 * Find the parameter-list body that starts at the first `(` at or after
 * `searchFrom`, returning the text between the parentheses.
 *
 * The scan is depth- and string-aware so a default value such as
 * `= "a)b"` or a nested `Array(1, 2)` cannot end the list early. Returns
 * `null` when the declaration has no parameter list at all (`Public Sub Bare`)
 * or when the parentheses never close.
 *
 * Anchoring on `searchFrom` (the end of the procedure NAME) rather than on
 * `lastIndexOf(')')` is what keeps a colon-separated single-line procedure
 * (`Public Sub X(): Debug.Print "(y)": End Sub`) from swallowing its own body
 * as a parameter list.
 */
function findParameterListBody(
  line: string,
  searchFrom: number,
): string | null {
  let open = -1;
  for (let i = searchFrom; i < line.length; i++) {
    const ch = line[i];
    if (ch === ' ' || ch === '\t') continue;
    if (ch === '(') open = i;
    break;
  }
  if (open < 0) return null;

  let depth = 0;
  let inString = false;
  for (let i = open; i < line.length; i++) {
    const ch = line[i];
    if (inString) {
      if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === '(') depth++;
    else if (ch === ')') {
      depth--;
      if (depth === 0) return line.slice(open + 1, i);
    }
  }
  return null;
}

/**
 * Split a parameter-list body on its top-level commas. Depth- and
 * string-aware, so `Optional s As String = "a,b"` stays one entry.
 */
export function splitParameterList(body: string): string[] {
  if (!body.trim()) return [];
  const out: string[] = [];
  let depth = 0;
  let inString = false;
  let start = 0;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (inString) {
      if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === '(') depth++;
    else if (ch === ')') depth--;
    else if (ch === ',' && depth === 0) {
      out.push(body.slice(start, i));
      start = i + 1;
    }
  }
  out.push(body.slice(start));
  return out;
}

/**
 * Parse one raw parameter fragment. A fragment the regex cannot make sense of
 * still yields an entry — dropping it would silently under-count the arity,
 * and a wrong arity is worse than a coarse one.
 */
function parseParameter(raw: string): VbaParameterInfo {
  const m = PARAM_RE.exec(raw);
  if (!m) {
    return {
      name: raw.trim(),
      type: null,
      byRef: true,
      optional: false,
      isArray: false,
      hasDefault: false,
    };
  }
  const mode = (m[2] ?? '').trim().toLowerCase();
  return {
    name: m[3] ?? m[4] ?? '',
    type: m[6] ?? m[7] ?? null,
    byRef: mode !== 'byval',
    optional: Boolean(m[1]),
    isArray: Boolean(m[5]),
    hasDefault: Boolean(m[8]),
  };
}

/** `true` when the fragment declares a `ParamArray`. */
function isParamArray(raw: string): boolean {
  return /^\s*ParamArray\s+/i.test(raw);
}

/**
 * Parse the signature of the procedure whose name ends at `nameEndIndex`.
 *
 * A `ParamArray` tail is excluded from `arity.required` (it accepts zero
 * arguments) and forces `arity.total` to `null` (it accepts any number). The
 * `ParamArray` parameter itself is still reported in `params` with
 * `optional: false`, because that flag mirrors the literal `Optional` keyword
 * — which VBA forbids on a `ParamArray`.
 */
export function parseSignature(
  line: string,
  nameEndIndex: number,
): VbaSignatureInfo {
  const body = findParameterListBody(line, nameEndIndex);
  if (body === null) {
    return { params: [], arity: { required: 0, total: 0 } };
  }
  const fragments = splitParameterList(body);
  const params: VbaParameterInfo[] = [];
  let unbounded = false;
  let required = 0;
  for (const raw of fragments) {
    const param = parseParameter(raw);
    params.push(param);
    if (isParamArray(raw)) unbounded = true;
    else if (!param.optional) required++;
  }
  return {
    params,
    arity: { required, total: unbounded ? null : params.length },
  };
}
