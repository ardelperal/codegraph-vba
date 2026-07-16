/**
 * Pure text helpers shared across the VBA extraction passes. None of these
 * touch extractor state — they are deterministic string transforms lifted
 * out of `VbaExtractor` verbatim so each pass module can import what it needs
 * without pulling in the orchestrator.
 */
import { isAccessEventName } from './events';

/**
 * Fold a VBA visibility keyword to the canonical lowercase enum, matching
 * the procedure convention: `Private` → 'private'; `Public`, `Global`,
 * `Friend`, or none → 'public' (VBA's default module-level `Const`/`Enum`
 * is Private, but we follow the same broader-than-private fold the proc
 * sweep uses so visibility is consistent across symbol kinds).
 */
export function foldVisibility(raw: string): 'public' | 'private' {
  return raw.trim().toLowerCase() === 'private' ? 'private' : 'public';
}

/**
 * Fix 2 (Issue #2): replace string-literal content with spaces so that
 * call-site patterns inside `"..."` spans are invisible to CALL_RE and
 * the statement-form detectors.  Column positions are preserved (each
 * character is replaced 1-for-1) so any col-based metadata stays correct.
 */
export function maskStringContent(line: string): string {
  let result = '';
  let i = 0;
  while (i < line.length) {
    const ch = line[i]!;
    if (ch === '"') {
      result += ' '; // opening quote masked
      i++;
      while (i < line.length) {
        const c = line[i]!;
        if (c === '"' && line[i + 1] === '"') {
          result += '  '; // doubled-quote escape masked (2 chars)
          i += 2;
        } else if (c === '"') {
          result += ' '; // closing quote masked
          i++;
          break;
        } else {
          result += ' '; // string content → space
          i++;
        }
      }
    } else {
      result += ch;
      i++;
    }
  }
  return result;
}

/**
 * #13 helper: escape a variable name for safe interpolation into the
 * self-reference RegExp built by `trackSqlVariableAssignment`. VBA
 * identifiers are alphanumeric+underscore only, so in practice nothing here
 * ever needs escaping — this guards against regex metacharacters anyway
 * rather than assume the input is always well-formed.
 */
export function escapeRegExpLiteral(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function parseConstDeclarations(
  body: string,
): Array<{ name: string; value: string | null }> {
  const declarations: Array<{ name: string; value: string | null }> = [];
  for (const part of splitOutsideVbaStrings(body, ',')) {
    const m =
      /^\s*(\p{L}[\p{L}\p{N}_]*)\s*(?:As\s+[^=]+?)?\s*=\s*(.+?)\s*$/iu.exec(part);
    if (!m) continue;
    const name = m[1] ?? '';
    const rawValue = (m[2] ?? '').trim();
    declarations.push({
      name,
      value: rawValue.startsWith('"') ? unwrapVbaStringLiteral(rawValue) : rawValue || null,
    });
  }
  return declarations;
}

export function splitOutsideVbaStrings(value: string, separator: string): string[] {
  const parts: string[] = [];
  let current = '';
  let inString = false;
  for (let i = 0; i < value.length; i++) {
    const ch = value[i];
    const next = value[i + 1];
    if (ch === '"') {
      current += ch;
      if (inString && next === '"') {
        current += next;
        i++;
        continue;
      }
      inString = !inString;
      continue;
    }
    if (!inString && ch === separator) {
      parts.push(current);
      current = '';
      continue;
    }
    current += ch;
  }
  parts.push(current);
  return parts;
}

/**
 * Issue #50 helper: return true iff `line.charAt(fromIndex..)` (after the
 * matched TempVars site, including any trailing whitespace) holds an `=`
 * (write-assignment) and not a `==` (VBA has no `==` operator, so a bare
 * check for `=` is safe). Used by `sweepTempVars` to classify a `TempVars`
 * access as read vs write from the local line context.
 *
 * We skip over trailing whitespace only — `.`, `(`, `<`, `>`, `*`, `+`
 * etc. all mean "this isn't an assignment target", and bare `=`
 * is the only access-suffix shape VBA syntax allows here. A line like
 * `TempVars!x = 1` (write) or `Debug.Print TempVars!x` (read) land in
 * the right bucket without further analysis.
 *
 * Strings should already be masked out of `line` (`maskStringContent`)
 * when this is called for the bang form (no string in scope anyway);
 * the paren form passes the ORIGINAL line — but for THAT form the
 * captured match ends at the closing `"` of the literal, so any `=`
 * that follows is unambiguously outside the string.
 */
export function detectAssignmentSuffix(line: string, fromIndex: number): boolean {
  let i = fromIndex;
  while (i < line.length) {
    const ch = line[i];
    if (ch === ' ' || ch === '\t') {
      i++;
      continue;
    }
    return ch === '=';
  }
  return false;
}

export function unwrapVbaStringLiteral(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed.startsWith('"')) return trimmed;
  let text = '';
  for (let i = 1; i < trimmed.length; i++) {
    const ch = trimmed[i];
    const next = trimmed[i + 1];
    if (ch === '"' && next === '"') {
      text += '"';
      i++;
      continue;
    }
    if (ch === '"') break;
    text += ch;
  }
  return text;
}

/**
 * Hueco 3 helper: parse an Access event-handler Sub name into its
 * `<ControlName>_<EventName>` components. Returns null when the name does
 * not match the convention, when the split yields an empty segment, or
 * when the control name is `Form` (form-level events are NOT control
 * handlers — `Form_Load` is the form's own lifecycle event, not a
 * command-button click).
 *
 * Splitting on the LAST underscore (rather than the first) lets
 * multi-word event names parse correctly: `ComandoAltaPM_BeforeDelConfirm`
 * yields control=`ComandoAltaPM`, event=`BeforeDelConfirm`, not
 * control=`ComandoAlta`, event=`PM_BeforeDelConfirm`.
 */
export function parseEventHandlerName(
  name: string,
): { controlName: string; eventName: string } | null {
  if (!name) return null;
  const lastUnderscore = name.lastIndexOf('_');
  if (lastUnderscore <= 0) return null; // no underscore OR starts with underscore
  const controlName = name.slice(0, lastUnderscore);
  const eventName = name.slice(lastUnderscore + 1);
  if (!controlName || !eventName) return null;
  // Form-level events live on the form, not on a control.
  if (controlName.toLowerCase() === 'form') return null;
  if (!isAccessEventName(eventName)) return null;
  return { controlName, eventName };
}
