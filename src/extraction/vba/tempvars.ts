/**
 * TempVars sweep (issue #50). Models each STATIC-LITERAL Access TempVars key
 * as a synthetic `class` placeholder node (cross-file id-stable) and emits one
 * `references` edge per reading/writing procedure with
 * `metadata.synthesizedBy: 'vba-tempvar'` and `metadata.access` ∈ {read, write}.
 */
import { generateNodeId } from '../tree-sitter-helpers';
import { detectAssignmentSuffix } from './text-utils';
import { VbaExtractorContext, ProcInfo } from './context';

/**
 * Issue #50: TempVars — Access's global key-value store for cross-form
 * state. Three scanner surfaces cover the four real idioms:
 *   - `TEMP_VAR_BANG_RE`  — `TempVars!clave`         (no parens; write or read)
 *   - `TEMP_VAR_PAREN_RE` — `TempVars("clave")`      (parens; write or read)
 *   - `TEMP_VAR_ADD_RE`   — `TempVars.Add "clave", v` (always a write)
 *
 * Bang vs paren split by line-source: the bang form has no string literals
 * in scope, so we scan the MASKED line (`maskStringContent` replaces
 * `"…"` content with spaces). The paren and Add forms have their key INSIDE
 * a `"…"` literal that gets blanked by the masker, so those regexes scan
 * the ORIGINAL (unmasked) line.
 *
 * Dynamic-key forms — `TempVars(strNombre)`, `TempVars("clave" & suffix)` —
 * are silently unmatched by all three regexes (none tolerate a function-call
 * or `&` arg). REQ-CODE-4 "unresolvable is silent" applies.
 */
const TEMP_VAR_BANG_RE = /\bTempVars!\s*(\p{L}[\p{L}\p{N}_]*)/gu;

/**
 * Issue #50 (cont.):
 * `TempVars("clave")` / `TempVars(  "clave"  )` capture. Scanned
 * over the original (unmasked) line — the literal lives INSIDE a
 * string and would be stripped by `maskStringContent`.
 */
const TEMP_VAR_PAREN_RE = /\bTempVars\s*\(\s*"([^"]+)"\s*\)/gu;

/**
 * Issue #50 (cont.):
 * `TempVars.Add "clave", value` capture. Always a write (the
 * `.Add` method inserts/updates the entry). Scanned over the
 * original (unmasked) line for the same string-literal reason as
 * `TEMP_VAR_PAREN_RE`.
 */
const TEMP_VAR_ADD_RE = /\bTempVars\.Add\s+"([^"]+)"\s*,/gi;

/**
 * Issue #50: emit one TempVar reading/writing site. Per call:
 *   - placeholder `class` node keyed on the synthetic
 *     `synthetic:tempvar/<key>` file path so cross-file extraction
 *     calls collapse to one node per key,
 *   - `references` edge from the calling `function` node
 *     (via `findOrCreateFunctionNodeId`) carrying
 *     `metadata.synthesizedBy: 'vba-tempvar'` AND
 *     `metadata.access: 'read' | 'write'`.
 *
 * Skips when there is no caller (module-level code) — REQ-CODE-4
 * "unresolvable/runtime reference is silent".
 */
function emitTempVarReference(
  ctx: VbaExtractorContext,
  caller: ProcInfo | undefined,
  key: string,
  lineNum: number,
  column: number,
  access: 'read' | 'write',
): void {
  if (!caller) return;
  if (!key) return;
  const syntheticFilePath = `synthetic:tempvar/${key}`;
  const targetId = generateNodeId(
    syntheticFilePath,
    'class', // placeholder kind — see SQL_TABLE_RE emitReference for precedent
    key,
    0, // stable; line-independent per Fix 5 / cross-form id stability
  );
  if (!ctx.synthTempVarNodeIds.has(targetId)) {
    ctx.synthTempVarNodeIds.add(targetId);
    ctx.nodes.push({
      id: targetId,
      kind: 'class',
      name: key,
      qualifiedName: key,
      filePath: syntheticFilePath,
      language: 'vba',
      startLine: lineNum,
      endLine: lineNum,
      startColumn: column,
      endColumn: column + key.length,
      updatedAt: Date.now(),
    });
  }
  ctx.edges.push({
    source: ctx.findOrCreateFunctionNodeId(caller),
    target: targetId,
    kind: 'references',
    provenance: 'heuristic',
    metadata: {
      synthesizedBy: 'vba-tempvar',
      // User-facing enum: lowercase single-token string values.
      access,
    },
    line: lineNum,
    column,
  });
}

/**
 * Issue #50: scan one line of VBA source for TempVars access sites and
 * emit one `references` edge per site. Bang form scans the MASKED line,
 * paren + Add forms scan the ORIGINAL line. Access is classified by the
 * line suffix after the match: a bare `=` next non-whitespace char is a
 * write (VBA has no `==`).
 */
export function sweepTempVars(
  ctx: VbaExtractorContext,
  maskedLine: string,
  originalLine: string,
  lineNum: number,
  caller: ProcInfo | undefined,
): void {
  // 1) Bang form — masked line. No string-literal interaction.
  const bangRe = new RegExp(TEMP_VAR_BANG_RE.source, TEMP_VAR_BANG_RE.flags);
  let bm: RegExpExecArray | null;
  while ((bm = bangRe.exec(maskedLine)) !== null) {
    const key = bm[1] ?? '';
    if (!key) continue;
    const access = detectAssignmentSuffix(maskedLine, bm.index + bm[0].length)
      ? 'write'
      : 'read';
    emitTempVarReference(ctx, caller, key, lineNum, bm.index, access);
  }

  // 2) Paren form — original line. The literal survives only here.
  const parenRe = new RegExp(TEMP_VAR_PAREN_RE.source, TEMP_VAR_PAREN_RE.flags);
  let pm: RegExpExecArray | null;
  while ((pm = parenRe.exec(originalLine)) !== null) {
    const key = pm[1] ?? '';
    if (!key) continue;
    const access = detectAssignmentSuffix(originalLine, pm.index + pm[0].length)
      ? 'write'
      : 'read';
    emitTempVarReference(ctx, caller, key, lineNum, pm.index, access);
  }

  // 3) Add form — original line. Always a write.
  const addRe = new RegExp(TEMP_VAR_ADD_RE.source, TEMP_VAR_ADD_RE.flags);
  let am: RegExpExecArray | null;
  while ((am = addRe.exec(originalLine)) !== null) {
    const key = am[1] ?? '';
    if (!key) continue;
    emitTempVarReference(ctx, caller, key, lineNum, am.index, 'write');
  }
}
