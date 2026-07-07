/**
 * Control-reference sweeps: `Me.<Control>` / `Me!<Control>` (hueco 1, issue
 * #44) and cross-form `Forms!<Form>` / `Forms("<Form>")!<Ctl>` bang refs
 * (issue #44). Both emit `UnresolvedReference`s (no synthetic `function`
 * node) the resolver later binds to the form's controls.
 */
import { VbaExtractorContext, ProcInfo } from './context';

/**
 * `Me.<ControlName>` / `Me!<ControlName>` reference capture — hole 1
 * of VBA control-modeling. Issue #44 extended it from `Me\.` to `Me[.!]`
 * so the bang form (default-collection shortcut) is captured
 * byte-identically to the dot form. Only the FIRST identifier after
 * `Me.`/`Me!` is captured; trailing `.Caption`/`.Value` are properties of
 * the control, not new symbols. Provenance: `vba-me-control`.
 */
const ME_CONTROL_RE = /\bMe[.!](\p{L}[\p{L}\p{N}_]*)/gu;

/**
 * Issue #44: `Forms!<FormName>` / `Forms("<FormName>")!<Ctl>` cross-form
 * reference capture — companion to `ME_CONTROL_RE`. `Forms` is in
 * `RUNTIME_RECEIVER_BLACKLIST`, so the generic CALL_RE path skips it; this
 * dedicated scanner surfaces the form → control binding as an
 * `UnresolvedReference` (`vba-forms-bang`), with NO synthetic `function`
 * node (W4 invariant). The bang alternative's `(?![.\w])` lookahead drops
 * `Forms!FormX.Foo` (a property access on the form, not a control access).
 *
 * Operates on the ORIGINAL (unmasked) line — the paren form
 * `Forms("FormX")!txtY` carries the form name INSIDE a string literal.
 */
const FORMS_BANG_RE =
  /\b(?:Forms!(\p{L}[\p{L}\p{N}_]*|\[[^\]]+\])(?![.\w])|Forms\(\s*(?:"((?:[^"]|"")*)"|(\p{L}[\p{L}\p{N}_]*|\[[^\]]+\]))\s*\)\s*!\s*(?:\p{L}[\p{L}\p{N}_]*|\[[^\]]+\]))/gu;

/**
 * Hueco 1: scan a line for `Me.<ControlName>` / `Me!<ControlName>`
 * patterns and emit one UnresolvedReference per occurrence, tagged
 * `metadata.synthesizedBy: 'vba-me-control'`. Operates on the masked
 * `callScanLine`. The +3 column offset skips the 3-char `Me.` / `Me!`
 * prefix (both are 3 chars).
 */
export function scanMeControlReferences(
  ctx: VbaExtractorContext,
  line: string,
  from: ProcInfo,
  lineNum: number,
): void {
  ME_CONTROL_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = ME_CONTROL_RE.exec(line)) !== null) {
    const controlName = m[1] ?? '';
    if (!controlName) continue;
    ctx.unresolvedReferences.push({
      fromNodeId: ctx.findOrCreateFunctionNodeId(from),
      referenceName: controlName,
      referenceKind: 'references',
      line: lineNum,
      column: m.index + 3, // +3 to skip the `Me.` / `Me!` prefix
      filePath: ctx.filePath,
      language: 'vba',
      metadata: { synthesizedBy: 'vba-me-control' },
    });
  }
}

/**
 * Issue #44: scan a line for cross-form bang references
 * (`Forms!<FormName>[!<Ctl>]` and `Forms("<FormName>")!<Ctl>`) and
 * emit ONE UnresolvedReference per match with `metadata.synthesizedBy
 * = 'vba-forms-bang'`. Operates on the ORIGINAL (unmasked) line. The form
 * identifier is unwrapped of surrounding `"` quotes and `[…]` brackets.
 */
export function scanFormsBang(
  ctx: VbaExtractorContext,
  line: string,
  from: ProcInfo,
  lineNum: number,
): void {
  FORMS_BANG_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = FORMS_BANG_RE.exec(line)) !== null) {
    // Group 1: bang form name (bare or `[bracketed]`); group 2: paren
    // form name in `"quotes"`; group 3: paren form name bare or
    // `[bracketed]`. Take whichever the regex alternative produced and
    // strip the surrounding decoration so the public referenceName is
    // the bare form identifier.
    const raw = m[1] ?? m[2] ?? m[3] ?? '';
    if (!raw) continue;
    const formName = raw
      .replace(/^"|"$/g, '')   // strip surrounding string-literal quotes
      .replace(/^\[|\]$/g, ''); // strip surrounding bracket decoration (#54)
    if (!formName) continue;
    ctx.unresolvedReferences.push({
      fromNodeId: ctx.findOrCreateFunctionNodeId(from),
      referenceName: formName,
      referenceKind: 'references',
      line: lineNum,
      column: m.index, // start of the `Forms` keyword
      filePath: ctx.filePath,
      language: 'vba',
      metadata: { synthesizedBy: 'vba-forms-bang' },
    });
  }
}
