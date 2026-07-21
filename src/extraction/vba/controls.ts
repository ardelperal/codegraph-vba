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

// Length of the `Me.` / `Me!` prefix that opens every captured control
// reference. Both forms are 3 chars; `ME_PREFIX_LEN` replaces the bare
// `+ 3` literal that previously appeared at multiple call sites.
const ME_PREFIX_LEN = 3;

/**
 * Issue #211: shared direct-assignment predicate. The same-line rule
 * the third branch (`control` for non-builtin Me.<Ctl>) already uses is
 * the right shape — `before` must be empty (or whitespace) AND `after`
 * must start with `=` — and lifts cleanly into the `builtIn` and bang
 * branches too. Returning `true` means the match IS the LHS of an
 * assignment (`Me.X = ...`), so callers should tag it `*-set`. Any
 * other shape (e.g. `If Me.X = ...`, `MsgBox Me.X`, `x = Me.X`) is a
 * read and must tag `*-get`. The function is intentionally narrow; a
 * cross-line assignment is out of scope (matches the prior comment).
 */
function isDirectAssignment(before: string, after: string): boolean {
  return /^\s*$/.test(before) && /^\s*=/.test(after);
}

const ACCESS_FORM_MEMBER_BLACKLIST = new Set([
  'requery', 'refresh', 'repaint', 'recalc', 'undo', 'dirty', 'newrecord',
  'currentrecord', 'recordset', 'recordsetclone', 'recordsource', 'controls',
  'name', 'caption', 'visible', 'filter', 'filteron', 'orderby', 'orderbyon',
  'parent', 'activecontrol', 'section', 'bookmark', 'form', 'hwnd', 'painting',
  'timerinterval',
]);

const seenMeControls = new WeakMap<VbaExtractorContext, Set<string>>();

function siblingLayoutPath(filePath: string): string | null {
  const normalized = filePath.replace(/\\/g, '/');
  const basename = normalized.slice(normalized.lastIndexOf('/') + 1);
  if (/^Form_.+\.cls$/i.test(basename)) {
    return normalized.replace(/\.cls$/i, '.form.txt');
  }
  if (/^Report_.+\.cls$/i.test(basename)) {
    return normalized.replace(/\.cls$/i, '.report.txt');
  }
  return null;
}

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
 *
 * Round-3 (issue #108): the `referenceKind` is keyed off the operator
 * (`.` vs `!`) at the match position. The same-line `=` heuristic
 * distinguishes reads from assignments: `x = Me.Name` → `property-get`,
 * `Me.Name = "X"` → `property-set`. Cross-line property-set is out of
 * scope; the `=` rule is satisfied when the assignment appears on the
 * SAME line at any position AFTER the matched control name (allowing
 * trailing `.Caption = value` chained assignments too — they still
 * flow a write into the captured control identifier).
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
    // Round-3: operator is the char at `m.index + 2` (skipping the `Me`).
    // Both `.` and `!` are 1 char so the captured name starts at
    // `m.index + ME_PREFIX_LEN`.
    const operator = line.charAt(m.index + 2); // '.' | '!'
    const isBang = operator === '!';
    const siblingPath = siblingLayoutPath(ctx.filePath);

    // Issue #140 is deliberately a separate, dot-only sweep for Access
    // form/report code-behind. Keeping `Me` in the generic runtime receiver
    // blacklist prevents runtime methods from becoming synthetic call nodes.
    // `.Control` inside `With Me` is not handled here: resolving an implicit
    // receiver requires block/data-flow tracking and is intentionally follow-up
    // work rather than a guessed edge.
    if (!isBang) {
      if (!siblingPath) continue;
      if (ACCESS_FORM_MEMBER_BLACKLIST.has(controlName.toLowerCase())) {
        // Keep issue #108's extraction-shape contract observable: built-in
        // form properties such as Me.Name remain property-get/property-set
        // unresolved refs. The dedicated resolver sees `builtIn: true` and
        // always declines them, so they still create no node or graph edge.
        const after = line.slice(m.index + ME_PREFIX_LEN + controlName.length);
        const before = line.slice(0, m.index);
        ctx.unresolvedReferences.push({
          fromNodeId: ctx.findOrCreateFunctionNodeId(from),
          referenceName: controlName,
          referenceKind: isDirectAssignment(before, after) ? 'property-set' : 'property-get',
          line: lineNum,
          column: m.index + ME_PREFIX_LEN,
          filePath: ctx.filePath,
          language: 'vba',
          metadata: { synthesizedBy: 'vba-me-control', siblingPath, builtIn: true },
        });
        continue;
      }
      const fromNodeId = ctx.findOrCreateFunctionNodeId(from);
      const dedupeKey = `${fromNodeId}\0${controlName.toLowerCase()}`;
      let seen = seenMeControls.get(ctx);
      if (!seen) {
        seen = new Set<string>();
        seenMeControls.set(ctx, seen);
      }
      if (seen.has(dedupeKey)) continue;
      seen.add(dedupeKey);

      const after = line.slice(m.index + ME_PREFIX_LEN + controlName.length);
      const before = line.slice(0, m.index);
      ctx.unresolvedReferences.push({
        fromNodeId,
        referenceName: controlName,
        referenceKind: 'references',
        line: lineNum,
        column: m.index + ME_PREFIX_LEN,
        filePath: ctx.filePath,
        language: 'vba',
        metadata: {
          synthesizedBy: 'vba-me-control',
          siblingPath,
          access: isDirectAssignment(before, after) ? 'write' : 'read',
        },
      });
      continue;
    }

    // Preserve the existing Me! default-collection behavior from issue #44.
    // Same-line get vs set heuristic: `before` must be empty/whitespace AND
    // `after` must start with `=`. Both predicates are the same `isDirectAssignment`
    // helper used by the dot branches (issue #211), so reads inside `If`,
    // `MsgBox`, or `IIf` correctly emit `bang-get` while `Me!X = v` emits
    // `bang-set`.
    const after = line.slice(m.index + ME_PREFIX_LEN + controlName.length);
    const before = line.slice(0, m.index);
    const isAssign = isDirectAssignment(before, after);
    const referenceKind: 'property-get' | 'property-set' | 'bang-get' | 'bang-set' = isBang
      ? isAssign
        ? 'bang-set'
        : 'bang-get'
      : isAssign
        ? 'property-set'
        : 'property-get';
    ctx.unresolvedReferences.push({
      fromNodeId: ctx.findOrCreateFunctionNodeId(from),
      referenceName: controlName,
      referenceKind,
      line: lineNum,
      column: m.index + ME_PREFIX_LEN,
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
 *
 * Round-3 (issue #108): cross-form bang is always a property access
 * (VBA semantics: there's no bang-call). The reference name is the FORM
 * module — the resolver matches it against the indexed `Form_*` nodes
 * downstream — so the kind is `'bang-get'` regardless of the read/write
 * direction. A direct cross-form bang assignment
 * (`Forms!FormX!Ctl = value`) is rare; round-3 emits `'bang-get'`
 * uniformly and the resolvers do not care about access direction here.
 */
export function scanFormsBang(
  ctx: VbaExtractorContext,
  line: string,
  maskedLine: string,
  from: ProcInfo,
  lineNum: number,
): void {
  FORMS_BANG_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = FORMS_BANG_RE.exec(line)) !== null) {
    // The original line is required for Forms("name") payloads, but the
    // keyword itself must still be executable code rather than prose inside
    // a string literal. maskStringContent preserves columns, so this check is
    // both cheap and exact for every regex match.
    if (maskedLine.slice(m.index, m.index + 5).toLowerCase() !== 'forms') continue;
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
      referenceKind: 'bang-get',
      line: lineNum,
      column: m.index, // start of the `Forms` keyword
      filePath: ctx.filePath,
      language: 'vba',
      metadata: { synthesizedBy: 'vba-forms-bang' },
    });
  }
}
