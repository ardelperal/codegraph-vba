/**
 * Issue #165 — strict behavior tests for `VbaExtractionRule.scan` and
 * `VbaExtractionRule.requires` semantics.
 *
 * The declarative rule table introduced in #153 (`src/extraction/vba/rules.ts`)
 * defines two optional fields that gate when a rule fires:
 *
 *  - `scan?: 'masked' | 'unmasked' | 'both'` — controls whether the rule
 *    matches against the string-literal-masked line (default) or the
 *    original line. In production today, every rule that sets
 *    `scan: 'masked'` lives in `call-sweep.ts`, and the dispatcher's
 *    `callScanLine` is the line after `maskStringContent()` has blanked
 *    string-literal content with spaces. A rule that ignores this field
 *    and runs on the raw line would falsely match the keyword shape of
 *    a name inside `"..."` (a long-standing real bug — `Set x = New Foo`
 *    inside a string argument used to be picked up as a `vba-set-new`
 *    reference).
 *
 *  - `requires?: 'class' | 'module' | 'inside-procedure' | 'outside-type-block' | 'inside-type-block' | 'outside-enum-block' | 'inside-enum-block'`
 *    — gates the rule by structural context. Production uses three
 *    flavours: `inside-procedure` (call-sweep, so `Set / With / WithEnd`
 *    never fire at module level), `outside-type-block` /
 *    `inside-type-block` (declarations, so `Event / Type / Declare` and
 *    `TypeMember / End Type` partition cleanly), and
 *    `outside-enum-block` / `inside-enum-block` (enums-consts, same
 *    partitioning for `Enum / Const / ProcStart / ProcEnd` vs
 *    `EnumMember / End Enum`).
 *
 * The existing rule-table suite (`__tests__/extraction-vba-rule-table.test.ts`)
 * only verifies the SHAPE of rules — id, description, pattern, emit — so
 * a refactor that flipped the gate semantics (e.g. made `set-new` fire at
 * module level, or let `Event <Name>` slip past inside a `Type … End
 * Type` block) would pass unchanged. This file pins the *behavior* of the
 * two gating fields directly, using the dispatchers and minimal synthetic
 * VBA fixtures — no Access binary, no full orchestrator, no SQLite.
 */
import { describe, it, expect } from 'vitest';
import {
  VbaExtractorContext,
  VbaClassifier,
  ProcInfo,
} from '../src/extraction/vba/context';
import { createCallsAndSqlClassifier } from '../src/extraction/vba/call-sweep';
import { createEventsTypesDeclaresClassifier } from '../src/extraction/vba/declarations';
import { createEnumsConstsClassifier } from '../src/extraction/vba/enums-consts';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { defineRule, matchRuleForScan } from '../src/extraction/vba/rules';

/** Run `cls.classifyLine(line, i, ctx)` for every entry in `lines`. */
function drive(
  cls: VbaClassifier,
  lines: readonly string[],
  ctx: VbaExtractorContext,
): void {
  for (let i = 0; i < lines.length; i++) {
    cls.classifyLine(lines[i] ?? '', i, ctx);
  }
}

/**
 * Seed `ctx.localProcs` so the call-sweep's `inside-procedure` gate can
 * find a matching `ProcInfo` for each `Sub` / `Function` in the fixture.
 * The call-sweep's gate keys on a closure-local `stack` that is pushed
 * ONLY when the proc-start block finds an entry in `ctx.localProcs` — in
 * production the procedures classifier runs first and populates that map,
 * but in these isolated dispatcher tests we drive the call-sweep alone,
 * so we mirror the population step by hand.
 */
function seedLocalProc(ctx: VbaExtractorContext, name: string, startLine: number): void {
  const proc: ProcInfo = {
    name,
    qualifiedName: name,
    kind: 'sub',
    visibility: 'public',
    startLine,
  };
  ctx.localProcs.set(name, [proc]);
}

// ============================================================================
// scan: 'masked' — string-literal masking (call-sweep)
// ============================================================================

describe('VbaExtractionRule.scan = masked — string-literal content must be invisible to the rule', () => {
  it('a `Set x = New <Type>` call at code level triggers the `set-new` rule (sanity baseline)', () => {
    const ctx = new VbaExtractorContext('mod.bas');
    const cls = createCallsAndSqlClassifier([]);
    seedLocalProc(ctx, 'Foo', 1);
    drive(
      cls,
      [
        'Sub Foo()',
        '    Set y = New RealClass',
        'End Sub',
      ],
      ctx,
    );
    const setNewEdges = ctx.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-set-new',
    );
    expect(setNewEdges).toHaveLength(1);
    const target = ctx.nodes.find((n) => n.id === setNewEdges[0]!.target);
    expect(target?.name).toBe('RealClass');
  });

  it('does NOT fire when `Set x = New <Type>` appears only inside a string literal', () => {
    // The masked-line dispatcher blanks `"..."` content with spaces, so a
    // `Set x = New RealClass` shape inside a string is invisible to the
    // rule. The line with the real call DOES fire. This is the regression
    // target: a refactor that switched the dispatcher to the raw line, or
    // dropped `maskStringContent()` from `callScanLine`, would falsely
    // emit a `vba-set-new` reference for the string-only line.
    const ctx = new VbaExtractorContext('mod.bas');
    const cls = createCallsAndSqlClassifier([]);
    seedLocalProc(ctx, 'Foo', 1);
    const lines = [
      'Sub Foo()',
      '    x = "Set y = New GhostType"', // shape only inside "..." — masked out
      '    Set y = New RealClass', // real call — fires
      'End Sub',
    ];
    drive(cls, lines, ctx);
    const setNewEdges = ctx.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-set-new',
    );
    // Exactly ONE reference, and it points at the real call's type, not
    // the string-only one.
    expect(setNewEdges).toHaveLength(1);
    const target = ctx.nodes.find((n) => n.id === setNewEdges[0]!.target);
    expect(target?.name).toBe('RealClass');
  });

  it('does NOT fire when `Set x = New <Type>` is masked into a leading-quote-only string fragment', () => {
    // The masker blanks an unterminated string literal too — so a line
    // that opens a `"` but never closes it must NOT be matched for the
    // `Set` shape that follows. This guards a regression where the
    // masker's `while (i < line.length)` loop is changed to require a
    // closing quote.
    const ctx = new VbaExtractorContext('mod.bas');
    const cls = createCallsAndSqlClassifier([]);
    seedLocalProc(ctx, 'Foo', 1);
    const lines = [
      'Sub Foo()',
      '    x = "Set y = New GhostType', // unterminated string — still masked
      '    Set y = New RealClass',      // real call — fires
      'End Sub',
    ];
    drive(cls, lines, ctx);
    const setNewEdges = ctx.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-set-new',
    );
    expect(setNewEdges).toHaveLength(1);
    const target = ctx.nodes.find((n) => n.id === setNewEdges[0]!.target);
    expect(target?.name).toBe('RealClass');
  });
});

// ============================================================================
// requires: 'inside-procedure' — call-sweep's per-procedure gating
// ============================================================================

describe('VbaExtractionRule.requires = inside-procedure — call-sweep rules stay silent at module level', () => {
  it('set-new does NOT fire for module-level `Set x = New <Type>`', () => {
    // The four `requires: 'inside-procedure'` rules in `call-sweep.ts` all
    // gate on the closure-local proc stack. Module-level VBA can't legally
    // carry a `Set` (real Access would reject it at parse), but the regex
    // pipeline MUST still respect the gate — otherwise every module-level
    // variable initialization in a Dysflow-exported `.bas` would synthesize
    // a `vba-set-new` reference to whatever the regex matched.
    const ctx = new VbaExtractorContext('mod.bas');
    const cls = createCallsAndSqlClassifier([]);
    seedLocalProc(ctx, 'Foo', 3);
    const lines = [
      'Set moduleVar = New ModuleType', // module-level — gate should silence it
      '',
      'Sub Foo()',
      '    Set procVar = New ProcType', // inside-procedure — fires
      'End Sub',
    ];
    drive(cls, lines, ctx);
    const setNewEdges = ctx.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-set-new',
    );
    const names = setNewEdges
      .map((e) => ctx.nodes.find((n) => n.id === e.target)?.name)
      .filter((n): n is string => !!n);
    expect(names).toEqual(['ProcType']);
  });

  it('with-start does NOT push onto `ctx.vbaWithStack` at module level', () => {
    const ctx = new VbaExtractorContext('mod.bas');
    const cls = createCallsAndSqlClassifier([]);
    seedLocalProc(ctx, 'Bar', 5);
    const lines = [
      'With Foo',                     // module-level — must NOT push
      '    .Member = 1',
      'End With',
      '',
      'Sub Bar()',
      '    With Foo',                 // inside-procedure — must push
      '        .Member = 1',
      '    End With',
      'End Sub',
    ];
    drive(cls, lines, ctx);
    // After processing the whole file the only surviving with-stack entry
    // must come from the `Sub Bar() / With Foo` pair; the module-level
    // `With Foo` was ignored AND the matching module-level `End With` was
    // also ignored (the same gate). The inside-procedure pair balances,
    // so the stack ends empty.
    expect(ctx.vbaWithStack).toEqual([]);
  });

  it('with-end at module level is a no-op (the gate keeps the stack empty)', () => {
    // Belt-and-braces: a stray `End With` at module level (with no opening
    // `With` in scope) must NOT pop past zero and corrupt the stack. The
    // gate fires first, so the rule never runs.
    const ctx = new VbaExtractorContext('mod.bas');
    const cls = createCallsAndSqlClassifier([]);
    seedLocalProc(ctx, 'Bar', 2);
    const lines = [
      'End With',                    // module-level — must NOT pop
      'Sub Bar()',
      '    With Foo',                // pushes 'foo'
      '    End With',                // pops 'foo'
      'End Sub',
    ];
    drive(cls, lines, ctx);
    expect(ctx.vbaWithStack).toEqual([]);
  });

  it('set-call does NOT fire at module level', () => {
    // Mirrors the set-new case for `Set x = <Factory>(...)`. Module-level
    // `Set x = SomeFactory(1)` is not legal VBA but the regex would match
    // it; the gate is what keeps the reference off the graph.
    const ctx = new VbaExtractorContext('mod.bas');
    const cls = createCallsAndSqlClassifier([]);
    // Register a same-file function return type so a real `Set x = <Name>(...)`
    // would have something to point at inside a procedure.
    ctx.functionReturnTypes.set('makeit', 'FactoryOutput');
    seedLocalProc(ctx, 'Bar', 2);
    const lines = [
      'Set moduleVar = MakeIt(1)',   // module-level — must NOT fire
      'Sub Bar()',
      '    Set procVar = MakeIt(1)',  // inside-procedure — must fire
      'End Sub',
    ];
    drive(cls, lines, ctx);
    const factoryEdges = ctx.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-factory-return',
    );
    const names = factoryEdges
      .map((e) => ctx.nodes.find((n) => n.id === e.target)?.name)
      .filter((n): n is string => !!n);
    expect(names).toEqual(['FactoryOutput']);
  });
});

// ============================================================================
// requires: 'outside-type-block' / 'inside-type-block' — declarations
// ============================================================================

describe('VbaExtractionRule.requires = outside-type-block / inside-type-block — declarations partition cleanly', () => {
  it('event-decl fires outside a type block', () => {
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEventsTypesDeclaresClassifier();
    drive(cls, ['Public Event MyEvent()'], ctx);
    expect(ctx.nodes.some((n) => n.kind === 'event' && n.name === 'MyEvent')).toBe(true);
  });

  it('event-decl does NOT fire inside a `Type … End Type` block', () => {
    // The classic regression target: a `Type X / Event Y / End Type`
    // shape (rare but valid VBA) MUST NOT synthesize an `event` node from
    // the inner `Event Y` line — `Event` is a declaration, not a type
    // member, and `event-decl` is gated `outside-type-block`.
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEventsTypesDeclaresClassifier();
    drive(
      cls,
      [
        'Public Type Foo',
        '    Event MyEvent', // ignored — inside-type-block
        'End Type',
      ],
      ctx,
    );
    expect(ctx.nodes.some((n) => n.kind === 'event')).toBe(false);
    // Sanity: the type block itself was created (so the gate was actually
    // engaged), but no member-line events leaked out.
    expect(ctx.nodes.some((n) => n.kind === 'type' && n.name === 'Foo')).toBe(true);
  });

  it('type-start does NOT fire inside an open type block (no nested Type)', () => {
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEventsTypesDeclaresClassifier();
    drive(
      cls,
      [
        'Public Type Outer',
        '    Type Inner', // ignored — inside-type-block
        '    X As Long',
        'End Type',
      ],
      ctx,
    );
    const types = ctx.nodes.filter((n) => n.kind === 'type').map((n) => n.name);
    expect(types).toEqual(['Outer']);
  });

  it('dll-declare does NOT fire inside a `Type … End Type` block', () => {
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEventsTypesDeclaresClassifier();
    drive(
      cls,
      [
        'Public Type Foo',
        '    Declare Sub WinApi Lib "user32" ()', // ignored
        'End Type',
      ],
      ctx,
    );
    expect(ctx.nodes.some((n) => n.kind === 'declare')).toBe(false);
  });

  it('type-member fires inside a type block', () => {
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEventsTypesDeclaresClassifier();
    drive(
      cls,
      [
        'Public Type Foo',
        '    MemberA As Long',
        '    MemberB As String',
        'End Type',
      ],
      ctx,
    );
    const members = ctx.nodes.filter((n) => n.kind === 'type_member').map((n) => n.name);
    expect(members.sort()).toEqual(['MemberA', 'MemberB']);
  });

  it('type-member does NOT fire outside any open type block', () => {
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEventsTypesDeclaresClassifier();
    drive(
      cls,
      [
        'MemberA As Long', // ignored — outside-type-block
        'MemberB As String',
      ],
      ctx,
    );
    expect(ctx.nodes.filter((n) => n.kind === 'type_member')).toEqual([]);
  });

  it('type-end only clears `vbaDeclTypeBlock` once it was set', () => {
    // Drive `End Type` without a matching `Type …`. The gate
    // (inside-type-block) keeps the rule silent, so `vbaDeclTypeBlock`
    // is never set and never touched — pin that with a follow-up `Type X`
    // that DOES set it, proving the previous `End Type` didn't break
    // the state.
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEventsTypesDeclaresClassifier();
    drive(
      cls,
      [
        'End Type',            // gate: ctx.vbaDeclTypeBlock is null → skip
        'Public Type Foo',     // opens a block
        '    MemberA As Long',
        'End Type',            // closes the block
      ],
      ctx,
    );
    const members = ctx.nodes.filter((n) => n.kind === 'type_member').map((n) => n.name);
    expect(members).toEqual(['MemberA']);
    expect(ctx.vbaDeclTypeBlock).toBeNull();
  });
});

// ============================================================================
// requires: 'outside-enum-block' / 'inside-enum-block' — enums / consts
// ============================================================================

describe('VbaExtractionRule.requires = outside-enum-block / inside-enum-block — enum/const rules partition cleanly', () => {
  it('enum-member fires inside an `Enum … End Enum` block', () => {
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEnumsConstsClassifier();
    drive(
      cls,
      [
        'Public Enum Color',
        '    Red = 1',
        '    Green = 2',
        'End Enum',
      ],
      ctx,
    );
    const members = ctx.nodes.filter((n) => n.kind === 'enum_member').map((n) => n.name);
    expect(members.sort()).toEqual(['Green', 'Red']);
  });

  it('enum-member does NOT fire outside an open enum block', () => {
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEnumsConstsClassifier();
    drive(
      cls,
      [
        'Red = 1', // ignored — outside-enum-block
        'Green = 2',
      ],
      ctx,
    );
    expect(ctx.nodes.filter((n) => n.kind === 'enum_member')).toEqual([]);
  });

  it('proc-start does NOT fire inside an enum block (no `Sub` lines emitted as enum members)', () => {
    // A `Sub MyHandler()` line inside an `Enum … End Enum` is malformed VBA
    // but the regex cascade would happily match it. The gate must keep it
    // out — otherwise an enum body would push onto `ctx.procStack` and
    // a later `Const` write would be mis-routed to a fake proc scope
    // (and the procedures sweep would later emit a phantom `function`
    // node for it). The gate is the outer defense; this test pins BOTH
    // the side-effect (procStack stays empty) and the downstream symptom
    // (no function node ever appears).
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEnumsConstsClassifier();
    drive(
      cls,
      [
        'Public Enum Color',
        '    Sub MyHandler()', // ignored — outside-enum-block gate
        '    Red = 1',
        'End Enum',
      ],
      ctx,
    );
    expect(ctx.procStack).toEqual([]);
    expect(ctx.nodes.some((n) => n.kind === 'function')).toBe(false);
    // The enum-member line still produced its node — proving the gate
    // didn't accidentally suppress the whole line.
    expect(ctx.nodes.some((n) => n.kind === 'enum_member' && n.name === 'Red')).toBe(true);
  });

  it('const-decl does NOT fire inside an `Enum … End Enum` block', () => {
    // `Const X = 1` inside `Enum … End Enum` would (without the gate)
    // emit a `constant` node AND write into the local-const scope
    // bucket. The gate silences both.
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEnumsConstsClassifier();
    drive(
      cls,
      [
        'Public Enum Color',
        '    Const X = 1', // ignored — outside-enum-block gate
        '    Red = 1',
        'End Enum',
      ],
      ctx,
    );
    expect(ctx.nodes.some((n) => n.kind === 'constant' && n.name === 'X')).toBe(false);
    // And the per-scope resolution bucket for the module scope is empty
    // (the `Const X = 1` line was entirely silenced — `setLocalConstInScope`
    // was never called, so the module bucket was never even created).
    expect(ctx.localConstants.get('module')?.has('x') ?? false).toBe(false);
  });

  it('const-decl fires at module level (outside-enum-block, outside any proc)', () => {
    // Baseline: the gate must NOT silence legitimate module-level Consts.
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEnumsConstsClassifier();
    drive(cls, ['Public Const MyConst = 42'], ctx);
    const c = ctx.nodes.find((n) => n.kind === 'constant' && n.name === 'MyConst');
    expect(c).toBeDefined();
    expect(c?.visibility).toBe('public');
  });

  it('enum-end only clears `vbaEnumBlock` once it was set', () => {
    const ctx = new VbaExtractorContext('m.bas');
    const cls = createEnumsConstsClassifier();
    drive(
      cls,
      [
        'End Enum', // gate: ctx.vbaEnumBlock is null → skip
        'Public Enum Color', // opens a block
        '    Red = 1',
        'End Enum', // closes the block
      ],
      ctx,
    );
    expect(ctx.nodes.some((n) => n.kind === 'enum_member' && n.name === 'Red')).toBe(true);
    expect(ctx.vbaEnumBlock).toBeNull();
  });
});

// ============================================================================
// End-to-end (orchestrator path) — sanity that gating holds through the
// full VbaExtractor (preprocessing + walker + finalize + module node).
// ============================================================================

describe('end-to-end: rule gating survives the full VbaExtractor walk', () => {
  function extract(src: string): ReturnType<VbaExtractor['extract']> {
    return new VbaExtractor('src/modules/modRuleGates.bas', src).extract();
  }

  it('module-level `Set x = New Foo` does not produce a `vba-set-new` reference edge', () => {
    // End-to-end mirror of the dispatcher-level test: drive the
    // orchestrator so the `vba-extractor.ts` walker order and the
    // `applyRaiseFanoutGate` post-pass both run, and confirm the gating
    // holds.
    const r = extract(
      [
        'Attribute VB_Name = "modRuleGates"',
        'Set moduleVar = New ModuleType', // module-level — must NOT fire
        '',
        'Sub Foo()',
        '    Set procVar = New ProcType', // inside-procedure — fires
        'End Sub',
      ].join('\n'),
    );
    const setNewEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-set-new',
    );
    const names = setNewEdges
      .map((e) => r.nodes.find((n) => n.id === e.target)?.name)
      .filter((n): n is string => !!n);
    expect(names).toEqual(['ProcType']);
  });

  it('`Event Y` inside a `Type X … End Type` block does not produce an event node', () => {
    const r = extract(
      [
        'Attribute VB_Name = "modRuleGates"',
        'Public Type Foo',
        '    Event MyEvent',
        'End Type',
      ].join('\n'),
    );
    expect(r.nodes.some((n) => n.kind === 'event')).toBe(false);
    expect(r.nodes.some((n) => n.kind === 'type' && n.name === 'Foo')).toBe(true);
  });

  it('`Const X = 1` inside an `Enum Y … End Enum` block does not produce a constant node', () => {
    const r = extract(
      [
        'Attribute VB_Name = "modRuleGates"',
        'Public Enum Color',
        '    Const X = 1',
        '    Red = 1',
        'End Enum',
      ].join('\n'),
    );
    expect(r.nodes.some((n) => n.kind === 'constant' && n.name === 'X')).toBe(false);
    expect(r.nodes.some((n) => n.kind === 'enum_member' && n.name === 'Red')).toBe(true);
  });

  it('`Set y = New GhostType` hidden inside a string literal does not produce a `vba-set-new` reference', () => {
    const r = extract(
      [
        'Attribute VB_Name = "modRuleGates"',
        'Sub Foo()',
        '    x = "Set y = New GhostType"', // masked out — must NOT fire
        '    Set y = New RealType', // real call — fires
        'End Sub',
      ].join('\n'),
    );
    const setNewEdges = r.edges.filter(
      (e) => e.kind === 'references' && e.metadata?.synthesizedBy === 'vba-set-new',
    );
    const names = setNewEdges
      .map((e) => r.nodes.find((n) => n.id === e.target)?.name)
      .filter((n): n is string => !!n);
    expect(names).toEqual(['RealType']);
  });
});

describe('VbaExtractionRule.scan dispatch', () => {
  const line = 'x = "Ghost"';
  const maskedLine = 'x =        ';
  const rule = (scan: 'masked' | 'unmasked' | 'both') =>
    defineRule({
      id: scan,
      description: scan,
      scan,
      pattern: /Ghost/,
      emit: () => null,
    });

  it('uses only the masked line for masked rules', () => {
    expect(matchRuleForScan(rule('masked'), line, maskedLine)).toBeNull();
  });

  it('uses the original line for unmasked rules', () => {
    expect(matchRuleForScan(rule('unmasked'), line, maskedLine)?.line).toBe(line);
  });

  it('falls back to the original line for both rules', () => {
    expect(matchRuleForScan(rule('both'), line, maskedLine)?.line).toBe(line);
  });
});
