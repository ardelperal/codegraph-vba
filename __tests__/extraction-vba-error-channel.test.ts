/**
 * Issue #261 (task E4 of `docs/vba-error-handling-plan.md`) — the module
 * variable this codebase actually propagates error messages through is
 * recognised as such.
 *
 * §2.1/§2.3 of the plan is the whole justification: **16 handlers out of
 * 3,774 re-raise**. `Err.Raise 1000` unwinds one frame and the house guard
 * `If Err.Number <> 1000` means *"an inner procedure already wrote a
 * human-readable message"*. The message itself travels through a field —
 * `m_Error`, `p_Error`, the public `Error` — which issue #251 already models
 * as `property-set` / `property-get` references onto a `variable` node. This
 * task adds exactly one thing to that: `metadata.errorChannel: true`, so the
 * hops that carry an error are distinguishable from the hops that carry
 * ordinary module state.
 *
 * Two invariants outrank every assertion below, and both are pinned here:
 *
 *   1. **No new node kind, no new edge kind, and not one extra row.** The flag
 *      is a field on rows #251 already emitted. `adds no rows` pins it.
 *   2. **The pattern is a NAME, never a substring.** `ErrorCount` must not
 *      match `Error` — a false positive here does not merely add noise, it
 *      claims that ordinary state is error propagation.
 *
 * The last block is the end-to-end one: the chain
 * `inner writes -> outer reads -> outer displays` must be traversable in a
 * real index, not merely present as two unrelated references. Its fixture is
 * modelled on `src/forms/Form_FormGestionRiesgos.cls` of the reference corpus,
 * where `EstablecerDatos` writes `m_Error` and `Form_Load` reads it and shows
 * it in a `MsgBox`.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import {
  DEFAULT_ERROR_CHANNEL_NAMES,
  compileErrorChannel,
  isErrorChannelName,
} from '../src/extraction/vba/error-channel';
import { clearProjectConfigCache, loadVbaConfig } from '../src/project-config';
import CodeGraph from '../src/index';
import type { Node, UnresolvedReference } from '../src/types';
import type { VbaExtractionOptions } from '../src/extraction/vba/options';

const CLASS_PATH = 'src/classes/Riesgo.cls';

/**
 * Extract a class module. The two header lines are prepended here so a
 * fixture's own first line is source line 3 — the convention
 * `extraction-vba-error-handler-region.test.ts` uses.
 */
function extract(body: string[], options?: VbaExtractionOptions) {
  return new VbaExtractor(
    CLASS_PATH,
    ['Attribute VB_Name = "Riesgo"', 'Option Explicit', ...body].join('\n'),
    options,
  ).extract();
}

function moduleVarRefs(
  result: ReturnType<typeof extract>,
  name?: string,
): UnresolvedReference[] {
  return result.unresolvedReferences.filter(
    (r) =>
      r.metadata?.synthesizedBy === 'vba-module-var' &&
      (name === undefined || r.referenceName === name),
  );
}

/** The one reference this procedure makes to `name` in the given direction. */
function refOf(
  result: ReturnType<typeof extract>,
  name: string,
  access: 'read' | 'write',
): UnresolvedReference | undefined {
  return moduleVarRefs(result, name).find((r) => r.metadata?.access === access);
}

/** Declared procedures only — call-target stubs carry `metadata.stub`. */
function procedure(nodes: Node[], name: string): Node | undefined {
  return nodes.find(
    (n) => n.kind === 'function' && n.metadata?.stub !== true && n.name === name,
  );
}

// ============================================================================
// 1. The two hops of the chain, on the corpus's own shapes
// ============================================================================

describe('Issue #261 — the error channel is flagged on reads and writes', () => {
  it('flags `m_Error = "…"` written inside a handler as a channel write, inside the handler', () => {
    const result = extract([
      'Private m_Error As String',
      '',
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Escribir',
      '    Exit Sub',
      'errores:',
      '    m_Error = "Riesgo.Guardar ha fallado: " & Err.Description',
      'End Sub',
    ]);

    const write = refOf(result, 'm_Error', 'write');
    expect(write).toBeDefined();
    expect(write!.referenceKind).toBe('property-set');
    expect(write!.metadata).toMatchObject({
      synthesizedBy: 'vba-module-var',
      access: 'write',
      errorChannel: true,
      // Issue #260's flag, which this line earns because it is the ONLY write
      // in the procedure and therefore not deduped away by an earlier one.
      inErrorHandler: true,
    });
  });

  it('flags the caller’s `If m_Error <> "" Then` guard as a channel read', () => {
    const result = extract([
      'Private m_Error As String',
      '',
      'Public Sub Comprobar()',
      '    If m_Error <> "" Then',
      '        Exit Sub',
      '    End If',
      'End Sub',
    ]);

    const read = refOf(result, 'm_Error', 'read');
    expect(read).toBeDefined();
    expect(read!.referenceKind).toBe('property-get');
    expect(read!.metadata).toMatchObject({
      synthesizedBy: 'vba-module-var',
      access: 'read',
      errorChannel: true,
    });
    // A read outside any handler region carries no #260 flag: absence is the
    // encoding, so the key must not appear at all.
    expect(read!.metadata).not.toHaveProperty('inErrorHandler');
  });

  it('flags every default channel name, and only as a whole name', () => {
    for (const name of DEFAULT_ERROR_CHANNEL_NAMES) {
      const result = extract([
        `Private ${name} As String`,
        '',
        'Public Sub Comprobar()',
        `    If ${name} <> "" Then Exit Sub`,
        'End Sub',
      ]);
      const read = refOf(result, name, 'read');
      expect(read, `${name} should be flagged`).toBeDefined();
      expect(read!.metadata?.errorChannel).toBe(true);
    }
  });
});

// ============================================================================
// 2. The precision gate — a name, not a substring
// ============================================================================

describe('Issue #261 — the channel is a name, never a substring', () => {
  it('does NOT flag a variable named `ErrorCount`', () => {
    const result = extract([
      'Private ErrorCount As Long',
      'Private m_Error As String',
      '',
      'Public Sub Contar()',
      '    ErrorCount = ErrorCount + 1',
      '    m_Error = ""',
      'End Sub',
    ]);

    const counted = moduleVarRefs(result, 'ErrorCount');
    expect(counted.length).toBeGreaterThan(0);
    for (const ref of counted) {
      expect(ref.metadata).not.toHaveProperty('errorChannel');
    }
    // …while the real channel in the same procedure IS flagged, so the
    // negative above cannot pass by the sweep simply not running.
    expect(refOf(result, 'm_Error', 'write')?.metadata?.errorChannel).toBe(true);
  });

  it.each([
    'ErrorCount',
    'm_ErrorLog',
    'HayErrorEnRiesgo',
    'ErroresRiesgoTexto',
    'gErrores',
  ])('does NOT flag the neighbouring name `%s`', (name) => {
    expect(isErrorChannelName(compileErrorChannel(), name)).toBe(false);
  });

  it('matches case-insensitively, because VBA does', () => {
    const channel = compileErrorChannel();
    expect(isErrorChannelName(channel, 'M_ERROR')).toBe(true);
    expect(isErrorChannelName(channel, 'error')).toBe(true);
  });
});

// ============================================================================
// 3. The `vba.errorChannel` knob
// ============================================================================

describe('Issue #261 — `vba.errorChannel` extends the built-in list', () => {
  it('flags a project-specific name AND keeps every default', () => {
    const options: VbaExtractionOptions = { errorChannel: ['lastFailure'] };
    const result = extract(
      [
        'Private lastFailure As String',
        'Private m_Error As String',
        '',
        'Public Sub Guardar()',
        '    lastFailure = ""',
        '    m_Error = ""',
        'End Sub',
      ],
      options,
    );

    expect(refOf(result, 'lastFailure', 'write')?.metadata?.errorChannel).toBe(true);
    expect(refOf(result, 'm_Error', 'write')?.metadata?.errorChannel).toBe(true);
  });

  it('compiles the configured names on top of the defaults, de-duplicated', () => {
    const channel = compileErrorChannel(['lastFailure', 'm_Error', '  ']);
    expect(channel.names).toEqual([...DEFAULT_ERROR_CHANNEL_NAMES, 'lastFailure']);
    expect(isErrorChannelName(channel, 'lastFailure')).toBe(true);
    expect(isErrorChannelName(channel, 'lastFailureCount')).toBe(false);
  });

  it('drops entries that are not bare identifiers rather than compiling them', () => {
    // No user-supplied regex, and nothing qualified: these run per identifier
    // per line, so a `.*` entry would be a backtracking hazard, not a feature.
    const channel = compileErrorChannel(['obj.Error', 'm_.*', '', 'ok_Name']);
    expect(channel.names).toEqual([...DEFAULT_ERROR_CHANNEL_NAMES, 'ok_Name']);
  });

  it('drives the handler-behaviour classifier from the SAME list', () => {
    // The knob must not leave `errorPolicy.behavior` reading a stale, separate
    // copy of the names: a handler whose only channel write is the configured
    // name is `channel`, not `unknown`.
    const body = [
      'Private lastFailure As String',
      '',
      'Public Sub Guardar()',
      '    On Error GoTo errores',
      '    Exit Sub',
      'errores:',
      '    lastFailure = "boom"',
      'End Sub',
    ];
    const withKnob = extract(body, { errorChannel: ['lastFailure'] });
    expect(
      (procedure(withKnob.nodes, 'Guardar')?.metadata?.errorPolicy as { behavior: string })
        .behavior,
    ).toBe('channel');

    const withoutKnob = extract(body);
    expect(
      (procedure(withoutKnob.nodes, 'Guardar')?.metadata?.errorPolicy as { behavior: string })
        .behavior,
    ).toBe('unknown');
  });
});

// ============================================================================
// 4. The invariant: a field on existing rows, never a new row
// ============================================================================

describe('Issue #261 — adds no rows', () => {
  const body = [
    'Private m_Error As String',
    'Private m_Count As Long',
    '',
    'Public Sub Guardar()',
    '    On Error GoTo errores',
    '    m_Count = m_Count + 1',
    '    Exit Sub',
    'errores:',
    '    m_Error = "boom"',
    'End Sub',
  ];

  it('emits the same nodes, edges and references with and without the knob', () => {
    const plain = extract(body);
    const configured = extract(body, { errorChannel: ['m_Count'] });

    expect(configured.nodes.length).toBe(plain.nodes.length);
    expect(configured.edges.length).toBe(plain.edges.length);
    expect(configured.unresolvedReferences.length).toBe(
      plain.unresolvedReferences.length,
    );
    expect(new Set(configured.nodes.map((n) => n.kind))).toEqual(
      new Set(plain.nodes.map((n) => n.kind)),
    );
    expect(new Set(configured.edges.map((e) => e.kind))).toEqual(
      new Set(plain.edges.map((e) => e.kind)),
    );

    // …and the ONLY observable difference is the flag on `m_Count`.
    expect(
      moduleVarRefs(plain, 'm_Count').every(
        (r) => r.metadata?.errorChannel === undefined,
      ),
    ).toBe(true);
    expect(
      moduleVarRefs(configured, 'm_Count').every(
        (r) => r.metadata?.errorChannel === true,
      ),
    ).toBe(true);
  });

  it('introduces no node kind and no edge kind of its own', () => {
    const result = extract(body);
    expect(result.nodes.some((n) => n.kind === ('label' as never))).toBe(false);
    for (const edge of result.edges) {
      expect(['contains', 'calls', 'references', 'type-member']).toContain(edge.kind);
    }
  });
});

// ============================================================================
// 5. The boundary this change deliberately does NOT cross
// ============================================================================

describe('Issue #261 — what the module-variable channel cannot see', () => {
  it('does not flag `Me.Error = …`, which is not a module-variable reference', () => {
    // `Me.`-qualified and object-qualified writes are how the corpus publishes
    // the message OUT of the object. The #251 sweep skips a dotted access by
    // design (it names a member of something else), so no reference exists to
    // flag. Pinned so the boundary is documented rather than discovered.
    const result = extract([
      'Public Error As String',
      'Private m_Error As String',
      '',
      'Public Sub Guardar()',
      '    Me.Error = m_Error',
      'End Sub',
    ]);
    expect(moduleVarRefs(result, 'Error')).toHaveLength(0);
    expect(refOf(result, 'm_Error', 'read')?.metadata?.errorChannel).toBe(true);
  });

  it('does not flag a ByRef `p_Error` parameter, which is not module state', () => {
    // The corpus's OTHER propagation mechanism: `Optional ByRef p_Error As
    // String`. It is a parameter, so #205 scoping keeps it out of the
    // module-variable sweep entirely and there is nothing for this task to
    // flag. `ExpedienteOperaciones.cls` uses only this form.
    const result = extract([
      'Private m_Error As String',
      '',
      'Private Function getHTML(Optional ByRef p_Error As String) As String',
      '    On Error GoTo errores',
      '    Exit Function',
      'errores:',
      '    p_Error = "boom"',
      'End Function',
    ]);
    expect(moduleVarRefs(result, 'p_Error')).toHaveLength(0);
  });
});

// ============================================================================
// 6. End to end — the chain must be TRAVERSABLE, not merely present
// ============================================================================

describe('Issue #261 — inner writes -> outer reads -> outer displays (integration)', () => {
  let dir: string;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cg-vba-errchan-'));
    clearProjectConfigCache();
  });

  afterEach(async () => {
    clearProjectConfigCache();
    await new Promise((resolve) => setTimeout(resolve, 100));
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      // ignore cleanup EPERM on Windows
    }
  });

  /**
   * The corpus shape, reduced: `Form_FormGestionRiesgos.cls`'s
   * `EstablecerDatos` writes `m_Error`, `Form_Load` calls it, reads `m_Error`
   * and shows it in a `MsgBox` from its handler.
   */
  const FORM = [
    'Attribute VB_Name = "Form_FormGestionRiesgos"',
    'Option Explicit',
    'Private m_Error As String',
    '',
    'Public Sub Form_Load()',
    '    On Error GoTo errores',
    '    EstablecerDatos',
    '    If m_Error <> "" Then',
    '        Err.Raise 1000',
    '    End If',
    '    Exit Sub',
    'errores:',
    '    If Err.Number <> 1000 Then',
    '        m_Error = "Al Form_Load se ha producido el error: " & Err.Description',
    '    End If',
    '    MsgBox m_Error, vbCritical, "Error"',
    'End Sub',
    '',
    'Private Sub EstablecerDatos()',
    '    On Error GoTo errores',
    '    Exit Sub',
    'errores:',
    '    m_Error = "Es un usuario no autorizado"',
    'End Sub',
  ].join('\n');

  it('connects the writer to the reader through the channel variable, and the reader displays', async () => {
    fs.writeFileSync(path.join(dir, 'Form_FormGestionRiesgos.cls'), FORM);

    const cg = CodeGraph.initSync(dir);
    await cg.indexAll();

    const channel = cg
      .getNodesByName('m_Error')
      .find((n) => n.kind === 'variable');
    expect(channel, 'the channel variable is a node').toBeDefined();

    // Hop 1 — the inner procedure WRITES the channel.
    const incoming = cg.getCallers(channel!.id);
    const write = incoming.find(
      (c) => c.edge.metadata?.access === 'write' && c.node.name === 'EstablecerDatos',
    );
    expect(write, 'EstablecerDatos writes the channel').toBeDefined();
    expect(write!.edge.kind).toBe('references');
    expect(write!.edge.metadata?.refKind).toBe('property-set');
    expect(write!.edge.metadata?.errorChannel).toBe(true);

    // Hop 2 — the outer procedure READS the same channel node. Same target,
    // so the two hops JOIN: this is the traversal, not two loose facts.
    const read = incoming.find(
      (c) => c.edge.metadata?.access === 'read' && c.node.name === 'Form_Load',
    );
    expect(read, 'Form_Load reads the channel').toBeDefined();
    expect(read!.edge.metadata?.refKind).toBe('property-get');
    expect(read!.edge.metadata?.errorChannel).toBe(true);

    // …and the caller relationship that makes the propagation real.
    const formLoad = read!.node;
    const callees = cg.getCallees(formLoad.id);
    expect(callees.some((c) => c.node.name === 'EstablecerDatos')).toBe(true);

    // Hop 3 — the reader DISPLAYS it. `MsgBox` is a VBA built-in, so it is not
    // a node; E3's `errorPolicy.behavior` is the graph's display marker, and
    // it reports both signals for this procedure (channel + display).
    expect(
      (formLoad.metadata?.errorPolicy as { behavior: string }).behavior,
    ).toBe('mixed');
    expect(
      (write!.node.metadata?.errorPolicy as { behavior: string }).behavior,
    ).toBe('channel');

    await cg.destroy();
  });

  it('reads `vba.errorChannel` out of codegraph.json', async () => {
    fs.writeFileSync(
      path.join(dir, 'codegraph.json'),
      JSON.stringify({ vba: { errorChannel: ['lastFailure', 'obj.Nope'] } }),
    );
    const vba = loadVbaConfig(dir);
    // The dotted entry is dropped by validation; the valid one survives.
    expect(vba.errorChannel).toEqual(['lastFailure']);

    fs.writeFileSync(
      path.join(dir, 'Fallos.bas'),
      [
        'Attribute VB_Name = "Fallos"',
        'Option Explicit',
        'Private lastFailure As String',
        '',
        'Public Sub Guardar()',
        '    lastFailure = "boom"',
        'End Sub',
      ].join('\n'),
    );

    const cg = CodeGraph.initSync(dir);
    await cg.indexAll();

    const channel = cg
      .getNodesByName('lastFailure')
      .find((n) => n.kind === 'variable');
    expect(channel).toBeDefined();
    const incoming = cg.getCallers(channel!.id);
    expect(
      incoming.some((c) => c.edge.metadata?.errorChannel === true),
    ).toBe(true);

    await cg.destroy();
  });
});
