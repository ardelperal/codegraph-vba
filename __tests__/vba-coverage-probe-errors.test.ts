/**
 * `scripts/vba-coverage-probe.mjs` — the `errorHandling` block (task E1 of
 * `docs/vba-error-handling-plan.md`).
 *
 * The probe is a measuring instrument and this block is the one that E2–E5
 * will be judged against, so every counter is pinned EXACTLY over a fixture
 * module carrying one procedure per protection class and one per handler
 * behaviour. A drift is either a real classifier change — which the roadmap
 * wants to see — or a bug. It must not pass silently.
 *
 * Two assertions carry more weight than the rest, because they are the two
 * mistakes the original throwaway scan made:
 *
 * - `classifies a p_Error write as channel, not as "other"` — the narrow
 *   hardcoded regex missed `p_Error` and dumped 1,899 bodies into an
 *   unclassified bucket that a 12-of-12 audit showed were channel writes.
 * - `classifies a body that both writes the channel and displays as mixed` —
 *   the original tested one signal after the other and let evaluation order
 *   silently pick a winner.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';
import { SqlQueryExtractor } from '../src/extraction/sql-query-extractor';
import {
  // @ts-expect-error — plain ESM script, no type declarations by design.
} from '../scripts/vba-coverage-probe.mjs';
// @ts-expect-error — plain ESM script, no type declarations by design.
import {
  runProbe,
  formatReport,
  maskVbaLine,
  joinContinuations,
  buildChannelWriteMatcher,
  analyzeVbaErrorHandling,
  DEFAULT_ERROR_CHANNEL_NAMES,
} from '../scripts/vba-coverage-probe.mjs';

const extractors = { VbaExtractor, VbaFormExtractor, SqlQueryExtractor };

/**
 * `modErrores.bas` — 15 procedures, laid out so every counter in the
 * `errorHandling` block has at least one contributor and no two counters can
 * be satisfied by the same procedure by accident.
 *
 *  1 GuardarCanal        handler   -> channel   (the house `p_Error` shape)
 *  2 GuardarDisplay      handler   -> display
 *  3 GuardarReraise      handler   -> reraise
 *  4 GuardarMixto        handler   -> mixed     (`Me.Error` write + MsgBox)
 *  5 GuardarDesconocido  handler   -> unknown   (empty handler body)
 *  6 SuprimirAbierto     resumeNext, scope never closed
 *  7 SuprimirCerrado     resumeNext, closed by `On Error GoTo 0`
 *  8 Corto               none, 1 statement line
 *  9 LargoSinIo          none, 7 statement lines, no I/O
 * 10 LargoConIo          none, 7 statement lines, DoCmd + a conversion
 * 11 SoloSalto           none — a label nobody targets is CONTROL FLOW
 * 12 Colgado             handler -> unknown, and a dangling `GoTo` target
 * 13 DosManejadores      handler -> channel, two `On Error GoTo` statements
 * 14 Literal             none — `"On Error GoTo errores"` inside a string
 * 15 ResetMenosUno       resumeNext, closed by `On Error GoTo -1`
 */
const BAS_SOURCE = [
  'Attribute VB_Name = "modErrores"',
  'Option Explicit',
  '',
  'Public Sub GuardarCanal()',
  '    On Error GoTo errores',
  '    Dim n As Long',
  '    n = 1',
  '    Exit Sub',
  'errores:',
  '    If Err.Number <> 1000 Then p_Error = "GuardarCanal: " & Err.Description',
  'End Sub',
  '',
  'Public Sub GuardarDisplay()',
  '    On Error GoTo errores',
  '    Dim n As Long',
  '    n = 2',
  '    Exit Sub',
  'errores:',
  '    MsgBox "Ha fallado"',
  'End Sub',
  '',
  'Public Sub GuardarReraise()',
  '    On Error GoTo errores',
  '    Dim n As Long',
  '    n = 3',
  '    Exit Sub',
  'errores:',
  '    Err.Raise 1000',
  'End Sub',
  '',
  'Public Sub GuardarMixto()',
  '    On Error GoTo errores',
  '    Dim n As Long',
  '    n = 4',
  '    Exit Sub',
  'errores:',
  '    Me.Error = "roto"',
  '    MsgBox "Ha fallado"',
  'End Sub',
  '',
  'Public Sub GuardarDesconocido()',
  '    On Error GoTo errores',
  '    Dim n As Long',
  '    n = 5',
  '    Exit Sub',
  'errores:',
  'End Sub',
  '',
  'Public Sub SuprimirAbierto()',
  '    On Error Resume Next',
  '    Dim n As Long',
  '    n = 6',
  'End Sub',
  '',
  'Public Sub SuprimirCerrado()',
  '    On Error Resume Next',
  '    Dim n As Long',
  '    n = 7',
  '    On Error GoTo 0',
  'End Sub',
  '',
  'Public Function Corto() As Long',
  '    Corto = 8',
  'End Function',
  '',
  'Public Sub LargoSinIo()',
  '    Dim a As Long',
  '    Dim b As Long',
  '    a = 1',
  '    b = 2',
  '    a = a + b',
  '    b = b + a',
  '    a = a + b',
  'End Sub',
  '',
  'Public Sub LargoConIo()',
  '    Dim a As Long',
  '    a = 1',
  '    DoCmd.OpenForm "FormMain"',
  '    a = 2',
  '    a = 3',
  '    a = CLng("4")',
  '    a = 5',
  'End Sub',
  '',
  'Public Sub SoloSalto()',
  '    Dim i As Long',
  '    For i = 1 To 3',
  '        If i = 2 Then GoTo siguiente',
  'siguiente:',
  '    Next i',
  'End Sub',
  '',
  'Public Sub Colgado()',
  '    On Error GoTo noExiste',
  '    Dim n As Long',
  '    n = 12',
  'End Sub',
  '',
  'Public Sub DosManejadores()',
  '    On Error GoTo primero',
  '    Dim n As Long',
  '    n = 13',
  '    On Error GoTo segundo',
  '    n = 14',
  '    Exit Sub',
  'primero:',
  '    p_Error = "primero"',
  '    Exit Sub',
  'segundo:',
  '    p_Error = "segundo"',
  'End Sub',
  '',
  'Public Sub Literal()',
  '    Dim s As String',
  '    s = "On Error GoTo errores"',
  '    Debug.Print s',
  'End Sub',
  '',
  'Public Sub ResetMenosUno()',
  '    On Error Resume Next',
  '    Dim n As Long',
  '    n = 15',
  '    On Error GoTo -1',
  'End Sub',
  '',
].join('\n');

describe('vba-coverage-probe — errorHandling', () => {
  let tmpDir: string;
  let basPath: string;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'vba-probe-errors-'));
    fs.mkdirSync(path.join(tmpDir, 'src'), { recursive: true });
    basPath = path.join(tmpDir, 'src', 'modErrores.bas');
    fs.writeFileSync(basPath, BAS_SOURCE);
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it('counts one procedure per declaration and never drifts from declaredProcedures', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errorHandling.proceduresTotal).toBe(15);
    // The starts come from the extractor's own non-stub function nodes, so
    // these two can only ever disagree if the probe is broken.
    expect(report.errorHandling.proceduresTotal).toBe(report.declaredProcedures);
  });

  it('splits procedures across the three protection classes', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errorHandling.proceduresByProtection).toEqual({
      // 1, 2, 3, 4, 5, 12, 13
      handler: 7,
      // 6, 7, 15
      resumeNext: 3,
      // 8, 9, 10, 11 (control-flow label only), 14 (string literal)
      none: 5,
    });
  });

  it('classifies every handler behaviour exclusively', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errorHandling.handlersByBehavior).toEqual({
      channel: 2, // GuardarCanal, DosManejadores
      display: 1, // GuardarDisplay
      reraise: 1, // GuardarReraise
      mixed: 1, // GuardarMixto
      unknown: 2, // GuardarDesconocido (empty body), Colgado (no body at all)
    });
    // The five buckets partition the handler procedures — nothing falls out.
    const total = Object.values(
      report.errorHandling.handlersByBehavior as Record<string, number>,
    ).reduce((a, b) => a + b, 0);
    expect(total).toBe(report.errorHandling.proceduresByProtection.handler);
  });

  it('reports the raw per-signal totals alongside the exclusive split', async () => {
    const report = await runProbe([tmpDir], { extractors });

    // Non-exclusive: GuardarMixto is counted under BOTH channel and display,
    // which is exactly the view the plan's §2.3 "970 handlers display" figure
    // is, and which the exclusive `mixed` bucket alone cannot express.
    expect(report.errorHandling.handlerSignals).toEqual({
      channel: 3, // GuardarCanal, GuardarMixto, DosManejadores
      display: 2, // GuardarDisplay, GuardarMixto
      reraise: 1, // GuardarReraise
    });
  });

  it('classifies a p_Error write as channel, not as "other"', async () => {
    // The regression this whole task exists for: the original scan's regex
    // missed `p_Error` and reported those bodies as unclassified.
    const records = analyzeVbaErrorHandling(basPath, BAS_SOURCE, [
      { startLine: 4, name: 'GuardarCanal' },
    ]);

    expect(records).toHaveLength(1);
    expect(records[0].behavior).toBe('channel');
    expect(records[0].signals).toEqual({
      channel: true,
      display: false,
      reraise: false,
    });
  });

  it('classifies a body that both writes the channel and displays as mixed', async () => {
    // Not "whichever branch was tested first".
    const report = await runProbe([tmpDir], { extractors });
    const records = analyzeVbaErrorHandling(basPath, BAS_SOURCE, [
      { startLine: 31, name: 'GuardarMixto' },
    ]);

    expect(records[0].procedure).toBe('GuardarMixto');
    expect(records[0].signals).toEqual({
      channel: true,
      display: true,
      reraise: false,
    });
    expect(records[0].behavior).toBe('mixed');
    expect(report.errorHandling.handlersByBehavior.mixed).toBe(1);
  });

  it('narrows unprotected procedures by size and then by I/O', async () => {
    const report = await runProbe([tmpDir], { extractors });

    // LargoSinIo and LargoConIo are both over five statement lines; Corto (1),
    // SoloSalto (4, the bare label does not count) and Literal (3) are not.
    expect(report.errorHandling.unprotectedOverFiveLoc).toBe(2);
    // `unprotectedTouchingIo` is a SUBSET of the above, matching §3.1's
    // cumulative 813 -> 310 -> 185 narrowing. Only LargoConIo touches I/O.
    expect(report.errorHandling.unprotectedTouchingIo).toBe(1);
  });

  it('counts every On Error GoTo site by label name', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errorHandling.handlerLabelNames).toEqual({
      errores: 5,
      noExiste: 1,
      primero: 1,
      segundo: 1,
    });
  });

  it('separates handler labels from pure control-flow labels', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errorHandling.labels).toEqual({
      // errores x5, siguiente, primero, segundo
      defined: 8,
      // everything but `siguiente`, which no On Error targets
      handlerTargets: 7,
      controlFlow: 1,
    });
  });

  it('counts only Resume Next scopes that no reset closes', async () => {
    const report = await runProbe([tmpDir], { extractors });

    // SuprimirAbierto stays open; SuprimirCerrado is closed by `GoTo 0` and
    // ResetMenosUno by `GoTo -1` (valid VBA, absent from the corpus, pinned
    // here so the zero-occurrence case is still covered).
    expect(report.errorHandling.resumeNextOpenScopes).toBe(1);
  });

  it('reports the dangling GoTo target with its file and procedure', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errorHandling.danglingGotoTargets).toEqual([
      { file: basPath, procedure: 'Colgado', label: 'noExiste' },
    ]);
  });

  it('counts procedures that swap handlers mid-body', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errorHandling.proceduresWithMultipleHandlers).toBe(1);
  });

  it('reports the raw construct counts behind the census table', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errorHandling.statements).toEqual({
      onErrorGoToLabel: 8,
      onErrorResumeNext: 3,
      onErrorGoToZero: 1,
      onErrorGoToMinusOne: 1,
      // the 8 labelled ones + `GoTo 0` + `GoTo -1` + SoloSalto's free jump
      gotoStatements: 11,
      resumeNext: 3,
      resumeLabel: 0,
      resumeBare: 0,
      errRaise: 1,
      errRaiseSentinel: 1,
      errNumber: 1,
      errDescription: 1,
      errClear: 0,
      errSource: 0,
    });
  });

  it('reports the classifier lists it actually used', async () => {
    const report = await runProbe([tmpDir], { extractors });

    expect(report.errorHandling.config.errorChannelNames).toEqual([
      'm_Error',
      'p_Error',
      'g_Error',
      'Error',
    ]);
    expect(report.errorHandling.config.displayCalls).toEqual([
      'MsgBox',
      'Debug.Print',
    ]);
    expect(report.errorHandling.config.ioMarkers.length).toBeGreaterThan(0);
  });

  it('is deterministic across runs over the same tree', async () => {
    const first = await runProbe([tmpDir], { extractors });
    const second = await runProbe([tmpDir], { extractors });

    expect(second.errorHandling).toEqual(first.errorHandling);
  });

  it('renders the error-handling tables into the markdown report', async () => {
    const report = await runProbe([tmpDir], { extractors });
    const md = formatReport(report);

    expect(md).toContain('### Error handling — protection');
    expect(md).toContain('### Error handling — handler behaviour');
    expect(md).toContain('| `mixed` | 1 |');
    expect(md).toContain('| writes the error channel | 3 |');
    expect(md).toContain('| `errores` | 5 |');
    expect(md).toContain('| dangling `GoTo` targets | 1 |');
  });
});

describe('vba-coverage-probe — the error channel is configurable', () => {
  const SOURCE = [
    'Attribute VB_Name = "modCanal"',
    '',
    'Public Sub Propio()',
    '    On Error GoTo errores',
    '    Dim n As Long',
    '    n = 1',
    '    Exit Sub',
    'errores:',
    '    lastFailure = Err.Description',
    'End Sub',
    '',
  ].join('\n');
  const PROCS = [{ startLine: 3, name: 'Propio' }];

  it('does not match a project-specific channel name by default', () => {
    const [record] = analyzeVbaErrorHandling('modCanal.bas', SOURCE, PROCS);

    expect(record.signals.channel).toBe(false);
    expect(record.behavior).toBe('unknown');
  });

  it('matches it once the name is added to the list', () => {
    const [record] = analyzeVbaErrorHandling('modCanal.bas', SOURCE, PROCS, {
      errorChannelNames: [...DEFAULT_ERROR_CHANNEL_NAMES, 'lastFailure'],
    });

    expect(record.signals.channel).toBe(true);
    expect(record.behavior).toBe('channel');
  });

  it('matches a name, never a substring', () => {
    const matches = buildChannelWriteMatcher(['Error']);

    expect(matches('Me.Error = "x"')).toBe(true);
    expect(matches('    obj.Error = "x"')).toBe(true);
    expect(matches('Error = "x"')).toBe(true);
    // `ErrorCount` is a different variable, and `m_Error` is a different
    // NAME — matching it here would need `m_Error` on the list, which is
    // exactly the point of the list.
    expect(matches('ErrorCount = 1')).toBe(false);
    expect(matches('m_Error = "x"')).toBe(false);
  });

  it('treats the channel in a condition as a read, not a write', () => {
    const matches = buildChannelWriteMatcher(['m_Error']);

    expect(matches('m_Error = "x"')).toBe(true);
    expect(matches('Set m_Error = Nothing')).toBe(true);
    expect(matches(' m_Error <> ""')).toBe(false);
    expect(matches('s = m_Error')).toBe(false);
  });
});

describe('vba-coverage-probe — error-handling scan discipline', () => {
  it('masks string literals and drops trailing comments', () => {
    expect(maskVbaLine('s = "On Error GoTo errores"')).toBe(
      's = ' + ' '.repeat('"On Error GoTo errores"'.length),
    );
    expect(maskVbaLine("n = 1 ' On Error GoTo errores")).toBe('n = 1 ');
    expect(maskVbaLine('n = 1 : Rem On Error GoTo errores')).toBe('n = 1 ');
    // A quote inside a comment must not reopen a literal.
    expect(maskVbaLine('n = 1 \' say "hi"')).toBe('n = 1 ');
  });

  it('sets nothing when the whole handler lives inside a string literal (#209)', () => {
    const source = [
      'Public Sub Literal()',
      '    Dim s As String',
      '    s = "On Error GoTo errores"',
      '    Debug.Print s',
      'End Sub',
      '',
    ].join('\n');
    const [record] = analyzeVbaErrorHandling('modLiteral.bas', source, [
      { startLine: 1, name: 'Literal' },
    ]);

    expect(record.protection).toBe('none');
    expect(record.behavior).toBeNull();
    expect(record.statements.onErrorGoToLabel).toBe(0);
    expect(record.definedLabels).toEqual([]);
  });

  it('joins line continuations before matching', () => {
    expect(joinContinuations(['On Error _', 'GoTo errores', 'x = 1'])).toEqual([
      'On Error GoTo errores',
      '',
      'x = 1',
    ]);
    // Chains collapse into the first physical line.
    expect(joinContinuations(['a _', 'b _', 'c'])).toEqual(['a b c', '', '']);
  });

  it('reads a handler split across a continuation', () => {
    const source = [
      'Public Sub Partido()',
      '    On Error _',
      '        GoTo errores',
      '    Dim n As Long',
      '    Exit Sub',
      'errores:',
      '    p_Error = Err.Description',
      'End Sub',
      '',
    ].join('\n');
    const [record] = analyzeVbaErrorHandling('modPartido.bas', source, [
      { startLine: 1, name: 'Partido' },
    ]);

    expect(record.protection).toBe('handler');
    expect(record.behavior).toBe('channel');
    expect(record.handlerTargets).toEqual(['errores']);
  });

  it('closes a colon-separated single-line procedure on its own line (#208)', () => {
    const source = [
      'Public Sub Corta(): On Error GoTo errores: Exit Sub: errores: End Sub',
      '',
      'Public Sub Siguiente()',
      '    Dim n As Long',
      'End Sub',
      '',
    ].join('\n');
    const records = analyzeVbaErrorHandling('modCorta.bas', source, [
      { startLine: 1, name: 'Corta' },
      { startLine: 3, name: 'Siguiente' },
    ]);

    // The whole procedure - header, handler and `End Sub` - lives on line 1,
    // so its scope must close there and leave `Siguiente` untouched.
    expect(records[0].protection).toBe('handler');
    expect(records[0].startLine).toBe(1);
    expect(records[0].endLine).toBe(1);
    expect(records[0].handlerTargets).toEqual(['errores']);
    expect(records[1].protection).toBe('none');
    expect(records[1].statements.onErrorGoToLabel).toBe(0);
  });

  it('keeps a colon-separated header open when End Sub is on a later line', () => {
    const source = [
      'Public Sub Corta(): On Error GoTo errores: Exit Sub',
      'errores:',
      '    m_Error = Err.Description',
      'End Sub',
      '',
      'Public Sub Siguiente()',
      '    Dim n As Long',
      'End Sub',
      '',
    ].join('\n');
    const records = analyzeVbaErrorHandling('modCorta.bas', source, [
      { startLine: 1, name: 'Corta' },
      { startLine: 6, name: 'Siguiente' },
    ]);

    // Only the header is colon-separated here, so the procedure genuinely runs
    // to its `End Sub` on line 4 and the handler label on line 2 is inside it.
    expect(records[0].protection).toBe('handler');
    expect(records[0].startLine).toBe(1);
    expect(records[0].endLine).toBe(4);
    expect(records[0].behavior).toBe('channel');
    expect(records[0].danglingTargets).toEqual([]);
    expect(records[1].protection).toBe('none');
    expect(records[1].statements.onErrorGoToLabel).toBe(0);
  });

  it('does not read a keyword-headed colon line as a label', () => {
    const source = [
      'Public Sub Bucle()',
      '    Dim i As Long',
      '    Select Case i',
      '        Case 1: i = 2',
      '    End Select',
      '    Do: i = i + 1: Loop While i < 3',
      'End Sub',
      '',
    ].join('\n');
    const [record] = analyzeVbaErrorHandling('modBucle.bas', source, [
      { startLine: 1, name: 'Bucle' },
    ]);

    expect(record.definedLabels).toEqual([]);
    expect(record.danglingTargets).toEqual([]);
  });
});
