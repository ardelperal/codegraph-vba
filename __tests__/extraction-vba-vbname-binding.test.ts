/**
 * extraction-vba-vbname-binding.test.ts
 *
 * Acceptance tests for issue #249 — Access code-behind binding must follow
 * the module, not the filename.
 *
 * Event-handler synthesis and the `Me.<Control>` sweep both gated on the FILE
 * BASENAME starting with `Form_` / `Report_`. When the filename and the
 * module's `Attribute VB_Name` disagree, the file still parsed and still
 * emitted its procedures, but the form came out with no event wiring and no
 * control references — and nothing errored or warned.
 *
 * The fixtures make that concrete. `Form_Expediente.cls` and `sample2.cls`
 * under `__tests__/fixtures/vba-vbname-binding/` are BYTE-IDENTICAL; the only
 * difference between them is the name on disk. So every difference these
 * tests find between the two extractions is the bug.
 *
 * The other half of the file is the guard that must NOT move. A plain
 * service class (`InformeRiesgoPDFServicio.cls`) whose methods end in real
 * Access event names must still produce zero control stubs — that guard is
 * what stopped ~550 spurious nodes in real projects, and widening it is the
 * failure mode this change had to avoid.
 *
 * Real files, real extractors, no mocking.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { VbaFormExtractor } from '../src/extraction/vba-form-extractor';
import type {
  Edge,
  ExtractionResult,
  Node,
  UnresolvedReference,
} from '../src/types';

const FIXTURE_DIR = path.join(__dirname, 'fixtures', 'vba-vbname-binding');
const NAMED_CLS = path.join(FIXTURE_DIR, 'Form_Expediente.cls');
const RENAMED_CLS = path.join(FIXTURE_DIR, 'sample2.cls');
const NAMED_TXT = path.join(FIXTURE_DIR, 'Form_Expediente.form.txt');
const REPORT_CLS = path.join(FIXTURE_DIR, 'sample_report.cls');
const SERVICE_CLS = path.join(FIXTURE_DIR, 'InformeRiesgoPDFServicio.cls');
const NEAR_MISS_CLS = path.join(FIXTURE_DIR, 'FormularioVentas.cls');

function extractCode(p: string): ExtractionResult {
  return new VbaExtractor(p, fs.readFileSync(p, 'utf8')).extract();
}

function fn(result: ExtractionResult, name: string): Node | undefined {
  return result.nodes.find((n) => n.kind === 'function' && n.name === name);
}

/** Every `event-handler` edge leaving the named Sub. */
function handlerEdges(result: ExtractionResult, subName: string): Edge[] {
  const sub = fn(result, subName);
  if (!sub) return [];
  return result.edges.filter(
    (e) => e.kind === 'event-handler' && e.source === sub.id,
  );
}

/**
 * The comparable shape of an event-handler edge: its event, its scope, and
 * the NAME + kind of whatever it points at. The node ids themselves are
 * deliberately excluded — they hash the sibling layout path, which differs
 * between the two files by design (the sibling lives beside the FILE).
 */
function handlerShape(result: ExtractionResult, subName: string) {
  return handlerEdges(result, subName)
    .map((e) => {
      const target = result.nodes.find((n) => n.id === e.target);
      return {
        eventName: e.metadata?.eventName,
        scope: e.metadata?.scope ?? null,
        targetKind: target?.kind,
        targetName: target?.name,
      };
    })
    .sort((a, b) => String(a.targetName).localeCompare(String(b.targetName)));
}

function meControlRefs(result: ExtractionResult): UnresolvedReference[] {
  return result.unresolvedReferences.filter(
    (r) => r.metadata?.synthesizedBy === 'vba-me-control',
  );
}

const named = extractCode(NAMED_CLS);
const renamed = extractCode(RENAMED_CLS);
const report = extractCode(REPORT_CLS);
const service = extractCode(SERVICE_CLS);
const nearMiss = extractCode(NEAR_MISS_CLS);

// =============================================================================
// The premise: the two .cls fixtures really are the same bytes. If this ever
// stops holding, every comparison below is measuring the wrong thing.
// =============================================================================

describe('issue #249 — fixture premise', () => {
  it('Form_Expediente.cls and sample2.cls are byte-identical', () => {
    expect(fs.readFileSync(RENAMED_CLS)).toEqual(fs.readFileSync(NAMED_CLS));
  });
});

// =============================================================================
// AC #1 — identical source under either filename produces the same set of
// event-handler edges, control-level and form-level alike.
// =============================================================================

describe('issue #249 — event-handler edges follow VB_Name', () => {
  it('the renamed file binds cmdGuardar_Click exactly as the named one does', () => {
    const expected = [
      {
        eventName: 'Click',
        scope: null,
        targetKind: 'form-instance-control',
        targetName: 'cmdGuardar',
      },
    ];
    expect(handlerShape(named, 'cmdGuardar_Click')).toEqual(expected);
    expect(handlerShape(renamed, 'cmdGuardar_Click')).toEqual(expected);
  });

  it('the renamed file binds the Form_Load lifecycle handler to a form-layout node', () => {
    expect(handlerShape(named, 'Form_Load')).toEqual([
      {
        eventName: 'Load',
        scope: 'form',
        targetKind: 'form-layout',
        targetName: 'Form_Expediente',
      },
    ]);
    // The layout NAME still comes from the file, because that is the sibling
    // that exists on disk next to it.
    expect(handlerShape(renamed, 'Form_Load')).toEqual([
      {
        eventName: 'Load',
        scope: 'form',
        targetKind: 'form-layout',
        targetName: 'sample2',
      },
    ]);
  });

  it('both files emit the same number of event-handler edges', () => {
    const count = (r: ExtractionResult) =>
      r.edges.filter((e) => e.kind === 'event-handler').length;
    expect(count(renamed)).toBe(count(named));
    expect(count(named)).toBe(2);
  });
});

// =============================================================================
// AC #1 (second half) — the `Me.<Control>` sweep.
// =============================================================================

describe('issue #249 — Me.<Control> references follow VB_Name', () => {
  it('the renamed file emits the same control references as the named one', () => {
    const names = (r: ExtractionResult) =>
      meControlRefs(r)
        .map((ref) => ref.referenceName)
        .sort();
    expect(names(named)).toEqual(['cboUsuario', 'txtNombre']);
    expect(names(renamed)).toEqual(['cboUsuario', 'txtNombre']);
  });

  it('the sibling layout path is derived from the FILE, not from VB_Name', () => {
    const siblingOf = (r: ExtractionResult) =>
      new Set(meControlRefs(r).map((ref) => ref.metadata?.siblingPath));
    expect(siblingOf(named)).toEqual(
      new Set([NAMED_TXT.replace(/\\/g, '/')]),
    );
    expect(siblingOf(renamed)).toEqual(
      new Set([
        path.join(FIXTURE_DIR, 'sample2.form.txt').replace(/\\/g, '/'),
      ]),
    );
  });
});

// =============================================================================
// AC #2 — the mismatch is diagnosable: `bindingSource: 'vb-name'`, and only
// on the mismatch.
// =============================================================================

describe('issue #249 — bindingSource marks the mismatch', () => {
  it('every synthesized edge from the renamed file carries bindingSource', () => {
    const edges = renamed.edges.filter((e) => e.kind === 'event-handler');
    expect(edges.length).toBeGreaterThan(0);
    for (const e of edges) {
      expect(e.metadata?.bindingSource).toBe('vb-name');
    }
  });

  it('every control reference from the renamed file carries bindingSource', () => {
    const refs = meControlRefs(renamed);
    expect(refs.length).toBeGreaterThan(0);
    for (const r of refs) {
      expect(r.metadata?.bindingSource).toBe('vb-name');
    }
  });

  it('the filename fast path stays unmarked', () => {
    for (const e of named.edges.filter((e) => e.kind === 'event-handler')) {
      expect(e.metadata?.bindingSource).toBeUndefined();
    }
    for (const r of meControlRefs(named)) {
      expect(r.metadata?.bindingSource).toBeUndefined();
    }
  });
});

// =============================================================================
// The filename fast path is untouched: the stub the code-behind emits still
// carries the id `VbaFormExtractor` produces for the real layout node, so the
// INSERT OR REPLACE convergence keeps working. Not hardcoded — the real form
// extractor runs on the real sibling file.
// =============================================================================

describe('issue #249 — the filename fast path still converges on the real layout node', () => {
  it('the Form_Load stub id matches the real form-layout node id', () => {
    const layout = new VbaFormExtractor(
      NAMED_TXT,
      fs.readFileSync(NAMED_TXT, 'utf8'),
    ).extract();
    const real = layout.nodes.find((n) => n.kind === 'form-layout');
    expect(real).toBeDefined();
    const edge = handlerEdges(named, 'Form_Load')[0];
    expect(edge?.target).toBe(real!.id);
  });
});

// =============================================================================
// The report half of the fallback: `Report_*` must pick `.report.txt`.
// =============================================================================

describe('issue #249 — the report fallback picks the report sibling', () => {
  it('binds Report_Open to a report-layout node beside the file', () => {
    expect(handlerShape(report, 'Report_Open')).toEqual([
      {
        eventName: 'Open',
        scope: 'form',
        targetKind: 'report-layout',
        targetName: 'sample_report',
      },
    ]);
    const edge = handlerEdges(report, 'Report_Open')[0];
    const target = report.nodes.find((n) => n.id === edge?.target);
    expect(target?.filePath).toBe(
      path.join(FIXTURE_DIR, 'sample_report.report.txt'),
    );
  });

  it('still binds the report control handler', () => {
    expect(handlerShape(report, 'txtTotal_Click')).toEqual([
      {
        eventName: 'Click',
        scope: null,
        targetKind: 'form-instance-control',
        targetName: 'txtTotal',
      },
    ]);
  });
});

// =============================================================================
// AC #3 — the guard that must not move. A plain service class binds to
// nothing, even though `Documento_Print` and `Cabecera_Format` end in real
// Access event names and would otherwise become control stubs.
// =============================================================================

describe('issue #249 — a plain service class still binds to nothing', () => {
  for (const [label, result] of [
    ['InformeRiesgoPDFServicio', service],
    ['FormularioVentas (Form-without-underscore near miss)', nearMiss],
  ] as const) {
    describe(label, () => {
      it('produces zero form-instance-control stubs', () => {
        expect(
          result.nodes.filter((n) => n.kind === 'form-instance-control'),
        ).toEqual([]);
      });

      it('produces zero layout stubs', () => {
        expect(
          result.nodes.filter(
            (n) => n.kind === 'form-layout' || n.kind === 'report-layout',
          ),
        ).toEqual([]);
      });

      it('produces zero event-handler edges', () => {
        expect(result.edges.filter((e) => e.kind === 'event-handler')).toEqual(
          [],
        );
      });

      it('produces zero Me.<Control> references', () => {
        expect(meControlRefs(result)).toEqual([]);
      });

      it('still emits its procedures', () => {
        expect(
          result.nodes.filter((n) => n.kind === 'function').length,
        ).toBeGreaterThan(0);
      });
    });
  }
});
