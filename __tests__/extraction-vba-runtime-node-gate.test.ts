/**
 * Issue #245 — the runtime-object stub-NODE gate.
 *
 * The extractor synthesizes a `function` node for every qualified call it
 * cannot resolve inside the current file. Two allowlists decided what was
 * runtime noise and they had drifted: the resolver's `RUNTIME_OBJECTS`
 * already knew `VBA`, `fso`, `Collection` and `ListBox`, but the extractor's
 * `RUNTIME_RECEIVER_BLACKLIST` did not — so those calls kept producing stub
 * nodes that leaked into symbol search and node counts. Both sets are now
 * derived from one literal, and `calls.ts` gates node synthesis on
 * `isRuntimeObject(receiverType)` before `generateNodeId`.
 *
 * The gate removes NODES, not REFERENCES: the `unresolved_refs` rows the
 * resolver's `declined-runtime` accounting is built on must be identical
 * before and after, which the "reference accounting" block below pins.
 */
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import {
  RUNTIME_OBJECTS,
  RUNTIME_RECEIVER_BLACKLIST,
  isRuntimeObject,
} from '../src/extraction/vba/runtime-objects';
import {
  RUNTIME_OBJECTS as RESOLUTION_RUNTIME_OBJECTS,
  isRuntimeObject as resolutionIsRuntimeObject,
} from '../src/resolution/vba-runtime-objects';

const CLS_HEADER = [
  'VERSION 1.0 CLASS',
  'BEGIN',
  '  MultiUse = -1',
  'END',
];

function extractBas(source: string[]) {
  return new VbaExtractor('src/modules/Caller.bas', source.join('\n')).extract();
}

/** Function nodes the extractor synthesized for an unresolved qualified call. */
function stubNodes(result: ReturnType<VbaExtractor['extract']>) {
  return result.nodes.filter(
    (n) => n.kind === 'function' && (n.metadata as { stub?: unknown } | undefined)?.stub === true,
  );
}

/** Real procedure declarations — function nodes WITHOUT `metadata.stub`. */
function declaredProcedures(result: ReturnType<VbaExtractor['extract']>) {
  return result.nodes.filter(
    (n) => n.kind === 'function' && (n.metadata as { stub?: unknown } | undefined)?.stub !== true,
  );
}

function unresolvedByKind(result: ReturnType<VbaExtractor['extract']>) {
  const counts: Record<string, number> = {};
  for (const ref of result.unresolvedReferences ?? []) {
    const kind = ref.referenceKind ?? '(none)';
    counts[kind] = (counts[kind] ?? 0) + 1;
  }
  return counts;
}

describe('Issue #245: runtime-object receivers never synthesize stub nodes', () => {
  it('`Call VBA.DoEvents` produces zero nodes and zero calls edges', () => {
    const result = extractBas([
      'Attribute VB_Name = "Caller"',
      'Option Explicit',
      '',
      'Public Sub Esperar()',
      '    Call VBA.DoEvents',
      'End Sub',
      '',
    ]);

    expect(stubNodes(result)).toHaveLength(0);
    expect(result.nodes.some((n) => n.name === 'VBA.DoEvents')).toBe(false);
    expect(result.edges.filter((e) => e.kind === 'calls')).toHaveLength(0);
    // The real procedure is untouched.
    expect(declaredProcedures(result).map((n) => n.name)).toEqual(['Esperar']);
  });

  it('a `New Collection` receiver stops producing `Collection.Add` stubs', () => {
    const result = extractBas([
      'Attribute VB_Name = "Caller"',
      'Option Explicit',
      '',
      'Public Sub Acumular()',
      '    Dim col As Collection',
      '    Set col = New Collection',
      '    col.Add "uno"',
      '    col.Add "dos"',
      'End Sub',
      '',
    ]);

    expect(stubNodes(result).map((n) => n.name)).toEqual([]);
  });

  it('paren-form runtime calls (`fso.FileExists(...)`, `VBA.Format(...)`) are gated too', () => {
    // Late-bound `fso` — no `Dim`, so the receiver survives as raw text and
    // reaches the gate as `fso`, exactly the corpus shape.
    const result = extractBas([
      'Attribute VB_Name = "Caller"',
      'Option Explicit',
      '',
      'Public Function Existe(ByVal ruta As String) As Boolean',
      '    Existe = fso.FileExists(ruta)',
      '    Debug.Print VBA.Format(Now, "yyyy")',
      'End Function',
      '',
    ]);

    expect(stubNodes(result).map((n) => n.name)).toEqual([]);
  });

  it('the gate keys on the receiver TYPE, so a user class keeps its stub', () => {
    // `Add` collides with `Collection.Add`; `MiClase` is not a runtime object,
    // so the stub the resolver later repoints must survive.
    const result = extractBas([
      'Attribute VB_Name = "Caller"',
      'Option Explicit',
      '',
      'Public Sub Acumular()',
      '    Dim m As MiClase',
      '    Set m = New MiClase',
      '    m.Add "uno"',
      'End Sub',
      '',
    ]);

    expect(stubNodes(result).map((n) => n.name)).toEqual(['MiClase.Add']);
  });

  it('a cross-module `.bas`-qualified call is not a runtime object and still stubs', () => {
    const result = extractBas([
      'Attribute VB_Name = "Caller"',
      'Option Explicit',
      '',
      'Public Sub Ir()',
      '    mdlCursor.MouseCursor 11',
      'End Sub',
      '',
    ]);

    expect(stubNodes(result).map((n) => n.name)).toEqual(['mdlCursor.MouseCursor']);
  });

  it('a class module declaring the colliding member still declares it (no gate on declarations)', () => {
    const result = new VbaExtractor(
      'src/classes/MiClase.cls',
      [
        ...CLS_HEADER,
        'Attribute VB_Name = "MiClase"',
        'Option Explicit',
        '',
        'Public Sub Add()',
        'End Sub',
        '',
        'Public Sub DoEvents()',
        'End Sub',
        '',
      ].join('\n'),
    ).extract();

    expect(declaredProcedures(result).map((n) => n.name).sort()).toEqual(['Add', 'DoEvents']);
    expect(stubNodes(result)).toHaveLength(0);
  });
});

describe('Issue #245: reference accounting is untouched', () => {
  // The gate must delete nodes only. Adding the runtime-call lines to a file
  // must change nothing in `unresolvedReferences` — same kinds, same counts —
  // otherwise the resolver's `declined-runtime` bookkeeping shifts.
  const withoutRuntimeCalls = [
    'Attribute VB_Name = "Caller"',
    'Option Explicit',
    '',
    'Public Sub Mixto()',
    '    Dim col As Collection',
    '    Set col = New Collection',
    '    ProcedimientoQueNoExiste 1',
    '    OtroQueFalta(2)',
    'End Sub',
    '',
  ];
  const withRuntimeCalls = [
    'Attribute VB_Name = "Caller"',
    'Option Explicit',
    '',
    'Public Sub Mixto()',
    '    Dim col As Collection',
    '    Set col = New Collection',
    '    col.Add "uno"',
    '    Call VBA.DoEvents',
    '    ProcedimientoQueNoExiste 1',
    '    OtroQueFalta(2)',
    'End Sub',
    '',
  ];

  it('runtime call lines contribute no unresolved references of their own', () => {
    expect(unresolvedByKind(extractBas(withRuntimeCalls))).toEqual(
      unresolvedByKind(extractBas(withoutRuntimeCalls)),
    );
  });

  it('the genuinely-missing user calls are still surfaced', () => {
    const names = (extractBas(withRuntimeCalls).unresolvedReferences ?? []).map(
      (r) => r.referenceName,
    );
    expect(names).toContain('ProcedimientoQueNoExiste');
    expect(names).toContain('OtroQueFalta');
  });
});

describe('Issue #245: one literal, two derived views', () => {
  it('the receiver blacklist is a strict subset of the canonical runtime-object set', () => {
    expect(RUNTIME_RECEIVER_BLACKLIST.size).toBeGreaterThan(0);
    for (const name of RUNTIME_RECEIVER_BLACKLIST) {
      expect(RUNTIME_OBJECTS.has(name.toLowerCase()), name).toBe(true);
    }
    expect(RUNTIME_OBJECTS.size).toBeGreaterThan(RUNTIME_RECEIVER_BLACKLIST.size);
  });

  it('the blacklist keeps its exact case-sensitive membership (widening it deletes unresolved rows)', () => {
    expect([...RUNTIME_RECEIVER_BLACKLIST].sort()).toEqual([
      'Application',
      'CommandBars',
      'CreateObject',
      'Debug',
      'DoCmd',
      'Err',
      'Fields',
      'Forms',
      'GetObject',
      'Modules',
      'References',
      'Reports',
      'Screen',
      'SysCmd',
    ]);
  });

  it('the names the two lists disagreed on are now in the canonical set', () => {
    for (const name of ['vba', 'fso', 'collection', 'listbox', 'combobox', 'textbox', 'dao']) {
      expect(RUNTIME_OBJECTS.has(name), name).toBe(true);
      expect(RUNTIME_RECEIVER_BLACKLIST.has(name), name).toBe(false);
    }
  });

  it('every canonical entry is stored lowercase and matches case-insensitively', () => {
    for (const name of RUNTIME_OBJECTS) {
      expect(name).toBe(name.toLowerCase());
      expect(isRuntimeObject(name.toUpperCase())).toBe(true);
      expect(isRuntimeObject(`[${name}]`)).toBe(true);
    }
    expect(isRuntimeObject('MiClase')).toBe(false);
    expect(isRuntimeObject(undefined)).toBe(false);
  });

  it('the resolution module re-exports the very same set and predicate', () => {
    expect(RESOLUTION_RUNTIME_OBJECTS).toBe(RUNTIME_OBJECTS);
    expect(resolutionIsRuntimeObject).toBe(isRuntimeObject);
  });

  it('src/extraction does not import from src/resolution (layering)', () => {
    // The canonical module is a leaf: no imports at all.
    const canonical = readFileSync(
      new URL('../src/extraction/vba/runtime-objects.ts', import.meta.url),
      'utf8',
    );
    expect(canonical).not.toMatch(/^\s*import\s/m);

    // Scoped to the VBA extraction subtree: the framework plumbing in
    // `src/extraction/index.ts` has its own long-standing dependency on
    // `src/resolution/types`, which is out of scope here.
    const extractionDir = fileURLToPath(new URL('../src/extraction/vba/', import.meta.url));
    const offenders: string[] = [];
    const walk = (dir: string) => {
      for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const full = join(dir, entry.name);
        if (entry.isDirectory()) walk(full);
        else if (entry.name.endsWith('.ts')) {
          const src = readFileSync(full, 'utf8');
          if (/from\s+['"][^'"]*resolution[^'"]*['"]/.test(src)) offenders.push(full);
        }
      }
    };
    walk(extractionDir);
    expect(offenders).toEqual([]);
  });
});
