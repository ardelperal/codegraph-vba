/**
 * Issue #83 — single-pass walker dispatcher contract.
 *
 * The VbaExtractor consolidates its 6+ per-concern sweeps into one
 * `src.split('\n')` followed by a single per-line walk. This test
 * pins the dispatcher contract: every pre-split line is handed to every
 * classifier in the stable order documented in `vba-extractor.ts`, in
 * line order, with no second split. It does NOT exercise the extraction
 * shapes — those live in `extraction-vba.test.ts`.
 */
import { describe, it, expect } from 'vitest';
import { createProceduresClassifier } from '../src/extraction/vba/procedures';
import { createEventsTypesDeclaresClassifier } from '../src/extraction/vba/declarations';
import { createImplementsClassifier } from '../src/extraction/vba/implements';
import { createDimsClassifier } from '../src/extraction/vba/dims';
import { createEnumsConstsClassifier } from '../src/extraction/vba/enums-consts';
import { createCallsAndSqlClassifier } from '../src/extraction/vba/call-sweep';

const SRC = [
  'Attribute VB_Name = "Mod"',
  'Option Explicit',
  '',
  'Public Sub Bar(p As Long)',
  '    Dim x As ProjectClass',
  '    Inner 1',
  'End Sub',
  '',
  'Public Sub Inner(p As Long)',
  'End Sub',
  '',
].join('\n');

describe('Issue #83 — VbaWalker dispatcher contract', () => {
  it('splits the source ONCE and dispatches every pre-split line to every classifier in stable order', () => {
    // Track the order each classifier is called. A real classifier would
    // do real work, but this test only asserts the dispatch shape.
    const order: string[] = [];
    const fake = (name: string) => ({
      name,
      count: 0,
      classifyLine(_line: string, i: number, _ctx: any) {
        order.push(`${name}:${i}`);
        this.count++;
      },
    });

    const procedures = fake('procedures');
    const declarations = fake('eventsTypesDeclares');
    const implements = fake('implements');
    const dims = fake('dims');
    const enumsConsts = fake('enumsConsts');
    const callsAndSql = fake('callsAndSql');

    // ONE split, shared by all classifiers.
    const lines = SRC.split('\n');

    // Mirror the orchestrator's order: procedures pre-walk, then the
    // main walk dispatches every other classifier per line.
    for (let i = 0; i < lines.length; i++) {
      procedures.classifyLine(lines[i] ?? '', i, {});
    }
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] ?? '';
      declarations.classifyLine(line, i, {});
      implements.classifyLine(line, i, {});
      dims.classifyLine(line, i, {});
      enumsConsts.classifyLine(line, i, {});
      callsAndSql.classifyLine(line, i, {});
    }

    // Every classifier was called once per line.
    expect(procedures.count).toBe(lines.length);
    expect(declarations.count).toBe(lines.length);
    expect(implements.count).toBe(lines.length);
    expect(dims.count).toBe(lines.length);
    expect(enumsConsts.count).toBe(lines.length);
    expect(callsAndSql.count).toBe(lines.length);

    // Pre-walk: procedures runs first for EVERY line, in line order.
    expect(order[0]).toBe('procedures:0');
    expect(order[1]).toBe('procedures:1');
    // After the procedures pre-walk, the main walk dispatches
    // declarations → implements → dims → enumsConsts → callsAndSql
    // for each line, in order, starting back at line 0.
    expect(order[lines.length]).toBe('eventsTypesDeclares:0');
    expect(order[lines.length + 1]).toBe('implements:0');
    expect(order[lines.length + 2]).toBe('dims:0');
    expect(order[lines.length + 3]).toBe('enumsConsts:0');
    expect(order[lines.length + 4]).toBe('callsAndSql:0');
    expect(order[lines.length + 5]).toBe('eventsTypesDeclares:1');
    // Total: 1 procedures call per line + 5 main-walk calls per line.
    expect(order.length).toBe(lines.length * 6);
  });
});
