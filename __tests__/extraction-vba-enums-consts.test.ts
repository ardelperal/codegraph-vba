/**
 * VbaExtractor — tests for `Enum` and `Const` declaration extraction.
 *
 * Dysflow exports the full module text into `.bas`/`.cls`, so a real
 * constants module (e.g. the project's own `constantes.bas` fixture) carries
 * `Public Const` lines and `Public Enum ... End Enum` blocks verbatim. Before
 * this slice the extractor parsed none of them, leaving the project's domain
 * dictionary (status enums, config constants) invisible in the graph.
 *
 * Contract (added requirements REQ-CODE-12 Enum, REQ-CODE-13 Const in
 * `openspec/specs/vba-code-extraction/spec.md`):
 *   - `Enum` → one `enum` node + one `enum_member` node per member, with a
 *     `contains` edge enum→member and module/class→enum.
 *   - `Const` → one `constant` node per declared name (multi-name lines emit
 *     one per name), with a `contains` edge module/class→constant.
 *   - Visibility folds like procedures: `Private` → 'private'; `Public`,
 *     `Global`, `Friend`, or no keyword → 'public'.
 *   - A file containing ONLY an Enum or Const (no Option-only) now emits the
 *     lazy module/class node, because it has real symbols.
 */
import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

describe('VbaExtractor — Enum declarations (REQ-CODE-12)', () => {
  const src = `Attribute VB_Name = "constantes"
Public Enum EnumTipoUsuario
    Administrador = 1
    Calidad = 2
    Secretaria = 3
End Enum`;

  it('emits an enum node with public visibility', () => {
    const r = extract('src/modules/constantes.bas', src);
    const en = r.nodes.find((n) => n.kind === 'enum' && n.name === 'EnumTipoUsuario');
    expect(en).toBeDefined();
    expect(en?.visibility).toBe('public');
    expect(en?.language).toBe('vba');
  });

  it('emits one enum_member node per member', () => {
    const r = extract('src/modules/constantes.bas', src);
    const members = r.nodes.filter((n) => n.kind === 'enum_member');
    const names = members.map((m) => m.name).sort();
    expect(names).toEqual(['Administrador', 'Calidad', 'Secretaria']);
  });

  it('qualifies enum_member names as Enum.Member', () => {
    const r = extract('src/modules/constantes.bas', src);
    const admin = r.nodes.find((n) => n.kind === 'enum_member' && n.name === 'Administrador');
    expect(admin?.qualifiedName).toBe('EnumTipoUsuario.Administrador');
  });

  it('emits a contains edge from the enum to each member', () => {
    const r = extract('src/modules/constantes.bas', src);
    const en = r.nodes.find((n) => n.kind === 'enum');
    const calidad = r.nodes.find((n) => n.kind === 'enum_member' && n.name === 'Calidad');
    const edge = r.edges.find(
      (e) => e.kind === 'contains' && e.source === en?.id && e.target === calidad?.id,
    );
    expect(edge).toBeDefined();
  });

  it('emits a contains edge from the module to the enum', () => {
    const r = extract('src/modules/constantes.bas', src);
    const mod = r.nodes.find((n) => n.kind === 'module');
    const en = r.nodes.find((n) => n.kind === 'enum');
    expect(mod).toBeDefined();
    const edge = r.edges.find(
      (e) => e.kind === 'contains' && e.source === mod?.id && e.target === en?.id,
    );
    expect(edge).toBeDefined();
  });

  it('does NOT emit enum members as constant nodes', () => {
    const r = extract('src/modules/constantes.bas', src);
    const consts = r.nodes.filter((n) => n.kind === 'constant');
    expect(consts).toHaveLength(0);
  });

  it('folds Private Enum visibility to private', () => {
    const r = extract('src/modules/m.bas', `Private Enum Hidden
    A = 1
End Enum`);
    const en = r.nodes.find((n) => n.kind === 'enum' && n.name === 'Hidden');
    expect(en?.visibility).toBe('private');
  });
});

describe('VbaExtractor — Const declarations (REQ-CODE-13)', () => {
  it('emits a constant node with public visibility for Public Const', () => {
    const r = extract('src/modules/c.bas', 'Public Const msoFileDialogOpen As Long = 1');
    const c = r.nodes.find((n) => n.kind === 'constant' && n.name === 'msoFileDialogOpen');
    expect(c).toBeDefined();
    expect(c?.visibility).toBe('public');
    expect(c?.language).toBe('vba');
  });

  it('folds Private Const visibility to private', () => {
    const r = extract('src/modules/c.bas', 'Private Const Secret = 42');
    const c = r.nodes.find((n) => n.kind === 'constant' && n.name === 'Secret');
    expect(c?.visibility).toBe('private');
  });

  it('captures a hex-valued const name', () => {
    const r = extract('src/modules/c.bas', 'Public Const STILL_ACTIVE = &H103');
    const c = r.nodes.find((n) => n.kind === 'constant' && n.name === 'STILL_ACTIVE');
    expect(c).toBeDefined();
  });

  it('emits one constant node per name on a multi-name Const line', () => {
    const r = extract('src/modules/c.bas', 'Const A = 1, B = 2, C = 3');
    const names = r.nodes.filter((n) => n.kind === 'constant').map((n) => n.name).sort();
    expect(names).toEqual(['A', 'B', 'C']);
  });

  it('emits a contains edge from the module to each constant', () => {
    const r = extract('src/modules/c.bas', 'Public Const Foo = 1');
    const mod = r.nodes.find((n) => n.kind === 'module');
    const c = r.nodes.find((n) => n.kind === 'constant' && n.name === 'Foo');
    expect(mod).toBeDefined();
    const edge = r.edges.find(
      (e) => e.kind === 'contains' && e.source === mod?.id && e.target === c?.id,
    );
    expect(edge).toBeDefined();
  });
});

describe('VbaExtractor — lazy module node now fires on Enum/Const symbols', () => {
  it('a const-only .bas emits a module node (REQ-CODE-10 narrowed to Option-only)', () => {
    const r = extract('src/modules/constantes.bas', `Attribute VB_Name = "constantes"
Option Explicit
Public Const X = 1`);
    const mod = r.nodes.find((n) => n.kind === 'module');
    expect(mod).toBeDefined();
    expect(mod?.name).toBe('constantes');
  });

  it('an Option-only .bas still emits nothing (no module node)', () => {
    const r = extract('src/modules/empty.bas', `Option Explicit
Option Compare Database`);
    const mod = r.nodes.find((n) => n.kind === 'module' || n.kind === 'class');
    expect(mod).toBeUndefined();
  });
});
