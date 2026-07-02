import { describe, it, expect } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';

function extract(filePath: string, source: string) {
  return new VbaExtractor(filePath, source).extract();
}

describe('VbaExtractor — roadmap #26 event declarations', () => {
  it('models Event declarations and RaiseEvent calls as first-class graph nodes and edges', () => {
    const src = [
      'Attribute VB_Name = "PedidoPublisher"',
      'Public Event PedidoGuardado(ByVal IdPedido As Long)',
      '',
      'Public Sub Guardar()',
      '  RaiseEvent PedidoGuardado(42)',
      'End Sub',
    ].join('\n');

    const r = extract('src/classes/PedidoPublisher.cls', src);
    const event = r.nodes.find((n) => n.kind === 'event' && n.name === 'PedidoGuardado');
    const caller = r.nodes.find((n) => n.kind === 'function' && n.name === 'Guardar');

    expect(event).toBeDefined();
    expect(event?.qualifiedName).toBe('PedidoPublisher.PedidoGuardado');
    expect(event?.visibility).toBe('public');
    expect(event?.signature).toBe('Public Event PedidoGuardado(ByVal IdPedido As Long)');
    expect(caller).toBeDefined();
    expect(r.edges).toContainEqual(
      expect.objectContaining({
        source: caller?.id,
        target: event?.id,
        kind: 'raises-event',
        metadata: expect.objectContaining({ eventName: 'PedidoGuardado' }),
      }),
    );
  });

  it('models WithEvents declarations as subscriber edges with the variable name attached', () => {
    const r = extract(
      'src/classes/FormListener.cls',
      'Private WithEvents m_Form As Form_Pedido',
    );

    const target = r.nodes.find((n) => n.name === 'Form_Pedido');
    expect(target).toBeDefined();
    expect(r.edges).toContainEqual(
      expect.objectContaining({
        target: target?.id,
        kind: 'subscribes-event',
        metadata: expect.objectContaining({
          synthesizedBy: 'vba-withevents',
          variableName: 'm_Form',
        }),
      }),
    );
  });
});

describe('VbaExtractor — roadmap #26 Type declarations', () => {
  it('models Type declarations with type_member nodes and type-member edges', () => {
    const src = [
      'Attribute VB_Name = "Tipos"',
      'Private Type TPedido',
      '  Id As Long',
      '  Nombre As String',
      'End Type',
    ].join('\n');

    const r = extract('src/modules/Tipos.bas', src);
    const type = r.nodes.find((n) => n.kind === 'type' && n.name === 'TPedido');
    const members = r.nodes.filter((n) => n.kind === 'type_member');

    expect(type).toBeDefined();
    expect(type?.qualifiedName).toBe('TPedido');
    expect(type?.visibility).toBe('private');
    expect(members.map((n) => n.name).sort()).toEqual(['Id', 'Nombre']);
    expect(members.find((n) => n.name === 'Id')?.metadata?.memberType).toBe('Long');
    expect(members.find((n) => n.name === 'Nombre')?.metadata?.memberType).toBe('String');
    for (const member of members) {
      expect(r.edges).toContainEqual(
        expect.objectContaining({
          source: type?.id,
          target: member.id,
          kind: 'type-member',
        }),
      );
    }
  });
});

describe('VbaExtractor — roadmap #26 Declare statements', () => {
  it('models Win32 Declare statements as declare nodes and keeps calls traceable', () => {
    const src = [
      'Attribute VB_Name = "WinApi"',
      'Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
      '',
      'Public Sub UseApi()',
      '  GetTickCount',
      'End Sub',
    ].join('\n');

    const r = extract('src/modules/WinApi.bas', src);
    const declaration = r.nodes.find((n) => n.kind === 'declare' && n.name === 'GetTickCount');
    const caller = r.nodes.find((n) => n.kind === 'function' && n.name === 'UseApi');

    expect(declaration).toBeDefined();
    expect(declaration?.metadata).toEqual(
      expect.objectContaining({
        dll: 'kernel32',
        declareKind: 'function',
        ptrSafe: true,
      }),
    );
    expect(r.edges).toContainEqual(
      expect.objectContaining({
        source: caller?.id,
        target: declaration?.id,
        kind: 'calls',
      }),
    );
  });
});
