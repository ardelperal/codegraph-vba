import { describe, expect, it } from 'vitest';
import { VbaExtractor } from '../src/extraction/vba-extractor';
import { ACCESS_EVENT_NAMES } from '../src/extraction/vba/events';
import { parseEventHandlerName } from '../src/extraction/vba/text-utils';

describe('Access event-handler whitelist', () => {
  it('exports the shared Access event names used by code-behind handlers', () => {
    expect(ACCESS_EVENT_NAMES).toEqual(expect.objectContaining({
      size: expect.any(Number),
    }));
    expect([...ACCESS_EVENT_NAMES]).toEqual(expect.arrayContaining([
      'Click',
      'AfterUpdate',
      'BeforeDelConfirm',
      'AfterDelConfirm',
      'DropButtonClick',
      'Updated',
      'CommandBeforeExecute',
      'PivotTableChange',
    ]));
  });

  it('rejects helper names while matching events case-insensitively', () => {
    expect(parseEventHandlerName('Guardar_Datos')).toBeNull();
    expect(parseEventHandlerName('cmdOk_click')).toEqual({
      controlName: 'cmdOk',
      eventName: 'click',
    });
  });

  it('preserves last-underscore splitting and the Form-level exclusion', () => {
    expect(parseEventHandlerName('Btn_Guardar_Click')).toEqual({
      controlName: 'Btn_Guardar',
      eventName: 'Click',
    });
    expect(parseEventHandlerName('Form_Load')).toBeNull();
  });

  it('does not synthesize ghost controls or edges for underscore helpers', () => {
    const result = new VbaExtractor('Form_Prueba.cls', `Attribute VB_Name = "Form_Prueba"
Private Sub cmdOk_Click()
End Sub
Public Sub Guardar_Datos()
End Sub
Private Sub Btn_Guardar_click()
End Sub`).extract();

    const controls = result.nodes
      .filter((node) => node.kind === 'form-instance-control')
      .map((node) => node.name);
    const eventEdges = result.edges.filter((edge) => edge.kind === 'event-handler');

    expect(controls).toEqual(['cmdOk', 'Btn_Guardar']);
    expect(eventEdges).toHaveLength(2);
    expect(eventEdges.map((edge) => edge.metadata?.eventName)).toEqual(['Click', 'click']);
  });

  it('keeps every legitimate event handler exercised by the real fixture', () => {
    const source = `Attribute VB_Name = "Form_FormNCAuditoriaMotivoEliminado"
Private Sub Form_Load()
End Sub
Private Sub MotivoBorrado_AfterUpdate()
End Sub
Private Sub ComandoGrabar_Click()
End Sub`;
    const result = new VbaExtractor(
      'Form_FormNCAuditoriaMotivoEliminado.cls',
      source,
    ).extract();

    expect(result.edges.filter((edge) => edge.kind === 'event-handler'))
      .toHaveLength(2);
    expect(result.nodes.filter((node) => node.kind === 'form-instance-control'))
      .toHaveLength(2);
  });
});
