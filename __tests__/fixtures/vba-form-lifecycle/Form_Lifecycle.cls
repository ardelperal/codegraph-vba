VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Lifecycle"
Option Compare Database
Option Explicit

' =============================================================================
' Form_Lifecycle — fixture for issue #247 (form-level lifecycle handlers).
'
' Holds one of each shape the procedures sweep has to tell apart:
'   Form_Open / Form_Load / Form_Unload — form-level lifecycle events, which
'       must bind to the SIBLING form-layout node, never to a control.
'   Form_Helper                          — an ordinary method that merely
'       starts with `Form_`. `Helper` is not an Access event name, so it must
'       produce NO event-handler edge at all.
'   cmdSave_Click                        — a plain control handler. Its node
'       id and its edge must be byte-identical to what they were before #247.
' =============================================================================

Private Sub Form_Open(Cancel As Integer)
    Me.cmdSave.Enabled = True
End Sub

Private Sub Form_Load()
    Dim helper As Object
    Set helper = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.cmdSave.Enabled = False
End Sub

Private Sub Form_Helper()
    ' Not an Access event — `Helper` is not in ACCESS_EVENT_NAMES.
    Me.cmdSave.Caption = "Guardar"
End Sub

Private Sub cmdSave_Click()
    Me.cmdSave.Caption = "Guardado"
End Sub
