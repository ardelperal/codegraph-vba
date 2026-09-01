VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Expediente"
Option Compare Database
Option Explicit

' =============================================================================
' Form_Expediente — fixture for issue #249 (code-behind binding keys on the
' filename instead of on the module's own name).
'
' `sample2.cls` in this same folder is a BYTE-IDENTICAL copy of this file.
' The only difference between the two is the name on disk, so any difference
' in what the extractor produces for them is the bug under test.
'
' Contents, one of each shape the binding decision has to reach:
'   cmdGuardar_Click — a control handler; needs a form-instance-control target.
'   Form_Load        — a form-level lifecycle handler; needs the layout target.
'   Me.cboUsuario / Me.txtNombre — control references the Me sweep must emit.
' =============================================================================

Private Sub Form_Load()
    Me.cboUsuario.Requery
End Sub

Private Sub cmdGuardar_Click()
    Me.txtNombre.SetFocus
End Sub
