VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TestForm"
Option Compare Database
Option Explicit

' =============================================================================
' Form_TestForm — synthetic fixture for control-modeling RED tests (huecos 1/2/3).
'
' Exercises three Access forms patterns the current extractor does NOT capture:
'   1. Me.<ControlName>.<Property> reference resolution (hueco 1).
'   2. .form.txt sibling that exposes control NAMES, not just TYPES (hueco 2).
'   3. Cross-file event-handler edge from <ControlName> to <ControlName>_Click
'      (hueco 3).
'
' Every control name referenced from Me.* MUST also appear in the sibling
' Form_TestForm.form.txt. If it doesn't, the fixture is broken.
' =============================================================================

Private Sub ComandoAltaPM_Click()
    Me.lblTitulo.Caption = "Hello"
    Me.txtDescripcion.Value = "World"
End Sub

Private Sub ComandoBajaPM_Click()
    Me.lblDescripcion.Caption = "Bye"
End Sub

Private Sub Form_Load()
    Dim helper As Object
    Set helper = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.txtCodigo.Value = vbNullString
End Sub

Private Sub EstablecerDatos()
    Me.lblTitulo.Caption = "X"
    Me.txtDescripcion.Value = "Y"
End Sub