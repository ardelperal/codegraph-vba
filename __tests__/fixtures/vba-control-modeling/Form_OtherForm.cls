VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_OtherForm"
Option Compare Database
Option Explicit

' =============================================================================
' Form_OtherForm — second synthetic fixture for control-modeling RED tests
' (hueco 5).
'
' Exists ONLY to make Form_Load disambiguation meaningful: when the user
' queries the index for "Form_Load" the engine must return TWO distinct hits
' (one per form), each qualified with its form prefix — never bare "Form_Load"
' matched against both forms without provenance.
' =============================================================================

Private Sub Form_Load()
    Me.lblTitulo.Caption = "Other"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.lblTitulo.Caption = vbNullString
End Sub