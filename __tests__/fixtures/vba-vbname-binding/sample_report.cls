VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Resumen"
Option Compare Database
Option Explicit

' =============================================================================
' Report_Resumen under a filename that lost the prefix — the report half of
' issue #249. The sibling extension the fallback picks must be `.report.txt`,
' not `.form.txt`.
' =============================================================================

Private Sub Report_Open(Cancel As Integer)
    Me.txtTotal.Visible = True
End Sub

Private Sub txtTotal_Click()
    Me.txtTotal.Visible = False
End Sub
