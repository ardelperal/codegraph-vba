VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Lifecycle"
Option Compare Database
Option Explicit

' =============================================================================
' Report_Lifecycle — the report half of the issue #247 fixture. `Report_Open`
' is the report's own lifecycle event and must bind to the sibling
' report-layout node; `Report_Helper` is not an Access event and must bind to
' nothing.
' =============================================================================

Private Sub Report_Open(Cancel As Integer)
    Me.txtTotal.Value = 0
End Sub

Private Sub Report_Helper()
    Me.txtTotal.Value = 1
End Sub

Private Sub txtTotal_Click()
    Me.txtTotal.Value = 2
End Sub
