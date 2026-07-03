VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_PayrollSummary"
Option Compare Database
Option Explicit

' =============================================================================
' Report_PayrollSummary — synthetic fixture for the Report_*.cls code-behind
' binding test (issue #41).
'
' The .cls body has TWO event handlers whose targets exist in the sibling
' .report.txt (txtTotal_Click → txtTotal, txtCount_Click → txtCount) plus one
' whose target does NOT exist in the sibling (txtMissing_Click → txtMissing)
' so the negative case can be checked from the same source.
' =============================================================================

Private Sub txtTotal_Click()
    Me.txtCount.Value = "1"
End Sub

Private Sub txtCount_Click()
    Me.txtTotal.Value = "0"
End Sub

Private Sub txtMissing_Click()
    ' No control named txtMissing exists in the sibling .report.txt — this
    ' still produces an event-handler edge to a synthesized stub node.
    Me.txtTotal.Value = "x"
End Sub