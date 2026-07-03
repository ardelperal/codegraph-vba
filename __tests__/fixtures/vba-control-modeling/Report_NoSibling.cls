VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_NoSibling"
Option Compare Database
Option Explicit

' =============================================================================
' Report_NoSibling — fixture for the negative case (issue #41 acceptance
' criterion 4): Report_*.cls body without a matching .report.txt sibling
' MUST NOT emit a synthesized event-handler edge to a non-existent
' .form.txt stub. The bug being fixed silently produced
' "Report_NoSibling.form.txt" stubs before.
' =============================================================================

Private Sub txtLonely_Click()
    Me.txtNonexistent.Value = "x"
End Sub