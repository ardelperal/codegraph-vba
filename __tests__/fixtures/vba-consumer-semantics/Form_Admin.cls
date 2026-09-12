VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Admin"
Option Compare Database
Option Explicit

' Effect coverage for the behavior-evidence contract:
'   btnPurge -> btnPurge_Click -> PurgeOrderLines (write + opens a form)
'   btnAudit -> wired by expression (=AuditNow()), not by naming convention.

Private Sub btnPurge_Click()
    Call PurgeOrderLines
End Sub
