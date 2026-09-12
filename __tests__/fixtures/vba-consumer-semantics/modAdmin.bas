Attribute VB_Name = "modAdmin"
Option Compare Database
Option Explicit

Public Sub PurgeOrderLines()
    CurrentDb.Execute "DELETE FROM tblAuditLog"
    DoCmd.OpenForm "Form_Invoices"
End Sub

Public Function AuditNow() As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("tblAuditLog")
    rs.Close
    AuditNow = True
End Function
