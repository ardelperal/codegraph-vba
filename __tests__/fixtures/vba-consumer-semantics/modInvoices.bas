Attribute VB_Name = "modInvoices"
Option Compare Database
Option Explicit

Public Sub SaveInvoiceTotals()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("qryInvoiceTotals")
    rs.Close
End Sub
