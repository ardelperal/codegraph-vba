Attribute VB_Name = "modOrders"
Option Compare Database
Option Explicit

Public gExportProcedureName As String

Public Sub SaveOrderTotals()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("qryOrderTotals")
    rs.Close
End Sub

Public Sub RefreshOrders()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("tblOrderHeaders")
    rs.Close
End Sub
