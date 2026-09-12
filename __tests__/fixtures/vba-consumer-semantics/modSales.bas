Attribute VB_Name = "modSales"
Option Compare Database
Option Explicit

Public Sub LoadSalesTotals()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("qrySalesTotals")
    rs.Close
End Sub

Public Sub FormatSalesRow()
End Sub
