VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Sales"
Option Compare Database
Option Explicit

' Report equivalent of the Form_Orders path:
'   Detalle -> Detalle_Format -> FormatSalesRow (modSales)
'   Report  -> Report_Open    -> LoadSalesTotals (modSales) -> qrySalesTotals

Private Sub Report_Open(Cancel As Integer)
    LoadSalesTotals
End Sub

Private Sub Detalle_Format(Cancel As Integer, FormatCount As Integer)
    FormatSalesRow
End Sub
