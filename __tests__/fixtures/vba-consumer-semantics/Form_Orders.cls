VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Orders"
Option Compare Database
Option Explicit

' Ground truth for the consumer-semantics corpus:
'   btnSave  -> btnSave_Click  -> SaveOrderTotals (modOrders) -> qryOrderTotals
'   Form     -> Form_Load      -> RefreshOrders   (modOrders)
'   btnExport-> btnExport_Click-> dynamic dispatch, deliberately unresolved.

Private Sub btnSave_Click()
    SaveOrderTotals
End Sub

Private Sub btnExport_Click()
    Application.Run gExportProcedureName
    ExportOrdersToExcel
End Sub

Private Sub Form_Load()
    RefreshOrders
End Sub
