VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Invoices"
Option Compare Database
Option Explicit

' Same control NAME as Form_Orders (btnSave) on purpose: consumers must keep
' the two layouts apart instead of answering with the first name match.

Private Sub btnSave_Click()
    SaveInvoiceTotals
End Sub
