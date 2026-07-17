VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_NoHandlers"
Option Compare Database
Option Explicit

' =============================================================================
' Form_NoHandlers — has a `WithEvents m_Notifier As Notifier` field but
' does NOT declare `m_Notifier_DataChanged` / `m_Notifier_ItemAdded`
' handler Subs. The post-extraction event-handler synthesis pass must
' NOT emit `event-handler` edges for this form's WithEvents binding.
' =============================================================================

Private WithEvents m_Notifier As Notifier

Private Sub SomeOtherSub()
    ' unrelated procedure
End Sub
