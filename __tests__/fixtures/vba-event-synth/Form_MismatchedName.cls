VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MismatchedName"
Option Compare Database
Option Explicit

' =============================================================================
' Form_MismatchedName — `WithEvents m_Other As Notifier` uses a variable
' name (`m_Other`) that does NOT match any handler Sub in this file.
' The expected handler Sub name is `m_Other_DataChanged`; the only
' handler this file declares is `m_Notifier_DataChanged` (deliberately
' named after a DIFFERENT variable to test the no-false-positive
' behaviour). The synthesis pass must NOT emit `event-handler` edges
' for this binding.
' =============================================================================

Private WithEvents m_Other As Notifier

Private Sub m_Notifier_DataChanged(ByVal id As Long)
    ' handler named after a DIFFERENT variable — does NOT match `m_Other_*`
    Debug.Print "won't fire: " & id
End Sub
