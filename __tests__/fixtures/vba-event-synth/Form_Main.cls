VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
Option Compare Database
Option Explicit

' =============================================================================
' Form_Main — subscribes to the Notifier's two events via the
' WithEvents naming convention. The presence of `m_Notifier_DataChanged`
' and `m_Notifier_ItemAdded` handler Subs is what makes the
' post-extraction event-handler synthesis pass materialize two
' `event-handler` edges from the raiser Subs in Notifier.cls to these
' two handlers.
' =============================================================================

Private WithEvents m_Notifier As Notifier

Private Sub m_Notifier_DataChanged(ByVal id As Long)
    ' react to data change
    Debug.Print "data changed: " & id
End Sub

Private Sub m_Notifier_ItemAdded(ByVal id As Long)
    ' react to item added
    Debug.Print "item added: " & id
End Sub
