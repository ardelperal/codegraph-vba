VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Notifier"
Option Compare Database
Option Explicit

' =============================================================================
' Notifier — declares two custom events and raises them from its public
' methods. Forms that hold a `WithEvents m_Notifier As Notifier` field are
' expected to provide handler Subs named `m_Notifier_DataChanged` and
' `m_Notifier_ItemAdded` (the standard WithEvents naming convention).
' =============================================================================

Public Event DataChanged(ByVal id As Long)
Public Event ItemAdded(ByVal id As Long)

Public Sub TriggerDataChanged(ByVal id As Long)
    RaiseEvent DataChanged
End Sub

Public Sub TriggerItemAdded(ByVal id As Long)
    RaiseEvent ItemAdded
End Sub
