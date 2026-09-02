VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormularioVentas"
Option Compare Database
Option Explicit

' =============================================================================
' The near-miss guard for issue #249: a VB_Name that starts with the letters
' `Form` but has no separating underscore. The trailing `_` is the whole
' discriminator of the Access code-behind convention, so this class is an
' ordinary one and must bind to nothing — even though `Detalle_Print` ends
' in a real Access event name.
' =============================================================================

Public Sub Detalle_Print()
    Me.txtImporte = 0
End Sub
