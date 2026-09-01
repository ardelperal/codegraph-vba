VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InformeRiesgoPDFServicio"
Option Compare Database
Option Explicit

' =============================================================================
' InformeRiesgoPDFServicio — the regression guard for issue #249.
'
' A plain service class: neither its filename nor its VB_Name carries a
' `Form_` / `Report_` prefix. Its methods are full of underscores, and
' `Documento_Print` / `Cabecera_Format` end in real Access event names, so
' the ONLY thing standing between this class and a pile of invented control
' stubs is the prefix guard. Letting that guard go synthesized roughly 550
' spurious stubs across real projects, so widening it is the failure mode
' #249 had to avoid: this fixture must keep producing nothing.
' =============================================================================

Private Sub Class_Initialize()
    Me.Estado = "listo"
End Sub

Public Sub Documento_Print()
    Me.txtCuerpo = GenerarHTML_Principal()
End Sub

Public Sub Cabecera_Format()
    Me.txtTitulo = "Informe de riesgo"
End Sub

Public Function GenerarHTML_Principal() As String
    GenerarHTML_Principal = GetEstilosCSS_PDF()
End Function

Public Function GetEstilosCSS_PDF() As String
    GetEstilosCSS_PDF = "body { margin: 0 }"
End Function
