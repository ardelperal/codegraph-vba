VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ARAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDAccionRealizada As String
Public IdAccionCorrectiva As String
Public NAccion As String
Public AccionRealizada As String
Public FechaAccionRealizada As String
Public FechaInicio As String
Public FechaFinPrevista As String
Public FechaFinReal As String
Public Estado As String
Public Notas As String
Public Responsable As String 'nombre completo

Private m_sIDAccionRealizadaCalculada As String
Private m_ObjAC As ACAuditoria
Private m_ePendienteDeRetipificar As EnumSino
Private m_eESTADOCalculado As EnumEstadoAR
Private m_sESTADOCalculadoTexto As String
Private m_sFechaInicialMinimaCalculada As String
Private m_eEstadoEnum As EnumEstadoAR
Private m_objDocumentos As Scripting.Dictionary
Private m_ObjResponsableObj As usuario


Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get ResponsableObj() As usuario
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjResponsableObj Is Nothing Then
        Set ResponsableObj = m_ObjResponsableObj
        Exit Property
    End If
    If Me.Responsable = "" Then
        Exit Property
    End If
    Set m_ObjResponsableObj = constructor.getUsuario(p_UsuarioRed:=Me.Responsable, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ResponsableObj = m_ObjResponsableObj
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ARAuditoria.ResponsableObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Documentos() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objDocumentos Is Nothing Then
        Set Documentos = m_objDocumentos
        Exit Property
    End If
    Set m_objDocumentos = constructor.getDocumentosAuditoria(p_IDAR:=Me.IDAccionRealizada, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Documentos = m_objDocumentos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ARAuditoria.Documentos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Documentos(ByVal objNewValue As Variant)

    Set m_objDocumentos = objNewValue

End Property
Public Property Get EstadoCalculado() As EnumEstadoAR
    
    
       
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    With Me
        If IsDate(.FechaInicio) Then
            If Not IsDate(.FechaFinPrevista) Then
                m_eESTADOCalculado = EnumEstadoAR.IRREGULAR
                EstadoCalculado = m_eESTADOCalculado
                Exit Property
            End If
            
        End If
        If IsDate(FechaFinPrevista) Then
            If Not IsDate(.FechaInicio) Then
                m_eESTADOCalculado = EnumEstadoAR.IRREGULAR
                EstadoCalculado = m_eESTADOCalculado
                Exit Property
            End If
            
        End If
        If IsDate(.FechaFinReal) Then
            If Not IsDate(.FechaFinPrevista) Then
                m_eESTADOCalculado = EnumEstadoAR.IRREGULAR
                EstadoCalculado = m_eESTADOCalculado
                Exit Property
            End If
            If Not IsDate(.FechaInicio) Then
                m_eESTADOCalculado = EnumEstadoAR.IRREGULAR
                EstadoCalculado = m_eESTADOCalculado
                Exit Property
            End If
        End If
        If Not IsDate(.FechaInicio) And Not IsDate(.FechaFinPrevista) Then
            m_eESTADOCalculado = EnumEstadoAR.REGISTRADA
            EstadoCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If IsDate(.FechaFinReal) Then
            m_eESTADOCalculado = EnumEstadoAR.FINALIZADA
            EstadoCalculado = m_eESTADOCalculado
            Exit Property
        End If
        If IsDate(.FechaFinPrevista) Then
            If CDate(.FechaFinPrevista) < CDate(Date) Then
                m_eESTADOCalculado = EnumEstadoAR.PTEREPLANIFICAR
            Else
                m_eESTADOCalculado = EnumEstadoAR.ACTIVA
            End If
            EstadoCalculado = m_eESTADOCalculado
            Exit Property
        End If
        
        
        
    End With
    m_eESTADOCalculado = EnumEstadoAR.REGISTRADA
    EstadoCalculado = m_eESTADOCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ARAuditoria.EstadoCalculado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get EstadoCalculadoTexto() As String
    
    Dim m_EstadoCalculado As EnumEstadoAR
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_EstadoCalculado = Me.EstadoCalculado
    If m_ObjEntorno.ColEstadosAR.Exists(CStr(m_EstadoCalculado)) Then
        m_sESTADOCalculadoTexto = m_ObjEntorno.ColEstadosAR(CStr(m_EstadoCalculado))
    Else
        m_sESTADOCalculadoTexto = "DESCONOCIDO"
    End If
    
    
    EstadoCalculadoTexto = m_sESTADOCalculadoTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ARAuditoria.EstadoCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function EstadoGrabar( _
                                Optional p_Estado As String, _
                                Optional ByRef p_Error As String _
                                ) As String
                
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If Not IsDate(p_Estado) Then
        p_Estado = Me.EstadoCalculadoTexto
    End If
    Me.Estado = p_Estado
    If p_Estado = "" Then
        m_SQL = "UPDATE TbNCAuditoriaAccionesRealizadas set ESTADO=Null " & _
                "WHERE IDAccionRealizada=" & Me.IDAccionRealizada & ";"
    Else
        m_SQL = "UPDATE TbNCAuditoriaAccionesRealizadas " & _
                "set ESTADO='" & Me.Estado & "' " & _
                "WHERE IDAccionRealizada=" & Me.IDAccionRealizada & ";"
    End If
    getdb().Execute m_SQL
    EstadoGrabar = p_Estado
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ARAuditoria.EstadoGrabar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


Public Property Get EstadoEnum() As EnumEstadoAR
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.Estado = "" Then
        Exit Property
    End If
    
    If m_ObjEntorno.ColEstadosARTexto.Exists(Me.Estado) Then
        m_eEstadoEnum = m_ObjEntorno.ColEstadosARTexto(Me.Estado)
    Else
        m_eEstadoEnum = "DESCONOCIDO"
    End If
    
   
    EstadoEnum = m_eEstadoEnum
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ARAuditoria.EstadoEnum ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get PendienteDeRetipificar() As EnumSino
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If CDate(Me.FechaFinPrevista) < CDate(Date) Then
        m_ePendienteDeRetipificar = EnumSino.Sí
    Else
        m_ePendienteDeRetipificar = EnumSino.No
    End If
    
    PendienteDeRetipificar = m_ePendienteDeRetipificar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ARAuditoria.PendienteDeRetipificar  ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get AC() As ACAuditoria
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_ObjAC Is Nothing Then
        Set AC = m_ObjAC
        Exit Property
    End If
    If Me.IdAccionCorrectiva = "" Then
        Exit Property
    End If
    Set m_ObjAC = constructor.getACAuditoria(p_IDAC:=Me.IdAccionCorrectiva, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set AC = m_ObjAC
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ARAuditoria.AC ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Get IDAccionRealizadaCalculada() As String

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    m_sIDAccionRealizadaCalculada = constructor.getID("TbNCAuditoriaAccionesRealizadas", "IDAccionRealizada", getdb(), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDAccionRealizadaCalculada = m_sIDAccionRealizadaCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ARAuditoria.IDAccionRealizadaCalculada ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property



Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "IDAccionRealizada"
        .Add "IDAccionCorrectiva"
        .Add "NAccion"
        .Add "AccionRealizada"
        .Add "FechaAccionRealizada"
        .Add "FechaInicio"
        .Add "FechaFinPrevista"
        .Add "FechaFinReal"
        .Add "ESTADO"
        .Add "Notas"
        .Add "Responsable"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método ARAuditoria.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDAccionRealizada" Then
        getPropiedad = Me.IDAccionRealizada
        Exit Function
    End If
    If m_NombrePropiedad = "IDAccionCorrectiva" Then
        getPropiedad = Me.IdAccionCorrectiva
        Exit Function
    End If
    If m_NombrePropiedad = "NAccion" Then
        getPropiedad = Me.NAccion
        Exit Function
    End If
    If m_NombrePropiedad = "AccionRealizada" Then
        getPropiedad = Me.AccionRealizada
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAccionRealizada" Then
        getPropiedad = Me.FechaAccionRealizada
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicio" Then
        getPropiedad = Me.FechaInicio
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinPrevista" Then
        getPropiedad = Me.FechaFinPrevista
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinReal" Then
        getPropiedad = Me.FechaFinReal
        Exit Function
    End If
    If m_NombrePropiedad = "ESTADO" Then
        getPropiedad = Me.Estado
        Exit Function
    End If
    If m_NombrePropiedad = "Notas" Then
        getPropiedad = Me.Notas
        Exit Function
    End If
    If m_NombrePropiedad = "Responsable" Then
        getPropiedad = Me.Responsable
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en ARAuditoria.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ARAuditoria.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDAccionRealizada" Then
        Me.IDAccionRealizada = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDAccionCorrectiva" Then
        Me.IdAccionCorrectiva = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NAccion" Then
        Me.NAccion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AccionRealizada" Then
        Me.AccionRealizada = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAccionRealizada" Then
        Me.FechaAccionRealizada = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicio" Then
        Me.FechaInicio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinPrevista" Then
        Me.FechaFinPrevista = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinReal" Then
        Me.FechaFinReal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ESTADO" Then
        Me.Estado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Notas" Then
        Me.Notas = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Responsable" Then
        Me.Responsable = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en ARAuditoria.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ARAuditoria.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

