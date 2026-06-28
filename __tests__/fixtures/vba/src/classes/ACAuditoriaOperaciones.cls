VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACAuditoriaOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public AC As ACAuditoria



Private Function AccionRepetida( _
                                p_AC As String, _
                                Optional p_Error As String _
                                ) As EnumSino
        
    Dim m_ID As Variant
    Dim m_ACAuditoria As ACAuditoria
    
    On Error GoTo errores
    If p_AC = "" Then
        p_Error = "No se ha introducido una Acción"
        Err.Raise 1000
    End If
    With Me.AC
        If .nc.ACs Is Nothing Then
            AccionRepetida = EnumSino.No
            Exit Function
        End If
        For Each m_ID In .nc.ACs
            Set m_ACAuditoria = .nc.ACs(m_ID)
            If m_ACAuditoria.AccionCorrectiva = p_AC Then
                AccionRepetida = EnumSino.Sí
                Exit Function
            End If
            Set m_ACAuditoria = Nothing
        Next
    End With
    
    
    AccionRepetida = EnumSino.No
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ACAuditoriaOperaciones.AccionRepetida ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
   
    
End Function
Public Function MotivoNoOK( _
                                Optional ByRef p_ObjACAlInicio As ACAuditoria, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_AccionRepetida As EnumSino
    Dim blnComprobar As Boolean
    On Error GoTo errores
        
    With Me.AC
        If .id = "" Then
            MotivoNoOK = "No se conoce la NC de la Acción Correctiva"
            Exit Function
        End If
        
        If .AccionCorrectiva = "" Then
            MotivoNoOK = "No se conoce la Acción Correctiva"
            Exit Function
        End If
        If Not p_ObjACAlInicio Is Nothing Then
            If p_ObjACAlInicio.AccionCorrectiva <> .AccionCorrectiva Then
                blnComprobar = True
                
            End If
        Else
            blnComprobar = True
        End If
        If .Responsable = "" Then
            MotivoNoOK = "Se ha introducido un responsable"
            Exit Function
        End If
        If blnComprobar = True Then
            m_AccionRepetida = AccionRepetida(.AccionCorrectiva, p_Error)
            If p_Error <> "" Then
                Err.Raise 100
            End If
            If m_AccionRepetida = EnumSino.Sí Then
                p_Error = "Existe otra acción correctiva con el mismo nombre para la misma No Conformidad de Auditoría"
                Err.Raise 1000
            End If
        End If
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método ACAuditoriaOperaciones.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function

Public Function Registrar( _
                            Optional ByRef p_ObjACAlInicio As ACAuditoria, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_EstadoNCAlInicio As EnumEstadoNC
    
    On Error GoTo errores
    m_MotivoNoOK = MotivoNoOK(p_ObjACAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.AC.AccionCorrectiva = Replace(Me.AC.AccionCorrectiva, "'", "")
    Me.AC.AccionCorrectiva = Replace(Me.AC.AccionCorrectiva, ";", ":")
    With Me.AC
        If p_ObjACAlInicio Is Nothing Then
            m_SQL = "TbNCAuditoriaAccionCorrectivas"
            .IdAccionCorrectiva = .IDAccionCorrectivaCalculada
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If .NAccion = "" Then
            
            
                .NAccion = .nc.NAccionCalculado
            
                p_Error = .nc.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
        Else
            m_SQL = "SELECT * FROM TbNCAuditoriaAccionCorrectivas " & _
                "WHERE IDAccionCorrectiva=" & .IdAccionCorrectiva & ";"
        End If
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
    
        If Not p_ObjACAlInicio Is Nothing Then
            If rcdDatos.EOF Then
                p_Error = "No se ha encontrado el registro"
                Err.Raise 1000
            End If
            
        End If
        If p_ObjACAlInicio Is Nothing Then
            m_EstadoNCAlInicio = .nc.EstadoCalculado
            rcdDatos.AddNew
                rcdDatos.Fields("IDAccionCorrectiva") = .IdAccionCorrectiva
                rcdDatos.Fields("ID") = .id
                rcdDatos.Fields("AccionCorrectiva") = .AccionCorrectiva
                rcdDatos.Fields("NAccion") = .NAccion
                If .Notas <> "" Then
                    rcdDatos.Fields("Notas") = .Notas
                End If
                If .Responsable <> "" Then
                    rcdDatos.Fields("Responsable") = .Responsable
                End If
                If .FechaAccionCorrectiva <> "" Then
                    rcdDatos.Fields("FechaAccionCorrectiva") = .FechaAccionCorrectiva
                End If
                If .FechaFinalUltima <> "" Then
                    rcdDatos.Fields("FechaFinalUltima") = .FechaFinalUltima
                End If
                If .FechaFinPrevistaUltima <> "" Then
                    rcdDatos.Fields("FechaFinPrevistaUltima") = .FechaFinPrevistaUltima
                End If
                If .FechaInicialMinima <> "" Then
                    rcdDatos.Fields("FechaInicialMinima") = .FechaInicialMinima
                End If
                .Estado = .EstadoCalculadoTexto
                If .Estado <> "" Then
                    rcdDatos.Fields("Estado") = .Estado
                End If
            rcdDatos.Update
            RegistrarLogAuditoria "Alta de Acción Correctiva", , Me.AC, , p_Error
            If m_EstadoNCAlInicio = EnumEstadoNC.REGISTRADA Then
                PintarIndicadores p_Reiniciando:=EnumSino.Sí 'desaparecerá una tarea de NC Registrada y apararecerá otra de Acciones sin tareas
            End If
        Else
            rcdDatos.Edit
                
                rcdDatos.Fields("AccionCorrectiva") = .AccionCorrectiva
                If .Notas <> "" Then
                    rcdDatos.Fields("Notas") = .Notas
                Else
                    rcdDatos.Fields("Notas") = Null
                End If
                If .Responsable <> "" Then
                    rcdDatos.Fields("Responsable") = .Responsable
                Else
                    rcdDatos.Fields("Responsable") = Null
                End If
                .Estado = .EstadoCalculadoTexto
                If .Estado <> "" Then
                    rcdDatos.Fields("Estado") = .Estado
                Else
                    rcdDatos.Fields("Estado") = Null
                End If
            rcdDatos.Update
            RegistrarLogAuditoria "Edición de Acción Correctiva", , Me.AC, p_ObjACAlInicio, p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        
        End If
        rcdDatos.Close
        Set rcdDatos = Nothing
        ActualizarDatosCalculados p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set .nc.ACs = Nothing
        
        
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ACAuditoriaOperaciones.Registrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function RegistrarNAccion( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    
    On Error GoTo errores
   
    With Me.AC
        If IsNumeric(.NAccion) Then
            Exit Function
        End If
        .NAccion = .nc.NAccionCalculado
        p_Error = .nc.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        
        m_SQL = "SELECT * FROM TbNCAuditoriaAccionCorrectivas " & _
                "WHERE IDAccionCorrectiva=" & .IdAccionCorrectiva & ";"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
    
        
        rcdDatos.Edit
            rcdDatos.Fields("NAccion") = .NAccion
            'rcdDatos.Fields("AccionCorrectiva") = .AccionCorrectiva
            
            
        rcdDatos.Update
        rcdDatos.Close
        Set rcdDatos = Nothing
       
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ACAuditoriaOperaciones.RegistrarNAccion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Eliminar(Optional ByRef p_Error As String) As String
    
    Dim m_ARAuditoria As ARAuditoria
    Dim m_ARs As Scripting.Dictionary
    Dim m_IDAR As Variant
    Dim m_SQL As String
    Dim m_AROp As ARAuditoriaOperaciones
    Dim m_EstadoNC As EnumEstadoNC
    
    On Error GoTo errores
    
    With Me.AC
        Set m_ARs = .ARs
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ARs Is Nothing Then
            Set m_AROp = New ARAuditoriaOperaciones
            For Each m_IDAR In m_ARs
                Set m_ARAuditoria = m_ARs(m_IDAR)
                Set m_AROp.AR = m_ARAuditoria
                m_AROp.Eliminar p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                Set m_ARAuditoria = Nothing
            Next
        End If
        
        m_SQL = "DELETE * " & _
                "FROM TbNCAuditoriaAccionCorrectivas " & _
                "WHERE IDAccionCorrectiva=" & .IdAccionCorrectiva & ";"
        getdb().Execute m_SQL
        Set .nc.ACs = Nothing
        ActualizarDatosNCAuditoria .nc, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_EstadoNC = .nc.EstadoCalculado
        If m_EstadoNC = EnumEstadoNC.REGISTRADA Then
            PintarIndicadores p_Reiniciando:=EnumSino.Sí 'desaparecerá una tarea de NC Registrada y apararecerá otra de Acciones sin tareas
        End If
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ACAuditoriaOperaciones.Eliminar ha devuelto el error " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function



Public Function ActualizarDatosCalculados( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
        
    On Error GoTo errores
    
    With Me.AC
        .EstadoGrabar , p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .FechaInicialMinimaGrabar , p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .FechaFinPrevistaUltimaGrabar , p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .FechaFinalUltimaGrabar , p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ACAuditoriaOperaciones.ActualizarDatosCalculados ha devuelto el error: " & Err.Description
    End If
End Function



Public Function ReplanificarAcciones( _
                                    Optional p_Fecha As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_ARAuditoria As ARAuditoria
    Dim m_ARAlInicio As ARAuditoria
    Dim m_ID As Variant
    Dim m_AROp As ARAuditoriaOperaciones
    On Error GoTo errores
    
    If Not IsDate(p_Fecha) Then
        p_Error = "No se ha indicado la fecha"
        Err.Raise 1000
    End If
    If CDate(p_Fecha) < CDate(Date) Then
        p_Error = "La fecha no puede ser anterior a la de hoy"
        Err.Raise 1000
    End If
    With Me.AC
        If .ARs Is Nothing Then
            Exit Function
        End If
        Set m_AROp = New ARAuditoriaOperaciones
        For Each m_ID In .ARs
            Set m_ARAuditoria = .ARs(m_ID)
            Set m_ARAlInicio = .ARs(m_ID)
            If m_ARAuditoria.EstadoCalculado <> EnumEstadoAR.PTEREPLANIFICAR Then
                GoTo siguiente
            End If
            m_ARAuditoria.FechaFinPrevista = p_Fecha
            
            With m_AROp
                Set .AR = m_ARAuditoria
                
                .Registrar m_ARAlInicio, p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End With
siguiente:
            Set m_ARAuditoria = Nothing
        Next
    End With
    ActualizarDatosACAuditoria Me.AC, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ACAuditoriaOperaciones.ReplanificarAcciones ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function RegistrarResponsable( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    
    On Error GoTo errores
   
    With Me.AC
        If .nc.RESPONSABLEIMPLANTACION = "" Then
            Exit Function
        End If
        
        .Responsable = .nc.RESPONSABLEIMPLANTACION
        
        
        m_SQL = "SELECT * FROM TbNCAuditoriaAccionCorrectivas " & _
                "WHERE IDAccionCorrectiva=" & .IdAccionCorrectiva & ";"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
    
        
        rcdDatos.Edit
            rcdDatos.Fields("Responsable") = .Responsable
            
            
            
        rcdDatos.Update
        rcdDatos.Close
        Set rcdDatos = Nothing
       
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ACAuditoriaOperaciones.RegistrarResponsable ha devuelto el error: " & Err.Description
    End If
End Function



