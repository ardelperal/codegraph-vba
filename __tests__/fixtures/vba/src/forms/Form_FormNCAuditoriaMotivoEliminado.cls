Option Compare Database
Option Explicit

Private m_Error As String



Private Sub Form_Load()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Public Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
   On Error GoTo errores
   
    If m_ObjNCAuditoriaActiva Is Nothing Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
        
    End If
    If m_ObjNCAuditoriaActiva.id = "" Then
        m_Error = "El formulario se ha abierto con parámetros insuficentes"
        Err.Raise 1000
    End If
    If Form_FormNCAuditoria.m_NCAlInicio Is Nothing Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    
    
    If Form_FormNCAuditoria.m_NCAlInicio.Borrado = True Then
        Me.ComandoGrabar.Enabled = True
    End If

    With Form_FormNCAuditoria.m_NCAlInicio
        
        If .MotivoBorrado <> "" Then
            Me.MotivoBorrado = .MotivoBorrado
            
        End If
        
    End With
    EstableceColorBordes
    Exit Function
    
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Sub EstableceColorBordes()
    
    On Error Resume Next
    
    Me.MotivoBorrado.BorderColor = COLOR_BORDE_CAMPO_RELLENO
    Me.MotivoBorrado.BorderWidth = ANCHO_BORDE_CAMPO_RELLENO
    
   
    With m_ObjNCAuditoriaActiva
        If .Borrado = True Then
            If Nz(Me.MotivoBorrado, "") = "" Then
                Me.MotivoBorrado.BorderColor = COLOR_BORDE_CAMPO_NORELLENO
                Me.MotivoBorrado.BorderWidth = ANCHO_BORDE_CAMPO_NORELLENO
            End If
            
        Else
            If Nz(Me.MotivoBorrado, "") <> "" Then
                Me.MotivoBorrado.BorderColor = COLOR_BORDE_CAMPO_NORELLENO
                Me.MotivoBorrado.BorderWidth = ANCHO_BORDE_CAMPO_NORELLENO
            End If
        End If
    End With
    
        
   
End Sub

Private Sub MotivoBorrado_AfterUpdate()
     On Error Resume Next
    
    EstableceColorBordes
End Sub
Private Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    On Error GoTo errores
    If Form_FormNCAuditoria.m_NCAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With Form_FormNCAuditoria.m_NCAlInicio
        
        If .MotivoBorrado <> Nz(Me.MotivoBorrado, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
       
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha devuelto el error: " & Err.Description
    End If
End Function
Public Function RellenarDatosObjeto(Optional ByRef p_Error As String) As String
    
    Dim ctl As Control
    Dim m_Valor As String
    
    On Error GoTo errores
   
    For Each ctl In Me.Controls
        If InStr(1, Nz(ctl.Tag, ""), "DATO") <> 0 Then
            m_ObjNCAuditoriaActiva.SetPropiedad ctl.Name, Nz(ctl.Value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
        End If
    Next
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarDatosObjeto ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Sub ComandoGrabar_Click()
    Dim m_NCOp As NCaUDITORIAOperaciones
    Dim m_HaHabidoCambios As Boolean
    
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If EsTecnico = EnumSino.Sí Then
        m_Error = "No tiene autorización para esa acción"
        Err.Raise 1000
    End If
    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_HaHabidoCambios Then
        m_Error = "No hay cambios que guardar"
        Err.Raise 1000
    End If
    RellenarDatosObjeto m_Error
    If m_Error <> "" Then
        Err.Raise 100
    End If
    Set m_NCOp = New NCaUDITORIAOperaciones
    With m_NCOp
        
        Set m_NCOp.nc = m_ObjNCAuditoriaActiva
        m_NCOp.ModificarMotivoBorrado m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    Form_FormNCAuditoria.m_NCAlInicio.MotivoBorrado = m_ObjNCAuditoriaActiva.MotivoBorrado
    EstablecerDatos
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub