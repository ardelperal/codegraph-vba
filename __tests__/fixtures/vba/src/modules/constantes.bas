Attribute VB_Name = "constantes"
Option Compare Database
Option Explicit



Public Const msoFileDialogFilePicker As Long = 3
Public Const msoFileDialogFolderPicker As Long = 4
Public Const msoFileDialogOpen As Long = 1
Public Const msoFileDialogSaveAs As Long = 2

Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STATUS_PENDING = &H103&

Public Enum EnumTipoUsuario
    Administrador = 1
    Calidad = 2
    Secretaria = 3
    Economia = 4
End Enum
Public Enum EnumOrden
    PorFechaInicio = 1
    PorFechaFinPrevista = 2
    PorNAccion = 3
End Enum
Public Enum EnumSino
    Sí = 1
    No = 2
End Enum

Public Enum EnumEstadoNC
    BORRADA = 1
    REGISTRADA = 2 'tiene al menos una ac con al menos una ar(ésta sin fechas)
    PLANIFICADA = 3
    ENEJECUCION = 4
    ENEJECUCIONFUERADEPLAZO = 5
    ACSSINTAREAS = 6
    Cerrada = 7
    CERRADAPTECE = 8
    CERRADAPTECECADUCADA = 9
    CERRADACENOCONFORME = 10
End Enum

Public Enum EnumEstadoAC
    ACTIVA = 1 'AL MENOS ALGUNA ACCIÓN TIENE FECHA DE INICIO Y NO TODAS FECHA FIN
    SINACCIONES = 2
    FINALIZADA = 3  'TODAS LAS ACCIONES AGOTADAS (CON FECHA DE FIN)
    PTEREPLANIFICAR = 4
    PTEREREGULARIZAR = 5 'ALGUNA ACCIÓN ES IRREGULAR (Tiene fecha de fin sin incio y cosas de esas)
    REGISTRADA = 6
End Enum
Public Enum EnumEstadoAR
    ACTIVA = 1 'FECHA DE INICIO RELLENA Y FECHA DE FIN NO
    FINALIZADA = 2 'FECHA DE FIN RELLENA
    PTEREPLANIFICAR = 3 'FECHA PREVISTA ANTERIOR AL DÍA DE HOY
    IRREGULAR = 4 'FECHA FIN PREVISTA SIN FECHA INICIO O FECHA FIN SIN LAS OTRAS
    REGISTRADA = 5 'cuando no hay ninguna fecha
End Enum

Public Enum EnumTipoTareaNCProyecto
    PORCOMPLETARDATOSOBLIGATORIOS = 1
    PORREGULARIZARFECHASACCIONES = 2
    PORREGISTRARACCIONES = 3
    PORREPLANIFICAR = 4
    PORCERRAR = 5
    PTECONTROLEFICACIA = 6
    APuntoDeCaducar = 7
    Caducada = 8
   
End Enum
Public Enum EnumTipoAnexoAuditoria
    Auditoria = 1
    nc = 2
    AR = 3
End Enum
