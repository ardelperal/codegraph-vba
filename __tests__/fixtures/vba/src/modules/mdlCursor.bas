Attribute VB_Name = "mdlCursor"
Option Compare Database
Option Explicit

Public Const IDC_SIZEWE = 32644&
Public Const IDC_ARROW = 32512&
#If VBA7 Then
    Public Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" _
            (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    
    Public Declare PtrSafe Function LoadCursorFromFile Lib "user32" Alias _
            "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
    
    Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#Else
    Public Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" _
            (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    
    Public Declare Function LoadCursorFromFile Lib "user32" Alias _
            "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
    
    Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If


Public Function MouseCursor(CursorType As Long)
    Dim lngRet As Long
    lngRet = LoadCursorBynum(0&, CursorType)
    lngRet = SetCursor(lngRet)
End Function

Public Function PointM(strPathToCursor As String)
    Dim lngRet As Long
    lngRet = LoadCursorFromFile(strPathToCursor)
    lngRet = SetCursor(lngRet)
End Function



