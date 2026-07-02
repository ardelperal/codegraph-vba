Attribute VB_Name = "modCallerDemo"
Option Compare Database
Option Explicit

' Fixture for the vba-graph-connectivity-fixes change (issues #12, #13):
'  - `x` is declared as a project class (ACAuditoriaOperaciones, defined in
'    src/classes/ACAuditoriaOperaciones.cls) and called TWICE via the
'    no-parens statement form at two different call sites, both targeting the
'    same real method — exercises the resolver's duplicate-collapse path
'    (F1): both call sites must converge on exactly ONE `(source,target,
'    calls)` edge row after resolution.
'  - `mdlCursor.MouseCursor(...)` is a `.bas`-qualified PAREN-form call
'    against a real `Public Function` already declared in mdlCursor.bas —
'    exercises the resolver's module-scoped `.bas` fallback path.
Public Sub CallDemo()
    Dim x As ACAuditoriaOperaciones
    Dim p_Error As String

    x.Registrar p_Error
    x.Registrar p_Error

    Call mdlCursor.MouseCursor(1)
End Sub
