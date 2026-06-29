Attribute VB_Name = "modTestHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modTestHelper — standard module for control-modeling RED tests (hueco 6).
'
' Exercises DoCmd.OpenForm "FormName" modeling. Today the extractor absorbs
' DoCmd.* into the runtime-receiver blacklist (line 637 of vba-extractor.ts),
' so the string-literal target is silently discarded. The RED test asserts the
' graph SHOULD capture an opens-form edge with target "FormTest".
' =============================================================================

Public Sub TestOpenForm()
    DoCmd.OpenForm "FormTest"
End Sub

Public Sub TestOpenFormWithArgs()
    DoCmd.OpenForm "FormTest", acNormal, , , acFormEdit
End Sub