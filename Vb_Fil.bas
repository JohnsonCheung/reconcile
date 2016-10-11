Attribute VB_Name = "Vb_Fil"
Option Compare Database

Sub DltFnAy(FnAy$(), Pth$, Optional IgnoreErr As Boolean)
For J% = 0 To UB(FnAy)
    DltFfnIfExist Pth & FnAy(J), IgnoreErr
Next
End Sub

Sub DltFfnIfExist(Ffn$, Optional IgnoreErr As Boolean)
If Fso.FileExists(Ffn) Then
    If IgnoreErr Then On Error Resume Next
    Kill Ffn
End If
End Sub

