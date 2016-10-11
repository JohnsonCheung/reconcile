Attribute VB_Name = "VbFs_Ffn"
Option Compare Database
Private Sub FfnPth__Tst()
Debug.Assert FfnPth("C:\AAA.bat") = "C:\"
End Sub
Property Get FfnPth$(Ffn$)
P% = InStrRev(Ffn$, "\")
FfnPth = Left(Ffn, P)
End Property
Property Get ReplExt$(Ffn$, Ext$)
ReplExt = CutExt(Ffn) & Ext
End Property
Property Get CutExt$(Ffn$)
P% = InStrRev(Ffn, ".")
If P = 0 Then
    CutExt = Ffn
Else
    CutExt = Left(Ffn, P - 1)
End If
End Property
Private Sub FfnFn__Tst()
Debug.Print FfnFn("c:\aaa\aaa\bb.fx") = "bb.fx"
Debug.Print FfnFn("bb.fx") = "bb.fx"
End Sub
Property Get FfnFn$(Ffn$)
FfnFn = TakAftRev(Ffn, "\")
End Property
