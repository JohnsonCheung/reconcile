Attribute VB_Name = "Ide_Res"
Option Compare Database

Private Sub ResStr__Tst()
Act = ResStrAy("AA", "Ide_Res")
Debug.Assert Sz(Act) = 2
Debug.Assert Act(0) = "asdf"
Debug.Assert Act(1) = "sdfdf"
End Sub

Property Get ResStrAy(ResNm$, Optional ResMdNm$ = "ResStrModule", Optional ByVal PjNm$) As String()
Dim Md As CodeModule
Set Md = PjMd(Pj(PjNm), ResMdNm)
Dim A As MthAtrOpt
A = MdMthAtr(Md, ResNm)
'DmpMthAtrOpt A
If A.Some Then
    ResStrAy = Pipe(A.MthAtr.Bdy, "SplitLines RmvBlankEle RmvFirstEle RmvLastEle RmvAyFirstChr")
End If
End Property

Property Get ResStr$(ResNm$, Optional ResMdNm$ = "ResStrModule", Optional ByVal PjNm$)
ResStr = JoinLine(ResStrAy(ResNm, ResMdNm, PjNm))
End Property
Private Sub Res_AA()
'asdf
'sdfdf
End Sub
