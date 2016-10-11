Attribute VB_Name = "Ide_MthInf"
Option Compare Database
Type Mth
    Md As CodeModule
    Nm As String
End Type
Property Get MthBegLno%(Md As CodeModule, MthNm$)
Dim A$()
A = MdCxtAy(Md)
Dim O$()
For J% = 0 To UB(A)
    If SrcLinMthNm(A(J)) = MthNm Then MthBegLno = J: Exit Property
Next
MthBegLno = 0
End Property
Property Get MthEndLno%(Md As CodeModule, BegLno%)
Dim M As MthAtrOpt
M = SrcLinMAO(Md.Lines(BegLno, 1))
If Not M.Some Then Er "{Md} {BegLno} is not a Mth-line", MdNm(Md), BegLno
A$ = "End " & M.MthAtr.MthTy
For J% = BegLno + 1 To Md.CountOfLines
    If IsPfx(Md.Lines(J, 1), A) Then MthEndLno = J: Exit Property
Next
Er "No MthEndLno for {Md}, {BegLno}, {BegLin}", MdNm(Md), BegLno, Md.Lines(BegLno, 1)
End Property
Property Get MdMthNmAy(Md As CodeModule, Optional LikNm$ = "*") As String()
Dim A$()
A = MdBdyAy(Md)
Dim O$()
For J% = 0 To UB(A)
    Push_NoDupNoBlank O, SrcLinMthNm(A(J))
Next
MdMthNmAy = O
End Property
Property Get CurMdMthNmAy(Optional LikNm$ = "*") As String()
CurMdMthNmAy = MdMthNmAy(CurMd, LikNm)
End Property
Sub ShwMth(P As Mth)
ShwMd P.Md
Dim K As vbext_ProcKind
L% = P.Md.ProcBodyLine(P.Nm, K)
'VBE.ActiveCodePane.CodePaneView = vbext_CodePaneview.vbext_cv_ProcedureView
VBE.ActiveCodePane.TopLine = L
End Sub
Private Sub MdMthNmAy__Tst()
Dim A$()
A = MdMthNmAy(CurMd)
Stop
End Sub
Property Get CurMdMth(MthNm$) As Mth
CurMdMth = Mth(CurMd, MthNm)
End Property
Sub CrtCurTstMth()
CrtTstMth CurMth
End Sub
Property Get CurMthNm$()
Dim M As CodeModule, K As vbext_ProcKind
L% = VBE.ActiveCodePane.TopLine
Set M = CurMd
CurMthNm = M.ProcOfLine(L, K)
End Property
Property Get CurMth() As Mth
CurMth = Mth(CurMd, CurMthNm)
End Property
Property Get Mth(Md As CodeModule, Nm$) As Mth
Dim O As Mth
Set O.Md = Md
O.Nm = Nm
Mth = O
End Property
Private Sub CrtCurTstMth__Tst()
End Sub
Property Get IsMth(P As Mth) As Boolean
Dim A$()
A = MdBdyAy(P.Md)
For J% = 0 To UB(A)
    L$ = A(J)
    B$ = SrcLinMthNm(L)
    If B = P.Nm Then IsMth = True: Exit Property
Next
End Property
Private Sub IsMth__Tst()
Debug.Assert IsMth(CurMdMth("IsMth__Tst")) = True
End Sub



