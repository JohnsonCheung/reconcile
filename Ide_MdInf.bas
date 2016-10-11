Attribute VB_Name = "Ide_MdInf"
Option Compare Database

Property Get CurMd() As CodeModule
On Error Resume Next
Set CurMd = Application.VBE.ActiveCodePane.CodeModule
End Property


Property Get CurMdNm$()
CurMdNm = MdNm(CurMd)
End Property
Property Get CurMdInf() As MdInf
Set CurMdInf = MdInf(CurMd)
End Property
Private Sub CurMdInf__Tst()
Dim A As MdInf
Set A = CurMdInf
Stop
End Sub
Property Get MdInf(Md As CodeModule) As MdInf
Dim O As New MdInf
Set O.Md = Md
Set MdInfo = O
End Property


Property Get MdBdyAy(Md As CodeModule) As String()
MdBdyAy = SplitLines(MdBdy(Md))
End Property

Property Get MdBdy$(Md As CodeModule)
Cnt% = Md.CountOfLines - Md.CountOfDeclarationLines
If Cnt = 0 Then Exit Property
MdBdy = Md.Lines(Md.CountOfDeclarationLines + 1, Cnt)
End Property
Property Get MdSrcFn$(Md As CodeModule)
MdSrcFn = MdNm(Md) & MdExt(Md)
End Property
Property Get MdExt$(Md As CodeModule)
Select Case MdTy(Md)
Case vbext_ComponentType.vbext_ct_ClassModule: O = ".cls"
Case vbext_ComponentType.vbext_ct_StdModule: O = ".bas"
Case Else: Er "Unexpected {MdTy}.  Expected is [vbext_ct_ClassModule | vbext_ct_StdModule]", MdTy(Md)
End Select
MdExt = O
End Property
Property Get MdTy(Md As CodeModule) As vbext_ComponentType
MdTy = MdCmp(Md).Type
End Property

Property Get MdExpCxt$(Md As CodeModule)
Ft$ = TmpFt
Md.Parent.Export Ft
MdExpCxt = FtStr(Ft, KillFt:=True)
End Property
Property Get MdPj(Md As CodeModule) As VBProject
Set MdPj = Md.Parent.Collection.Parent
End Property
Property Get MdCxt$(Md As CodeModule)
N% = Md.CountOfLines
If N = 0 Then Exit Property
MdCxt = Md.Lines(1, Md.CountOfLines)
End Property
Property Get MdDcl$(Md As CodeModule)
N% = Md.CountOfDeclarationLines
If N = 0 Then Exit Property
MdDcl = Md.Lines(1, N)
End Property

Property Get MdCxtAy(Md As CodeModule) As String()
MdCxtAy = SplitLines(MdCxt(Md))
End Property
Private Sub MdDclAy__Tst()
BrwAy CurMdInfo.DclAy
End Sub

Property Get MdDclAy(Md As CodeModule) As String()
MdDclAy = SplitLines(MdDcl(Md))
End Property

Private Sub MdBdyAy__Tst()
BrwAy MdBdyAy(CurMd)
End Sub
Property Get MdCmp(Md As CodeModule) As VBComponent
Set MdCmp = Md.Parent
End Property

Property Get MdNm$(Md As CodeModule)
On Error Resume Next
MdNm = Md.Parent.Name
End Property

Private Sub MdNm__Tst()
Debug.Assert CurMdInf.Nm = "Ide_Md"
End Sub
