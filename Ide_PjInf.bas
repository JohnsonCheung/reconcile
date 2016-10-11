Attribute VB_Name = "Ide_PjInf"
Option Compare Database

Property Get CurPjMd(Nm$) As CodeModule
Set CurPjMd = CurPj.VBComponents(Nm).CodeModule
End Property

Property Get PjMdAy(Pj As VBProject) As CodeModule()
Dim O() As CodeModule
Dim Cmp As VBComponent, T As vbext_ComponentType
For Each I In Pj.VBComponents
    Set Cmp = I
    T = Cmp.Type
    If T = vbext_ComponentType.vbext_ct_StdModule Or T = vbext_ComponentType.vbext_ct_ClassModule Then Push O, Cmp.CodeModule
Next
PjMdAy = O
End Property

Property Get Pj(Nm$) As VBProject
If Nm = "" Then
    Set Pj = CurPj
Else
    Set Pj = Application.VBE.VBProjects(Nm)
End If
End Property

Property Get PjPth$(Pj As VBProject)
PjPth = FfnPth(Pj.FileName)
End Property

Property Get PjFn$(Pj As VBProject)
PjFn = FfnFn(Pj.FileName)
End Property

Property Get PjSrcPth$(Pj As VBProject)
P$ = PjPth(Pj)
F$ = PjFn(Pj)
PjSrcPth = P & "Src\" & F & "\"
End Property

Property Get CurPj() As VBProject
Set CurPj = VBE.ActiveVBProject
End Property

Property Get PjMdByNm(PjNm$, MdNm$) As CodeModule
Set PjMdByNm = PjMd(Pj(PjNm), MdNm)
End Property

Property Get PjMd(Pj As VBProject, Nm$) As CodeModule
Set PjMd = Pj.VBComponents(Nm).CodeModule
End Property

Property Get PjTmpMdNmAy(Pj As VBProject) As String()
PjTmpMdNmAy = PjMdNmAy(Pj, "TmpMd*")
End Property

Property Get CurPjTmpMdNmAy() As String()
CurPjTmpMdNmAy = PjTmpMdNmAy(CurPj)
End Property

Private Sub CurPjTmpMdNmAy__Tst()
DmpAy CurPjTmpMdNmAy
End Sub

Property Get PjMdNmAy(Pj As VBProject, Optional Lik$ = "*") As String()
Dim Cmp As VBComponent, O$()
For Each Cmp In Pj.VBComponents
    Select Case Cmp.Type
    Case vbext_ct_ClassModule, vbext_ct_StdModule:
        N$ = Cmp.Name
        If N Like Lik Then Push O, Cmp.Name
    End Select
Next
PjMdNmAy = O
End Property
