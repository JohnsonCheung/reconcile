Attribute VB_Name = "Ide_MdOpr"
Option Compare Database
Type Md
    Pj As VBProject
    Nm As String
End Type

Sub ShwMd(Md As CodeModule)
Md.CodePane.Show
End Sub
Sub ReplMdCxt(Md As CodeModule, Cxt$)
N% = Md.CountOfLines
If N >= 1 Then
    Md.DeleteLines 1, N
End If
Md.InsertLines 1, Cxt$
End Sub
Private Sub ClrMdBdy__Tst()
Dim Md As CodeModule
Set Md = NewTmpMd()
AppMdLines Md, Join(Array("Property Get A()", "'AA", "End Property"), vbCrLf)
ShwMd Md
RmvMd Md
RmvTmpMdInCurPj
End Sub
Sub RmvMd(Md As CodeModule)
MdPj(Md).VBComponents.Remove MdCmp(Md)
End Sub
Sub ClrMdBdy(Md As CodeModule)
BLno& = Md.CountOfDeclarationLines + 1
Cnt = Md.CountOfLines - BLno + 1
If Cnt >= 1 Then Md.DeleteLines BLno, Cnt
End Sub

Sub AppMdLines(Md As CodeModule, Lines$)
Md.InsertLines Md.CountOfLines + 1, Lines
End Sub

Private Sub ShwMd__Tst()
ShwMd CurPjMd("Ide_SrcLin")
End Sub
Property Get NewTmpCls() As CodeModule
Dim Cmp As VBComponent
Set Cmp = CurPj.VBComponents.Add(vbext_ct_ClassModule)
Cmp.Name = "TmpCls" & TimStmpNo
Set NewTmpCls = Cmp.CodeModule
End Property
Property Get NewTmpMd() As CodeModule
Dim Cmp As VBComponent
Set Cmp = CurPj.VBComponents.Add(vbext_ct_StdModule)
Cmp.Name = "TmpMd" & TimStmpNo
Set NewTmpMd = Cmp.CodeModule
End Property

Property Get IsMd(P As Md) As Boolean
Dim Cmp As VBComponent
For Each Cmp In P.Pj.VBComponents
    If Cmp.Name = P.Nm Then IsMd = True: Exit Property
Next
End Property

Property Get Md(Pj As VBProject, Nm$) As Md
Dim O As Md
Set O.Pj = Pj
O.Nm = Nm
Md = O
End Property
