Attribute VB_Name = "Ide_MdCpy"
Option Compare Database

Sub CpyMd(FmMd As Md, Pj As VBProject)
Nm$ = FmMd.Nm
If Not IsMd(FmMd) Then Er "Cannot copy!  Given {Md} not exist in {Pj}", FmMd.Nm, Pj.Name
If IsMd(Md(Pj, Nm)) Then Er "{Md} exist in {Pj}", Nm, Pj.Name
Tmp$ = TmpFt("CpyMd")
PjMd(FmMd.Pj, Nm).Parent.Export Tmp
Pj.VBComponents.Import Tmp
End Sub
