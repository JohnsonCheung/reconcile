Attribute VB_Name = "Ide_Pj"
Option Compare Database

Property Get NzPj(Pj As VBProject) As VBProject
If IsNothing(Pj) Then
    Set NzPj = Application.VBE.ActiveVBProject
Else
    Set NzPj = Pj
End If
End Property

