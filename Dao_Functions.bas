Attribute VB_Name = "Dao_Functions"
Option Compare Database

Function NzDb(D As Database) As Database
If IsNothing(D) Then
    Set NzDb = CurrentDb
Else
    Set NzDb = D
End If
End Function
Sub SetCurFbPth()
ChDir CurFbPth
End Sub
Property Get CurFbPth$()
CurFbPth = FfnPth(CurrentDb.Name)
End Property
Private Sub CurFbPth__Tst()
Debug.Print CurFbPth
End Sub
