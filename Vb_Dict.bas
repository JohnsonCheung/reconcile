Attribute VB_Name = "Vb_Dict"
Option Compare Database

Sub CvDict(P As Dictionary, OA$(), OB$())
Erase OA
Erase OB
If P.Count = 0 Then Exit Sub
For Each K In P.Keys
    Push OA, K
    Push OB, P(K)
Next
End Sub
