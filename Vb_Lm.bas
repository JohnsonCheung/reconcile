Attribute VB_Name = "Vb_Lm"
Option Compare Database
Property Get CvLm(Lm$) As Dictionary
Dim A$(), O As New Dictionary
A = SplitLvm(Lm)
For Each I In A
    With Brk(CStr(I), "=")
        O.Add .S1, .S2
    End With
Next
Set CvLm = O
End Property
Property Get DictLm$(Dict As Dictionary, Optional Sep$ = ";", Optional Brk$ = "=")
Dim O$()
For Each K In Dict.Keys
    ToBeCoded
Next
End Property
Private Sub BrkLm__Tst()
Dim A$(), B$()
BrkLm "a=b;c=1", A, B
Debug.Assert Sz(A) = 2
Debug.Assert Sz(B) = 2
Debug.Assert A(0) = "a"
Debug.Assert A(1) = "c"
Debug.Assert B(0) = "b"
Debug.Assert B(1) = "1"
End Sub
Sub BrkLm(Lm$, OA$(), OB$())
CvDict CvLm(Lm), OA, OB
End Sub
