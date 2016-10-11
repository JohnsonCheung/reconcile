Attribute VB_Name = "Vb_AyMap"
Option Compare Database

Property Get AyMap_PI(Ay, FctNm$, P1) As Variant()
U% = UB(Ay)
If U = -1 Then Exit Property
Dim O()
ReDim O(U)
For J& = 0 To U
    O(J) = Run(FctNm, P1, Ay(J))
Next
AyMap_PI = O
End Property
Private Sub AyMap_PI_StrAy__Tst()
Dim A$(2)
A(0) = "Fld1"
A(1) = "Fld2"
A(2) = "Fld3"
BrwAy AyMap_PI_StrAy(A, "Fmt", "[{0}]=Trim([{0}])")
End Sub
Property Get AyMap_PI_StrAy(Ay, FctNm$, P1) As String()
U% = UB(Ay)
If U = -1 Then Exit Property
Dim O$()
ReDim O(U)
For J& = 0 To U
    O(J) = Run(FctNm, P1, Ay(J))
Next
AyMap_PI_StrAy = O
End Property
