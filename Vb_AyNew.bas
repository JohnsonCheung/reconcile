Attribute VB_Name = "Vb_AyNew"
Option Compare Database
Property Get StrAy(ParamArray Ap()) As String()
Dim Av()
Av = Ap
StrAy = CvAyToStr(Av)
End Property
Property Get TrmAy(Ay) As String()
U& = UB(Ay)
If U = -1 Then Exit Property
ReDim O$(U)
For J& = 0 To UB(Ay)
    O(J) = Trim(Ay(J))
Next
TrmAy = O
End Property
Property Get NewVarAy(U)
If U = -1 Then Exit Property
If U < -1 Then Er "NewVarAy: {U} cannot be < -1", U
Dim O()
ReDim O(U)
NewVarAy = O
End Property
Property Get AddAyPfx(Ay, Pfx$) As String()
U& = UB(Ay)
If U = -1 Then Exit Property
ReDim O$(U)
For J& = 0 To UB(Ay)
    O(J) = Pfx & Ay(J)
Next
AddAyPfx = O
End Property
Property Get AddAySfx(Ay, Sfx$) As String()
U& = UB(Ay)
If U = -1 Then Exit Property
ReDim O$(U)
For J& = 0 To UB(Ay)
    O(J) = Ay(J) & Sfx
Next
AddAySfx = O
End Property

Property Get NewStrAy(U) As String()
If U = -1 Then Exit Property
If U < -1 Then Er "NewStrAy: {U} cannot be < -1", U
Dim O$()
ReDim O(U)
NewStrAy = O
End Property
Property Get NewIntAy(U) As Integer()
If U = -1 Then Exit Property
If U < -1 Then Er "NewIntAy: {U} cannot be < -1", U
Dim O%()
ReDim O(U)
NewIntAy = O
End Property

