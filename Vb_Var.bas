Attribute VB_Name = "Vb_Var"
Option Compare Database
Private Sub ZZZ_IsAv()
Dim Av()
Debug.Assert IsAv(Av) = True
Dim V
Debug.Assert IsAv(V) = False
End Sub
Private Sub ZZZ_AssertAv()
AssertAv Av
End Sub
Sub AssertAv(Av)
If Not IsAv(Av) Then Er "Given {Av} is not Array-of-variant", TypeName(Av)
End Sub
Property Get IsAv(V)
IsAv = VarType(V) = vbArray + vbVariant
End Property
Property Get VarStr$(V)
On Error GoTo X
VarStr = V
Exit Property
X:
VarStr = "Conver Var to Str has err[" & Err.Description & "]"
End Property
Sub AssignAv(Av, O0, Optional O1, Optional O2, Optional O3, Optional O4, Optional O5)
AssertAv Av
Assign Av(0), O0
If IsMissing(O1) Then Exit Sub Else Assign Av(1), O1
If IsMissing(O2) Then Exit Sub Else Assign Av(2), O2
If IsMissing(O3) Then Exit Sub Else Assign Av(3), O3
If IsMissing(O4) Then Exit Sub Else Assign Av(4), O4
If IsMissing(O5) Then Exit Sub Else Assign Av(5), O5
End Sub
Sub Assign(V, O)
If IsObject(V) Then
    Set O = V
Else
    O = V
End If
End Sub
