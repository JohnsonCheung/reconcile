Attribute VB_Name = "Vb_Functions"
Option Compare Database
Public Env As New Env
Property Get IsNothing(V) As Boolean
IsNothing = TypeName(V) = "Nothing"
End Property
Sub Er(ErMsgStr$, ParamArray Ap())
Dim Av()
Av = Ap
A$ = ErMsg_Av(ErMsgStr, Av)
BrwStr A
Err.Raise 1, , A
End Sub
Function ToBeCoded()
Stop
End Function
Property Get Min(A, ParamArray Ap())
Dim Av()
Av = Ap
O = A
For J = 0 To UB(Av)
    If Av(J) < O Then O = Av(J)
Next
Min = O
End Property

Property Get Max(A, ParamArray Ap())
Dim Av()
Av = Ap
O = A
For J = 0 To UB(Av)
    If Av(J) > O Then O = Av(J)
Next
Max = O
End Property

Private Sub Max__Tst()
Debug.Assert Max(1, 2, 4, 0) = 4
Debug.Assert Min(1, 2, 4, 0) = 0
End Sub

Property Get IsBlank(V) As Boolean
If IsStr(V) Then
    IsBlank = Trim(V) = ""
    Exit Property
End If
IsBlank = True
If IsEmpty(V) Then Exit Property
If IsMissing(V) Then Exit Property
If IsNull(V) Then Exit Property
IsBlank = False

End Property

Property Get IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Property
