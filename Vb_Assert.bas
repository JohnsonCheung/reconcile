Attribute VB_Name = "Vb_Assert"
Option Compare Database

Sub AssertEqStr(S1$, S2$)
ToBeCoded
End Sub
Property Get ChkEqStr(S1$, S2$) As String()

End Property
Sub AssertEqLinAy(Ay1$(), Ay2$())

End Sub
Property Get ChkEqLinAy(Ay1$(), Ay2$())

End Property

Sub AssertEqAy(Ay1, Ay2)
AssertChk = ChkAyEq(Ay1, Ay2)
End Sub
Property Get ChkAyEq(Ay1, Ay2) As String()
AssertAy Ay1, "Ay1", "ChkAyEq"
AssertAy Ay2, "Ay2", "ChkAyEq"
Dim O$()
If Sz(Ay1) <> Sz(Ay2) Then Push O, FmtQQ("Size diff: [?] / [?]", Sz(Ay1), Sz(Ay2)): GoTo E
For J& = 0 To UB(Ay1)
    If Ay1(J) <> Ay2(J) Then Push O, FmtQQ("Ele-? diff: [?] / [?]", J, Ay1(J), Ay2(J)): GoTo E
Next
Exit Property
E: ChkAyEq = O
End Property

