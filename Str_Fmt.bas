Attribute VB_Name = "Str_Fmt"
Option Compare Database
Private Sub FmtNm__Tst()
M$ = "Public"
T$ = "Property Get"
N$ = "FmtQQ"
S$ = "$"
P$ = "QQStr$, Paramarray Ap()"
A$ = ""
ActStr$ = FmtNm("{Modifier} {MthTy} {Nm}{Sfx}({Prm}){AsRetTy}", M, T, N, S, P, A)
ExpStr$ = "Public Property Get FmtQQ$(QQStr$, Paramarray Ap())"
Debug.Assert ActStr = ExpStr
End Sub
Property Get FmtNm$(NmStr$, ParamArray Ap())
Dim Av()
Av = Ap
O$ = NmStr
For J% = 0 To UB(Av)
    Dim A As LMR
    A = BrkMacro(O)
    O = A.L_Part & Av(J) & A.R_Part
Next
FmtNm = O
End Property

Property Get FmtQQAv$(QQStr$, Av())
O$ = QQStr
For J = 0 To UB(Av)
    With Brk(O, "?", NoTrim:=True)
        O = .S1 & Av(J) & .S2
    End With
Next
FmtQQAv = O
End Property
Property Get FmtQQ$(QQStr$, ParamArray Ap())
Dim Av()
Av = Ap
FmtQQ = FmtQQAv(QQStr, Av)
End Property
Private Sub Fmt__Tst()
Debug.Assert Fmt("aaa{0}bbbb{1}c", 1, 2) = "aaa1bbbb2c"
End Sub

Property Get Fmt$(FmtStr$, ParamArray Ap())
Dim Av()
Av = Ap
Fmt = FmtAv(FmtStr, Av)
End Property
Property Get FmtAv$(FmtStr$, Av())
U% = UB(Av)
If U = -1 Then FmtAv = FmtStr: Exit Property
O$ = FmtStr
For J% = 0 To U
    O = Replace(O, "{" & J & "}", Av(J))
Next
FmtAv = O
End Property
Private Sub FmtQQ__Tst()
Debug.Assert FmtQQ("aa?bb?cc", 1, 2) = "aa1bb2cc"
End Sub

Property Get JoinLine$(Ay)
JoinLine = Join(Ay, vbCrLf)
End Property

Private Sub JoinLine__Tst()
Dim A
Dim Ay$(2)
Ay(0) = "1"
Ay(1) = "2"
Ay(2) = "3"
A = Ay
Debug.Assert JoinLine(A) = "1" & vbCrLf & "2" & vbCrLf & "3"
End Sub


