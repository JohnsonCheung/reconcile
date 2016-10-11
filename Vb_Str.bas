Attribute VB_Name = "Vb_Str"
Option Compare Database

Sub BrwStr(S$)
Ft$ = TmpFt
WrtStr S, Ft
BrwFt Ft, KillFt:=True
End Sub
Private Property Get TimStmp__Cnt&()
Static X&
X = X + 1
TimStmp__Cnt = X
End Property
Property Get RmvFirstChr$(S)
RmvFirstChr = Mid(CStr(S), 2)
End Property
Private Sub NamedAp__Tst()
DmpAy NamedAp("A b C ddd e", 1, 2, 3)
End Sub
Property Get NamedAp(NameLvs$, ParamArray Ap()) As String()
Dim Av()
Av = Ap
NamedAp = NamedAv(NameLvs, Av)
End Property
Property Get NamedAv(NameLvs$, Av()) As String()
Dim O$()
Dim A$()
A = SplitLvs(NameLvs)
U2& = UB(Av)
U1& = UB(A)
U& = Max(U1, U2)
ReSzAy O, U
L% = MaxAyLen(A) + 1
For J& = 0 To U
    If J <= U1 Then
        Nm$ = A(J) & Space(L - Len(A(J)))
    Else
        Nm = ""
    End If
    If J <= U2 Then
        V$ = VarStr(Av(J))
    Else
        V = ""
    End If
    O(J) = Nm & " = [" & V & "]"
Next
NamedAv = O
End Property


Property Get BrkQuoteStr(QStr$) As S1S2
If QStr = "" Then Exit Property
L% = Len(QStr)
If L = 1 Then BrkQuoteStr = NewS1S2(QStr, QStr): Exit Property
If L = 2 Then BrkQuoteStr = NewS1S2(Left(QStr, 1), Right(QStr, 1)): Exit Property
P% = InStr(QStr, "*")
If P = 0 Then Er "BrkQuoteStr: Invalid {QStr}", QStr
BrkQuoteStr = NewS1S2(Left(QStr, P - 1), Mid(QStr, P + 1))
End Property
Property Get NewS1S2(S1$, S2$) As S1S2
Dim O As S1S2
O.S1 = S1
O.S2 = S2
NewS1S2 = O
End Property
Property Get NzStr$(S$, S1$)
If IsEmptyStr(S$) Then
    NzStr = S1
Else
    NzStr = S
End If
End Property

Property Get Quote$(S, QStr$)
With BrkQuoteStr(QStr)
    Quote = .S1 & S & .S2
End With
End Property

Property Get LastChr$(S)
LastChr = Right(CStr(S), 1)
End Property

Property Get SplitNChrListStr(NChrListStr$, N%) As String()
A$ = NChrListStr
Dim O$()
For J = 0 To Len(A) / N - 1
    Push_NoDupNoBlank O, Trim(Mid(A, J * N + 1, N))
Next
SplitNChrListStr = O
End Property

Property Get RmvLastChr$(S)
A$ = S
RmvLastChr = Left(A, Len(A) - 1)
End Property

Property Get IsEmptyStr(S) As Boolean
If IsStr(S) Then
    IsEmptyStr = TrmStr(CStr(S)) = ""
ElseIf IsNull(S) Then
    IsEmptyStr = True
Else
    ThwNotStr S
End If
End Property

Property Get TrmStr$(S$)
TrmStr = Trim(S)
End Property

Sub ThwNotStr(S)
Er "Given {S} is not string, but {Type}", ToStr(S), TypeName(S)
End Sub

Sub AssertIsStr(S)
If Not IsStr(S) Then ThwNotStr S
End Sub
Property Get IsNum(V) As Boolean
Dim T As VbVarType
T = VarType(V)
Select Case T
Case vbByte, vbCurrency, vbSingle, vbDouble, vbInteger, vbLong
    IsNum = True
End Select
End Property
Property Get ToStr$(V)
If IsNum(V) Then
    ToStr = V
ElseIf IsDte(V) Then
    ToStr = Format(V, "YYYY-MM-DD HH:MM:SS")
Else
    ToStr = TypeName(V)
End If
End Property
Property Get IsDte(S) As Boolean
IsDte = VarType(S) = vbDate
End Property

Property Get IsStr(S) As Boolean
IsStr = VarType(S) = vbString
End Property
Private Sub TakBeg__Tst()
Debug.Assert TakBeg("sldf()", "(") = "sldf"
End Sub
Property Get TakBeg$(S$, BrkStr$)
TakBeg = Brk(S, BrkStr).S1
End Property
Property Get IsSfx(S$, Sfx$) As Boolean
IsSfx = Right(S, Len(Sfx)) = Sfx
End Property
Property Get IsPfx(S$, Pfx$) As Boolean
IsPfx = Left(S, Len(Pfx)) = Pfx
End Property
Property Get RmvPfx$(S$, Pfx$)
If IsPfx(S, Pfx) Then
    RmvPfx = Mid(S, Len(Pfx) + 1)
Else
    RmvPfx = S
End If
End Property
Private Sub TimStmpNo__Cnt__Tst()
Debug.Print TimStmpNo__Cnt%
Debug.Print TimStmpNo__Cnt%
Debug.Print TimStmpNo__Cnt%
End Sub
Private Property Get TimStmpNo__Cnt%()
Static X%
If X = 999 Then
    X = 1
Else
    X = X + 1
End If
TimStmpNo__Cnt = X
End Property
Private Sub TakBef__Tst()
Debug.Assert TakBef("aaaa.bbb", ".") = "aaaa"
Debug.Assert TakBef("aaaa.bbb", ".", True) = "aaaa."
Debug.Assert TakBef("aaaa...bbb", "...") = "aaaa"
Debug.Assert TakBef("aaaa...bbb", "...", True) = "aaaa..."
Debug.Assert TakBef("aaaa...bbb", "/") = ""
Debug.Assert TakBef("aaaa...bbb", "/", True) = ""
End Sub
Property Get TakBef$(S$, SubStr$, Optional InclSubStr As Boolean)
P& = InStr(S, SubStr)
If P = 0 Then Exit Property
If InclSubStr Then
    TakBef = Left(S, P + Len(SubStr) - 1)
Else
    TakBef = Left(S, P - 1)
End If
End Property
Property Get TakAft$(S$, SubStr$, Optional InclSubStr As Boolean)
P& = InStr(S, SubStr)
If P = 0 Then Exit Property
If InclSubStr Then
    TakAft = Mid(S, P)
Else
    TakAft = Mid(S, P + Len(SubStr))
End If
End Property
Property Get TakAftRev$(S$, SubStr$, Optional InclSubStr As Boolean)
P& = InStrRev(S, SubStr)
If P = 0 Then Exit Property
If InclSubStr Then
    TakAftRev = Mid(S, P)
Else
    TakAftRev = Mid(S, P + Len(SubStr))
End If
End Property
Property Get TakBefRev$(S$, SubStr$, Optional InclSubStr As Boolean)
P& = InStrRev(S, SubStr)
If P = 0 Then Exit Property
If InclSubStr Then
    TakBefRev = Left(S, P + Len(SubStr) - 1)
Else
    TakBefRev = Left(S, P - 1)
End If
End Property
Property Get SplitLvm(Lvm$) As String()
SplitLvm = SplitLv(Lvm, ";")
End Property
Property Get SplitLvvbar(Lvvbar$) As String()
SplitLvvbar = SplitLv(Lvvbar, "|")
End Property

Property Get SplitLv(Lv$, Sep$) As String()
SplitLv = TrmAy(Split(Lv, Sep))
End Property

Property Get TimStmpNo#()
A$ = Format(Now(), "YYMMDDHHMMSS")
T# = CDbl(A)
B# = T * 1000
C# = TimStmpNo__Cnt
D# = B + C
TimStmpNo = D
End Property
Property Get TimStmp$()
TimStmp = Format(Now(), "YYYY-MM-DD(HHMMSS)") & TimStmp__Cnt
End Property
Property Get SplitLines(Lines$) As String()
SplitLines = Split(Lines, vbCrLf)
End Property
Sub WrtStr(Str$, Ft$)
F% = FreeFile(1)
Open Ft For Output As F%
Print #F%, Str
Close #F
End Sub

Property Get SplitLvs(Lvs$) As String()
SplitLvs = Split(TrmDblSpc(Lvs), " ")
End Property

Property Get TrmDblSpc$(S$)
O = Trim(S)
While InStr(O, "  ") > 0
    O = Replace(O, "  ", " ")
Wend
TrmDblSpc = O
End Property
