Attribute VB_Name = "Vb_AyOpr"
Option Compare Database
Private Sub AddAy__Tst()
Act = AddAy(Array(1, 2), Array(3, 4))
Debug.Assert Sz(Act) = 4
Debug.Assert Act(0) = 1
Debug.Assert Act(1) = 2
Debug.Assert Act(2) = 3
Debug.Assert Act(3) = 4
End Sub
Private Sub ExpandToAy__Tst()
DmpAy ExpandToAy("AP[M B MINV MPAY]XX")
End Sub
Property Get ExpandToAy(S$) As String()
P1% = InStr(S, "[")
P2% = InStr(S, "]")
Pfx$ = Left(S, P1 - 1)
Sfx$ = Mid(S, P2 + 1)
Dim Ay$()
Ay = SplitLvs(Mid(S, P1 + 1, P2 - P1 - 1))
If Pfx <> "" Then Ay = AddAyPfx(Ay, Pfx)
If Sfx <> "" Then Ay = AddAySfx(Ay, Sfx)
ExpandToAy = Ay
End Property
Private Sub RmvEleAt__Tst()
Dim O$()
'=======================
For J% = 0 To 4
    Push O, J
Next
A = RmvEleAt(O)  '<====
Debug.Assert Sz(A) = 4
For J% = 0 To 3
    Debug.Assert A(J) = J + 1
Next
'=======================
Erase O
For J% = 0 To 4
    Push O, J
Next
A = RmvEleAt(O, 1) '<====
Debug.Assert Sz(A) = 4
Debug.Assert A(0) = 0
For J% = 1 To 3
    Debug.Assert A(J) = J + 1
Next
'=======================
Erase O
For J% = 0 To 4
    Push O, J
Next
A = RmvEleAt(O, 4) '<====
Debug.Assert Sz(A) = 4
For J% = 0 To 3
    Debug.Assert A(J) = J
Next
End Sub

Property Get RmvEleAt(Ay, Optional At&)
Dim O
O = Ay
U& = UB(Ay)
If U >= 0 Then ReDim Preserve O(U - 1)
For J = At + 1 To UB(Ay)
    O(J - 1) = Ay(J)
Next
RmvEleAt = O
End Property
Property Get RmvFirstEle(Ay)
Dim O
O = Ay
Erase O
U& = UB(Ay)
ReSzAy O, U - 1
For J& = 1 To U
    O(J - 1) = Ay(J)
Next
RmvFirstEle = O
End Property

Property Get RmvAyFirstChr(Ay) As String()
Dim O$()
U& = UB(Ay)
ReSzAy O, U
For J& = 0 To U
    O(J) = RmvFirstChr(Ay(J))
Next
RmvAyFirstChr = O
End Property

Property Get RmvLastEle(Ay)
Dim O
O = Ay
U& = UB(Ay)
If U > 0 Then
    ReDim Preserve O(U - 1)
Else
    Erase O
End If
RmvLastEle = O
End Property

Property Get AddAy(Ay, ParamArray AyAp())
O = Ay
Dim Av()
Av = AyAp
For J% = 0 To UB(Av)
    PushAy O, Av(J)
Next
AddAy = O
End Property

Private Sub TranposeAy__Tst()
Dim A()
ReDim A(1 To 2, 1 To 4)
For R% = 1 To 2
    For C% = 1 To 4
        A(R, C) = R + C * 100
    Next
Next
Act = TransposeAy(A)
Stop
End Sub
Property Get TransposeAy(Ay)
Dim O()
O = Ay
NR& = UBound(Ay, 1)
NC& = UBound(Ay, 2)
Erase O
ReDim O(1 To NC, 1 To NR)
For R& = 1 To NR
    For C& = 1 To NC
        O(C, R) = Ay(R, C)
    Next
Next
TransposeAy = O
End Property

Sub BrwAy(Ay, Optional Pfx$ = "BrwAy", Optional KeepTmpFt As Boolean = False)
F$ = TmpFt(Pfx)
WrtAy Ay, F
BrwFt F, KeepTmpFt
End Sub
Property Get InsAy(Ay, Itm, Optional At&)
O = Ay
Erase O
U& = UB(Ay)
ReSzAy O, U + 1
For J& = 0 To At - 1
    O(J) = Ay(J)
Next
O(At) = Itm
For J = At To UB(Ay)
    O(At + 1) = Ay(J)
Next
InsAy = O
End Property
Sub WrtAy(Ay, Ft$)
F% = FreeFile(1)
Open Ft For Output As F
For J = 0 To UB(Ay)
    Print #F, Ay(J)
Next
Close #F
End Sub
Private Sub AssertIsAy__Tst()
AssertIsAy 1
End Sub
Sub AssertIsAy(V)
If IsArray(V) Then Exit Sub
Er "Given {V} is not array", TypeName(V)
End Sub
Property Get JoinComma$(Ay)
JoinComma = Join(Ay, ",")
End Property
Private Sub CutAy__Tst()
ActAy = CutAy(Array(0, 1, 2, 3, 4), 2, 3)
ExpAy = Array(2, 3)
AssertEqAy ActAy, ExpAy
End Sub
Property Get CutAyByGpIdx(Ay, I As GpIdx)
CutAyByGpIdx = CutAy(Ay, I.BIdx, I.EIdx)
End Property

Property Get CutAy(Ay, BIdx, EIdx)
O = Ay
Erase O
For J& = BIdx To EIdx
    Push O, Ay(J)
Next
CutAy = O
End Property

Private Sub BrwAy__Tst()
Dim A$(2)
A(0) = "lskfj"
A(1) = "sdlfkj"
A(2) = "lksdjfsd"
BrwAy A
End Sub


Sub InsAyAt(Ay, At&, I)
N& = Sz(Ay)
ReDim Preserve Ay(N)
For J& = N - 1 To At Step -1
    Ay(J + 1) = Ay(J)
Next
Ay(At) = I
End Sub

Private Sub InsAyAt__Tst()
Ay = Array(1, 2, 3)
InsAyAt Ay, 1, "A"
Debug.Assert Sz(Ay) = 4
Debug.Assert Ay(0) = 1
Debug.Assert Ay(1) = "A"
Debug.Assert Ay(2) = 2
Debug.Assert Ay(3) = 3
End Sub

Sub AssertAy(Ay, Optional GivenXXX$ = "value", Optional Src$ = "AssertAy")
If Not IsArray(Ay) Then Er Src & ": given [" & GivenXXX & "] is not array, but {type}", TypeName(Ay)
End Sub


Sub PushAy(OAy, Ay)
For J = 0 To UB(Ay)
    Push OAy, Ay(J)
Next
End Sub

Sub PushAy_NoDup(OAy, Ay)
For J = 0 To UB(Ay)
    Push_NoDup OAy, Ay(J)
Next
End Sub

Sub DmpAy(Ay)
For J = 0 To UB(Ay)
    Debug.Print Ay(J)
Next
End Sub

Sub Push_NoDup(Ay, I)
If AyHas(Ay, I) Then Exit Sub
Push Ay, I
End Sub

Sub Push_NoDupNoBlank(Ay, I)
If Trim(I) = "" Then Exit Sub
Push_NoDup Ay, I
End Sub
Sub Push_NoBlank(OAy, I)
If IsBlank(I) Then Exit Sub
Push OAy, I
End Sub
Property Get QuoteAy(Ay, QStr$) As String()
U& = UB(Ay)
O = NewStrAy(U)
For J& = 0 To U
    O(J) = Quote(Ay(J), QStr)
Next
QuoteAy = O
End Property
Private Sub SetLastEle__Tst()
Dim A$()
ReDim A(12)
SetLastEle A, "A"
For J% = 0 To 11
    Debug.Assert A(J) = ""
Next
Debug.Assert A(12) = "A"
Debug.Assert UB(A) = 12

Dim Ay()
On Error GoTo X
SetLastEle Ay, 1
Stop
X:
End Sub
Sub SetLastEle(Ay, Itm)
N& = Sz(Ay)
If N = 0 Then Er FmtQQ("Ay has no last element.  TypeName(Ay)=[?]", TypeName(Ay))
If IsObject(Itm) Then
    Set Ay(N - 1) = Itm
Else
    Ay(N - 1) = Itm
End If
End Sub
Private Sub MinusAy__Tst()
Act = MinusAy(Array(1, 2, 3, 4, 5, 6, 8), Array(3, 4), Array(9, 8))
Debug.Assert Sz(Act) = 4
Debug.Assert Act(0) = 1
Debug.Assert Act(1) = 2
Debug.Assert Act(2) = 5
Debug.Assert Act(3) = 6
End Sub
Property Get MinusAy(Ay, ParamArray AyAp())
Dim Av()
Av = AyAp
If Sz(Av) = 0 Then MinusAy = Ay: Exit Function
O = Ay
If Sz(Av) = 1 Then
    Erase O
    Ay0 = Av(0)
    For J& = 0 To UB(Ay)
        If Not AyHas(Ay0, Ay(J)) Then Push O, Ay(J)
    Next
    MinusAy = O
    Exit Property
End If
For J = 0 To UB(Av)
    O = MinusAy(O, Av(J))
Next
MinusAy = O
End Property

