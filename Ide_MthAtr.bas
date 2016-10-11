Attribute VB_Name = "Ide_MthAtr"
Option Compare Database
Private Sub MdMthAtrAy__Tst()
Dim A() As MthAtr
A = MdMthAtrAy(CurMd, InclBdy:=True)
Stop
End Sub

Property Get MdMthAtrAy(Md As CodeModule, Optional InclBdy As Boolean) As MthAtr()
Dim Ay$(), M As MthAtrOpt, O() As MthAtr
Ay = MdBdyAy(Md)
For J% = 0 To UB(Ay)
    M = SrcLinMAO(Ay(J))
    If M.Some Then
        If InclBdy Then
            A$ = ZMthBdy(Ay, J, M.MthAtr.MthTy)
            M.MthAtr.Bdy = A
        End If
        PushMthAtr O, M.MthAtr
    End If
Next
MdMthAtrAy = O
End Property

Property Get MthAtrStr$(P As MthAtr)
M$ = P.Modifier
T$ = P.MthTy
N$ = P.Nm
S$ = P.MthNmSfx
Prm$ = ""
A$ = P.RetTy
MthAtrStr = FmtNm("{Modifier} {MthTy} {Nm}{Sfx}({Prm}){AsRetTy}", M, T, N, S, Prm, A)
End Property

Private Property Get ZMthBdy$(BdyAy$(), BegIdx%, MthTy$)
BIdx% = ZMthBdy__BIdx(BdyAy, BegIdx)
EIdx% = ZMthBdy__EIdx(BdyAy, BegIdx, MthTy)
ZMthBdy = JoinLine(CutAy(BdyAy, BIdx, EIdx))
End Property

Private Sub ZMthBdy__BIdx__Tst()
Dim A$()
Push A, "aa"
Push A, ""
Push A, "'"
Push A, ""
Act% = ZMthBdy__BIdx(A, 3)
Debug.Assert Act = 2
End Sub

Private Property Get ZMthBdy__BIdx%(BdyAy$(), BegIdx%)
For J% = BegIdx - 1 To 0 Step -1
    L$ = Trim(BdyAy(J))
    If L = "" Then GoTo Nxt1
    If Left(L, 1) = "'" Then GoTo Nxt1
    For I = J + 1 To BegIdx
        L = Trim(BdyAy(I))
        If L <> "" Then
            ZMthBdy__BIdx = I
            Exit Property
        End If
    Next
    Er "Imposssibe to reach here"
Nxt1:
Next
ZMthBdy__BIdx = 0
End Property
Private Sub MdMthAtr__Tst()
Dim Act As MthAtrOpt
Act = MdMthAtr(CurPjMd("Ide_MdSrt"), "MdMthAtr__Tst")
Stop
End Sub
Private Sub AssertMthTy(MthTy$)
Select Case MthTy
Case "Sub", "Function", "Property Get", "Property Let", "Property Set"
Case Else
    Er "Given {MthTy} is invalid.  Invalid value = Function | Sub | Property XXX", MthTy
End Select
End Sub

Private Sub AssertPrpTy(PrpTy$)
Select Case MthTy
Case "Get", "Let", "Set"
Case Else
    Er "Given {PrpTy} is invalid.  Invalid value = Get | Let | Set", PrpTy
End Select
End Sub
Property Get MthAtrOptToAy(P As MthAtrOpt, Optional Nm$ = "MthAtrOpt") As String()
Dim O$()
Push O, "MthAtrOpt-Name = [" & Nm & "]"
If P.Some Then
    PushTabAy O, MthAtrToAy(P.MthAtr)
Else
    PushTab O, "None"
End If
MthAtrOptToAy = O
End Property
Property Get MthAtrToAy(P As MthAtr) As String()
With P
    MthAtrToAy = NamedAp("Modifier MthTy Nm MthNmSfx RetTy Bdy", .Modifier, .MthTy, .Nm, .MthNmSfx, .RetTy, .Bdy)
End With
End Property
Private Sub DmpMthAtrOpt__Tst()
BrwAy MthAtrOptToAy(MdMthAtr(CurPjMd("Dao_RunSql"), "SqlInt"), "XX-Name")
End Sub
Sub DmpMthAtrOpt(P As MthAtrOpt)
DmpAy MthAtrOptToAy(P)
End Sub
Property Get MdMthAtr(Md As CodeModule, MthNm$, Optional PrpTy$, Optional ExclBdy As Boolean) As MthAtrOpt
If PrpTy <> "" Then AssertPrpTy PrpTy
Dim Ay$()
Ay = MdBdyAy(Md)
Dim O As MthAtrOpt
For J% = 0 To UB(Ay)
    O = SrcLinMAO(Ay(J))
    If Not MdMthAtr__IsSel1(O, MthNm) Then GoTo NxtLin
    If Not MdMthAtr__IsSel2(O, PrpTy) Then GoTo NxtLin
    If Not ExclBdy Then
        O.MthAtr.Bdy = ZMthBdy(Ay, J, O.MthAtr.MthTy)
    End If
    MdMthAtr = O
    Exit Property
NxtLin:
Next
End Property

Private Property Get MdMthAtr__IsSel1(P As MthAtrOpt, MthNm$) As Boolean
If Not P.Some Then Exit Property
If P.MthAtr.Nm <> MthNm Then Exit Property
T$ = P.MthAtr.MthTy
AssertMthTy T
MdMthAtr__IsSel1 = True
End Property

Private Property Get MdMthAtr__IsSel2(P As MthAtrOpt, PrpTy$) As Boolean
If PrpTy = "" Then MdMthAtr__IsSel2 = True: Exit Property
T$ = P.MthAtr.MthTy
If Not IsPfx(T, "Property") Then Exit Property
MdMthAtr__IsSel2 = T = "Property " & PrpTy
End Property

Private Property Get ZMthBdy__EIdx%(BdyAy$(), BegIdx%, MthTy$)
Select Case MthTy
Case "Function", "Sub": XXX$ = MthTy
Case Else
    If Not IsPfx(MthTy, "Property") Then Er "Invalid {MthTy}.  Must be [Function] | [Sub] | [Property XXX]", MthTy
    XXX = "Property"
End Select
Pfx$ = "End " & XXX
For J% = BegIdx + 1 To UB(BdyAy)
    If IsPfx(BdyAy(J), Pfx) Then ZMthBdy__EIdx = J: Exit Property
Next
Er "No {End XXX} BdyAy scanning from {BegIdx}", Pfx, BegIdx
End Property

Sub PushMthAtr(Ay() As MthAtr, M As MthAtr)
N% = MthAtrSz(Ay)
ReDim Preserve Ay(N)
Ay(N) = M
End Sub

Property Get MthAtrUB%(Ay() As MthAtr)
MthAtrUB = MthAtrSz(Ay) - 1
End Property

Property Get MthAtrSz%(Ay() As MthAtr)
On Error Resume Next
MthAtrSz = UBound(Ay) + 1
End Property

