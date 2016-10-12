Attribute VB_Name = "EDI_BrkEDI"
Option Compare Database
Private Const ZZV_A_EDIPth = "C:\Users\cheungj\Desktop\reconciliation\EDI\"
Private Const ZZV_A_Fv = ZZV_A_EDIPth & "SPO_KERRY_0000019776_20161006002522.csv"

Sub AAAA()
ZZZ_Sq1
End Sub

Sub BrwEDIFv(Fv$)
BrwEDIBrk BrkEDIFv(Fv)
End Sub
Sub BrwEDIFld(EDIFv$)
Sq2 = BrkEDIFv(EDIFv)(1)
BrwAy CvSqRow(Sq2, 1)
End Sub
Sub BrwEDIBrk(EDIBrk())
Dim Av(): Av = EDIBrk
BrwSq Av(1), "Sq2"
For J% = 0 To UB(Av(0))
    BrwSq Av(0)(J), "Sq1-of-" & J & "-of-" & UB(Av(0))
Next
End Sub
Private Property Get B_EDILinesRecNm$(EDITy$)
'Each EDI has is own RecNm (HdrRecNm and DetRecNm) for `Lines`
'All `EDI-Lines` record will be put as sq in Ws[Lines]
'All `EDI-Header` record will be put as `multiple columns` in Ws[Header]
'By an EDINm, return XXXX so that XXXXH will be the `Lines`-Hdr-RecNm
'                         and     XXXXD will be the `Lines`-Det-RecNm
Select Case EDITy
Case "DE1": O = "DES"
Case "DE2": O = "P2S"
Case "SPO": O = "BOM"
Case "IVM": O = "IVM"
Case "IRP": O = "INV"
Case "LPD": O = "BOM"
Case "IMN": O = "IMN"
Case "PMU": O = "PMU"
Case "HANMOV": O = "HAN"
Case Else: Er "Invalid {EDITy}.  Valid type are [DE1 DE2 SPO IVM IRP LPD IMN PMU HANMOV", EDITy
End Select
B_EDILinesRecNm = O
End Property

Private Property Get B_EDITy$(Fv$)
A$ = FfnFn(Fv)
B_EDITy = Brk(A, "_").S1
End Property

Property Get BrkEDIFv(Fv$) As Variant()
BrkEDIFv = BrkEDI(FtAy(Fv), B_EDITy(Fv))
End Property

Private Property Get ZZV_EDILinesRecNm$()
ZZV_EDILinesRecNm = B_EDILinesRecNm(ZZV_EDITy)
End Property

Private Property Get ZZV_EDITy$()
ZZV_EDITy = B_EDITy(ZZV_A_Fv)
End Property

Private Sub ZZZ_BrkEDIFv()
BrwEDIBrk BrkEDIFv(ZZV_A_Fv)
End Sub

Property Get BrkEDI(InpAy$(), EDITy$) As Variant()
Dim LinTyAy$():           LinTyAy = B_LinTyAy(RmvBlankEle(InpAy))
Dim LinTyGp As Gp:        LinTyGp = Gp(LinTyAy)
Dim EDILinesRecNm$: EDILinesRecNm = B_EDILinesRecNm(EDITy)
Dim Av():    Av = B_Ay1_Ay2(InpAy, EDILinesRecNm)
Dim Ay1$(): Ay1 = Av(0)
Dim Ay2$(): Ay2 = Av(1)
Dim Sq1():  Sq1 = B_Sq1(Ay1)
Dim Sq2:    Sq2 = B_Sq2(Ay2)
BrkEDI = Array(Sq1, Sq2)
End Property

Private Property Get B_Sq1Mged(Ws1_Sq())
B_Sq1Mged = MgeSqAv(Ws1_Sq)
End Property

Private Property Get B_LinTyGp(LinTyAy$()) As Gp
B_LinTyGp = Gp(LinTyAy)
End Property

Private Property Get B_Ay1_Ay2(InpAy$(), EDILinesRecNm$) As Variant()
Dim O1$(), O2$()
R1$ = Brk(InpAy(0), ";").S1
R2$ = RmvLastChr(R1) & "T"
U% = UB(InpAy)
A1$ = EDILinesRecNm & "H"
A2$ = EDILinesRecNm & "D"
For J% = 0 To U
    L$ = InpAy(J)
    If IsPfx(L, A1) Then Push O2, L: GoTo Nxt
    If IsPfx(L, A2) Then Push O2, L: GoTo Nxt
    If IsPfx(L, R2) Then
        B_Ay1_Ay2__FixO2Ele0 O2
        B_Ay1_Ay2 = Array(O1, O2)
        Exit Property
    End If
    Push O1, L
Nxt:
Next
BrwAy InpAy, "InpAy"
Er "Impossible to reach here: In [InpAy], it should have {End-Rec-Type-Name} which is determined by {Beg-Rec-Type-Name}", R2, R1
End Property

Private Sub B_Ay1_Ay2__FixO2Ele0(O2$())
O2(0) = Join(TrmAy(ReplAyTab(Split(O2(0), ";"))), ";")
End Sub

Private Property Get B_LinTyAy(InpAy) As String()
Dim O$()
U& = UB(InpAy)
ReSzAy O, U
For J% = 0 To U
    If InpAy(J) = "HBOT" Then
        O(J) = "HBOT"
    Else
        O(J) = Brk(CStr(InpAy(J)), ";").S1
    End If
Next
B_LinTyAy = O
End Property


Private Sub ZZZ_EDILinesRecNm()
Debug.Assert ZZV_EDILinesRecNm = "BOM"
End Sub

Private Sub ZZZ_LinTyAy()
BrwAy ZZV_LinTyAy
End Sub
Private Sub ZZZ_InpAy()
BrwAy ZZV_InpAy
End Sub

Private Sub ZZZ_LinTyGp()
BrwGp ZZV_LinTyGp
End Sub

Private Sub ZZZ_Sq2()
BrwSq ZZV_Sq2
End Sub


Private Sub ZZZ_ZSq()
A1$ = "A;B;C;D;E"
A2$ = "1;2;3;4;5"
A3$ = "6;7;8;9;10"
BrwSq ZSq(StrAy(A1, A2, A3))
End Sub


Private Sub ZZZ_Sq1()
A = ZZV_Sq1
For J% = 0 To UB(A)
    BrwSq A(J), "Sq-" & J & "-of-" & UB(A)
Next
End Sub

Private Property Get ZZV_InpAy() As String()
ZZV_InpAy = RmvBlankEle(FtAy(ZZV_A_Fv))
End Property

Private Property Get ZZV_LinTyGp() As Gp
ZZV_LinTyGp = B_LinTyGp(ZZV_LinTyAy)
End Property

Private Property Get ZZV_Ay2() As String()
ZZV_Ay2 = B_Ay1_Ay2(ZZV_InpAy, ZZV_EDILinesRecNm)(1)
End Property

Private Property Get ZZV_Ay1() As String()
ZZV_Ay1 = B_Ay1_Ay2(ZZV_InpAy, ZZV_EDILinesRecNm)(0)
End Property

Private Property Get ZZV_LinTyAy() As String()
ZZV_LinTyAy = B_LinTyAy(ZZV_InpAy)
End Property

Private Property Get ZZV_Sq2()
ZZV_Sq2 = B_Sq2(ZZV_Ay2)
End Property

Private Property Get ZZV_Sq1() As Variant()
ZZV_Sq1 = B_Sq1(ZZV_Ay1)
End Property

Private Property Get B_Sq1(Ay1$()) As Variant()
Dim A$()
Dim O()
Push A, Ay1(0)
Pfx$ = RmvLastChr(Brk(A(0), ";").S1) & "D"
For J% = 1 To UB(Ay1)
    L$ = Ay1(J)
    If IsPfx(L, Pfx) Then
        Push A, L
    Else
        Push O, ZSq(A)
        Pfx$ = RmvLastChr(Brk(L, ";").S1) & "D"
        Erase A
        Push A, L
    End If
Next
Push O, ZSq(A)
B_Sq1 = O
End Property

Private Property Get B_Sq2(Ay2$())
NR% = Sz(Ay2)
NC% = Sz(Split(Ay2(0), ";"))
Dim O$()
    ReDim O(1 To NR, 1 To NC)
Dim B$()
For R% = 1 To NR
    B = Split(Ay2(R - 1), ";")
    For C% = 1 To Sz(B)
        O(R, C) = B(C - 1)
    Next
Next
B_Sq2 = O
End Property
Private Property Get ZSq(A$())
N% = Sz(Split(A(0), ";"))
Dim O$()
ReDim O(1 To N, 1 To Sz(A))
For J% = 1 To Sz(A)
    Dr = Split(A(J - 1), ";")
    For I% = 1 To Sz(Dr)
        O(I, J) = Dr(I - 1)
    Next
Next
ZSq = O
End Property




