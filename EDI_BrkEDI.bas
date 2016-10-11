Attribute VB_Name = "EDI_BrkEDI"
Option Compare Database
Private Const ZZV_A_EDIPth = "C:\Users\cheungj\Desktop\reconciliation\EDI\"
Private Const ZZV_A_Fv = ZZV_A_EDIPth & "SPO_KERRY_0000019776_20161006002522.csv"
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

Private Property Get ZZV_EDILinesRecNm$()
ZZV_EDILinesRecNm = B_EDILinesRecNm(ZZV_EDITy)
End Property

Private Property Get ZZV_EDITy$()
A$ = FfnFn(ZZV_A_Fv)
ZZV_EDITy = Brk(A, "_").S1
End Property

Private Sub ZZZ_BrkEDI()
Dim Act(): Act = BrkEDI(ZZV_InpAy)
Dim ActSq1(): ActSq1 = Act(0)
Dim ActSq2: ActSq2 = Act(1)
BrwSq ActSq2, "Sq2"
For J% = 0 To UB(ActSq1)
    BrwSq ActSq1(J), "Sq1-of-" & J & "-of-" & UB(ActSq1)
Next
Stop
End Sub

Property Get BrkEDI(InpAy$(), EDINm$) As Variant()
Dim LinTyAy$():    LinTyAy = B_LinTyAy(RmvBlankEle(InpAy))
Dim LinTyGp As Gp: LinTyGp = Gp(LinTyAy)
Dim Ay1$():            Ay1 = B_Ay1(InpAy, LinTyGp)
Dim Ay2$():            Ay2 = B_Ay2(InpAy, LinTyGp)
Dim Sq1():             Sq1 = B_Sq1(Ay1)
Dim Sq2:               Sq2 = B_Sq2(Ay2)
BrkEDI = Array(Sq1, Sq2)
End Property
Private Property Get B_Sq1Mged(Ws1_Sq())
B_Sq1Mged = MgeSqAv(Ws1_Sq)
End Property
Private Property Get B_LinTyGp(LinTyAy$()) As Gp
B_LinTyGp = Gp(LinTyAy)
End Property

Private Property Get B_Ay2(InpAy$(), LinTyGp As Gp, EDILinesRecNm$) As String()
U& = GpUB(LinTyGp)
B& = GpIdx(LinTyGp, U - 2).BIdx
E& = GpIdx(LinTyGp, U - 1).EIdx
B_Ay2 = CutAy(InpAy, B, E)
End Property

Private Property Get B_Ay1(InpAy$(), LinTyGp As Gp, EDILinesRecNm$) As String()
U% = GpUB(LinTyGp)
E& = LinTyGp.EIdx(U - 3)
B_Ay1 = CutAy(InpAy, 0, E)
End Property

Private Property Get B_LinTyAy(InpAy) As String()
Dim O$()
U& = UB(InpAy)
ReSzAy O, U
For J% = 0 To U
    O(J) = Brk(CStr(InpAy(J)), ";").S1
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
BrwSq ZSq(A1, A2)
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
ZZV_Ay2 = B_Ay2(ZZV_InpAy, ZZV_LinTyGp)
End Property

Private Property Get ZZV_Ay1() As String()
ZZV_Ay1 = B_Ay1(ZZV_InpAy, ZZV_LinTyGp)
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
Dim O()
Push O, ZSq(Ay1(0), Ay1(1))
For J% = 2 To UB(Ay1) Step 2
    Push O, ZSq(Ay1(J), Ay1(J + 1))
Next
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
    For C% = 1 To NC
        O(R, C) = B(C - 1)
    Next
Next
B_Sq2 = O
End Property

Private Property Get ZSq(A$, B$)
A1 = Split(A, ";")
B1 = Split(B, ";")
N% = Sz(A1)
Dim O$()
ReDim O(1 To N, 1 To 2)
For J% = 1 To N
    O(J, 1) = A1(J - 1)
    O(J, 2) = B1(J - 1)
Next
ZSq = O
End Property




