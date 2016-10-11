Attribute VB_Name = "EDI_CvEDI"
Option Compare Database
Private C_InpAy$()
Private C_Ws0 As Worksheet

Private C_Ws1_At() As Range
Private C_Ws1_Sq()

Private C_Wb As Workbook
Private C_Wb_CusPrp As Dictionary

Private C_Ws2_At As Range
Private C_Ws2_Sq
Private C_Ws2_FmtSpecAy() As String
Const ZZV_A_Fv$ = "C:\Users\cheungj\Desktop\reconciliation\EDI\DE1_KERRY_0000019724_20160930180401.csv"

Private Sub ZZZ_CvEDIPth()
CvEDIPth FfnPth(ZZV_A_Fv)
End Sub

Sub CvEDIPth(EDIPth$, Optional KillCsv As Boolean)
Dim FnAy$()
Dim Xls As New Excel.Application
FnAy = PthFnAy(EDIPth, "*.csv")
Dim Fv$
Dim Selected As Boolean
For J% = 0 To UB(FnAy)
    Selected = False
    If IsPfx(FnAy(J), "DE1") Then Selected = True
    If IsPfx(FnAy(J), "HANMOV") Then Selected = True
    If Selected Then
        Debug.Print "CvEDIPth: ", J, UB(FnAy), FnAy(J)
        Fv = EDIPth & FnAy(J)
        CvEDI Fv, KillCsv, Xls
    End If
Next
Xls.Quit
Set Xls = Nothing
End Sub

Private Property Get B_Ws1_SqMged(Ws1_Sq())
B_Ws1_SqMged = MgeSqAv(Ws1_Sq)
End Property

Private Property Get B_Wb_CusPrp(Wb_CusPrpNmLvm$, Ws1_SqMged) As Dictionary
Dim A$()
    A = SplitLvm(Wb_CusPrpNmLvm)
Dim O As New Dictionary
For J% = 0 To UB(A)
    O.Add A(J), LookupTwoColSq(Ws1_SqMged, A(J))
Next
End Property

Private Property Get B_EDITy$(Fn$)
B_EDITy = Brk2(Fn, "_").S1
End Property

Private Property Get B_LinTyGp(LinTyAy$()) As Gp
B_LinTyGp = Gp(LinTyAy)
End Property


Private Property Get B_Ws1(Wb As Workbook) As Worksheet
Set B_Ws1 = Wb.Sheets("Header")
End Property

Private Property Get B_Ws1_At(Ws1 As Worksheet, Ws1_Sq()) As Range()
U% = UB(Ws1_Sq)
Dim O() As Range
ReDim O(U)
R1& = 1
If True Then    ' Multiple Columns
    For J% = 0 To U
        C% = 1 + J * 2
        Set O(J) = WsRC(Ws1, 1, C)
    Next
Else
    R& = 1
    For J = 0 To U
        Set O(J) = WsRC(Ws1, R, 1)
        R = R + UBound(Ws1_Sq(J), 1)
    Next
End If
B_Ws1_At = O
End Property

Private Property Get B_Ws2(Wb As Workbook) As Worksheet
Set B_Ws2 = Wb.Sheets("Lines")
End Property

Private Property Get B_Ws2_FmtSpecAy(EDITy$) As String()
B_Ws2_FmtSpecAy = ResStrAy(EDITy, "EDI_FmtSpec")
End Property

Private Sub Fmt_Wb()
'SetWbCusPrp C_Wb, C_Wb_CusPrp
End Sub

Private Sub ZZZ_Ws1_At()
Dim Act() As Range
Act = ZZV_Ws1_At
Stop
End Sub

Private Sub ZZZ_Fmt_Ws1()
Set C_Ws1 = ZZV_Ws1
C_Ws1_Sq = ZZV_Ws1_Sq
C_Ws1_At = ZZV_Ws1_At
Fmt_Ws1
Excel.Application.Visible = True
Stop
End Sub

Sub AA()
ZZZ_CvEDI
End Sub
Private Sub ZZZ_CvEDI()
CvEDI ZZV_A_Fv
End Sub

Private Property Get B_Wb(Xls As Excel.Application, Fx$) As Workbook
If IsOpnWb(Xls, Fx) Then
    Set B_Wb = Xls.Workbooks(Fx)
Else
    Set B_Wb = NewWb(Fx, "EDI;Header;Lines", Xls)
End If
End Property
Private Sub BBB(Fv$, Xls As Excel.Application)
Dim EDITy$:                    EDITy = B_EDITy(FfnFn(Fv))
Dim InpAy$():                  InpAy = RmvBlankEle(FtAy(Fv))

Dim Fx$:                          Fx = ReplExt(Fv, ".xlsx")
Dim Wb As Workbook:           Set Wb = B_Wb(Xls, Fx)
Dim Ws0 As Worksheet:        Set Ws0 = B_Ws0(Wb)
Dim Ws1 As Worksheet:        Set Ws1 = B_Ws1(Wb)
Dim Ws2 As Worksheet:        Set Ws2 = B_Ws2(Wb)

Dim Ws2_FmtSpecAy$():  Ws2_FmtSpecAy = B_Ws2_FmtSpecAy(EDITy)

Dim Av():         Av = BrkEDI(InpAy, EDITy)
Dim Ws1_Sq(): Ws1_Sq = Av(0)
Dim Ws2_Sq:   Ws2_Sq = Av(1)

Dim Ws1_At() As Range:        Ws1_At = B_Ws1_At(Ws1, Ws1_Sq)
Dim Ws2_At As Range:      Set Ws2_At = Ws2.Range("A3")

Dim Wb_CusPrpNmLvm$:        Wb_CusPrpNmLvm = B_Wb_CusPrpNmLvm(EDITy)
Dim Wb_CusPrp As Dictionary: Set Wb_CusPrp = B_Wb_CusPrp(Wb_CusPrpNmLvm, Ws1_Sq)


C_InpAy = InpAy
Set C_Ws0 = Ws0

C_Ws1_At = Ws1_At
C_Ws1_Sq = Ws1_Sq

Set C_Ws2_At = Ws2_At
C_Ws2_Sq = Ws2_Sq
C_Ws2_FmtSpecAy = Ws2_FmtSpecAy

Set C_Wb = Wb
Set C_Wb_CusPrp = Wb_CusPrp
End Sub

Private Property Get B_Ws0(Wb As Workbook) As Worksheet
Set B_Ws0 = Wb.Sheets("EDI")
End Property

Private Property Get B_Wb_CusPrpNmLvm$(EDITy$)
Select Case EDITy
Case "DE1": B_Wb_CusPrpNmLvm = A_Wb_CusPrpNmLvm_DE1
Case "SPO": B_Wb_CusPrpNmLvm = A_Wb_CusPrpNmLvm_SPO
Case "DE2": B_CusPrpNmLvm = A_Wb_CusPrpNmLvm_DE2
Case "IRP": B_CusPrpNmLvm = A_Wb_CusPrpNmLvm_IRM
Case Else: Stop
End Select
End Property

Sub CvEDI(Fv$, Optional KillCsv As Boolean, Optional Xls As Excel.Application)
Dim X As Excel.Application
Set X = NzXls(Xls)
BBB Fv, X
Fmt_Ws0
Fmt_Ws1
Fmt_Ws2
Fmt_Wb
X.Visible = True
C_Wb.Close True
If KillCsv Then Kill Fv
If IsNothing(Xls) Then X.Quit: Set X = Nothing
End Sub

Private Sub Fmt_Ws0()
PutSq WsA1(C_Ws0), AyOneColSq(C_InpAy)
End Sub

Private Sub Fmt_Ws1()
Dim At As Range
Dim Sq
For J% = 0 To UB(C_Ws1_At)
    Set At = C_Ws1_At(J)
    Sq = C_Ws1_Sq(J)
    Fmt_Ws1Sq At, Sq
Next
End Sub

Private Sub ZZZ_Fmt_Ws1Sq()
Dim Ws As Worksheet
Set Ws = NewWs
Dim At As Range
Set At = Ws.Range("A1")
Dim EDITy$
Dim Av(): Av = BrkEDI(ZZV_InpAy, ZZV_EDITy)
Dim Sq1(): Sq1 = Av(0)
Dim Sq2: Sq2 = Av(1)
Fmt_Ws1Sq At, Sq1(0)
At.Application.Visible = True
End Sub

Private Property Get ZZV_EDITy$()
ZZV_EDITy = B_EDITy(FfnFn(ZZV_A_Fv))
End Property

Private Property Get ZZV_InpAy() As String()
ZZV_InpAy = FtAy(ZZV_A_Fv)
End Property

Private Sub Fmt_Ws1Sq(At As Range, Sq)
PutSq At, Sq                    '<== PutSq
RgeRCRC(At, 1, 1, 1, 2).EntireColumn.AutoFit        '<== AutoFit Column

Dim R1 As Range
Dim R2 As Range
    Set R1 = RgeRCRC(At, 1, 1, 1, 2)
    Set R2 = RgeRCRC(At, 2, 2, UBound(Sq, 1), 2)
    R1.Interior.Color = 14395790     '<== Color
    R2.Interior.Color = 15917520     '<== Color
Dim R As Range
    Set R = ReSzRge(At, Sq)
SetBdr ReSzRge(At, Sq)          '<== Set Bdr
RgeRR(R, 2, R.Rows.Count).EntireRow.OutlineLevel = 2
RgeWs(R).OutLine.SummaryRow = xlSummaryAbove
End Sub

Private Function ZZV_Ws1() As Worksheet
Set ZZV_Ws1 = B_Ws1(ZZV_Wb)
End Function

Private Function ZZV_Wb() As Workbook
Set ZZV_Wb = NewWb(ZZV_Fx, "EDI;Header;Lines")
End Function

Private Function ZZV_Fx$()
ZZV_Fx = ReplExt(ZZV_A_Fv, ".xlsx")
End Function

Private Property Get ZZV_Ws1_Sq() As Variant()
ZZV_Ws1_Sq = BrkEDI(ZZV_InpAy, ZZV_EDITy)(0)
End Property

Private Property Get ZZV_Ws2_Sq() As Variant()
ZZV_Ws2_Sq = BrkEDI(ZZV_InpAy, ZZV_EDITy)(1)
End Property

Private Function ZZV_Ws1_At() As Range()
ZZV_Ws1_At = B_Ws1_At(ZZV_Ws1, ZZV_Ws1_Sq)
End Function

Private Sub Fmt_Ws2()
Dim At As Range: Set At = C_Ws2_At
Dim Sq: Sq = C_Ws2_Sq
Dim SPEC$(): SPEC = C_Ws2_FmtSpecAy
Dim ListObj As ListObject: Set ListObj = PutSq(At, Sq, CrtListObj:=True)
FmtLO ListObj, SPEC
End Sub

Private Sub ZZZ_InpAy()
BrwAy ZZV_InpAy
End Sub

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


