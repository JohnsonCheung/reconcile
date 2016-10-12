Attribute VB_Name = "Xls_FmtLO"
Option Compare Database
Private A_LO As ListObject
Private B_SpecDict As Dictionary
Private B_FldNmAy$()
Private B_Alias_To_FldNm_Dict As Dictionary
Private B_Alias_To_Cno_Dict As Dictionary
Private C_Freeze_Cell As Range
Private C_Num_Rge() As Range ' All data in these range, needs to convert to number
Private C_Wdt%()
Private C_Wdt_ColAy()
Private C_Colr&()
Private C_Colr_ColAy()
Private C_Formula$()
Private C_Formula_ColAy()
Private C_Dte_Rge() As Range
Private C_NumFmt$()
Private C_NumFmt_ColAy()
Private C_Tot_Formula$()
Private C_Tot_Rge() As Range
Private C_YYYYMMDD_Rge() As Range
Private C_HAlign() As XlHAlign
Private C_HAlign_ColAy()
Private C_HdrHgt_Row As Range
Private C_HdrHgt_Factor%
Private C_OutLine%()
Private C_OutLine_ColAy()
Private C_WrapTxt_Rge() As Range
Private Const ZZV__Fv$ = "C:\Users\cheungj\Desktop\reconciliation\EDI\HANMOV_KERRY_109457_20160922124310_V3.csv" '
Private Const ZZV__SpecRes$ = "HANMOV"
'LPD_KERRY_0000019644_20160928233032.csv
'DE1_KERRY_0000019724_20160930180401.csv
'HANMOV_KERRY_109457_20160922124310_V3.csv
Private Property Get CC_Freeze() As Variant()
A$ = ZSpecRightVal("Freeze")
On Error Resume Next
Set O = LOWs(A_LO).Range(A)
CC_Freeze = Array(O)
End Property
Sub CvRgeFromYYYYMMDD(Rge As Range)
Sq = Rge.Value
If Rge.Count = 1 Then
    Rge.Value = CvYYYYMMDD(Sq)
    Exit Sub
End If
For J& = 1 To UBound(Sq, 1)
    For I& = 1 To UBound(Sq, 2)
        Sq(J, I) = CvYYYYMMDD(Sq(J, I))
    Next
Next
Rge.Value = Sq
Rge.NumberFormat = "YYYY/MM/DD;;#"
End Sub

Private Property Get BB_Alias_To_Cno_Dict(Alias_To_FldNm_Dict As Dictionary) As Dictionary
Dim D As Dictionary: Set D = Alias_To_FldNm_Dict
Dim L As ListObject: Set L = A_LO

Dim O As New Dictionary
U% = UB(F)
For Each A In D
    FldNm$ = D(A)
    I% = L.ListColumns(FldNm).Index
    O.Add A, I
Next
Set BB_Alias_To_Cno_Dict = O
End Property

Private Property Get CC_Num() As Variant()
A$ = ZSpecRightVal("Num")
CC_Num = Array(ZAliasLvs_Col(A))
End Property

Private Property Get CC_YYYYMMDD() As Variant()
A$ = ZSpecRightVal("YYYYMMDD")
CC_YYYYMMDD = Array(ZAliasLvs_Col(A))
End Property

Private Property Get CC_WrapTxt() As Variant()
A$ = ZSpecRightVal("WrapTxt")
CC_WrapTxt = Array(ZAliasLvs_Col(A))
End Property

Private Property Get CC_Colr() As Variant()
Dim O1&()
Dim O2()
Dim A$()
A = ZSpecItm("Colr")
U% = UB(A)
ReSzAy O1, U
ReSzAy O2, U
For J% = 0 To U
    With Brk(A(J), "|")
        O1(J) = CvKW_Colr(.S1)
        O2(J) = ZAliasLvs_Col(.S2)
    End With
Next
CC_Colr = Array(O1, O2)
End Property

Private Property Get ZAliasLvs_EntCol(AliasLvs$) As Range()
Dim A$(): A = SplitLvs(AliasLvs)
U% = UB(A)
Dim O() As Range: ReDim O(U)
Dim Rge As Range: Set Rge = A_LO.DataBodyRange
For J% = 0 To U
    If Not B_Alias_To_Cno_Dict.Exists(A(J)) Then Stop
    C% = B_Alias_To_Cno_Dict(A(J))
    Set O(J) = RgeC(Rge, C).EntireColumn
Next
ZAliasLvs_EntCol = O
End Property

Private Property Get ZAliasLvs_Col(AliasLvs$, Optional InclTot As Boolean) As Range()
Dim A$(): A = SplitLvs(AliasLvs)
U% = UB(A)
If U = -1 Then Exit Property
Dim O() As Range: ReDim O(U)
Dim Rge As Range: Set Rge = A_LO.DataBodyRange
For J% = 0 To U
    If Not B_Alias_To_Cno_Dict.Exists(A(J)) Then Stop
    C% = B_Alias_To_Cno_Dict(A(J))
    R1& = IIf(InclTot, -2, 1)
    R2& = Rge.Rows.Count
    Set O(J) = RgeCRR(Rge, C, R1, R2)
Next
ZAliasLvs_Col = O
End Property

Private Sub ZZZ_CC_Freeze()
ZZX_Set_AB_Var
AssignAv CC_Freeze, C_Freeze_Cell
Dim Act As Range
Set Act = C_Freeze_Cell
Stop
End Sub

Private Sub ZZV___LPD()

End Sub

Sub FmtLO(LO As ListObject, SpecAy$())
If Sz(SpecAy) = 0 Then
    Dim Ws As Worksheet
    Set Ws = LO.Parent
    Ws.Cells.EntireColumn.AutoFit
    FreezeAt Ws.Range("A2")
    Exit Sub
End If
Set A_LO = LO
B_FldNmAy = LOFldNmAy(A_LO)
Set B_SpecDict = BB_SpecDict(SpecAy)
Set B_Alias_To_FldNm_Dict = BB_Alias_To_FldNm_Dict(B_FldNmAy, B_SpecDict)
Set B_Alias_To_Cno_Dict = BB_Alias_To_Cno_Dict(B_Alias_To_FldNm_Dict)
BBB
SetColr
SetDte
SetFormula
SetHAlign
SetHdrHgt
SetNum      ' Convert to Num use val
SetNumFmt
SetFreeze
SetOutLine
SetTot
SetWdt
SetWrapTxt
SetYYYYMMDD
End Sub

Private Sub SetColr()
Dim A&(): A = C_Colr
Dim C(): C = C_Colr_ColAy
Dim Rge As Range
For J% = 0 To UB(A)
    Colr& = A(J)
    For Each I In C(J)
        Set Rge = I
        Rge.Interior.Color = Colr
    Next
Next

'With Rge
'    .Pattern = xlSolid
'    .PatternColorIndex = xlAutomatic
'    .ThemeColor = xlThemeColorAccent1
'    .TintAndShade = TintAndShape#
'    .PatternTintAndShade = 0
'End With
End Sub

Private Sub ZZZ_CC_Colr()
ZZX_Set_AB_Var
AssignAv CC_Colr, C_Colr, C_Colr_ListColAy
End Sub

Private Sub ZZZ_SetTot()
ZZZ_CC_Tot
SetTot
End Sub
Private Sub ZZZ_CC_Tot()
ZZX_Set_AB_Var
AssignAv CC_Tot, C_Tot_Formula, C_Tot_Rge
End Sub

Private Sub ZZZ_SetColr()
ZZZ_CC_Colr
SetColr
End Sub
Private Sub SetHdrHgt()
Dim H As Range: Set H = C_HdrHgt_Row
                   F% = C_HdrHgt_Factor
If IsNothing(H) Then Exit Sub
If F <= 0 Then Exit Sub
If F > 5 Then F = 5
H.EntireRow.RowHeight = H.RowHeight * F
H.HorizontalAlignment = XlHAlign.xlHAlignLeft
H.VerticalAlignment = XlVAlign.xlVAlignTop
H.WrapText = True
End Sub
Private Sub SetFreeze()
FreezeAt C_Freeze_Cell
End Sub

Private Sub FmtLO__Tst()
FmtLO ZZV_A_LO, ZZV__SpecAy
A_LO.Application.Visible = True
A_LO.Application.WindowState = xlMaximized
End Sub

Private Sub ZZZ_BB_Alias_To_FldNm_Dict()
ZZX_Set_A_Var
Dim Act As Dictionary
    Set Act = ZZV_B_Alias_To_FldNm_Dict
Stop
End Sub

Private Sub ZZZ_BB_FldNmAy()
ZZX_Set_A_Var
Dim Act$()
    Act = ZZV_B_FldNmAy
Stop
End Sub

Private Sub ZZZ_BB_Alias_To_Cno()
ZZX_Set_A_Var
Dim Act As Dictionary
    Set Act = ZZV_B_Alias_To_Cno_Dict
Stop
End Sub
Private Sub ZZZ_BB_SpecDict()
ZZX_Set_A_Var
Dim Act As Dictionary
    Set Act = ZZV_B_SpecDict
Stop
End Sub
Private Sub SetWdt()
Dim W%(): W = C_Wdt
Dim C():  C = C_Wdt_ColAy
Dim R As Range
LOWs(A_LO).Cells.EntireColumn.AutoFit
For J% = 0 To UB(W)
    Wdt% = W(J)
    For Each I In C(J)
        Set R = I
        R.ColumnWidth = Wdt
    Next
Next
End Sub

Private Property Get BB_Alias_To_FldNm_Dict(FldNmAy$(), SpecDict As Dictionary) As Dictionary
Dim A$(): If SpecDict.Exists("AliasL") Then A = SpecDict("AliasL")
Dim D As New Dictionary ' Field name to Alias
    For J% = 0 To UB(A)
        With Brk(A(J), "|")
            D.Add .S1, .S2
        End With
    Next
    
Dim F$(): F = FldNmAy
Dim O As New Dictionary
Dim Alias$
For J% = 0 To UB(F)
    FldNm$ = F(J)
    If D.Exists(FldNm) Then
        Alias = D(FldNm)
    Else
        Alias = FldNm
    End If
    O.Add Alias, FldNm
Next
Set BB_Alias_To_FldNm_Dict = O
End Property

Private Sub BBB()
mFreeze_ = CC_Freeze
mNum____ = CC_Num
mColr___ = CC_Colr
mDte____ = CC_Dte
mFormula = CC_Formula
mHAlign_ = CC_HAlign
mHdrHgt_ = CC_HdrHgt
mNumFmt_ = CC_NumFmt
mOutLine = CC_OutLine
mTot____ = CC_Tot
mWdt____ = CC_Wdt
mWrapTxt = CC_WrapTxt
mYYYYMMDD = CC_YYYYMMDD
AssignAv mFreeze_, C_Freeze_Cell
AssignAv mColr___, C_Colr, C_Colr_ColAy
AssignAv mDte____, C_Dte_Rge
AssignAv mFormula, C_Formula, C_Formula_ColAy
AssignAv mHAlign_, C_HAlign, C_HAlign_ColAy
AssignAv mHdrHgt_, C_HdrHgt_Factor, C_HdrHgt_Row
AssignAv mNumFmt_, C_NumFmt, C_NumFmt_ColAy
AssignAv mOutLine, C_OutLine, C_OutLine_ColAy
AssignAv mTot____, C_Tot_Formula, C_Tot_Rge
AssignAv mWdt____, C_Wdt, C_Wdt_ColAy
AssignAv mWrapTxt, C_WrapTxt
AssignAv mNum____, C_Num_Rge
AssignAv mYYYYMMDD, C_YYYYMMDD_Rge
End Sub


Private Property Get CC_Tot() As Variant()
Dim A$(): A = ZSpecItm("Tot")
U% = UB(A)
Dim O1$() ' Formula
Dim O2() As Range
Dim AliasAy$()
Dim Tot As XlTotalsCalculation
Dim HRow As Range
Set HRow = A_LO.HeaderRowRange
For J% = 0 To U
    With Brk(A(J), "|")
        AliasAy = SplitLvs(.S2)
        Tot = CvKW_TotCal(.S1)
    End With
    For I% = 0 To UB(AliasAy)
        FldNm$ = B_Alias_To_FldNm_Dict(AliasAy(I))
        Select Case Tot
        Case XlTotalsCalculation.xlTotalsCalculationAverage
            Formula1$ = FmtQQ("=Average(?[?])", A_LO.Name, FldNm)
            Formula2$ = FmtQQ("=Subtotal(101,?[?])", A_LO.Name, FldNm)
        Case XlTotalsCalculation.xlTotalsCalculationCount
            Formula1$ = FmtQQ("=CountA(?[?])", A_LO.Name, FldNm)
            Formula2$ = FmtQQ("=Subtotal(103,?[?])", A_LO.Name, FldNm)
        Case Else
            Formula1$ = FmtQQ("=Sum(?[?])", A_LO.Name, FldNm)
            Formula2$ = FmtQQ("=Subtotal(109,?[?])", A_LO.Name, FldNm)
        End Select
        R& = HRow.Row
        C% = B_Alias_To_Cno_Dict(AliasAy(I))
        Set Cell1 = RgeRC(HRow, -1, C)
        Set Cell2 = RgeRC(HRow, 0, C)
        Push O1, Formula1
        Push O2, Cell1
        Push O1, Formula2
        Push O2, Cell2
    Next
Next
CC_Tot = Array(O1, O2)
End Property

Private Property Get ZFnd_FldNmAy(AliasLvs$) As String()
Dim A$()
    A = SplitLvs(AliasLvs)
U% = UB(A)
Dim O$()
ReDim O(U)
For J% = 0 To U
    O(J) = B_Alias_To_FldNm_Dict(A(J))
Next
ZFnd_FldNmAy = O
End Property

Private Property Get ZFnd_ListCol(AliasLvs$) As ListColumn()
Dim FldNmAy$()
    FldNmAy = ZFnd_FldNmAy(AliasLvs)
U% = UB(FldNmAy)
Dim O() As ListColumn
    ReDim O(U)
For J% = 0 To U
    FldNm$ = FldNmAy(J)
    Set O(J) = A_LO.ListColumns(FldNm)
Next
ZFnd_ListCol = O
End Property


Private Property Get CC_Dte() As Variant()
A$ = ZSpecRightVal("Dte")
CC_Dte = Array(ZAliasLvs_Col(A))
End Property

Private Property Get CC_Formula() As Variant()
Dim A$(): A = ZSpecItm("Formula")
U% = UB(A)
Dim O1$(), O2()
ReSzAy O1, U
ReSzAy O2, U
For J% = 0 To U
    With Brk(A(J), "|")
        O1(J) = .S1
        O2(J) = ZAliasLvs_Col(.S2)
    End With
Next
CC_Formula = Array(O1, O2)
End Property
Private Property Get BB_SpecDict(Ay$()) As Dictionary
Dim O As New Dictionary
For J% = 0 To UB(Ay)
    BB_SpecDict__AddLine O, Ay(J)
Next
Set BB_SpecDict = O
End Property

Private Sub BB_SpecDict__AddLine(ODict As Dictionary, Line$)
Dim Ay$()
With Brk2(Line, "|")
    If .S1 = "" Then Exit Sub
    If ODict.Exists(.S1) Then
        Ay = ODict(.S1)
        Push Ay, .S2
        ODict(.S1) = Ay
    Else
        Push Ay, .S2
        ODict.Add .S1, Ay
    End If
End With
End Sub

Private Property Get CC_HAlign() As Variant()
Dim HasTot As Boolean
    HasTot = Sz(ZSpecItm("Tot")) >= 2
Dim H() As XlHAlign
Dim C()
Dim A$(): A = ZSpecItm("HAlign")
U& = UB(A)
ReSzAy H, U
ReSzAy C, U
For J% = 0 To U
    With Brk(A(J), "|")
        H(J) = CvKW_HAlign(.S1)
        C(J) = ZAliasLvs_Col(.S2, InclTot:=HasTot)
    End With
Next
CC_HAlign = Array(H, C)
End Property

Private Property Get ZZV_Sq2()
ZZV_Sq2 = BrkEDIFv(ZZV__Fv)(1)
End Property

Private Sub ZZX_Set_AB_Var()
ZZX_Set_A_Var
ZZX_Set_B_Var
End Sub

Private Sub ZZX_Set_A_Var()
Set A_LO = ZZV_A_LO
A_LO.Application.Visible = True
End Sub

Private Sub ZZX_Set_B_Var()
Set B_SpecDict = ZZV_B_SpecDict
B_FldNmAy = ZZV_B_FldNmAy
Set B_Alias_To_FldNm_Dict = ZZV_B_Alias_To_FldNm_Dict
Set B_Alias_To_Cno_Dict = ZZV_B_Alias_To_Cno_Dict
End Sub
Private Sub SetNum()
Dim R() As Range: R = C_Num_Rge
Dim Rge As Range
For J% = 0 To UB(R)
    Set Rge = R(J)
    CvRgeToNum Rge
Next
End Sub
Private Sub ZZZ_CC_Num()
ZZX_Set_AB_Var
A = CC_Num
C_Num_Rge = A(0)
End Sub

Private Sub ZZZ_SetNum()
ZZZ_CC_Num
SetNum
Stop
End Sub

Private Sub ZZZ_CC_Wdt()
ZZX_Set_AB_Var
A = CC_Wdt
C_Wdt = A(0)
C_Wdt_ColAy = A(1)
Stop
End Sub
Private Sub ZZZ_CC_HAlign()
ZZX_Set_AB_Var
A = CC_HAlign
C_HAlign = A(0)
C_HAlign_ColAy = A(1)
Stop
End Sub

Private Sub ZZZ_SetHAlign()
ZZZ_CC_HAlign
SetHAlign
End Sub

Private Sub ZZZ_SetWdt()
ZZZ_CC_Wdt
SetWdt
End Sub


Private Sub ZZZ_SetHdrHgt()
ZZX_Set_AB_Var
Av = CC_HdrHgt
C_HdrHgt_Factor = Av(0)
Set C_HdrHgt_Row = Av(1)
SetHdrHgt
Stop
End Sub


Private Sub ZZZ_FldNmAy()
ZZX_Set_AB_Var
BrwAy ZZV_B_FldNmAy
End Sub

Private Sub ZZZ_A_LO()
Dim LO As ListObject
Set LO = ZZV_A_LO
LO.Application.Visible = True
End Sub

Private Sub ZZZ_Sq2()
BrwSq ZZV_Sq2
End Sub

Private Sub ZZZ_ZSpecItm()
Set A_SpecDict = ZZV_B_SpecDict
Dim A$(): A = ZSpecItm("Wdt")
Debug.Assert A(0) = "10  | AA BB"
A = ZSpecItm("Alias")
Stop
End Sub

Private Property Get ZZV_B_SpecDict() As Dictionary
Set ZZV_B_SpecDict = BB_SpecDict(ZZV__SpecAy)
End Property

Private Property Get ZZV_A_LO() As ListObject
Static Ws As Worksheet
If Not IsVdtWs(Ws) Then
    Set Ws = NewWs
    PutSq Ws.Range("A3"), ZZV_Sq2, CrtListObj:=True
End If
Set ZZV_A_LO = Ws.ListObjects(1)
End Property

Private Property Get ZZV_B_FldNmAy() As String()
ZZV_B_FldNmAy = BB_FldNmAy(ZZV_A_LO)
End Property

Private Property Get ZZV_B_Alias_To_FldNm_Dict() As Dictionary
Set ZZV_B_Alias_To_FldNm_Dict = BB_Alias_To_FldNm_Dict(ZZV_B_FldNmAy, ZZV_B_SpecDict)
End Property

Private Property Get ZZV_B_Alias_To_Cno_Dict() As Dictionary
Set ZZV_B_Alias_To_Cno_Dict = BB_Alias_To_Cno_Dict(ZZV_B_Alias_To_FldNm_Dict)
End Property

Private Property Get BB_FldNmAy(LO As ListObject) As String()
A = LO.HeaderRowRange.Value
BB_FldNmAy = TrmAy(CvAyToStr(CvSqRow(A, 1)))
End Property

Private Property Get ZSpecRightVal$(Itm$)
A = ZSpecItm(Itm)
If Sz(A) = 0 Then Exit Property
ZSpecRightVal = Trim(A(0))
End Property

Private Property Get CC_HdrHgt() As Variant()
Dim HdrHgt_Factor%
    HdrHgt_Factor = Val(ZSpecRightVal("HdrHgt"))

Dim HdrHgt_Row As Range
    Set HdrHgt_Row = A_LO.HeaderRowRange
CC_HdrHgt = Array(HdrHgt_Factor, HdrHgt_Row)
End Property

Private Property Get CC_NumFmt() As Variant()
Dim HasTot As Boolean: HasTot = Sz(ZSpecItm("Tot")) >= 2
Dim A$(): A = ZSpecItm("NumFmt")
U% = UB(A)
Dim O1$(), O2()
ReSzAy O1, U
ReSzAy O2, U
For J% = 0 To U
    With Brk(A(J), "|")
        O1(J) = .S1
        O2(J) = ZAliasLvs_Col(.S2, InclTot:=HasTot)
    End With
Next
CC_NumFmt = Array(O1, O2)
End Property

Private Property Get CC_OutLine() As Variant()
Dim A$(): A = ZSpecItm("Outline")
U% = UB(A)
Dim O1%(), O2()
ReSzAy O1, U
ReSzAy O2, U
For J% = 0 To U
    With Brk(A(J), "|")
        O1(J) = .S1
        O2(J) = ZAliasLvs_Col(.S2)
    End With
Next
CC_OutLine = Array(O1, O2)
End Property

Private Property Get CC_Wdt() As Variant()
Dim A$()
A = ZSpecItm("Wdt")
U% = UB(A)
W = NewIntAy(U)
C = NewVarAy(U)
For J% = 0 To U
    With Brk(A(J), "|")
        W(J) = .S1
        C(J) = ZAliasLvs_EntCol(.S2)
    End With
Next
CC_Wdt = Array(W, C)
End Property

Private Sub SetHAlign()
Dim A() As XlHAlign: A = C_HAlign
Dim C():             C = C_HAlign_ColAy
If Sz(A) = 0 Then Exit Sub
Dim Rge As Range
For J% = 0 To UB(A)
    RgeAy = C(J)
    For Each I In RgeAy
        Set Rge = I
        If Rge.Row = 4 Then Stop
        Rge.HorizontalAlignment = A(J)
        Rge.Application.Visible = True
    Next
Next
End Sub

Property Get ZCol(LO As ListObject, ColNm$) As Range
Set ZCol = LO.ListColumns(ColNm).Range
End Property

Private Property Get ZRgeAy(ColNm$(), Optional ExclTot As Boolean) As Range
Dim R As Range
For J% = 0 To UB(ColNm)
    Set R = LO.ListColumns(ColNm(J)).Range
Next
ToBeCoded
End Property

Private Property Get ZColAy(ColNm$()) As Range
Dim A() As Range
A = ZRgeAy(ColNm)
For J% = 0 To UB(A)
    Set A(J) = A(J).EntireColumn
Next
ZColAy = A
End Property

Private Sub SetWrapTxt()
Dim R() As Range
R = C_WrapTxt_Rge
If Sz(R) = 0 Then Exit Sub
Dim Rge As Range
For Each I In R
    Set Rge = I
    Rge.WrapText = True
Next
End Sub

Private Sub SetOutLine()
Dim OL%(): OL = C_OutLine
Dim C(): C = C_OutLine_ColAy
If Sz(OL) = 0 Then Exit Sub
Dim Rge As Range
For J% = 0 To UB(OL)
    Lvl% = OL(J)
    For Each I In C(J)
        Set Rge = I
        Rge.EntireColumn.OutlineLevel = Lvl
    Next
Next
SetSummaryLeft LOWs(A_LO)
End Sub

Private Sub SetFormula()
Dim F$(): F = C_Formula
Dim C(): C = C_Formula_ColAy
If Sz(F) = 0 Then Exit Sub
Dim Rge As Range
For J% = 0 To UB(F)
    Formula$ = F(J)
    For Each I In C
        Set Rge = I
        Rge.Formula = Formula
    Next
Next
End Sub

Private Sub SetDte()
Dim R() As Range: R = C_Dte_Rge
If Sz(R) = 0 Then Exit Sub
Dim Rge As Range
For Each I In R
    Set Rge = I
    Rge.NumberFormat = "YYYY/MM/DD;;#"
Next
End Sub

Private Property Get ZIsPfx(OLin$, Pfx$) As Boolean
If Not IsPfx(OLin, Pfx) Then Exit Property
OLin = RmvPfx(OLin, Pfx)
ZIsPfx = True
End Property
Private Sub ZZZ_SetYYYYMMDD()
ZZX_Set_AB_Var
ZZZ_CC_YYYYMMDD
SetYYYYMMDD
End Sub
Private Sub SetYYYYMMDD()
Dim R() As Range
R = C_YYYYMMDD_Rge

If Sz(R) = 0 Then Exit Sub
Dim Rge As Range
For Each I In R
    Set Rge = I
    CvRgeFromYYYYMMDD Rge
Next
End Sub

Private Sub SetNumFmt()
Dim N$(): N = C_NumFmt
Dim C():  C = C_NumFmt_ColAy
Dim Rge As Range
For J% = 0 To UB(N)
    NumFmt$ = N(J)
    For Each I In C(J)
        Set Rge = I
        Rge.NumberFormat = NumFmt
    Next
Next
End Sub

Private Sub SetTot()
Dim F$(): F = C_Tot_Formula
Dim C() As Range: C = C_Tot_Rge
For J% = 0 To UB(F)
    C(J).Formula = F(J)
    C(J).HorizontalAlignment = XlHAlign.xlHAlignCenter
Next
End Sub

Private Property Get ZSpecItm(Itm$) As String()
K$ = Itm & "L"
If Not B_SpecDict.Exists(K) Then Exit Property
ZSpecItm = B_SpecDict(K)
End Property

Private Property Get ZZV__SpecAy() As String()
ZZV__SpecAy = ResStrAy(ZZV__SpecRes, "EDI_FmtSpec")
End Property

Private Sub ZZZ_ZSpecRightVal()
ZZX_Set_AB_Var
Debug.Assert ZSpecRightVal("HdrHgt") = "3"
End Sub

Private Sub ZZZ_ZAliasLvs_Col()
ZZX_Set_AB_Var
Dim Act() As Range
Act = ZAliasLvs_Col("TstDte ExpDte")
A1$ = Act(0).Address
A2$ = Act(1).Address
Stop
End Sub

Private Sub ZZZ_CC_YYYYMMDD()
ZZX_Set_AB_Var
Av = CC_YYYYMMDD
C_YYYYMMDD_Rge = Av(0)
End Sub


