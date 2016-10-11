Attribute VB_Name = "Xls_Ws"
Option Compare Database
Sub SetSummaryAbove(Ws As Worksheet)
Ws.Activate
R& = WsLastCell(Ws).Row + 1
WsRC(Ws, R, 1).Select
WsRC(Ws, R, 1).Activate
Ws.OutLine.SummaryRow = xlSummaryAbove
Ws.Range("A1").Select
End Sub
Sub SetSummaryLeft(Ws As Worksheet)
Ws.Activate
R& = WsLastCell(Ws).Row + 1
WsRC(Ws, R, 1).Select
WsRC(Ws, R, 1).Activate
Ws.OutLine.SummaryColumn = xlSummaryOnLeft
Ws.Range("A1").Select
End Sub

Property Get WsLastCell(Ws As Worksheet) As Range
Set WsLastCell = Ws.Cells.SpecialCells(xlCellTypeLastCell)
End Property

Property Get WsVeryLastCell(Ws As Worksheet) As Range
Set WsVeryLastCell = WsRC(Ws, WsMaxRno(Ws), WsMaxCno(Ws))
End Property

Property Get WsMaxRno&(Ws As Worksheet)
WsMaxRno = Ws.Cells.Rows.Count
End Property

Property Get WsMaxCno&(Ws As Worksheet)
WsMaxCno = Ws.Cells.Columns.Count
End Property

Property Get WsC1C2(Ws As Worksheet, C As C1C2) As Range
Set WsC1C2 = WsCC(Ws, C.C1, C.C2)
End Property
Property Get WsWb(Ws As Worksheet) As Workbook
Set WsWb = Ws.Parent
End Property
Property Get WsR1R2(Ws As Worksheet, R As R1R2) As Range
Set WsR1R2 = WsRR(Ws, R.R1, R.R2)
End Property
Property Get WsRR(Ws As Worksheet, R1, R2) As Range
Set WsRR = Ws.Range(Ws.Cells(R1, 1), Ws.Cells(R2, 1)).EntireRow
End Property
Sub SetRowOutLine(Ws As Worksheet, R() As R1R2, Optional Lvl% = 2)
For J% = 0 To R1R2UB(R)
    WsR1R2(Ws, R(J)).OutlineLevel = Lvl
Next
End Sub
Property Get WsLO(Ws As Worksheet) As ListObject
Set WsLO = Ws.ListObjects(1)
End Property
Sub SetColOutLine(Ws As Worksheet, C() As C1C2, Optional Lvl% = 2)
For J% = 0 To UB(R)
    WsC1C2(Ws, C(J)).OutlineLevel = Lvl
Next
End Sub
Property Get IsVdtWs(Ws As Worksheet) As Boolean
On Error GoTo X
A = Ws.Name
IsVdtWs = True
Exit Property
X:
End Property

Sub SetRowHgtTriple(Ws As Worksheet, R)
Dim Rge As Range
Set Rge = WsR(Ws, R)
Rge.RowHeight = Rge.RowHeight * 3
End Sub
Property Get WsRCRC(Ws As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = Ws.Range(Ws.Cells(R1, C1), Ws.Cells(R2, C2))
End Property

Property Get WsRC(Ws As Worksheet, R, C) As Range
Set WsRC = Ws.Cells(R, C)
End Property

Property Get WsCC(Ws As Worksheet, C1, C2) As Range
Set WsCC = Ws.Range(Ws.Cells(1, C1), Ws.Cells(1, C2)).EntireColumn
End Property
Property Get NewWs(Optional Wb As Workbook, Optional WsNm$, Optional AtBeg As Boolean, Optional AftWsNm$, Optional Visible As Boolean) As Worksheet
If IsNothing(Wb) Then Set Wb = NewWb
If AtBeg Then
    Set O = Wb.Sheets.Add(Wb.Sheets(1))
ElseIf AftWs = "" Then
    Set O = Wb.Sheets.Add(, LastWs(Wb))
Else
    Set O = Wb.Sheets.Add(, Wb.Sheets(AftWs))
End If
If WsNm <> "" Then O.Name = WsNm
Set NewWs = O
If Visible Then Wb.Application.Visible = True
End Property
Property Get WsA1(Ws As Worksheet) As Range
Set WsA1 = Ws.Cells(1, 1)
End Property
Property Get WsA1SqRge(Ws As Worksheet) As Range
Dim A1 As Range
Set A1 = WsA1(Ws)
R& = AtDownEndR(A1)
C% = AtRightEndC(A1)
WsA1SqRge = WsRCRC(Ws, 1, 1, R, C)
End Property
Property Get WsR(Ws As Worksheet, R) As Range
Dim Rge As Range
Set Rge = Ws.Rows(R)
Set WsR = Rge.EntireRow
End Property
Property Get WsC(Ws As Worksheet, C) As Range
Dim Rge As Range
Set Rge = Ws.Columns(C)
Set WsC = Rge.EntireColumn
End Property

