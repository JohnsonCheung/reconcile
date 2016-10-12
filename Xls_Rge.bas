Attribute VB_Name = "Xls_Rge"
Option Compare Database
Property Get RgeWs(Rge As Range) As Worksheet
Set RgeWs = Rge.Parent
End Property
Property Get RgeRCRC(Rge As Range, R1, C1, R2, C2) As Range
Set RgeRCRC = RgeWs(Rge).Range(Rge(R1, C1), Rge(R2, C2))
End Property
Property Get ReSzRge(At As Range, Sq) As Range
R& = UBound(Sq, 1)
C% = UBound(Sq, 2)
Set ReSzRge = RgeRCRC(At, 1, 1, R, C)
End Property
Property Get RgeRR(R As Range, R1, R2) As Range
Set RgeRR = RgeRCRC(R, R1, 1, R2, R.Columns.Count)
End Property
Property Get RgeCC(R As Range, C1, C2) As Range
Set RgeCC = RgeRCRC(R, 1, C1, R.Rows.Count, C2)
End Property
Function PutSq(At As Range, Sq, Optional CrtListObj As Boolean, Optional NoListObjHeader As Boolean)
Dim R As Range
Set R = ReSzRge(At, Sq)
R.Value = Sq
If CrtListObj Then
    Set PutSq = RgeWs(R).ListObjects.Add(xlSrcRange, R, XlListObjectHasHeaders:=IIf(NoListObjHeader, xlNo, xlYes))
Else
    Set PutSq = R
End If
End Function
Property Get RgeCRR(Rge As Range, C, R1, R2) As Range
Set RgeCRR = RgeRCRC(Rge, R1, C, R2, C)
End Property
Private Sub ZZZ_PutSq()
Dim Ws As Worksheet
    Set Ws = NewWs(Visible:=True)
PutSq Ws.Range("A1"), SampleSq1, CrtListObj:=True
End Sub
Sub ResetNumWarn(Rge As Range)
Dim Cell As Range
For Each Cell In Rge
ToBeCoded
Next
End Sub
Property Get RgeRC(Rge As Range, R, C) As Range
Set RgeRC = Rge(R, C)
End Property
Property Get RgeC(Rge As Range, C) As Range
Set RgeC = RgeRCRC(Rge, 1, C, Rge.Rows.Count, C)
End Property
Sub CvRgeToNum(Rge As Range)
Sq = Rge.Value
If Rge.Count = 1 Then
    Rge.Value = Val(Sq)
    Exit Sub
End If

For J& = 1 To UBound(Sq, 1)
    For I& = 1 To UBound(Sq, 2)
        Sq(J, I) = Val(Sq(J, I))
    Next
Next
Rge.Value = Sq
End Sub
Property Get AtDownEndR&(At As Range)
If IsEmpty(AtDownCell(At).Value) Then
    AtDownEndR = At.Column
Else
    AtDownEndR = At.End(xlToDown).Column
End If
End Property
Property Get At(Rge As Range) As Range
Set At = Rge(1, 1)
End Property
Property Get AtRightEndC%(At As Range)
If IsEmpty(AtRightCell(At).Value) Then
    AtRightEndC = Rge.Column
Else
    AtRightEndC = At.End(xlToRight).Column
End If
End Property
Property Get AtRightCell(At As Range, Optional NRight% = 1) As Range
Set AtRightCell = At(1, NRight + 1)
End Property
Property Get AtLeftCell(At As Range, NLeft%) As Range
Set AtLeftCell = At(1, -NLeft + 1)
End Property
Private Sub ZZZ_AtUpCell()
Dim Act As Range
Dim Ws As Worksheet
Set Ws = NewWs

Set Act = AtUpCell(WsRC(Ws, 12, 34), 0)
Debug.Assert Act.Column = 34
Debug.Assert Act.Row = 12

Set Act = AtUpCell(WsRC(Ws, 12, 34), 1)
Debug.Assert Act.Column = 34
Debug.Assert Act.Row = 11

Set Act = AtUpCell(WsRC(Ws, 12, 34), 2)
Debug.Assert Act.Column = 34
Debug.Assert Act.Row = 10

Set Act = AtUpCell(WsRC(Ws, 12, 34), -1)
Debug.Assert Act.Column = 34
Debug.Assert Act.Row = 13

Ws.Application.Quit
Stop
End Sub

Private Sub ZZZ_FreezeAt()
Dim Ws As Worksheet
Dim At As Range
Set Ws = NewWs(Visible:=True)
Set At = Ws.Range("A2")
FreezeAt At
End Sub

Sub FreezeAt(At As Range)
Dim Wb As Workbook: Set Wb = RgeWs(At).Parent
Dim Win As Excel.Window: Set Win = Wb.Windows(1)
Win.WindowState = xlMaximized
WsRC(At.Parent, 1, 1).Activate
WsRC(At.Parent, 1, 1).Select
At.Select
Win.FreezePanes = True
End Sub
Property Get AtUpCell(At As Range, Optional NUp& = 1) As Range
Set AtUpCell = At(1 - NUp, 1)
End Property
Property Get AtDownCell(At As Range, Optional NDown& = 1) As Range
Set AtDownCell = At(NDown + 1, 1)
End Property
Property Get RgeRCC(Rge As Range, R, C1, C2) As Range
Set RgeRCC = RgeRCRC(Rge, R, C1, R, C2)
End Property
Sub SetBdr(Rge As Range)
With Rge.Borders(xlInsideHorizontal)
    .LineStyle = XlLineStyle.xlContinuous
    .Weight = xlThin
End With
With Rge.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlMedium
End With
With Rge.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlMedium
End With
With Rge.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlMedium
End With
With Rge.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlMedium
End With
End Sub
Property Get RgeR(Rge As Range, R) As Range
Set RgeR = RgeRCRC(Rge, R, 1, R, Rge.Columns.Count)
End Property
