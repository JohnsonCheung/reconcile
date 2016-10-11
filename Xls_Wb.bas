Attribute VB_Name = "Xls_Wb"
Option Compare Database
Private Sub NewWb__Tst()
Dim W As Workbook
Set W = NewWb("XX")
Stop
End Sub
Private Sub SetWbCusPrp__Tst()
Dim Wb As Workbook
Set Wb = NewWb
SetWbCusPrp Wb, CvLm("aa=1;bb=3;cc=3")
End Sub
Sub SetWbCusPrp(Wb As Workbook, CusPrp As Dictionary)
Dim A As CustomProperties
Set A = Wb.CustomDocumentProperties
For Each K In CusPrp.Keys
    Nm$ = K
    V = CusPrp(Nm)
    A.Add Nm, V
Next
End Sub
Property Get NewWb(Optional Fx$, Optional WsNmLvm$ = "Sheet1", Optional Xls As Excel.Application) As Workbook
If IsOpnWb(Xls, Fx) Then
    Set NewWb = Excel.Application.Workbooks(FfnFn(Fx))
    Exit Property
End If
Dim O As Workbook
Set O = NzXls(Xls).Workbooks.Add
Dim WsNm$()
WsNm = SplitLvm(WsNmLvm)
KeepFirstWs O
O.Sheets(1).Name = WsNm(0)
For J% = 1 To UB(WsNm)
    Set Ws = NewWs(O, WsNm(J))
Next

If Fx <> "" Then
    DltFfnIfExist Fx
    O.SaveAs Fx, ConflictResolution:=xlLocalSessionChanges
End If
Set NewWb = O
End Property

Private Sub IsOpnWb__Tst()
Dim Xls As Excel.Application
Debug.Assert IsOpnWb(Xls, "c:\temp\aa.xlsx") = False
End Sub
Property Get IsOpnWb(Xls As Excel.Application, Fx$) As Boolean
If IsNothing(Xls) Then Exit Property
Dim Wb As Workbook
For Each Wb In Xls.Workbooks
    If Wb.FullName = Fx Then IsOpnWb = True: Exit Property
Next
End Property

Sub KeepFirstWs(Wb As Workbook)
Dim Ws As Worksheet
For J% = Wb.Sheets.Count To 2 Step -1
    Set Ws = Wb.Worksheets(J)
    Ws.Delete
Next
End Sub

Property Get LastWs(Wb As Workbook) As Worksheet
Set LastWs = Wb.Sheets(Wb.Sheets.Count)
End Property
