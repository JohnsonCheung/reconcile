Attribute VB_Name = "Xls_Functions"
Option Compare Database

Function NzXls(Xls As Excel.Application) As Excel.Application
If IsNothing(Xls) Then
    Set NzXls = New Excel.Application
Else
    Set NzXls = Xls
End If
End Function
