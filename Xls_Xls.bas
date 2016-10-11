Attribute VB_Name = "Xls_Xls"
Option Compare Database

Sub QuitXls(Xls As Excel.Application)
On Error Resume Next
Xls.Quit
Set Xls = Nothing
End Sub
