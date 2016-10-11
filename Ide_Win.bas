Attribute VB_Name = "Ide_Win"
Option Compare Database

Sub ClsAllWin()
Dim W As Window
For Each W In VBE.Windows
    W.Close
Next
End Sub
