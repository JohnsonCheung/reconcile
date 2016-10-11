Attribute VB_Name = "Vb_Ft"
Option Compare Database

Sub BrwFt(Ft$, Optional KillFt As Boolean = False)
Shell "NotePad " & Ft, vbMaximizedFocus
If KillFt Then Kill Ft
End Sub
Property Get FtStr$(Ft$, Optional KillFt As Boolean = False)
Dim S As TextStream
Set S = Fso.GetFile(Ft).OpenAsTextStream(ForReading)
FtStr = S.ReadAll
S.Close
If KillFt Then Kill Ft
End Property
Property Get FtAy(Ft$, Optional KillFt As Boolean = False) As String()
FtAy = SplitLines(FtStr(Ft, KillFt))
End Property
