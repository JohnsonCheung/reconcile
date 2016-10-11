Attribute VB_Name = "Xls_LO"
Option Compare Database
Property Get LOFldNmAy(LO As ListObject) As String()
Dim O$()
N% = LO.ListColumns.Count
ReDim O(N - 1)
For J = 0 To N - 1
    O(J) = LO.ListColumns(J + 1).Name
Next
LOFldNmAy = O
End Property

Property Get LOWs(LO As ListObject) As Worksheet
Set LOWs = LO.Parent
End Property
Property Get LOWb(LO As ListObject) As Workbook
Set LOWb = LOWs(LO).Parent
End Property

