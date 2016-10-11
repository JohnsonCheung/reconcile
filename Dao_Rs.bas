Attribute VB_Name = "Dao_Rs"
Option Compare Database

Property Get RsAy(Rs As Recordset) As Variant()
Dim O()
With Rs
    While Not .EOF
        Push O, .Fields(0).Value
        .MoveNext
    Wend
End With
RsAy = O
End Property

Property Get RsInt%(Rs As Recordset)
AssertRsEOF Rs
RsInt = Rs.Fields(0).Value
End Property
Sub AssertRsEOF(Rs As Recordset)
If Rs.EOF Then Er "Given Rs is EOF"
End Sub
Property Get RsLng&(Rs As Recordset)
AssertRsEOF Rs
RsLng = Rs.Fields(0).Value
End Property
