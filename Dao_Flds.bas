Attribute VB_Name = "Dao_Flds"
Option Compare Database
Private Sub FldsNmAy__Tst()
BrwAy FldsNmAy(CurrentDb.TableDefs("MGE").Fields)
Stop
End Sub
Property Get FldsNmAy(P As DAO.Fields, Optional FirstNFld%) As String()
U% = P.Count - 1
If FirstNFld > 0 Then U = FirstNFld - 1
If U = -1 Then Exit Property
ReDim O$(U)
For J% = 0 To U
   O(J) = P(J).Name
Next
FldsNmAy = O
End Property

Property Get FldsTxtFldAy(P As DAO.Fields) As String()
U% = P.Count - 1
If U = -1 Then Exit Property
Dim O$()
For J% = 0 To U
    T% = P(J).Type
    If T = DAO.DataTypeEnum.dbText Or T = DAO.DataTypeEnum.dbMemo Then
        Push O, P(J).Name
    End If
Next
FldsTxtFldAy = O
End Property
Property Get FldsLik(P As DAO.Fields, Lik$) As DAO.Field()
Dim O() As DAO.Field
For J% = 0 To P.Count - 1
    If P(J).Name Like Lik Then Push O, P(J)
Next
FldsLik = O
End Property
