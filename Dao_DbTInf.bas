Attribute VB_Name = "Dao_DbTInf"
Option Compare Database
Type DbT
    D As Database
    T As String
End Type
Property Get DbT(D As Database, T$) As DbT
Dim O As DbT
Set O.D = D
O.T = T
DbT = O
End Property

Property Get DbTFlds(P As DbT) As Fields
If IsQry(P.D, P.T) Then
    Set DbTFlds = P.D.QueryDefs(P.T).Fields
Else
    Set DbTFlds = P.D.TableDefs(P.T).Fields
End If
End Property

Property Get DbTFldNmAy(P As DbT, Optional FirstNFld% = 0) As String()
DbTFldNmAy = FldsNmAy(DbTFlds(P), FirstNFld)
End Property

Property Get DbTNFld%(P As DbT)
DbTNFld = P.D.TableDefs(P.T).Fields.Count
End Property


Property Get IsDbTOpn(P As DbT) As Boolean
If Not IsDbTOpn(P) Then Exit Property
ToBeCoded
End Property
Property Get DbTFldAy(P As DbT, Optional Lik$ = "*") As DAO.Field()
DbTFldAy = FldsLik(P.D.TableDefs(P.T).Fields, Lik)
End Property
Private Sub DbTTxtFldAy__Tst()
BrwAy DbTTxtFldAy(CurT("MGE"))
End Sub
Property Get IsDbTCurDb(P As DbT) As Boolean
IsDbTCurDb = CurrentDb.Name = P.D.Name
End Property
Property Get DbTTxtFldAy(P As DbT) As String()
DbTTxtFldAy = FldsTxtFldAy(DbTFlds(P))
End Property


Property Get CurT(T$) As DbT
CurT = DbT(CurrentDb, T)
End Property
Private Property Get CurTTxtFldAy__Tst()
BrwAy CurTTxtFldAy("MGE")
End Property
Property Get CurTTxtFldAy(T$) As String()
CurTTxtFldAy = DbTTxtFldAy(CurT(T))
End Property


Private Sub CurTFldNmAy__Tst()
BrwAy CurTFldNmAy("MGE")
End Sub
Property Get CurTFldNmAy(T$) As String()
CurTFldNmAy = DbTFldNmAy(CurT(T))
End Property


