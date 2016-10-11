Attribute VB_Name = "Dao_DbQ"
Option Compare Database
Type DbQ
    D As Database
    Q As String
End Type
Property Get DbQ(D As Database, Q$) As DbQ
Dim O As DbQ
Set O.D = D
O.Q = Q
DbQ = O
End Property

Sub CrtDbQ(Q As DbQ, Sql$)
If IsDbQ(Q) Then
    Q.D.QueryDefs(Q.Q).Sql = Sql
Else
    Q.D.QueryDefs.Append Q.D.CreateQueryDef(Q.Q, Sql)
End If
End Sub

Sub DrpDbQ(Q As DbQ)
Q.D.QueryDefs.Delete Q.Q
End Sub
Sub DnlDbQ(Q As DbQ, T$)
DrpDbTIfExist DbT(Q.D, T)
Q.D.Execute FmtQQ("Select * into [?] from [?]", T, Q.Q)
End Sub
Property Get CurDbQ(Q$) As DbQ
CurDbQ = DbQ(CurrentDb, Q)
End Property
Private Sub CrtPssQ__Tst()
CrtPssQ CurPssQ("TmpQry", DsnCnnStr(DsnFfn(Godiva.B6F))), "Select * from IIC"
End Sub
Sub CrtPssQ(Q As PssQ, Sql$)
CrtDbQ DbQ(Q.D, Q.Q), Sql
Q.D.QueryDefs(Q.Q).Connect = Q.CnnStr
End Sub

Property Get IsPssQ(Q As DbQ)
If Not IsDbQ(Q) Then Exit Property
IsPssQ = Q.D.QueryDefs(Q.Q).Type = QueryDefTypeEnum.dbQSQLPassThrough
End Property

Property Get CurQ(Q$) As DbQ
Dim O As DbQ
Set O.D = CurrentDb
O.Q = Q
CurQ = O
End Property
