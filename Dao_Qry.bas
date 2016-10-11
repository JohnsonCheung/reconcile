Attribute VB_Name = "Dao_Qry"
Option Compare Database

Sub CrtQry(QryNm$, Sql$)
DbCrtQry CurrentDb, QryNm, Sql
End Sub
Sub DbCrtQry(D As Database, QryNm$, Sql$)
If IsQry(D, QryNm) Then
    D.QueryDefs(QryNm).Sql = Sql
Else
    D.QueryDefs.Append D.CreateQueryDef(QryNm, Sql)
End If
End Sub
Function IsQry(D As Database, QryNm$) As Boolean
Dim Q As QueryDef
For Each Q In D.QueryDefs
    If Q.Name = QryNm Then IsQry = True: Exit Function
Next
End Function
