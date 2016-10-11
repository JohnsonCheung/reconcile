Attribute VB_Name = "Dao_Sql"
Option Compare Database

Property Get SqlSel$(T$, F_Lvc$, W$)
SqlSel = FmtQQ("Select ? from ??", F_Lvc, T, Phrase_Where(W))
End Property

Property Get SqlTblWhere$(T$, Optional Where$)
If Where <> "" Then W$ = " Where " & Where
SqlTblWhere = FmtQQ("Select * from ??", T, W)
End Property

Property Get SqlAy(Sql$, Optional D As Database, Optional IgnoreEr As Boolean) As Variant()
SqlAy = RsAy(NzDb(D).OpenRecordset(Sql))
End Property

Property Get SqlStrAy(Sql$, Optional D As Database, Optional IgnoreEr As Boolean) As String()
SqlStrAy = CvAyToStr(SqlAy(Sql, D, IgnoreEr))
End Property

Sub RunSqlAy(SqlAy$(), Optional D As Database, Optional IgnoreEr As Boolean)
For J% = 0 To UB(SqlAy)
    RunSql SqlAy(J), D, IgnoreEr
Next
End Sub

Sub RunSql(Sql$, Optional D As Database, Optional IgnoreEr As Boolean)
If IgnoreEr Then
    On Error GoTo X
    NzDb(D).Execute Sql
    Exit Sub
X:
    Debug.Print "DbRunSql: There is error in running Sql:"
    Debug.Print Sql
    Debug.Print Err.Description
    Exit Sub
End If
NzDb(D).Execute Sql
End Sub

Sub RunSqlQQ(SqlQQ$, ParamArray Ap())
Dim Av()
Av = Ap
RunSql FmtQQAv(SqlQQ, Av)
End Sub

Property Get SqlInt%(Sql$, Optional D As Database)
SqlInt = RsInt(NzDb(D).OpenRecordset(Sql))
End Property

Property Get SqlLng&(Sql$, Optional D As Database)
SqlLng = RsLng(NzDb(D).OpenRecordset(Sql))
End Property


