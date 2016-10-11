Attribute VB_Name = "Dao_DbInf"
Option Compare Database

Property Get DbTblNmAy(Optional Lik$ = "*", Optional Db As Database) As String()
Sql$ = FmtQQ("Select Name from MSysObjects where Name not like 'MSYS*' and name not like '~*' and Name Like '?' and Type=1", Lik)
DbTblNmAy = SqlStrAy(Sql, Db)
End Property
Property Get DbQryNmAy(Optional Lik$ = "*", Optional Db As Database) As String()
Sql$ = FmtQQ("Select Name from MSysObjects where  Name not like 'MSYS*' and name not like '~*' and  Name Like '?' and Type=5", Lik)
DbQryNmAy = SqlStrAy(Sql, Db)
End Property

