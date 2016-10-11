Attribute VB_Name = "Dao_Lnk"
Option Compare Database
Private Sub LnkFb__Tst()
SetCurFbPth
Fb$ = "GTRADATA1.accdb"
LnkFb CurT("B6AML"), Fb, "AML"
End Sub
Sub LnkFb(T As DbT, Fb$, Optional TblNm$)
DrpDbTIfExist T
Debug.Print Fb
Cnn$ = FmtQQ(";DATABASE=?", Fb)
Dim Tbl As TableDef
Set Tbl = T.D.CreateTableDef(T.T)
Tbl.Connect = Cnn
Tbl.SourceTableName = NzStr(TblNm, T.T)
T.D.TableDefs.Append Tbl
End Sub
Sub LnkFx(T As DbT, Fx$, Optional WsNm$)
DrpDbTIfExist T

End Sub
Sub LnkFv(T As DbT, Fv$)
DrpDbTIfExist T
End Sub
Sub LnkFt(T As DbT, Ft$)
DrpDbTIfExist T

End Sub

