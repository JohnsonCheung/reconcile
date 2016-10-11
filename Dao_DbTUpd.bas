Attribute VB_Name = "Dao_DbTUpd"
Option Compare Database

Sub CpyDbT(P As DbT, ToTblNm$, Optional Where$)
DrpDbTIfExist DbT(P.D, ToTblNm)
If Where <> "" Then W$ = " where " & Where
P.D.Execute FmtQQ("select * into [?] from [?]?", ToTblNm, P.T, W)
End Sub

Sub CpyDbTStru(P As DbT, ToTblNm$)
DrpDbTIfExist DbT(P.D, ToTblNm)
P.D.Execute FmtQQ("select * into [?] from [?] where false", ToTblNm, P.T)
End Sub

Sub DrpDbT(P As DbT)
On Error GoTo Er
P.D.Execute "Drop Table [" & P.T & "]"
Exit Sub
Er: Er "Cannot Drop {Table} @ {Db} due to {Er}, {ErNo}", P.T, P.D.Name, Err.Description, Err.Number
End Sub

Private Sub DrpDbTIfExist__Tst()
DrpDbTIfExist CurT("Asf")
End Sub

Sub DrpDbTIfExist(P As DbT)
On Error GoTo Er
P.D.Execute "Drop Table [" & P.T & "]"
Exit Sub
Er: If Err.Number <> 3376 Then Er "Cannot Drop {Table} @ {Db} due to {Er}, {ErNo}", P.T, P.D.Name, Err.Description, Err.Number
End Sub

Sub InsDbT(P As DbT, FldNmLvs$, ParamArray Ap())
Dim Av()
Av = Ap
InsDbT_Av P, FldNmLvs, Av
End Sub

Sub InsDbT_Av(P As DbT, FldNmLvs$, Av())
U% = UB(Av)
ColNmAy = SplitLvs(FldNmLvs)
UFld = UB(ColNmAy)
With P.D.TableDefs(T).OpenRecordset
    For J = 0 To UB(Av(0))
        .AddNew
        For I = 0 To UFld
            .Fields(ColNmAy(I)).Value = Av(I)(J)
        Next
        .Update
    Next
    .Close
End With
End Sub

Sub InsCurT(T$, FldNmLvs$, ParamArray Ap())
Dim Av()
Av = Ap
InsDbT_Av CurT(T), FldNmLvs, Av
End Sub

Sub DrpCurT(T$)
DrpDbT CurT(T)
End Sub



