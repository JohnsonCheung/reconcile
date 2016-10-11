Attribute VB_Name = "Dao_DbTCol"
Option Compare Database

Property Get DbTCol(P As DbT, ColNm$, Optional Where$) As Variant()
Dim O()
With P.D.OpenRecordset(SqlSel(P.T, ColNm, Where))
    While Not .EOF
        Push O, .Fields(0).Value
        .MoveNext
    Wend
    .Close
End With
DbTCol = O
End Property


Sub DbTCol_Ap(P As DbT, ColNmLvs$, ParamArray OColAp())
Dim OAv()
OAv = OColAp
DbTCol_Av P, ColNmLvs, "", OAv
For J% = 0 To UB(OAv)
    OColAp(J) = OAv(J)
Next
End Sub

Property Get CurTCol(T$, ColNm$, Optional Where$) As Variant()
CurTCol = DbTCol(CurT(T), ColNm, Where)
End Property

Private Sub DbTCol_Av(P As DbT, ColNmLvs$, Where$, OAv())
Fld = SplitLvs(ColNmLvs)
UFld% = UB(Fld)
Sql = SqlSel(P.T, Join(Fld, ","), Where)
With P.D.OpenRecordset(Sql)
    While Not .EOF
        For J% = 0 To UFld
            Push OAv(J), .Fields(Fld(J)).Value
        Next
        .MoveNext
    Wend
    .Close
End With
End Sub


Private Sub DbTCol_Av__Tst()
Dim U$(), W$()
Dim Av()
Av = Array(U, W)
DbTCol_Av CurT("MGE"), "AC NET", "", Av
Stop
End Sub


Private Sub DbTCol_Ap__Tst()
Dim AC$(), NET@()
DbTCol_Ap CurT("MGE"), "AC NET", AC, NET
Stop
End Sub

Property Get DbTCol_Lng(P As DbT, ColNm$, Optional Where$ = "") As Long()
DbTCol_Lng = CvAyToLng(DbTCol(P, ColNm, Where))
End Property

Property Get DbTCol_Str(P As DbT, ColNm$, Optional Where$ = "") As String()
DbTCol_Str = CvAyToStr(DbTCol(P, ColNm, Where))
End Property

Private Sub DbTCol_Str__Tst()
BrwAy DbTCol(CurT("MGE"), "NET", "NET>1000")
End Sub


Sub DbTCol_Ap_Where(P As DbT, ColNmLvs$, Where$, ParamArray OColAp())
Dim OAv()
OAv = OColAp
DbTCol_Av P, ColNmLvs, Where, OAv
For J% = 0 To UB(OAv)
    OColAp(J) = OAv(J)
Next
End Sub


