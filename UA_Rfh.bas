Attribute VB_Name = "UA_Rfh"
Option Compare Database
Const TblLst$ = "UsrWhs UsrTxTy UsrPgm Usr Pgm Whs TxTy"
Const Where_Active = "ZID='SC'"

Sub RfhUsrAuthDb()
Tmp_UsrWhs
Tmp_UsrPgm
Tmp_UsrPgmPfx
Tmp_UsrTxTy
Tmp_Usr
Tmp_Co
Tmp_TxTy
Tmp_Whs
Tmp_Pgm

TblLst_Delete_All_Rec
Tbl_Ins_FmTmp "Usr"
Tbl_Ins_FmTmp "Pgm"
Tbl_Ins_FmTmp "Whs"
Tbl_Ins_FmTmp "Co"
Tbl_Ins_FmTmp "TxTy"
Tbl_Ins_FmTmp "UsrWhs"
Tbl_Ins_FmTmp "UsrCo"
Tbl_Ins_FmTmp "UsrPgm"
Tbl_Ins_FmTmp "UsrTxTy"
TblLst_Drop_TmpTbl
End Sub

Private Sub Tbl_Ins_FmTmp(T$)
RunSql FmtQQ("Insert into ? Select * from [#?]", T, T)
End Sub

Private Sub TblLst_Ay__Tst()
DmpAy TblLst_Ay
End Sub

Private Property Get TblLst_Ay() As String()
TblLst_Ay = Split(TblLst)
End Property

Private Sub TblLst_Drop_TmpTbl()
For Each T In TblLst_Ay
    RunSql FmtQQ("Drop Table [#?]", T)
Next
End Sub

Private Sub TblLst_Delete_All_Rec()
For Each T In TblLst_Ay
    RunSql FmtQQ("Delete From ?", T)
Next
End Sub

Private Sub Tmp_UsrTxTy()
Dim Usr$(), TxTyStr$()
Dim U$(), W$()
UsrTxTy_Ay_ZSC Usr, TxTyStr
UsrTxTy_Ay_Norm U, W, Usr, TxTyStr
UsrTxTy_CrtTmp U, W
End Sub

Private Sub Tmp_UsrPgmPfx()
End Sub

Private Sub Tmp_Usr()
End Sub

Private Sub Tmp_Whs()
RunSql "Select Distinct WhsId into [#Whs] from [#UsrWhs]"
End Sub
Private Sub Tmp_Co()
Tmp__Distinct "Co"
End Sub

Private Sub Tmp_Pgm()
RunSql "Select Distinct PgmId into [#Pgm] from [#UsrPgm]"
End Sub

Private Sub Tmp_UsrPgm()
Dim Usr$(), Pgm$()
Dim U$(), P$()
UsrPgm_Ay_ZSC Usr, Pgm
UsrPgm_Ay_Norm U, P, Usr, Pgm
UsrPgm_CrtTmp U, P
End Sub

Private Sub Tmp_UsrWhs()
Dim Usr$(), WhsStr$()
Dim U$(), W$()
UsrWhs_Ay_ZSC Usr, WhsStr
UsrWhs_Ay_Norm U, W, Usr, WhsStr
UsrWhs_CrtTmp U, W
End Sub
Private Sub Tmp_UsrCo()
Dim Usr$(), CoAy()
Dim U$(), C$()
UsrCo_Ay_ZSC Usr, CoAy
UsrCo_Ay_Norm U, C, Usr, CoAy
UsrCo_CrtTmp U, C
End Sub

Private Sub UsrWhs_CrtTmp(Usr$(), Whs$())
DrpCurT "#UsrWhs"
RunSql "Select UsrId,WhsId into [#UsrWhs] from UsrWhs where False"
InsCurT "#UsrWhs", "UsrId WhsId", Usr, Whs
End Sub
Private Sub UsrTxTy_CrtTmp(Usr$(), TxTy$())
DrpCurT "#UsrTxTy"
RunSql "Select UsrId,TxTyId into [#UsrTxTy] from UsrTxTy where False"
InsCurT "#UsrTxTy", "UsrId TxTyId", Usr, TxTy
End Sub
Private Sub UsrPgm_CrtTmp(Usr$(), Pgm$())
DrpCurT "#UsrPgm"
RunSql "Select UsrId,PgmId into [#UsrPgm] from UsrPgm where False"
InsCurT "#UsrPgm", "UsrId PgmId", Usr, Pgm
End Sub

Private Sub UsrCo_CrtTmp(Usr$(), Co$())
DrpCurT "#UsrCo"
RunSql "Select UsrId,CoId into [#UsrCo] from UsrCo where False"
InsCurT "#UsrCo", "UsrId CoId", Usr, Co
End Sub

Private Sub Tmp_TxTy()
Tmp__Distinct "TxTy"
End Sub
Private Sub Tmp__Distinct(T$)
DrpCurT "#" & T
RunSql FmtQQ("Select Distinct ?Id into [#?] from [#Usr?]", T, T, T)
End Sub
Private Sub Rfh_ZSC()
Run "Select * into ZSC from BP45USFGOP_ZSC"
End Sub

Private Sub UsrWhs_Ay_ZSC__Tst()
Dim U$(), W$()
UsrWhs_Ay_ZSC U, W
Stop
End Sub

Private Sub UsrWhs_Ay_Norm__Tst()
Dim U$(), W$()
UsrWhs_Ay_ZSC U, W
Dim OU$(), OW$()
UsrWhs_Ay_Norm OU, OW, U, W
Stop
Stop
End Sub

Private Sub UsrTxTy_Ay_ZSC__Tst()
Dim U$(), T$()
UsrTxTy_Ay_ZSC U, T
Stop
End Sub

Private Sub UsrTxTy_Ay_Norm__Tst()
Dim U$(), T$()
UsrTxTy_Ay_ZSC U, T
Dim OU$(), OT$()
UsrTxTy_Ay_Norm OU, OT, U, T
Stop
End Sub
Private Property Get Tbl_ZSC() As DbT
Dim O As DbT
O.T = "ZSC"
Set O.D = CurrentDbt
Tbl_ZSC = O
End Property
Private Sub UsrWhs_Ay_ZSC(OUsr$(), OWhsStr$())
DbTCol_Ap_Where Tbl_ZSC, "ZPRF ZWHSE", Where_Active, OUsr, OWhsStr
OUsr = TrmAy(OUsr)
OWhsStr = TrmAy(OWhsStr)
End Sub

Private Sub UsrPgm_Ay_ZSC(OUsr$(), OPgmStr$())
DbTCol_Ap_Where Tbl_ZSC, "ZPRF ZCLP", Where_Active, OUsr, OPgmStr
OUsr = TrmAy(OUsr)
OPgmStr = TrmAy(OPgmStr)
End Sub

Private Sub UsrTxTy_Ay_ZSC(OUsr$(), OTxTyStr$())
DbTCol_Ap_Where Tbl_ZSC, "ZPRF ZTREF", Where_Active, OUsr, OTxTyStr
OUsr = TrmAy(OUsr)
OTxTyStr = TrmAy(OTxTyStr)
End Sub


Private Sub UsrCo_Ay_ZSC(OUsr$(), OCoAy())
Erase OUsr
Erase OCoAy
S$ = SqlSel("ZSC", "*", Where_Active)
Dim Rs As Recordset
Set Rs = CurrentDb.OpenRecordset(S)
With Rs
    While Not .EOF
        Push OUsr, Trim(.Fields("ZPRF").Value)
        Push OCoAy, UsrCo_Rs_CoAy(Rs)
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub UsrCo_Ay_ZSC__Tst()
Dim U$(), CoAy()
UsrCo_Ay_ZSC U, CoAy
Stop
End Sub

Private Sub UsrPgm_Ay_ZSC__Tst()
Dim U$(), PgmStr$()
UsrPgm_Ay_ZSC U, PgmStr
Stop
End Sub

Private Sub UsrCo_Ay_Norm__Tst()
Dim U$(), CoAy()
UsrCo_Ay_ZSC U, CoAy
Dim OUsr$(), OCo$()
UsrCo_Ay_Norm OUsr, OCo, U, CoAy
Stop
End Sub

Private Sub UsrPgm_Ay_Norm__Tst()
Dim U$(), P$()
UsrPgm_Ay_ZSC U, P
Dim OUsr$(), OPgm$()
UsrPgm_Ay_Norm OUsr, OPgm, U, P
Stop
End Sub


Private Sub UsrCo_Ay_Norm(OUsr$(), OCo$(), U$(), CoAy())
Erase OUsr
Erase OCo
For J% = 0 To UB(U)
    UsrM$ = U(J)
    CoAyM = CoAy(J)
    For I% = 0 To UB(CoAyM)
        CoM$ = CoAyM(I)
        Push OUsr, UsrM
        Push OCo, CoM
    Next
Next
End Sub

Private Property Get UsrCo_Rs_CoAy(ZSC_Rs As Recordset) As String()
Dim O$(), F$
For J% = 1 To 40
    F = "ZCMP" & Format(J, "00")
    Push_NoBlank O, ZSC_Rs.Fields(F).Value
Next
UsrCo_Rs_CoAy = O
End Property

Private Property Get UsrPgm_Rs_PgmAy(ZSC_Rs As Recordset) As String()
Dim O$(), F$
For J% = 1 To 40
    F = "ZCMP" & Format(J, "00")
    Push_NoBlank O, ZSC_Rs.Fields(F).Value
Next
UsrPgm_Rs_PgmAy = O
End Property

Private Sub UsrWhs_Ay_Norm(OUsr$(), OWhs$(), Usr$(), WhsStr$())
Dim A$()
Erase OUsr
Erase OWhs
U& = UB(Usr)
'A = Whs_Col()
Dim O$()
For J = 0 To U
    A = Split_WhsStr(WhsStr(J))
    UsrM$ = Trim(Usr(J))
    For I = 0 To UB(A)
        WhsM = A(I)
        Push OUsr, UsrM
        Push OWhs, WhsM
    Next
Next
End Sub
Private Sub UsrPgm_Ay_Norm(OUsr$(), OPgm$(), Usr$(), PgmStr$())
Dim A$()
Erase OUsr
Erase OPgm
U& = UB(Usr)
'A = Pgm_Col()
Dim O$()
For J = 0 To U
    A = Split_PgmStr(PgmStr(J))
    UsrM$ = Trim(Usr(J))
    For I = 0 To UB(A)
        PgmM = A(I)
        Push OUsr, UsrM
        Push OPgm, PgmM
    Next
Next
End Sub


Private Sub Split_WhsStr__Tst()
Dim Act$()
Act = Split_WhsStr("123456")
Debug.Assert Sz(Act) = 3
Debug.Assert Act(0) = "12"
Debug.Assert Act(1) = "34"
Debug.Assert Act(2) = "56"
End Sub

Private Sub Split_PgmStr__Tst()
Dim Act$()
Act = Split_PgmStr("123456")
Debug.Assert Sz(Act) = 2
Debug.Assert Act(0) = "123"
Debug.Assert Act(1) = "456"
End Sub

Private Property Get Split_WhsStr(WhsStr$) As String()
Split_WhsStr = SplitNChrListStr(WhsStr, 2)
End Property

Private Property Get Split_PgmStr(PgmStr$) As String()
Split_PgmStr = SplitNChrListStr(PgmStr, 3)
End Property

Private Sub UsrTxTy_Ay_Norm(OUsr$(), OTxTy$(), Usr$(), TxTyStr$())
Dim A$()
Erase OUsr
Erase OTxTy
U& = UB(Usr)
Dim O$()
For J = 0 To U
    A = Split_TxTyStr(TxTyStr(J))
    UsrM$ = Usr(J)
    For I = 0 To UB(A)
        TxTyM = A(I)
        Push OUsr, UsrM
        Push OTxTy, TxTyM
    Next
Next
End Sub

Private Sub Split_TxTyStr__Tst()
Dim Act$()
Act = Split_TxTyStr("123456")
Debug.Assert Sz(Act) = 3
Debug.Assert Act(0) = "12"
Debug.Assert Act(1) = "34"
Debug.Assert Act(2) = "56"
End Sub

Private Property Get Split_TxTyStr(TxTyStr$) As String()
Split_TxTyStr = SplitNChrListStr(TxTyStr, 2)
End Property

Private Sub All__Tst()
Split_TxTyStr__Tst
End Sub
