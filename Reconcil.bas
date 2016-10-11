Attribute VB_Name = "Reconcil"
Option Compare Database
Public Godiva As New GodivaConst
Private Const ZZCrtTmp_and_Trm_and_CvDte_$ = "ZZCrtTmp_and_Trm_and_CvDte"

Private Sub NRec_Dmp_B6RecInpTbl()
'Dmp_NRec_Lvs Join(DbTblNmAy("Inp_*"), " ")
End Sub

Private Sub GenReconcile()
DoCmd.SetWarnings False
Log "Start"
Gen_AP
Gen_AR
Gen_IN
Log "End"
End Sub

Private Property Get ZTmpQry() As PssQ
ZTmpQry = CurPssQ("TmpQry", Godiva.B6FDsnCnnStr)
End Property

Property Get NRec&(Tbl$)
CrtPssQ ZTmpQry, FmtQQ("Select Count(*) from ?", Tbl)
B6NRec = SqlLng("Select * from TmpQry")
End Property

Private Sub NRec_B6RecTbl()
NRec_Dmp "ITH ILI RAR GSB APL AML"
End Sub

Private Sub NRec_Dmp(TblLvs$)
A = SplitLvs(TblLvs)
'NRec% = B6NRec_Lvs(TblLvs)
For J% = 0 To UB(A)
'Debug.Print A(J), NRec(J)
Next
End Sub

Property Get NRec_B6Lvs(TblLvs$) As Long()
A = SplitLvs(TblLvs)
Dim O&()
U& = UB(A)
ReSzAy O, U
For J% = 0 To U
'    O(J) = B6NRec(CStr(A(J)))
Next
NRec_B6Lvs = O
End Property

Sub ZLog(Msg$)
DoCmd.RunSql FmtQQ("Insert into Log (Evt) Values('?')", Msg)
End Sub

Private Sub RecCount()
'25/9/2016 15:54:13
'25/9/2016 16:01:17
'ITH 165656
'ILI 18709
'RAR 10105
'GSB 8991
'APL 22335
'AML 22813
'INP_INM 165551
'INP_INB 2702
'INP_ARM 10105
'INP_ARB 402
'INP_GLM 48490
'INP_GLB 8991
'INP_APB 718
'INP_APMINV 22335
'INP_APMPAY 22813
End Sub

Private Sub ZTrmTxt(Tmp$)
TrmTxtFld1 "#" & Tmp
End Sub
Sub ZCrtInp(B6RecItmNm$)
A$ = B6RecItmNm
Log FmtQQ("Start Download qryB6Rec_? Into INP_?", A, A)
DnlDbQ CurQ("qryB6REC_" & A), "INP_" & A
End Sub

Private Sub ZAddCol_K_SRC(Itm$)
A$ = ""
RunSql FmtQQ("Alter Table [#?] add column K_SRC TEXT(8)", Itm)
RunSql FmtQQ("Update [#?] Set K_SRC='?'", Itm, A)
End Sub
Private Sub ZAddCol_K_SRC_Dnl_TmpTbl_AddCol_K_SRC()

RunSql "Alter Table [#INB] add column K_SRC TEXT(8)"
RunSql "Alter Table [#INM] add column K_SRC TEXT(8)"
RunSql "Alter Table [#GLM] add column K_SRC TEXT(8)"
RunSql "Alter Table [#GLB] add column K_SRC TEXT(8)"
RunSql "Alter Table [#ARM] add column K_SRC TEXT(8)"
RunSql "Alter Table [#ARB] add column K_SRC TEXT(8)"
RunSql "Alter Table [#APMPAY] add column K_SRC TEXT(8)"
RunSql "Alter Table [#APMINV] add column K_SRC TEXT(8)"
RunSql "Alter Table [#APB] add column K_SRC TEXT(8)"
RunSql "Update [#INB] Set K_SRC='INBAL'"
RunSql "Update [#INM] Set K_SRC='INMTD'"
RunSql "Update [#ARB] Set K_SRC='INBAL'"
RunSql "Update [#ARM] Set K_SRC='INMTD'"
RunSql "Update [#GLB] Set K_SRC='*GLBAL'"
RunSql "Update [#GLM] Set K_SRC='*GLMTD'"
RunSql "Update [#APB] Set K_SRC='APBAL'"
RunSql "Update [#APMINV] Set K_SRC='APMTDINV'"
RunSql "Update [#APMPAY] Set K_SRC='APMTDPAY'"
End Sub

Private Sub Mge_APM()
CpyDbT CurT("#GLM"), "#GLMAP", "SUBLEDGER='AP'"
'----------------
RunSql "Alter Table [#GLMAP] add column VND_NO LONG, K_DOC_NO LONG, K_NET_USD currency,K_YEAR INT, K_PERIOD INT;"
RunSql "Update [#GLMAP] set VND_NO = CLNG(DOC_RF1) where DOC_RF1<>''"
RunSql "Update [#GLMAP] set K_DOC_NO = CLNG(DOC_REFNO) where DOC_REFNO<>''"
RunSql "Update [#GLMAP] set K_NET_USD = CCUR(NET_USD), K_YEAR=YEAR, K_PERIOD =PERIOD"
RunSql "Alter Table [#GLMAP] drop column YEAR,PERIOD,DOC_RF1,DOC_REFNO"

'----------------
CpyDbT CurT("#APMINV"), "#APMINV1", "YEAR(GL_DATE)=2016"
RunSql "Alter Table [#APMINV1] add column K_DOC_NO LONG, K_NET_USD currency,  K_YEAR INT, K_PERIOD INT"
RunSql "Update [#APMINV1] set " & _
                          "K_DOC_NO = DOC_NO," & _
                          "K_NET_USD = ROUND(INVAMT_LINE_USD,2)," & _
                          "K_YEAR = YEAR(GL_DATE)," & _
                          "K_PERIOD = MONTH(GL_DATE)"
RunSql "Alter Table [#APMINV1] DROP column DOC_NO"
'----------------
CpyDbT CurT("#APMPAY"), "#APMPAY1", "YEAR(PAY_DATE)=2016"
RunSql "Alter Table [#APMPAY1] add column K_DOC_NO LONG, K_NET_USD currency, K_YEAR INT, K_PERIOD INT"
RunSql "Update [#APMPAY1] set " & _
                          "K_DOC_NO = DOC_NO," & _
                          "K_NET_USD = ROUND(-AP_AMT_USD,2)," & _
                          "K_YEAR = YEAR(PAY_DATE)," & _
                          "K_PERIOD = MONTH(PAY_DATE)"
RunSql "Alter Table [#APMPAY1] DROP column DOC_NO"
'----------------
MgeDbT CurT("#Mge_APM"), "#GLMAP #APMINV1 #APMPAY1"
'========================
RunSql "Alter Table [#Mge_APM] add column GL_ACDESC TEXT(30)"
RunSql "Update [#Mge_APM] a inner join INP_ZACDESC b on a.GL_AC = b.GL_AC set a.GL_ACDESC = b.GL_ACDESC"

RunSql "Alter Table [#Mge_APM] add column AC_NO text(50),AC_NAME text(50)"
RunSql "Update [#Mge_APM] set AC_NO = GL_AC & ': ' & GL_ACDESC, AC_NAME = GL_ACDESC & ' (' & GL_AC & ')'"
RunSql "Alter Table [#Mge_APM] drop column GL_AC, GL_ACDESC"

'-------------------------------
RunSql "Update [#Mge_APM] a inner join INP_ZVNDNAME b on a.VND_NO = b.VND_NO set a.VND_NAME = b.VND_NAME"

RunSql "Alter Table [#Mge_APM] add column K_VND_NO text(50),K_VND_NAME text(50)"
RunSql "Update [#Mge_APM] set K_VND_NO = VND_NO & ': ' & VND_NAME, K_VND_NAME = VND_NAME & ' (' & VND_NO & ')'"
RunSql "Alter Table [#Mge_APM] drop column VND_NO, VND_NAME"

CurrentDb.TableDefs("#Mge_APM").Fields("K_DOC_NO").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_APM").Fields("K_VND_NO").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_APM").Fields("K_VND_NAME").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_APM").Fields("AC_NO").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_APM").Fields("AC_NAME").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_APM").Fields("K_NET_USD").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_APM").Fields("K_YEAR").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_APM").Fields("K_PERIOD").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_APM").Fields("K_SRC").OrdinalPosition = 1
CurrentDb.TableDefs.Refresh

End Sub

Private Sub Mge_APB()
'----------------
CpyFld CurT("#APMGL"), "#APMGL1", ""

RunSql "Alter Table [#APMGL1] add column VND_NO LONG, K_DOCSEQNO LONG, K_COCUR_NET currency, K_SRC TEXT(6);"
RunSql "Update [#APMGL1] set VND_NO = CLNG(DOC_RF1) where DOC_RF1<>''"
RunSql "Update [#APMGL1] set K_DOCSEQNO = CLNG(DOC_REFNO) where DOC_REFNO<>''"
RunSql "Update [#APMGL1] set K_COCUR_NET = CCUR(COCUR_NET),K_SRC = '*GL'"

'----------------
CpyFld CurT("#APB"), "#APB1", _
    "DOC_SEQNO             AS K_DOC_SEQNO," & _
    "ROUND(COCUR_INVBAL,2) AS K_COCUR_NET," & _
    "'AP BAL'              AS K_SRC,"
'----------------
MgeDbT CurT("#Mge_APB"), "#APMGL1 #APB1"
'========================
RunSql "Drop Table [#APMGL1]"
RunSql "Drop Table [#APB1]"
'========================
'AddDesFld CurT("#Mge_APB"), "INP_ZACDESC", "GL_AC", "GL_ACDESC"
'AddDesFld CurT("#Mge_APB"), "INP_ZVNDNAME", "VND_NO", "VND_NAME"
'========================
'CmbFld CurT("#Mge_APB"), Array( _
'    Array(50, "K_GL_AC     = GL_AC & ': ' & GL_ACDESC"), _
'    Array(50, "K_GL_ACDESC = GL_AC & ': ' & GL_ACDESC"), _
'    Array(50, "K_VNDNAME   = VND_NAME & ' (' & VND_NO & ')'"))

RunSql "Alter table [#Mge_APB] add column K_GL_AC TEXT(50), K_GL_ACDESC TEXT(50), K_VNDNO TEXT(50), K_VNDNAME TEXT(50)"
RunSql "update [#Mge_APB] set K_GL_AC    = GL_AC & ': ' & GL_ACDESC," & _
                           " K_GL_ACDESC  = GL_ACDESC & ' (' & GL_AC & ')'," & _
                           " K_VNDNO   = VND_NO & ': ' & VND_NAME ," & _
                           " K_VNDNAME = VND_NAME & ' (' & VND_NO & ')'"
                           
RunSql "Alter table [#Mge_APB] drop column GL_AC, GL_ACDESC, VND_NO, VND_NAME"
'========================
'ReSeqFld CurT("#Mge_APB"), ""
End Sub

Private Sub Mge_ARM()
CpyDbT CurT("#GLM"), "#GLMAR", "SUBLEDGER='AR'"

End Sub
Private Sub Mge_ARB()
CpyDbT CurT("#GLM"), "#GLMAR", "SUBLEDGER='AR'"
End Sub

Private Sub Mge_INB()
CpyDbT CurT("#GLM"), "#GLMIN", "SUBLEDGER='IN'"
'----------------
RunSql "Alter Table [#GLMIN] add column SKU TEXT(15), K_NET_USD currency;"
RunSql "Update [#GLMIN] set SKU = TRIM(GL_A3) WHERE JNL_HDR_DESC='INVENTORY RECEIPT PURCH ORDER' AND JNL_SRC='IN'"
RunSql "Update [#GLMIN] set SKU = TRIM(DOC_RF1) WHERE JNL_HDR_DESC<>'INVENTORY RECEIPT PURCH ORDER' AND JNL_SRC='IN'"

RunSql "Update [#GLMIN] set K_NET_USD = CCUR(NET_USD)"
'----------------
CpyDbT CurT("#INB"), "#INB1"
RunSql "Alter Table [#INB1] add column K_NET_USD currency"
RunSql "Update [#INB1] set K_NET_USD = ROUND(-USD,2)"
'----------------
MgeDbT CurT("#Mge_INB"), "#GLMIN #INB1"
'========================
RunSql "Alter Table [#Mge_INB] add column GL_ACDESC TEXT(30), SKU_DESC TEXT(30)"
RunSql "Update [#Mge_INB] a inner join INP_ZACDESC b on a.GL_AC = b.GL_AC set a.GL_ACDESC = b.GL_ACDESC"
RunSql "Update [#Mge_INB] a inner join INP_ZSKUDESC b on a.SKU = b.SKU set a.SKU_DESC = b.SKU_DESC"
'========================
RunSql "Alter table [#Mge_INB] add column K_AC_NO TEXT(50),K_AC_NAME TEXT(50),K_SKU TEXT(50)"
RunSql "update [#Mge_INB] set K_AC_NO    = GL_AC & ': ' & GL_ACDESC," & _
                           " K_AC_NAME  = GL_ACDESC & ' (' & GL_AC & ')'," & _
                           " K_SKU = SKU & ' : ' & SKU_DESC"
                           
RunSql "Alter table [#Mge_INB] drop column GL_AC, GL_ACDESC, SKU, SKU_DESC"
'========================
CurrentDb.TableDefs("#Mge_INB").Fields("K_SKU").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_INB").Fields("K_AC_NO").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_INB").Fields("K_AC_NAME").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_INB").Fields("K_NET_USD").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_INB").Fields("K_SRC").OrdinalPosition = 1
CurrentDb.TableDefs.Refresh
End Sub
Private Sub Mge_INM()
CpyDbT CurT("#GLM"), "#GLMIN", "SUBLEDGER='IN'"
'----------------
RunSql "Alter Table [#GLMIN] add column SKU TEXT(15), K_NET_USD currency, TX_TYPE TEXT(2);"
RunSql "Update [#GLMIN] set SKU = TRIM(GL_A3),   TX_TYPE = TRIM(LEFT(GL_A4,2)) WHERE JNL_HDR_DESC='INVENTORY RECEIPT PURCH ORDER' AND JNL_SRC='IN'"
RunSql "Update [#GLMIN] set SKU = TRIM(DOC_RF1), TX_TYPE = TRIM(LEFT(GL_A3,2)) WHERE JNL_HDR_DESC<>'INVENTORY RECEIPT PURCH ORDER' AND JNL_SRC='IN'"
RunSql "Update [#GLMIN] set K_NET_USD = CCUR(NET_USD)"
'----------------
CpyDbT CurT("#INM"), "#INM1"
RunSql "Alter Table [#INM1] add column K_NET_USD currency, YEAR INT, PERIOD BYTE, DOC_DATE DATE"
RunSql "Update [#INM1] set K_NET_USD = ROUND(-USD_ITH,2) where NOT TX_TYPE IN ('C','#')"
RunSql "Update [#INM1] set YEAR = YEAR(GL_DATE)"
RunSql "Update [#INM1] set PERIOD = MONTH(GL_DATE)"
RunSql "Update [#INM1] set DOC_DATE = GL_DATE"
'----------------
MgeDbT CurT("#Mge_INM"), "#GLMIN #INM1"
'========================
RunSql "Alter Table [#Mge_INM] add column GL_ACDESC TEXT(30), SKU_DESC TEXT(30)"
RunSql "Update [#Mge_INM] a inner join INP_ZACDESC b on a.GL_AC = b.GL_AC set a.GL_ACDESC = b.GL_ACDESC"
RunSql "Update [#Mge_INM] a inner join INP_ZSKUDESC b on a.SKU = b.SKU set a.SKU_DESC = b.SKU_DESC"
'========================
RunSql "Alter table [#Mge_INM] add column K_AC_NO TEXT(50),K_AC_NAME TEXT(50),K_SKU TEXT(50)"
RunSql "update [#Mge_INM] set K_AC_NO    = GL_AC & ': ' & GL_ACDESC," & _
                           " K_AC_NAME  = GL_ACDESC & ' (' & GL_AC & ')'," & _
                           " K_SKU = SKU & ' : ' & SKU_DESC"
                           
RunSql "Alter table [#Mge_INM] drop column GL_AC, GL_ACDESC, SKU, SKU_DESC"
'========================
CurrentDb.TableDefs("#Mge_INM").Fields("K_SKU").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_INM").Fields("K_AC_NO").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_INM").Fields("K_AC_NAME").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_INM").Fields("K_NET_USD").OrdinalPosition = 1
CurrentDb.TableDefs("#Mge_INM").Fields("K_SRC").OrdinalPosition = 1
CurrentDb.TableDefs.Refresh
End Sub

Private Sub ZAddCol_JNL_SRC(Ledger$)
RunSql FmtQQ("Alter Table [#GLM?] add column JNL_SRC TEXT(2)", Ledger)
RunSql FmtQQ("update [#GLM?] set JNL_SRC=LEFT(JNL_NO,2)", Ledger)
End Sub

Private Sub DnlB6_AP()
ZRun_Itm "ZZDnl", "AP[BGL MGL B MINV MPAY]"
End Sub

Private Sub ZRun_Itm(FnNm$, Tp$)
Dim A$()
A = ExpandToAy(Tp)
For J% = 0 To UB(A)
    Run FnNm, A(J)
Next
End Sub
Private Sub ZZDnl__Tst()
ZZDnl "APMGL"
End Sub
Sub ZZDnl(Itm$)
DoCmd.SetWarnings False
DrpDbT CurT("INP_" & Itm)
RunSql FmtQQ("Select * into [INP_?] from [qryB6Rec_?]", Itm, Itm)
End Sub
Private Sub Tmp_AP()
A$ = ZZCrtTmp_and_Trm_and_CvDte_
ZRun_Itm A, "AP[BGL MGL B MINV MPAY]"
End Sub
Private Sub Tmp_APMGL()
A$ = ZZCrtTmp_and_Trm_and_CvDte_
ZRun_Itm A, "AP[MGL]"

End Sub

Private Sub Tmp_AR()
A$ = ZZCrtTmp_and_Trm_and_CvDte_
ZRun_Itm A, "AR[BGL MGL B]"
End Sub

Private Sub DnlB6_AR()
A$ = ZZCrtTmp_and_Trm_and_CvDte_
ZRun_Itm A, "AR[BGL MGL B M]"
End Sub

Private Sub DnlB6_IN()
ZRun_Itm "ZZDnl", "IN[BGL MGL B M]"
End Sub

Private Sub Gen_AP()
DnlB6_AP
Tmp_AP
Mge_APB
Mge_APM
End Sub

Private Sub Gen_AR()
DnlB6_AR
Mge_ARB
Mge_ARM
End Sub

Private Sub Gen_IN()
DnlB6_IN
Mge_INB
Mge_INM
End Sub
Private Sub ZCvDte(Tmp$)
Cv400DteFld1 "#" & Tmp
End Sub
Sub ZZCrtTmp_and_Trm_and_CvDte(Tmp$)
ZCrtTmp Tmp
ZTrmTxt Tmp
ZCvDte Tmp
End Sub

Private Sub ZCrtTmp(Tmp$)
CpyDbT CurT("INP_" & Tmp), "#" & Tmp
End Sub


