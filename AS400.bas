Attribute VB_Name = "AS400"
Option Compare Database
Type PssQ
    D As Database
    Q As String
    CnnStr As String
End Type
Public Ip$
Public Lib$
Sub Brw400Tbl(Tbl$, Optional Lib$, Optional Ip$)

End Sub

Property Get PssQ(D As Database, Q$, CnnStr$) As PssQ
Dim O As PssQ
Set O.D = D
O.Q = Q
O.CnnStr = CnnStr
PssQ = O
End Property
Property Get DsnTp$()
Dim O$()
Push O, "[ODBC]"
Push O, "DRIVER=iSeries Access ODBC Driver"
Push O, "CONNTYPE=2"
Push O, "TRANSLATE=1"
Push O, "LANGUAGEID=ENU"
Push O, "SYSTEM=?"
Push O, "DbQ = ?"
DsnTp = JoinLine(O)
End Property
Private Property Get DsnPth$()
DsnPth = Env.MyDtaSrcPth
End Property
Property Get DsnFfn$(Lib$)
DsnFfn = DsnPth & DsnFn(Lib)
End Property
Private Property Get DsnFn$(Lib$)
DsnFn = Lib & "_ReadOnly.dsn"
End Property
Sub CrtAs400Dsn(Ip$, Lib$)
WrtStr FmtQQ(DsnTp, Ip, Lib), DsnPth & DsnFn(Lib)
End Sub
Property Get CurPssQ(Q$, CnnStr$) As PssQ
CurPssQ = PssQ(CurrentDb, Q, CnnStr)
End Property

Sub DnlB6Tbl(T$, Optional B6Pfx$ = "B6_", Optional Where$)
Dnl400Tbl CurT(B6Pfx & T), Godiva.Ip, Godiva.B6F, T
End Sub

Private Sub DnlAs4Tbl__Tst()
Dnl400Tbl CurT("IIC"), Godiva.Ip, Godiva.B6F, "IIC", "IID='IC'"
End Sub
Sub Dnl400Tbl(T As DbT, Ip$, Lib$, Optional Tbl$, Optional Where$)
Debug.Print T.T
Dim Q As PssQ
Q = PssQ(T.D, "TmpQry", DsnCnnStr(DsnFfn(Lib)))
CrtPssQ Q, SqlTblWhere(NzStr(Tbl, T.T), Where)
DrpDbTIfExist T
T.D.Execute FmtQQ("Select * into [?] from TmpQry", T.T)
End Sub
Property Get DsnCnnStr$(DsnFfn$)
DsnCnnStr = "ODBC;FILEDSN=" & DsnFfn
End Property
Private Sub TrmDbTTxtFld__Tst()
Dim Mge As DbT
Dim Tmp As DbT
Mge = CurT("MGE")
Tmp = CurT("#Tmp")
CpyDbT Mge, Tmp.T

TrmTxtFld Tmp
BrwDbT Tmp
Stop
ClsDbT Tmp
DrpDbT Tmp
End Sub

Sub TrmTxtFld(P As DbT)
TxtFld = DbTTxtFldAy(P)
B = AyMap_PI_StrAy(TxtFld, "Fmt", "[{0}]=Trim([{0}])")
A = Join(B, ",")
Sql$ = FmtQQ("update [?] set ?", P.T, A)
P.D.Execute Sql
End Sub
Sub TrmTxtFld1(T$)
TrmTxtFld CurT(T)
End Sub
Private Sub DbT400DteFldNmAy__Tst()
DmpAy DbT400DteFldNmAy(CurT("#ARB"))
End Sub

Property Get DbT400DteFldNmAy(P As DbT) As String()
Dim A() As DAO.Field
A = DbTFldAy(P, "*_DATE")
Dim O$()
For J% = 0 To UB(A)
    If A(J).Type = dbDecimal Then
        Push O, A(J).Name
    End If
Next
DbT400DteFldNmAy = O
End Property
Private Sub Cv400DteFld__Tst()
Cv400DteFld CurT("#ARB")
End Sub
Sub Cv400DteFld(P As DbT)
'Any field is Long and *_DATE, WILL assume to be YYYYMMDD, it will convert to DATE
A = DbT400DteFldNmAy(P)
For J% = 0 To UB(A)
    Cv400OneDteFld P, CStr(A(J))
Next
End Sub
Sub Cv400DteFld1(T$)
Cv400DteFld CurT(T)
End Sub
Private Sub Cv400OneDteFld__Tst()
CpyDbT CurT("INP_GLM"), "#GLM"
Cv400OneDteFld CurT("#GLM"), "DOC_DATE"
End Sub
Sub Cv400OneDteFld(P As DbT, FldNm$)
PP% = P.D.TableDefs(P.T).Fields(FldNm).OrdinalPosition
P.D.Execute FmtQQ("Alter table [?] add column TmpYY TEXT(4),TmpMM TEXT(2),TmpDD TEXT(2)", P.T)
P.D.Execute FmtQQ("Update [?] set TmpYY=Left(?,4),TmpMM=Mid(?,5,2),TmpDD=Mid(?,7,2)", P.T, FldNm, FldNm, FldNm)
P.D.Execute FmtQQ("alter table [?] drop column [?]", P.T, FldNm)
P.D.Execute FmtQQ("alter table [?] add column [?] date", P.T, FldNm)
P.D.Execute FmtQQ("Update [?] set [?]=DateSerial(TmpYY,TmpMM,TmpDD) where Not IsNull(TmpYY) and Not IsNull(TmpDD) and Not IsNull(TmpMM)", P.T, FldNm)
P.D.Execute FmtQQ("Alter table [?] drop column TmpYY,TmpMM,TmpDD", P.T)
P.D.TableDefs.Refresh
P.D.TableDefs(P.T).Fields(FldNm).OrdinalPosition = PP
End Sub


Private Sub Stru()
'AP
'APCMPY as COMPANY,
'AHSV04 as AC,
'APVNDR as VND_NO,
'VNDNAM as VND_NAME,
'PHDCYR as DOC_YEAR,
'PHDCPX as DOC_PFX,
'PHDCSQ as DOC_SEQ,
'APINV as INV_NO,
'APCOUT as NET,
'APHCUR As CUR
'
'AR
'RCOMP as COMPANY,
'AHSV04 as AC,
'RCUST as CUS_NO,
'CNME as CUS_NAME,
'ARDYR as DOC_YEAR,
'ARDPFX as DOC_PFX,
'ARDTYP as DOC_TYPE,
'ARDOCN as DOC_SEQ,
'RAR.RINVC as DOC_RINVC,
'RREF as REF_NO,
'-RREM*RCNVFC as NET,
'RCURR As CUR
'
'GL
'CRSG01 AS COMPANY,
'LHDRAM-LHCRAM AS NET ,
'CRSG04 AS AC, LHYEAR AS YEAR,
'LHPERD AS PERIOD,
'LHDREF AS GL_DOCREF,
'LHDDAT AS GL_DOCDATE,
'LHJRF1 AS GL_RF1,
'LHJRF2 AS GL_RF2,
'LHDATE AS GL_DATE,
'LHJNEN AS JNL_NO,
'LHJNLN AS JNL_LINNO,
'LHLSTS AS JNL_POST,
'LHLDES As JNL_LINDES
'
'IN
'ILI.LPROD as PROD,
'IIM.IDESC as PROD_DESC,
'IIM.IUMS as PROD_UOM,
'ILI.LWHS as WHS,
'ILI.LLOC as LOC,
'ILI.LLOT as LOT_NO,
'LOPB-LISSU+LADJU+LRCT as ONHAND,
'CMF.CFTLVL + CMF.CFPLVL as CST,
'-((LOPB-LISSU+LADJU+LRCT) * (CMF.CFTLVL + CMF.CFPLVL)) as NET,
'GAH.AHSV04 as AC,

End Sub
