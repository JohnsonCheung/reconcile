Attribute VB_Name = "Dao_DbTMge"
Option Compare Database
Private A_Tar As DbT
Private A_TarD As Database
Private A_TarT$
Private A_FmTbl$()
Private Sub MgeDbT__Tst()
MgeDbT CurT("MgeAP"), "INP1_AP INP1_GL"
End Sub
Sub MgeDbT(Tar As DbT, TblNmLvs$)
'Create {Tar} by merging {TblNmLvs}
A_Tar = Tar
A_TarT = Tar.T
Set A_TarD = Tar.D
A_FmTbl = SplitLvs(TblNmLvs)
DrpDbTIfExist A_Tar
RunSqlAy SqlAy_DrpTmp, A_TarD, IgnoreEr:=True
A = SqlAy
For J = 0 To UB(A)
    A_TarD.Execute A(J)
Next
End Sub
Private Property Get SqlAy() As String()
SqlAy = AddAy(SqlAy_CrtTmp, SqlAy_CrtMge, SqlAy_Ins, SqlAy_DrpTmp)
End Property
Private Sub SqlAy_CrtTmp__Tst()
ZSetPrm
DmpAy SqlAy_CrtTmp
End Sub
Private Sub SqlAy_CrtMge__Tst()
ZSetPrm
DmpAy SqlAy_CrtMge
End Sub

Private Sub SqlAy_DrpTmp__Tst()
ZSetPrm
DmpAy SqlAy_DrpTmp
End Sub
Private Sub ZSetPrm()
A_Tar = CurT("MgeAP")
Set A_TarD = A_Tar.D
A_TarT = A_Tar.T
A_FmTbl = Split("INP1_AR INP1_GL")
End Sub
Private Property Get SqlAy_CrtTmp() As String()
O = NewStrAy(UFmTbl)
tARfLD = FmTblFld(0)
O(0) = FmtQQ("Select * into [#0] from [?]", A_FmTbl(0))
For J% = 1 To UFmTbl
    SelFld = MinusAy(FmTblFld(J), tARfLD)
    Sel$ = JoinComma(QuoteAy(SelFld, "[]"))
    O(J) = FmtQQ("Select ? into [#?] from [?] where false", Sel, J, A_FmTbl(J))
    PushAy tARfLD, SelFld
Next
SqlAy_CrtTmp = O
End Property
Private Property Get FmDbT(J%) As DbT
FmDbT = DbT(A_TarD, A_FmTbl(J))
End Property
Property Get FmTblFld(J%) As String()
FmTblFld = DbTFldNmAy(FmDbT(J))
End Property
Private Property Get TmpTblNmAy() As String()
O = NewStrAy(UFmTbl)
For J% = 0 To UFmTbl
    O(J) = "#" & J
Next
TmpTblNmAy = O
End Property
Private Property Get SqlAy_CrtMge() As String()
Fm$ = JoinComma(QuoteAy(TmpTblNmAy, "[]"))
Dim O$(0)
O(0) = FmtQQ("Select * into [?] from ?", A_TarT, Fm)
SqlAy_CrtMge = O
End Property
Private Property Get SqlAy_Ins() As String()
O = NewStrAy(UFmTbl)
For J = 0 To UFmTbl
    O(J) = FmtQQ("Insert into [?] select * from [?]", A_TarT, A_FmTbl(J))
Next
SqlAy_Ins = O
End Property
Private Sub SqlAy_Ins__Tst()
ZSetPrm
DmpAy SqlAy_Ins
End Sub

Private Property Get SqlAy_DrpTmp() As String()
O = NewStrAy(UFmTbl)
For J = 0 To UFmTbl
    O(J) = FmtQQ("Drop Table [#?]", J)
Next
SqlAy_DrpTmp = O
End Property
Private Property Get UFmTbl%()
UFmTbl = UB(A_FmTbl)
End Property




