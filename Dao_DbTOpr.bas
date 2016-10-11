Attribute VB_Name = "Dao_DbTOpr"
Option Compare Database

Sub BrwDbT(P As DbT)
If IsDbTCurDb(P) Then DoCmd.OpenTable P.T: Exit Sub
Dim A As New Application
A.OpenCurrentDatabase P.D.Name
A.DoCmd.OpenTable P.T
Stop
End Sub
Private Sub BrwDbT__Tst()
BrwDbT CurT("MGE")
End Sub

Sub ClsDbT(P As DbT)
If IsDbTOpn(P) Then
    DoCmd.Close AcObjectType.acTable, P.T
End If
End Sub
Sub BrwCurT(T$)
BrwDbT CurT(T)
End Sub
Sub ClsCurT(T$)
ClsDbT CurT(T)
End Sub
Private Sub CpyFld__Tst()
DrpDbTIfExist CurT("#APB_X")
CpyFld CurT("#APB"), "#APB_X", "K_INV_NO=INV_NO; K_NET = COCUR_INVBAL"
End Sub
Sub CpyFld(P As DbT, TarTbl$, Asist$)
FmTbl = P.T
Sql$ = FmtQQ("Select  ?,a.* into [?] from [?] A", AsList, TarTbl, FmTbl)
P.D.Execute Sql
End Sub
