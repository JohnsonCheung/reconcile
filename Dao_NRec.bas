Attribute VB_Name = "Dao_NRec"
Option Compare Database
Sub DmpNRec_Lvs(TblLvs$)
A = SplitLvs(TblLvs)
For J% = 0 To UB(A)
    Debug.Print A(J), TblNRec(CStr(A(J)))
Next
End Sub
Sub DmpNRec(T$)
Debug.Print T$, TblNRec(T)
End Sub
Property Get TblNRec&(T$)
TblNRec = SqlLng(FmtQQ("Select count(*) from [?]", T))
End Property
