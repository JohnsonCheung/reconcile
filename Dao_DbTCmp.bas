Attribute VB_Name = "Dao_DbTCmp"
Option Compare Database
Private A_Db As Database
Private A_T1 As DbT
Private A_T2 As DbT
Private A_1$
Private A_2$
Private A_Nm1$
Private A_Nm2$

Private A_Key$()

Function CmpDbT(A As DbT, B$, Optional NKey%, Optional KeyLvs$, Optional ByVal T1Nm$, Optional ByVal T2Nm$) As String()
If NKey = 0 And KyFldLvs = "" Then Er "{NKey} or {KeyLvs} must be given", NKey, KeyLvs
A_T1 = A
A_T2 = DbT(A_Db, B)
Set A_Db = A.D
A_1 = A.T
A_2 = B
A_Key = X_Key(NKey, KeyLvs)
A_Nm1$ = NzStr(T1Nm, A_1)
A_Nm2$ = NzStr(T2Nm, A_2)
Dim O$()
PushAy O, Cmp_NFld
PushAy O, Cmp_KeyFld
If Sz(O) > 0 Then CmpDbT = O: Exit Function
PushAy O, Cmp_NRec
PushAy O, Cmp_MissInA
PushAy O, Cmp_MissInB
PushAy O, Cmp_DiffRec
If Sz(O) = 0 Then Exit Function
CmpDbT = InsAy(O, FmtQQ("Table [?] and [?] are diff", A_Nm1, A_Nm2))
End Function

Private Property Get X_Key(NKey%, KeyLvs$) As String()
If NKey > 0 Then
    A_Key = DbTFldNmAy(A_T1, NKey)
Else
    A_Key = SplitLvs(KeyLvs)
End If
End Property

Private Property Get Cmp_NFld() As String()
Dim N1%, N2%
N1 = DbTNFld(A_T1)
N2 = DbTNFld(A_T2)
If N1 <> N2 Then Push O, FmtQQ("Table [?] and [?] has different # of fields [?] [?]", A_Nm1, A_Nm2, N1, N2)
End Property

Private Property Get Cmp_KeyFld() As String()
Dim F1$(), F2$(), A$()
F1 = DbTFldNmAy(A_T1)
F2 = DbTFldNmAy(A_T2)

A = MinusAy(A_Key, F1)
If Sz(A) > 0 Then
End If

A = MinusAy(A_Key, F2)
If Sz(A) > 0 Then
End If
Cmp_KeyFld = O
'If IsContainAy(
End Property

Private Property Get Cmp_NRec() As String()

End Property

Private Property Get Cmp_MissInA() As String()

End Property

Private Property Get Cmp_MissInB() As String()

End Property
Private Property Get Cmp_DiffRec() As String()

End Property

