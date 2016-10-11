Attribute VB_Name = "Xls_Pair"
Option Compare Database
Type R1R2
    R1 As Long
    R2 As Long
End Type
Type C1C2
    C1 As Integer
    C2 As Integer
End Type
Property Get R1R2UB&(R() As R1R2)
R1R2UB = R1R2Sz(R) - 1
End Property
Property Get R1R2Sz&(R() As R1R2)
On Error Resume Next
R1R2Sz = UBound(R) + 1
End Property
Sub PushR1R2(OAy() As R1R2, R1, R2)
N& = R1R2Sz(OAy)
ReDim Preserve OAy(N)
Dim M As R1R2
M.R1 = R1
M.R2 = R2
OAy(N) = M
End Sub
Property Get C1C2UB&(R() As C1C2)
C1C2UB = C1C2Sz(R)
End Property
Property Get C1C2Sz&(R() As C1C2)
On Error Resume Next
C1C2Sz = UBound(R) + 1
End Property
Sub PushC1C2(OAy() As C1C2, C1, C2)
N& = C1C2Sz(OAy)
ReDim Preserve OAy(N)
Dim M As C1C2
M.C1 = C1
M.C2 = C2
OAy(N) = M
End Sub
