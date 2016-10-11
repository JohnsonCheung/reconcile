Attribute VB_Name = "Ide_MthOpr"
Option Compare Database

Sub CrtTstMth(M As Mth)
MthNm$ = M.Nm & "__Tst"
If IsSfx(M.Nm, "__Tst") Then GoTo X
If IsMth(Mth(M.Md, MthNm)) Then GoTo X
L$ = "Private Sub " & MthNm & "()" & vbCrLf & "End Sub"
AppMdLines M.Md, L
X: ShwMth Mth(M.Md, MthNm)
End Sub

Sub RmvMth(Md As CodeModule, MthNm$)
N% = 0
Do
    If N > 3 Then Stop
    Beg% = MthBegLno(Md, MthNm)
    If Beg = 0 Then Exit Sub
    E% = MthEndLno(Md, Beg)
    Cnt% = E - Beg + 1
    If Cnt = 0 Then Exit Sub
    Md.DeleteLines Beg, Cnt
    N = N + 1
Loop
End Sub

