Attribute VB_Name = "Ide_PjOpr"
Option Compare Database

Sub RmvTmpMd(Pj As VBProject)
Dim A$()
A = PjTmpMdNmAy(Pj)
For J% = 0 To UB(A)
    RmvMd PjMd(Pj, A(J))
Next
End Sub

Sub RmvTmpMdInCurPj()
RmvTmpMd CurPj
End Sub
Sub ExpPj(Pj As VBProject)
Pth$ = PjSrcPth(Pj)
CrtPthIfNotExist Pth
ClrPth Pth
Dim Md As CodeModule
For Each I In PjMdAy(Pj)
    Set Md = I
    Ffn$ = Pth & MdSrcFn(Md)
    MdCmp(Md).Export Ffn
Next
End Sub
Private Sub AAA()
CommitPj CurPj
End Sub
Private Sub ZZZ_ExpPj()
ExpPj CurPj
End Sub
Sub CommitPj(Pj As VBProject, Optional Msg$ = "Commit")
ExpPj Pj
P$ = PjSrcPth(Pj)
Dim A$(): A = CommitPj__BatchContentAy(P, Msg)
B$ = PjPth(Pj) & "Commit.bat"
WrtAy A, B
Shell "C:\Users\cheungj\AppData\Local\Programs\Git\git-cmd.exe " & B
End Sub
Private Property Get CommitPj__BatchContentAy(PjSrcPth$, Msg$) As String()
Dim O$()
Drive = Left(PjSrcPth, 2)
CD$ = "CD " & PjSrcPth
GitInit_IfNeeded$ = "IF NOT EXIST .git git init"
CheckIn$ = ""
Push O, Drive
Push O, CD
Push O, GitInit_IfNeeded
Push O, "git add *"
Push O, "git commit -m """ & Msg & """"
Push O, "pause"
CommitPj__BatchContentAy = O
End Property
Sub CommitCurPj()
CommitPj CurPj
End Sub
