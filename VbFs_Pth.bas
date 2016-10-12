Attribute VB_Name = "VbFs_Pth"
Option Compare Database

Property Get NzPth$(Pth$)
NzPth = CvPth(NzStr(Pth, CurDir))
End Property

Property Get CvPth$(Pth$)
If LastChr(Pth) <> "\" Then
    CvPth = Pth & "\"
Else
    CvPth = Pth
End If
End Property
Sub OpnPth(Pth$)
Shell "explorer """ & Pth & """", vbMaximizedFocus
End Sub
Property Get PthFfnAy(Pth$, Optional SPEC$ = "*.*") As String()
P$ = EnsureBackSlash(Pth)
PthFfnAy = AddAyPfx(PthFnAy(P, SPEC), P)
End Property
Property Get PthFnAy(Pth$, Optional SPEC$ = "*.*") As String()
Dim O$()
A$ = Dir(EnsureBackSlash(Pth) & SPEC)
While A <> ""
    If A <> ".." And A <> "." Then Push O, A
    A = Dir
Wend
PthFnAy = O
End Property

Property Get EnsureBackSlash$(Pth$)
If LastChr(Pth) = "\" Then
    EnsureBackSlash = Pth
Else
    EnsureBackSlash = Pth & "\"
End If
End Property

Sub ClrPth(Pth$, Optional IgnoreErr As Boolean)
Dim FnAy$(): FnAy = PthFnAy(Pth)
DltFnAy FnAy, Pth, IgnoreErr
End Sub
Sub CrtPthIfNotExist(Pth$)
If IsPth(Pth) Then Exit Sub
CrtPthEachSeg Pth
End Sub
Private Sub ZZZ_IsPth()
Debug.Assert IsPth("C:")
Debug.Assert IsPth("C:\")
Debug.Assert IsPth("C:\Temp")
Debug.Assert IsPth("C:\Temp\")
Debug.Assert IsPth("C:\TEMP\")
Debug.Assert IsPth("C:\temp\")
Debug.Assert IsPth("C:\lsjdf\") = False
End Sub
Property Get IsPth(Pth$) As Boolean
IsPth = Fso.FolderExists(Pth)
End Property
Private Sub ZZZ_CrtPthEachSeg()
CrtPthEachSeg "C:\temp\a\b\c\d\e\"
End Sub
Private Sub ZZZ_AssertBackSlash()
AssertBackSlash "A:\"
AssertBackSlash "A:"
End Sub
Sub AssertBackSlash(Pth$)
If Right(Pth, 1) <> "\" Then Er "Given {Pth} must end with \", Pth
End Sub
Sub CrtPthEachSeg(Pth$)
AssertBackSlash Pth
Dim A$(): A = Split(Pth, "\")
B$ = A(0)
For J% = 1 To UB(A) - 1
    B = B & "\" & A(J)
    If Not IsPth(B) Then MkDir B
Next
End Sub
