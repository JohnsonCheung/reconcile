VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MdInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Md As CodeModule
Property Get BdyAy() As String()
BdyAy = SplitLines(Bdy)
End Property

Property Get Bdy$()
Cnt% = Md.CountOfLines - Md.CountOfDeclarationLines
If Cnt = 0 Then Exit Property
Bdy = Md.Lines(Md.CountOfDeclarationLines + 1, Cnt)
End Property

Property Get ExpCxt$()
Ft$ = TmpFt
Md.Parent.Export Ft
ExpCxt = FtStr(Ft, KillFt:=True)
End Property
Property Get Pj() As VBProject
Set Pj = Md.Parent.Collection.Parent
End Property
Property Get Cxt$()
N% = Md.CountOfLines
If N = 0 Then Exit Property
Cxt = Md.Lines(1, Md.CountOfLines)
End Property
Property Get Dcl$()
N% = Md.CountOfDeclarationLines
If N = 0 Then Exit Property
Dcl = Md.Lines(1, N)
End Property

Property Get CxtAy() As String()
CxtAy = SplitLines(MdCxt(Md))
End Property
Private Sub DclAy__Tst()
BrwAy CurMdInfo.DclAy
End Sub

Property Get DclAy() As String()
DclAy = SplitLines(Dcl)
End Property

Private Sub BdyAy__Tst()
BrwAy CurMdInf.BdyAy
End Sub
Property Get Cmp() As VBComponent
Set Cmp = Md.Parent
End Property

Property Get Nm$()
Nm = Md.Parent.Name
End Property

Private Sub Nm__Tst()
Debug.Assert CurMdInf.Nm = "Ide_Md"
End Sub

