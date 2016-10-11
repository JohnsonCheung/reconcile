VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Property Get MdSortedBdy$(Md As CodeModule)
Dim A() As MthAtr
A = MdMthAtrAy(Md, InclBdy:=True)
U% = MthAtrUB(A)
K = NewStrAy(U)
For J% = 0 To U
    K(J) = MthAtrSrtKey(A(J))
Next
I = SrtAyToIdx(K)
O = NewStrAy(U)
For J% = 0 To U
    O(J) = vbCrLf & A(I(J)).Bdy
Next
MdSortedBdy = JoinLine(O)
End Property

Property Get MthAtrSrtKey$(P As MthAtr)
Modifier% = MthAtrSrtKey__Modifier(P.Modifier)
MthAtrSrtKey = FmtQQ("?|?", Modifier, P.Nm)
End Property

Sub SrtCurMd()
SrtMd CurMd
End Sub

Sub SrtMd(Md As CodeModule)
A$ = MdSortedBdy(Md)
ClrMdBdy Md
AppMdLines Md, A
End Sub

Private Sub MdSortedBdy__Tst()
BrwStr MdSortedBdy(CurMd)
End Sub

Private Property Get MthAtrSrtKey__Modifier%(Modifier$)
M$ = Modifier
If M = "Public" Or M = "" Then
    O% = 0
ElseIf M = "Friend" Then
    O = 1
ElseIf M = "Private" Then
    O = 2
Else
    Er "Invalid {Modifier}.  It should be ['' Public Friend Private]"
End If
MthAtrSrtKey__Modifier = O
End Property

Private Sub SrtMd__Tst()
SrtMd CurPjMd("Class2")
End Sub
