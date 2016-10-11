Attribute VB_Name = "Ide_MdSrt"
Option Compare Database

Sub SrtMd(Md As CodeModule)
A$ = MdSortedBdy(Md)
ClrMdBdy Md
AppMdLines Md, A
End Sub
Sub SrtCurMd()
SrtMd CurMd
End Sub
Private Sub SrtMd__Tst()
SrtMd CurPjMd("Class2")
End Sub

Property Get MthAtrSrtKey$(P As MthAtr)
Modifier% = MthAtrSrtKey__Modifier(P.Modifier)
MthAtrSrtKey = FmtQQ("?|?", Modifier, P.Nm)
End Property
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

Private Sub MdSortedBdy__Tst()
BrwStr MdSortedBdy(CurPjMd("Ide_MdSrt"))
End Sub

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




