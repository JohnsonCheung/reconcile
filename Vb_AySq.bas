Attribute VB_Name = "Vb_AySq"
Option Compare Database
Type SqSz
    NR As Long
    NC As Long
End Type
Property Get MgeSqAv(SqAv())
O = SqAv(0)
For J% = 1 To UB(SqAv)
    O = AddTwoSq(O, SqAv(J))
Next
MgeSqAv = 0
End Property
Property Get AyOneColSq(Ay)
N& = Sz(Ay)
If N = 0 Then Exit Property
Dim O
O = Ay
Erase O
ReDim O(1 To N, 1 To 1)
For J& = 1 To N
    O(J, 1) = Ay(J - 1)
Next
AyOneColSq = O
End Property
Property Get LookupTwoColSq(TwoColSq, Key)
'Assume D2 is a table of 2 column with 1st column as key and 2nd column as value
For R& = LBound(TwoColSq, 1) To UBound(TwoColSq, 1)
    If TwoColSq(R, 1) = Key Then LookupSq = TwoColSq(R, 2): Exit Property
Next
End Property
Private Sub AddTwoSq__Tst()
Dim A()
    ReDim A(1 To 2, 1 To 3)
    For R% = 1 To 2
        For C% = 1 To 3
            A(R, C) = R * 100 + C
        Next
    Next
Dim B()
    ReDim B(1 To 10, 1 To 3)
    For R% = 1 To 10
        For C% = 1 To 3
            B(R, C) = R * 1000 + C
        Next
    Next
BrwSq AddTwoSq(A, B)
End Sub
Private Sub AddSq__Tst()
Dim A(), B()
A = SampleSq1
B = SampleSq2
Act = AddSq(A, B, B, A)
BrwSq Act
End Sub
Property Get SampleSq1()
Dim A()
    ReDim A(1 To 2, 1 To 3)
    For R% = 1 To 2
        For C% = 1 To 3
            A(R, C) = R * 100 + C
        Next
    Next
SampleSq1 = A
End Property
Property Get SampleSq2()
Dim B()
    ReDim B(1 To 10, 1 To 3)
    For R% = 1 To 10
        For C% = 1 To 3
            B(R, C) = R * 1000 + C
        Next
    Next
SampleSq2 = B
End Property
Property Get AddTwoSq(Sq1, Sq2)
NC& = UBound(Sq1, 2)
If NC <> UBound(Sq2, 2) Then Er FmtQQ("Cannot PushSq due Sq1-Col <> Sq2-Col. Sq1-Sz=[?] Sq2-Sz=[?]", SqSzStr(Sq1), SqSzStr(Sq2))
NR1& = UBound(Sq1, 1)
NR2& = UBound(Sq2, 1)
Dim O
O = Sq1
ReDim O(1 To NR1 + NR2, 1 To NC)
For R& = 1 To NR1
    For C& = 1 To NC
        O(R, C) = Sq1(R, C)
    Next
Next
For R& = 1 To NR2
    For C& = 1 To NC
        O(R + NR1, C) = Sq2(R, C)
    Next
Next
AddTwoSq = O
End Property

Property Get AddSq(Sq, ParamArray SqAp())
Dim SqAv()
SqAv = SqAp
AddSq = AddSqAv(Sq, SqAv)
End Property
Private Sub BrwSq__Tst()
BrwSq SampleSq1
End Sub
Sub BrwSqPair(P As SqPair, Optional Pfx$ = "SqPair")
BrwAy SqPairStrAy(P), Pfx
End Sub
Property Get SqPairStrAy(P As SqPair) As String()
SqPairStrAy = AddAy(SqStrAy(P.Sq1), SqStrAy(P.Sq2))
End Property
Sub BrwSq(Sq, Optional Pfx$ = "BrwSq")
BrwAy SqStrAy(Sq), Pfx
End Sub
Property Get SqStrAy(Sq, Optional Sep$ = vbTab & " | ") As String()
NC& = UBound(Sq, 2)
Dim Row$()
Dim O$()
ReDim Row(NC - 1)
For R& = 1 To UBound(Sq, 1)
    For C& = 1 To NC
        Row(C - 1) = Sq(R, C)
    Next
    Push O, Join(Row, Sep)
Next
SqStrAy = O
End Property
Property Get CvSqRow(Sq, R&)
Dim O
O = Sq
Erase O
U& = UBound(Sq, 2)
ReDim O(U - 1)
For J& = 1 To U
    O(J - 1) = Sq(R, J)
Next
CvSqRow = O
End Property
Property Get AddSqAv(Sq, SqAv())
Dim O
O = Sq
For J% = 0 To UB(SqAv)
    O = AddTwoSq(O, SqAv(J))
Next
AddSqAv = O
End Property
Property Get SqSzStr$(Sq)
With SqSz(Sq)
    SqSzStr = FmtQQ("(?,?)", .NR, .NC)
End With
End Property

Property Get SqSz(Sq) As SqSz
AssertIsSq Sq

End Property

Property Get IsSq(Sq) As Boolean
On Error GoTo X
If UBound(Sq, 1) < 0 Then GoTo X
If UBound(Sq, 2) < 0 Then GoTo X
X:
End Property

Sub AssertIsSq(Sq, Optional Msg$ = "Given Sq is not a square, but a {TypeName}.  Sq is a 2-dimension-array with lower bound started from.")
If Not IsSq(Sq) Then
    Er Msg, TypeName(Sq)
End If
End Sub
