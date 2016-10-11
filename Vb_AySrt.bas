Attribute VB_Name = "Vb_AySrt"
Option Compare Database
Private Sub SrtAyAsc__Tst()
Ay = Array(2, 4, 1, 3, 5, 0)
Act = SrtAyAsc(Ay)
Debug.Assert Sz(Act) = 6
For J% = 0 To 5
    Debug.Assert Act(J) = J
Next
End Sub

Private Sub SrtAyDes__Tst()
Ay = Array(2, 4, 1, 3, 5, 0)
Act = SrtAyDes(Ay)
Debug.Assert Sz(Act) = 6
For J% = 0 To 5
    Debug.Assert Act(J) = 5 - J
Next
End Sub

Property Get SrtAyAsc(Ay)
SrtAyAsc = SrtAy(Ay)
End Property

Property Get SrtAyDes(Ay)
SrtAyDes = SrtAy(Ay, True)
End Property

Property Get SrtAyToIdx(Ay, Optional IsDes As Boolean) As Long()
Dim O&()
Dim At&
For J& = 0 To UB(Ay)
    Itm = Ay(J)
    At = Idx_At(Ay, O, Itm, IsDes)
    InsAyAt O, At, J
Next
SrtAyToIdx = O
End Property

Private Property Get Idx_At&(Ay, IdxAy&(), Itm, IsDes As Boolean)
U& = UB(IdxAy)
If IsDes Then
    For J& = 0 To U
        If Itm > Ay(IdxAy(J)) Then Idx_At = J: Exit Property
    Next
    Idx_At = U + 1
Else
    For J& = 0 To U
        If Itm < Ay(IdxAy(J)) Then Idx_At = J: Exit Property
    Next
    Idx_At = U + 1
End If
End Property
Property Get SrtAyAscToIdx(Ay) As Long()
SrtAyAscToIdx = SrtAyToIdx(Ay)
End Property
Property Get SrtAyDesToIdx(Ay) As Long()
SrtAyDesToIdx = SrtAyToIdx(Ay, IsDes = True)
End Property

Property Get SrtAy(Ay, Optional IsDes As Boolean = False)
I = SrtAyToIdx(Ay, IsDes)
O = Ay
For J& = 0 To UB(I)
    O(J) = Ay(I(J))
Next
SrtAy = O
End Property
