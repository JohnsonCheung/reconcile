Attribute VB_Name = "Vb_AyGp"
Option Compare Database
Type GpIdx
    BIdx As Long
    EIdx As Long
End Type
Type Gp
    BIdx() As Long
    EIdx() As Long
    Itm() As Variant
End Type
Private Sub Gp__Tst()
Dim A$()
A = Split("A A B B B C D D D D D E")
Dim Act As Gp
Act = Gp(A)
Stop
End Sub
Property Get NewGp(Optional U& = -1) As Gp
Dim O As Gp
ReSzAy O.BIdx, U
ReSzAy O.EIdx, U
ReSzAy O.Itm, U
NewGp = O
End Property
Property Get GpIdx(Gp As Gp, Idx&) As GpIdx
Dim O As GpIdx
O.BIdx = Gp.BIdx(Idx)
O.EIdx = Gp.EIdx(Idx)
GpIdx = O
End Property
Property Get Gp(Ay) As Gp
U& = UB(Ay)
If U = -1 Then Exit Property
Dim O As Gp, Last As Variant
O = NewGp
Last = Ay(0)
PushGp O, Last, 0, -1
For J& = 1 To U
    If Ay(J) <> Last Then
        Last = Ay(J)
        SetLastEle O.EIdx, J - 1
        PushGp O, Ay(J), J, -1
    End If
Next
SetLastEle O.EIdx, U
Gp = O
End Property

Private Sub GpStrAy__Tst()
Dim A$()
A = Split("A A B B B C D D D D D E")
BrwGp Gp(A)
End Sub
Sub BrwGp(Gp As Gp)
BrwAy GpStrAy(Gp)
End Sub
Property Get GpStrAy(Gp As Gp) As String()
U& = GpUB(Gp)
If U = -1 Then Exit Property
Dim O$()
ReDim O(U)
For J& = 0 To U
    O(J) = GpStr(Gp, J)
Next
GpStrAy = O
End Property
Property Get GpStr$(Gp As Gp, Idx&)
GpStr = Gp.Itm(Idx) & ":" & Gp.BIdx(Idx) & ":" & Gp.EIdx(Idx)
End Property
Sub PushGp(Gp As Gp, Itm, BIdx&, EIdx&)
Push Gp.Itm, Itm
Push Gp.BIdx, BIdx
Push Gp.EIdx, EIdx
End Sub

Property Get GpUB&(Gp As Gp)
GpUB = UB(Gp.Itm)
End Property

Property Get GpSz&(Gp As Gp)
GpSz = Sz(Gp.Itm)
End Property
