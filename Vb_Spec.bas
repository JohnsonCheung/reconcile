Attribute VB_Name = "Vb_Spec"
Option Compare Database
Property Get BrkSpec(SPEC$()) As Dictionary
For J& = 0 To UB(SPEC)
    With Brk(SPEC(J), ";")
        If .S1 = ";" Then
        End If
    End With
Next
End Property
Private Property Get ZSpecAy(SpecAy$(), Pfx$) As String()
Dim O$()
For J& = 0 To UB(SpecAy)
    If IsPfx(SpecAy(J), Pfx) Then
        Push O, ZRmvPfx(SpecAy(J), Pfx)
    End If
Next
ZSpecAy = O
End Property
Private Property Get ZRmvPfx$(L$, Pfx$)
ZRmvPfx = RmvPfx(Trim(RmvPfx(L, Pfx)), ";")
End Property

