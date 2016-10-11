Attribute VB_Name = "Vb_AySubSet"
Option Compare Database

Property Get AySubSet_ByPfx(Ay, Pfx$)
Dim O$()
For J& = 0 To UB(Ay)
    If IsPfx(CStr(Ay(J)), Pfx) Then Push O, Ay(J)
Next
AySubSet_ByPfx = O
End Property
