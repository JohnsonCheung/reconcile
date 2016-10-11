VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GodivaConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Ip$
Public B6F$
Public B4F$
Public B6UF$
Public B4UF$
Public Q_HKQYLIB$
Public Q_GODHKLIB$
Public B6FDsnCnnStr$
Private Sub Class_Initialize()
Ip = "172.16.10.100"
B6F = "HKSBPCF64"
Q_HKQYLIB = "HKQYLIB"
Q_GODHKLIB = "GODHKLIB"
WrtDsnFfn
B6FDsnCnnStr = DsnCnnStr(DsnFfn(B6F))
End Sub

Private Sub WrtDsnFfn()
WrtDsnFfn_One B6F

End Sub

Private Sub WrtDsnFfn_One(Lib$)

End Sub
