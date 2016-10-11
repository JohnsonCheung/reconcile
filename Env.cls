VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Env"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Property Get MyDtaSrcPth$()
MyDtaSrcPth = HomPth & "Documents\My Data Sources\"
End Property
Property Get HomPth$()
HomPth = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\"
End Property
