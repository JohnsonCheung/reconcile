Attribute VB_Name = "Ide_MdBldTstAll"
Option Compare Database
Sub BldAllTstFct(Optional Md As CodeModule)
RmvMth Md, "All__Tst"
AppMdLines Md, AllTstFctLines(Md)
End Sub
Private Property Get AllTstFctLines(Md As CodeModule)
A = MdMthNmAy(Md, "*__Tst")
End Property
Private Sub AllTstFctLines__Tst()

End Sub
