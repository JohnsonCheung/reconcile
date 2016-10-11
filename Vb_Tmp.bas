Attribute VB_Name = "Vb_Tmp"
Option Compare Database
Private X_Fso As Scripting.FileSystemObject
Public Property Get Fso() As Scripting.FileSystemObject
If IsNothing(X_Fso) Then Set X_Fso = New Scripting.FileSystemObject
Set Fso = X_Fso
End Property
Property Get TmpPth$()
TmpPth = Fso.GetSpecialFolder(TemporaryFolder) & "\"
End Property

Sub ClrTmpPth()
ClrPth TmpPth, IgnoreErr:=True
End Sub

Property Get TmpPthFnAy() As String()
TmpPthFnAy = PthFnAy(TmpPth)
End Property
Property Get TmpPthFfnAy(Optional SPEC$ = "*.*") As String()
TmpPthFfnAy = PthFfnAy(TmpPth, SPEC)
End Property

Property Get TmpFt$(Optional Pfx$ = "", Optional Ext$ = ".txt")
If Pfx <> "" Then A$ = Pfx & "-"
TmpFt = TmpPth & A & TimStmp & Ext
End Property
