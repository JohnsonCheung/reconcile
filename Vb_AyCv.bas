Attribute VB_Name = "Vb_AyCv"
Option Compare Database

Property Get CvAyToStr(Ay) As String()
U& = UB(Ay)
Dim O$()
ReSzAy O, U
For J& = 0 To U
    O(J) = Nz(Ay(J), "")
Next
CvAyToStr = O
End Property
Property Get CvAyToInt(Ay) As Integer()
If IsIntAy(Ay) Then IntAy = Ay: Exit Property
Dim O&()
CvAyToInt = CvAy(Ay, O)
End Property

Property Get CvAyToLng(Ay) As Long()
If IsLngAy(Ay) Then LngAy = Ay: Exit Property
Dim O&()
CvAyToLng = CvAy(Ay, O)
End Property

Property Get CvAyToDte(Ay) As Date()
If IsDteAy(Ay) Then DteAy = Ay: Exit Property
Dim O() As Date
CvAyToDte = CvAy(Ay, O)
End Property

Property Get CvAy(Ay, OAy)
AssertIsAy Ay
AssertIsAy OAy
U& = UB(Ay)
Erase OAy
If U = -1 Then Exit Property
ReDim OAy(U)
For J& = 0 To U
    OAy(J) = Ay(J)
Next
CvAy = OAy
End Property
Property Get CvAyToBool(Ay) As Boolean()
Dim O() As Boolean
CvAyToBool = CvAy(Ay, O)
End Property
Sub ReSzAy(Ay, U)
If U = -1 Then Erase Ay: Exit Sub
ReDim Preserve Ay(U)
End Sub
