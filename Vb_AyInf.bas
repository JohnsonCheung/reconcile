Attribute VB_Name = "Vb_AyInf"
Option Compare Database
Type AyPair
    Ay1 As Variant
    Ay2 As Variant
End Type
Type SqPair
    Sq1 As Variant
    Sq2 As Variant
End Type
Property Get IsIntAy(Ay) As Boolean
IsIntAy = TypeName(Ay) = "Integer()"
End Property
Sub PushTabAy(OAy, Ay)
For J& = 0 To UB(Ay)
    PushTab OAy, Ay(J)
Next
End Sub
Sub PushNTabAy(NTab%, OAy, Ay)
For J& = 0 To UB(Ay)
    PushNTab NTab, OAy, Ay(J)
Next
End Sub
Property Get IsLngAy(Ay) As Boolean
IsLngAy = TypeName(Ay) = "Long()"
End Property

Property Get MaxAyLen&(Ay)
Dim O&
For J& = 0 To UB(Ay)
    If Len(Ay(J)) > O Then O = Len(Ay(J))
Next
MaxAyLen = O
End Property

Private Sub IsIntAy__Tst()
Dim A%()
Dim B()
Dim C%
Debug.Assert IsIntAy(A) = True
Debug.Assert IsIntAy(B) = False
Debug.Assert IsIntAy(C) = False
End Sub
Sub PushTab(Ay, I)
PushNTab 1, Ay, I
End Sub
Sub PushNTab(NTab%, Ay, I)
Push Ay, VBA.String(NTab, vbTab) & I
End Sub

Sub Push(Ay, I)
N = Sz(Ay)
ReDim Preserve Ay(N)
If IsObject(I) Then
    Set Ay(N) = I
Else
    Ay(N) = I
End If
End Sub

Function IsDteAy(Ay) As Boolean
IsDteAy = TypeName(Ay) = "Date()"
End Function

Function IsBoolAy(Ay) As Boolean
IsBoolAy = TypeName(Ay) = "Boolean()"
End Function

Property Get RmvBlankEle(Ay)
Dim O
O = Ay
Erase O
For J& = 0 To UB(Ay)
    If Not IsBlank(Ay(J)) Then Push O, Ay(J)
Next
RmvBlankEle = O
End Property
Function IsAyEq(Ay1, Ay2) As Boolean
AssertAy Ay1, "Ay1", "IsAyEq"
AssertAy Ay2, "Ay2", "IsAyEq"
End Function

Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function

Function UB&(Ay)
UB = Sz(Ay) - 1
End Function
Property Get LastEle(Ay)
N& = Sz(Ay)
If N = 0 Then Er FmtQQ("Ay has no last element.  TypeName(Ay)=[?]", TypeName(Ay))
LastEle = Ay(N - 1)
End Property

Property Get AyIdx&(Ay, I)
For J = 0 To UB(Ay)
    If Ay(J) = I Then Ay_Idx = J: Exit Property
Next
AyIdx = -1
End Property
Property Get AyHas(Ay, I) As Boolean
AyHas = AyIdx(Ay, I) >= 0
End Property

Property Get MacroAy(MacroStr$, Optional Incl_Bracket As Boolean = False) As String()
Dim O$()
S$ = MacroStr
Dim A$
Do
    MacroAy__X A, S
    If A = "" Then MacroAy = O: Exit Property
    If Incl_Bracket Then A = "{" & A & "}"
    Push_NoDup O, A
Loop
End Property

Private Sub MacroAy__Tst()
Dim Act$()
Rest$ = "skldf{123}aaa{bb}ccc"
Act = MacroAy(Rest)
Debug.Assert Sz(Act) = 2
Debug.Assert Act(0) = "123"
Debug.Assert Act(1) = "bb"
End Sub

Private Sub MacroAy__X(OMacro$, ORest$)
X: OMacro = ""
P1& = InStr(ORest, "{")
If P1 = 0 Then Exit Sub
P2& = InStr(P1 + 1, ORest, "}")
If P2 = 0 Then Exit Sub
If P1 >= P2 Then Exit Sub
OMacro = Mid(ORest, P1 + 1, P2 - P1 - 1)
ORest = Mid(ORest, P2 + 1)
End Sub

Private Sub MacroAy__X___Tst()
Dim Macro$
Dim Rest$
Rest = "skldf{123}aaa{bb}ccc"
MacroAy__X Macro, Rest
Debug.Assert Macro = "123"
Debug.Assert Rest = "aaa{bb}ccc"
End Sub
