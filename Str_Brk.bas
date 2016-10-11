Attribute VB_Name = "Str_Brk"
Option Compare Database
Type S1S2
    S1 As String
    S2 As String
End Type
Type LMR
    L_Part As String
    M_Part As String
    R_Part As String
End Type
Property Get Brk2(S$, BrkStr$, Optional NoTrim As Boolean) As S1S2
P& = InStr(S, BrkStr)
Dim O As S1S2
If P = 0 Then
    O.S2 = S
Else
    O.S1 = Left(S, P - 1)
    O.S2 = Mid(S, P + Len(BrkStr))
End If
If Not NoTrim Then O = TrimS1S2(O)
Brk2 = O
End Property
Property Get TrimS1S2(P As S1S2) As S1S2
TrimS1S2.S1 = Trim(P.S1)
TrimS1S2.S2 = Trim(P.S2)
End Property
Property Get Brk1(S$, BrkStr$, Optional NoTrim As Boolean) As S1S2
P& = InStr(S, BrkStr)
Dim O As S1S2
If P = 0 Then O.S1 = S: Brk1 = O: Exit Property
O.S1 = Left(S, P - 1)
O.S2 = Mid(S, P + Len(BrkStr))
If Not NoTrim Then O = TrimS1S2(O)
Brk1 = O
End Property

Property Get Brk(S$, BrkStr$, Optional NoTrim As Boolean) As S1S2
P& = InStr(S, BrkStr)
If P = 0 Then Er "Brk: {S} does not contain {BrkStr}", S, BrkStr
Dim O As S1S2
O.S1 = Left(S, P - 1)
O.S2 = Mid(S, P + Len(BrkStr))
If Not NoTrim Then O = TrimS1S2(O)
Brk = O
End Property

Private Sub Brk__Tst()
Debug.Assert Brk("AAA?BBB", "?").S1 = "AAA"
Debug.Assert Brk("AAA?BBB", "?").S2 = "BBB"
Debug.Assert Brk("AAA?BBB", "??").S1 = "AAA"
Debug.Assert Brk("AAA?BBB", "??").S2 = "BBB"
End Sub

Private Sub BrkMacro__Tst()
Dim Act As LMR
Act = BrkMacro("aa{bbb} cccc ")
Debug.Assert Act.L_Part = "aa"
Debug.Assert Act.M_Part = "bbb"
Debug.Assert Act.R_Part = " cccc "
End Sub
Sub DmpLMR(P As LMR)
Debug.Print "Left  -Part =[" & P.L_Part & "]"
Debug.Print "Middle-Part =[" & P.M_Part & "]"
Debug.Print "Right -Part =[" & P.R_Part & "]"
End Sub

Property Get BrkMacro(MacroStr$) As LMR
S$ = MacroStr
P1& = InStr(S, "{"): If P1 = 0 Then GoTo Er
P2& = InStr(P1, S, "}"): If P2 = 0 Then GoTo Er
Dim O As LMR
O.L_Part = Left(S, P1 - 1)
O.M_Part = Mid(S, P1 + 1, P2 - P1 - 1)
O.R_Part = Mid(S, P2 + 1)
BrkMacro = O
Exit Property
Er: Er "{MacroStr} does not have {xxx}", S
End Property
