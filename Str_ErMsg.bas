Attribute VB_Name = "Str_ErMsg"
Option Compare Database
Property Get ErMsg$(ErMsgStr$, ParamArray Ap())
Dim Av()
Av = Ap
ErMsg = ErMsg_Av(ErMsgStr, Av)
End Property

Private Sub ErMsg__Tst()
Act$ = ErMsg("aaa {bb} ddd {cc} 11", 1, 2)
E$ = "aaa {bb} ddd {cc} 11" & vbCrLf & vbCrLf & "{bb} = [1]" & vbCrLf & "{cc} = [2]"
Debug.Assert Act = E
End Sub

Property Get ErMsg_Av$(ErMsgStr$, Av())
Dim O$()
Push O, ErMsgStr
Push O, ""
Push O, ZVarValStr(ErMsgStr, Av)
ErMsg_Av = JoinLine(O)
End Property

Private Property Get ZVarValStr$(ErMsgStr$, Av())
Dim O$()
Dim Ay$()
Ay = MacroAy(ErMsgStr)
For J = 0 To Min(UB(Ay), UB(Av))
    Push O, "{" & Ay(J) & "} = [" & Av(J) & "]"
Next
ZVarValStr = JoinLine(O)
End Property

