Attribute VB_Name = "Ide_SrcLin"
Option Compare Database
Const C_Private = "Private "
Const C_Friend = "Friend "
Const C_Public = "Public "
Const C_Sub = "Sub "
Const C_Fn = "Property Get "
Const C_Get = "Property Get "
Const C_Set = "Property Set "
Const C_Let = "Property Let "
Type MthAtr
    Nm As String
    Modifier As String ' Private / Public / Friend
    MthTy As String ' Sub / Property Get / Property Get / Property Let / Property Set
    MthNmSfx As String ' ! @ # $ % ^ &
    RetTy As String ' Type Name
    Bdy As String   ' Optional
End Type
Type MthAtrOpt
    MthAtr As MthAtr
    Some As Boolean
End Type

Private Property Get ZModifier$(OLin$)
If IsFriend(OLin) Then ZModifier = "Friend": Exit Property
If IsPrivate(OLin) Then ZModifier = "Private": Exit Property
If IsPublic(OLin) Then ZModifier = "Public": Exit Property
End Property
Private Sub SrcLinMthNm__Tst()
Debug.Assert SrcLinMthNm("Property Get SrcLinMthNm$(L$)") = "SrcLinMthNm"
Debug.Assert SrcLinMthNm("Public Property Get SrcLinMthNm$(L$)") = "SrcLinMthNm"
Debug.Assert SrcLinMthNm("Private Property Get SrcLinMthNm$(L$)") = "SrcLinMthNm"
Debug.Assert SrcLinMthNm("Friend Property Get SrcLinMthNm$(L$)") = "SrcLinMthNm"
Debug.Assert SrcLinMthNm("Sub SrcLinMthNm$(L$)") = "SrcLinMthNm"
Debug.Assert SrcLinMthNm("Public Sub SrcLinMthNm$(L$)") = "SrcLinMthNm"
Debug.Assert SrcLinMthNm("Private Sub SrcLinMthNm$(L$)") = "SrcLinMthNm"
Debug.Assert SrcLinMthNm("Friend Sub SrcLinMthNm$(L$)") = "SrcLinMthNm"
End Sub

Property Get SrcLinMthNm$(L$)
S$ = L
If Not IsMthLin(S) Then Exit Property
SrcLinMthNm = ZMthNm(S)
End Property
Private Property Get IsMthSfxChr(C$) As Boolean
IsMthSfxChr = True
If C = "$" Then Exit Property
If C = "@" Then Exit Property
If C = "#" Then Exit Property
If C = "!" Then Exit Property
If C = "%" Then Exit Property
If C = "^" Then Exit Property
IsMthSfxChr = False
End Property
Private Property Get IsPublic(OLin$) As Boolean
If IsPfx(OLin, C_Public) Then OLin = RmvPfx(OLin, C_Public): IsPublic = True
End Property
Private Property Get IsPrivate(OLin$) As Boolean
If IsPfx(OLin, C_Private) Then OLin = RmvPfx(OLin, C_Private): IsPrivate = True
End Property
Private Property Get IsFriend(OLin$) As Boolean
If IsPfx(OLin, C_Friend) Then OLin = RmvPfx(OLin, C_Friend): IsFriend = True
End Property
Private Property Get IsModifier(OLin$) As Boolean
IsModifier = True
If IsPublic(OLin) Then Exit Property
If IsPrivate(OLin) Then Exit Property
If IsFriend(OLin) Then Exit Property
IsModifier = False
End Property

Private Property Get IsFct(OLin$) As Boolean
If IsPfx(OLin, C_Fn) Then OLin = RmvPfx(OLin, C_Fn): IsFct = True
End Property
Private Property Get IsSub(OLin$) As Boolean
If IsPfx(OLin, C_Sub) Then OLin = RmvPfx(OLin, C_Sub): IsSub = True
End Property

Private Property Get IsMthLin(OLin$) As Boolean
A = IsModifier(OLin)
IsMthLin = True
If IsFct(OLin) Then Exit Property
If IsSub(OLin) Then Exit Property
If IsGet(OLin) Then Exit Property
If IsLet(OLin) Then Exit Property
If IsSet(OLin) Then Exit Property
IsMthLin = False
End Property

Private Property Get IsGet(OLin$) As Boolean
If IsPfx(OLin, C_Get) Then IsGet = True: OLin = RmvPfx(OLin, C_Get)
End Property
Private Property Get IsLet(OLin$) As Boolean
If IsPfx(OLin, C_Let) Then IsLet = True: OLin = RmvPfx(OLin, C_Let)
End Property
Private Property Get IsSet(OLin$) As Boolean
If IsPfx(OLin, C_Set) Then IsSet = True: OLin = RmvPfx(OLin, C_Set)
End Property

Private Property Get ZMthNm$(OLin$)
P% = InStr(OLin, "(")
If P = 0 Then Er "{OLin} does not have [(]", OLin
A$ = Left(OLin, P - 1)
C$ = Mid(OLin, P - 1, 1)
If IsMthSfxChr(C) Then
    ZMthNm = RmvLastChr(A)
    OLin = Mid(OLin, P - 1)
Else
    ZMthNm = A
    OLin = Mid(OLin, P)
End If
End Property
Private Property Get ZMthTy$(OLin$)
If IsFct(OLin) Then ZMthTy = "Property Get": Exit Property
If IsSub(OLin) Then ZMthTy = "Sub": Exit Property
If IsGet(OLin) Then ZMthTy = "Property Get": Exit Property
If IsSet(OLin) Then ZMthTy = "Property Set": Exit Property
If IsLet(OLin) Then ZMthTy = "Property Let": Exit Property
End Property

Private Property Get ZPrpTy$(OLin$)

End Property

Private Property Get ZMthNmSfx$(OLin$)
C$ = Left(OLin, 1)
If Not IsMthSfxChr(C) Then Exit Property
ZMthNmSfx = C
OLin = Mid(OLin, 2)
End Property

Property Get SomeMthAtr(Modifier$, MthTy$, Nm$, RetTy$, Optional Bdy$) As MthAtrOpt
Dim O As MthAtrOpt
O.Some = True
O.MthAtr = MthAtr(Modifier, MthTy, Nm, RetTy, Bdy)
SomeMthAtr = O
End Property
Property Get IsMthAtrOptEq(A As MthAtrOpt, B As MthAtrOpt) As Boolean
If A.Some <> B.Some Then Exit Property
If Not IsMthAtrEq(A.MthAtr, B.MthAtr) Then Exit Property
IsMthAtrOptEq = True
End Property
Property Get IsMthAtrEq(A As MthAtr, B As MthAtr, Optional InclBdy As Boolean) As Boolean
If A.Modifier <> B.Modifier Then Exit Property
If A.MthTy <> B.MthTy Then Exit Property
If A.Nm <> B.Nm Then Exit Property
If A.RetTy <> B.RetTy Then Exit Property
If InclBdy Then
    If A.Bdy <> B.Bdy Then Exit Property
End If
IsMthAtrEq = True
End Property
Property Get MthAtr(Modifier$, MthTy$, Nm$, MthNmSfx$, RetTy$, Optional Bdy$) As MthAtr
Dim O As MthAtr
O.Modifier = Modifier
O.MthTy = MthTy
O.Nm = Nm
O.MthNmSfx = MthNmSfx
O.RetTy = RetTy
O.Bdy = Bdy
MthAtr = O
End Property
Private Sub SrcLinMOA__Tst()
Dim A As MthAtrOpt, B As MthAtrOpt
A = SrcLinMAO("Property Get BrkSrcLin$(L$)"):         B = SomeMthAtr("", "Property Get", "BrkSrcLin", "$"):         Debug.Assert IsMthAtrOptEq(A, B)
A = SrcLinMAO("Public Property Get BrkSrcLin$(L$)"):  B = SomeMthAtr("Public", "Property Get", "BrkSrcLin", "$"):   Debug.Assert IsMthAtrOptEq(A, B)
A = SrcLinMAO("Private Property Get BrkSrcLin$(L$)"): B = SomeMthAtr("Private", "Property Get", "BrkSrcLin", "$"):  Debug.Assert IsMthAtrOptEq(A, B)
A = SrcLinMAO("Friend Property Get BrkSrcLin$(L$)"):  B = SomeMthAtr("Friend", "Property Get", "BrkSrcLin", "$"):   Debug.Assert IsMthAtrOptEq(A, B)
A = SrcLinMAO("Sub BrkSrcLin$(L$)"):              B = SomeMthAtr("", "Sub", "BrkSrcLin", "$"):              Debug.Assert IsMthAtrOptEq(A, B)
A = SrcLinMAO("Public Sub BrkSrcLin$(L$)"):       B = SomeMthAtr("Public", "Sub", "BrkSrcLin", "$"):        Debug.Assert IsMthAtrOptEq(A, B)
A = SrcLinMAO("Private Sub BrkSrcLin$(L$)"):      B = SomeMthAtr("Private", "Sub", "BrkSrcLin", "$"):       Debug.Assert IsMthAtrOptEq(A, B)
A = SrcLinMAO("Friend Sub BrkSrcLin$(L$)"):       B = SomeMthAtr("Friend", "Sub", "BrkSrcLin", "$"):        Debug.Assert IsMthAtrOptEq(A, B)
End Sub

Property Get SrcLinMAO(ByVal L$) As MthAtrOpt
Modifier$ = ZModifier(L)
MthTy$ = ZMthTy(L)
If MthTy = "" Then Exit Property
Nm$ = ZMthNm(L)
MthNmSfx = ZMthNmSfx(L)
Dim M As MthAtr
M.Modifier = Modifier
M.MthTy = MthTy
M.Nm = Nm
M.MthNmSfx = MthNmSfx

Dim O As MthAtrOpt
O.Some = True
O.MthAtr = M
SrcLinMAO = O
End Property
Property Get MthAtrStr$(P As MthAtr)

End Property

