Attribute VB_Name = "Dao_IsXXX"
Option Compare Database

Property Get IsDbT(T As DbT) As Boolean
IsDbT = Not T.D.OpenRecordset(FmtQQ("Select Name from MSYSObjects where Name = '?' and Type=1", T.T)).EOF
End Property
Private Sub IsDbQ__Tst()
Debug.Assert IsDbQ(CurDbQ("QRY_GLAP")) = True
End Sub
Property Get IsDbQ(Q As DbQ) As Boolean
IsDbQ = Not Q.D.OpenRecordset(FmtQQ("Select Name from MSYSObjects where Name = '?' and Type=5", Q.Q)).EOF
End Property

