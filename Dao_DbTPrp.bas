Attribute VB_Name = "Dao_DbTPrp"
Option Compare Database
Private Property Get Tbl(T As DbT) As TableDef
Set Tbl = T.D.TableDefs(T.T)
End Property

Property Get DbTCnnStr$(T As DbT)
DbTCnnStr = Tbl(T).Connect
End Property
Property Get DbTAtr(T As DbT) As TableDefAttributeEnum
DbTAtr = Tbl(T).Attributes
End Property
