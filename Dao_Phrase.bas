Attribute VB_Name = "Dao_Phrase"
Option Compare Database

Property Get Phrase_Where$(Where$)
If Where = "" Then Exit Property
Phrase_Where = " Where " & Where
End Property

Property Get Phrase_OrdBy$(OrdBy$)
If OrdBy = "" Then Exit Property
Phrase_OrdBy = " Order by " & OrdBy
End Property
