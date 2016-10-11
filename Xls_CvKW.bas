Attribute VB_Name = "Xls_CvKW"
Option Compare Database

Property Get CvKW_TotCal(KW_TotCal$) As XlTotalsCalculation
Dim O As XlTotalsCalculation
Select Case KW_TotCal
Case "*Tot": O = xlTotalsCalculationSum
Case "*Avg": O = xlTotalsCalculationAverage
Case "*Cnt": O = xlTotalsCalculationCount
Case Else:  O = xlTotalsCalculationSum
End Select
CvKW_TotCal = O
End Property
Property Get CvKW_HAlign(KW_HAlign$) As XlHAlign
Dim O As XlTotalsCalculation
Select Case KW_TotCal
Case "*Center": O = XlHAlign.xlHAlignCenter
Case "*Left": O = XlHAlign.xlHAlignLeft
Case "*Right": O = XlHAlign.xlHAlignRight
Case Else:  O = XlHAlign.xlHAlignCenter
End Select
CvKW_HAlign = O
End Property
Property Get CvKW_Colr&(KW_Colr$)
Dim O&
Select Case KW_Colr
Case "*Green": O = 9359529
Case "*Yellow": O = 65535
Case "*Red": O = 255
Case "*Blue": O = 15652797
Case Else: O = Val(KW_Colr)
End Select
CvKW_Colr = O
End Property
