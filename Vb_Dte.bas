Attribute VB_Name = "Vb_Dte"
Option Compare Database

Private Sub CvYYYYMMDD__Tst()
Debug.Assert CvYYYYMMDD(20161213) = DateSerial(2016, 12, 13)
End Sub

Property Get CvYYYYMMDD(YYYYMMDD) As Date
On Error GoTo X
A$ = YYYYMMDD
YYYY% = Left(A, 4)
MM% = Mid(A, 5, 2)
DD = Mid(A, 7, 2)
CvYYYYMMDD = DateSerial(YYYY, MM, DD)
X:
End Property
