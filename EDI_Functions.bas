Attribute VB_Name = "EDI_Functions"
Option Compare Database

Property Get EDIPth$()
EDIPth = "C:\Users\cheungj\Desktop\reconciliation\EDI\"
End Property
Sub OpnEDIPth()
OpnPth EDIPth
End Sub
Property Get SampleEDIFv_HANMOV$()
SampleEDIFv_HANMOV = PthFfnAy(EDIPth, "HANMOV*.csv")(0)
End Property

Property Get SampleEDIFv_SPO$()
SampleEDIFv_SPO = PthFfnAy(EDIPth, "SPO*.csv")(0)
End Property

Property Get SampleEDIFv_DE1$()
SampleEDIFv_DE1 = PthFfnAy(EDIPth, "DE1*.csv")(0)
End Property

Property Get SampleEDIFv_DE2$()
SampleEDIFv_DE2 = PthFfnAy(EDIPth, "DE2*.csv")(0)
End Property

Property Get SampleEDIFv_IRP$()
SampleEDIFv_IRP = PthFfnAy(EDIPth, "IRP*.csv")(0)
End Property

Property Get SampleEDIFv_IMN$()
SampleEDIFv_IMN = PthFfnAy(EDIPth, "IMN*.csv")(0)
End Property

Property Get SampleEDIFv_LPD$()
SampleEDIFv_LPD = PthFfnAy(EDIPth, "LPD*.csv")(0)
End Property

Property Get SampleEDIFv_PMU$()
SampleEDIFv_PMU = PthFfnAy(EDIPth, "PMU*.csv")(0)
End Property

Property Get DistEDITy() As String()
A = PthFnAy(EDIPth, "*.csv")
Dim O$()
For Each I In A
    Push_NoDup O, Brk(CStr(I), "_").S1
Next
DistEDITy = O
End Property
