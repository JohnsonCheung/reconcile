Attribute VB_Name = "Vb_Pipe"
Option Compare Database

Property Get Pipe(Inp, FnNmLvs$)
Dim A$()
T = Inp
A = SplitLvs(FnNmLvs)
For J% = 0 To UB(A)
    T = Run(A(J), T)
Next
Pipe = T
End Property
