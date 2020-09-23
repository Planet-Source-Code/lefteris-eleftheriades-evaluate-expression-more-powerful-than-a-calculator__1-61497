Attribute VB_Name = "DerivedMath"
Option Explicit
'Derived Math Functions
Public Function Sec(X As Double) As Double
  Sec = 1 / Cos(X)
End Function

Public Function Cosec(X As Double) As Double
  Cosec = 1 / Sin(X)
End Function

Public Function Cot(X As Double) As Double
  Cot = 1 / Tan(X)
End Function

Public Function Arcsin(X As Double) As Double
  Arcsin = Atn(X / Sqr(-X * X + 1))
End Function

Public Function Arccos(X As Double) As Double
  Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

Public Function Arcsec(X As Double) As Double
  Arcsec = Atn(X / Sqr(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1))
End Function

Public Function Arccosec(X As Double) As Double
  Arccosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
End Function

Public Function Arccot(X As Double) As Double
  Arccot = Atn(X) + 2 * Atn(1)
End Function

Public Function HSin(X As Double) As Double
  HSin = (Exp(X) - Exp(-X)) / 2
End Function

Public Function HCos(X As Double) As Double
  HCos = (Exp(X) + Exp(-X)) / 2
End Function

Public Function HTan(X As Double) As Double
  HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
End Function

Public Function HSec(X As Double) As Double
  HSec = 2 / (Exp(X) + Exp(-X))
End Function

Public Function HCosec(X As Double) As Double
  HCosec = 2 / (Exp(X) - Exp(-X))
End Function

Public Function HCotan(X As Double) As Double
  HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
End Function

Public Function HArcsin(X As Double) As Double
  HArcsin = Log(X + Sqr(X * X + 1))
End Function

Public Function HArccos(X As Double) As Double
  HArccos = Log(X + Sqr(X * X - 1))
End Function

Public Function HArctan(X As Double) As Double
  HArctan = Log((1 + X) / (1 - X)) / 2
End Function

Public Function HArcsec(X As Double) As Double
  HArcsec = Log((Sqr(-X * X + 1) + 1) / X)
End Function

Public Function HArccosec(X As Double) As Double
  HArccosec = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
End Function

Public Function HArccotan(X As Double) As Double
  HArccotan = Log((X + 1) / (X - 1)) / 2
End Function
