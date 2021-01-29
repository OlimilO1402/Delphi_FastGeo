Attribute VB_Name = "ModConsts"
Option Explicit
'Double Precision!
'
'(**********[ Orientation constants ]**********)
'ist jetzt Enum in FGTypes.cls
'Public Const RightHandSide        As Long = -1
'Public Const LeftHandSide         As Long = 1
'Public Const Clockwise            As Long = -1
'Public Const CounterClockwise     As Long = 1
'Public Const CollinearOrientation As Long = 0
'Public Const AboveOriention       As Long = 1
'Public Const BelowOrientation     As Long = -1
'Public Const CoplanarOrientation  As Long = 0

'(************[ Epsilon constants ]*************)
Public Const Epsilon_High      As Double = 1E-16
Public Const Epsilon_Medium    As Double = 0.000000000001
Public Const Epsilon_Low       As Double = 0.00000001
Public Const Epsilon           As Double = Epsilon_Medium

'{$IFDEF FASTGEO_Double_PRECISION}
Public Const Infinity          As Double = 1E+30
'{$ENDIF}

'(*******[ Random resolution constants ]********)

Public Const RandomResolutionInt = 1000000000
Public Const RandomResolutionFlt = RandomResolutionInt * 1#

'const PI2       as Double=  6.283185307179586476925286766559000
'const PIDiv180  as Double=  0.017453292519943295769236907684886
'const w180DivPI as Double= 57.295779513082320876798154814105000

Public Const PI2       As Double = 6.28318530717959
Public Const PIDiv180  As Double = 1.74532925199433E-02
Public Const C180DivPI As Double = 57.2957795130823


Public Function UBoundDim(ByRef Arr() As Variant) As Long
tryE: On Error GoTo Catch
  UBoundDim = UBound(Arr)
  Exit Function
Catch:
  UBoundDim = 0
End Function

Public Function MinL(LngVal1 As Long, LngVal2 As Long) As Long
  If LngVal1 < LngVal2 Then MinL = LngVal1 Else MinL = LngVal2
End Function
Public Function MaxL(LngVal1 As Long, LngVal2 As Long) As Long
  If LngVal1 > LngVal2 Then MaxL = LngVal1 Else MaxL = LngVal2
End Function

Public Function MinD(DblVal1 As Double, DblVal2 As Double) As Double
  If DblVal1 < DblVal2 Then MinD = DblVal1 Else MinD = DblVal2
End Function
Public Function MaxD(DblVal1 As Double, DblVal2 As Double) As Double
  If DblVal1 > DblVal2 Then MaxD = DblVal1 Else MaxD = DblVal2
End Function

Public Function ArcCos(D As Double) As Double
  ArcCos = (3.14159265358979 / 2) - Atn(D / (Sqr(1 - D ^ 2)))
End Function
Public Function ArcSin(D As Double) As Double
  ArcSin = Atn(D / (Sqr(1 - D ^ 2)))
End Function
Public Function ArcTan(D As Double) As Double
  ArcTan = Atn(D)
End Function

Public Function Trunc(DblVal As Double) As Double
  Trunc = CDbl(CLng(DblVal))
End Function
