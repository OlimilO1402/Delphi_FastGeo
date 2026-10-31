Attribute VB_Name = "FGDSClosestPoint"
Option Explicit
'Double Precision!

Public Sub CPOnTriangleFromPoint(Graphic As TFastGEOGraphics)
Dim i As Long
Dim TmpPC As TColor, TmpPW As Long
Dim aT As TTriangle2D
Dim aP() As TPoint2D
Dim cP() As TPoint2D
Const MAX_POINTS As Long = 30
'Const AnimT As Long = 0 '10
Const D As Long = 6
  Call Graphic.Clear(clWhite)
  TmpPC = Graphic.Pen.Color
  TmpPW = Graphic.Pen.Width
  ReDim aP(0 To MAX_POINTS - 1)
  ReDim cP(0 To MAX_POINTS - 1)
  Call GenerateRandomTriangle(5, 5, Graphic.Width - 10, Graphic.Height - 10, aT)
  'aT.P(1).x = Graphic.Width / 2 * 1:       aT.P(1).y = Graphic.Height / 6 * 1
  'aT.P(2).x = Graphic.Width / D * (D - 1): aT.P(2).y = Graphic.Height / D * (D - 1)
  'aT.P(3).x = Graphic.Width / D * 1:       aT.P(3).y = Graphic.Height / D * (D - 1)
  
  Call GenerateRandomPoints2DXY(5, 5, Graphic.Width - 10, Graphic.Height - 10, aP)
  For i = 0 To UBound(aP)
    cP(i) = ClosestPointOnTriangleFromPoint2D(aT, aP(i))
  Next
  
  Graphic.Pen.Width = 3
  Call Graphic.DrawT(aT)
  Graphic.Pen.Width = 1
  For i = 0 To UBound(aP)
    Graphic.Pen.Color = clBlack
    Call Graphic.DrawCirclePR(aP(i), 4)
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clRed
    Call Graphic.DrawCirclePR(cP(i), 4)
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clLime
    Call Graphic.DrawLineP(aP(i), cP(i))
    'Call Sleep(AnimT)
  Next
  Graphic.Pen.Color = TmpPC
  Graphic.Pen.Width = TmpPW
End Sub
Public Sub CPOnQuadixFromPoint(Graphic As TFastGEOGraphics)
Dim i As Long
Dim TmpPC As TColor, TmpPW As Long
Dim aQ As TQuadix2D
Dim aP() As TPoint2D
Dim cP() As TPoint2D
Const MAX_POINTS As Long = 30
'Const AnimT As Long = 0 '10
Const D As Long = 6
  Call Graphic.Clear(clWhite)
  TmpPC = Graphic.Pen.Color
  TmpPW = Graphic.Pen.Width
  ReDim aP(0 To MAX_POINTS - 1)
  ReDim cP(0 To MAX_POINTS - 1)
  Call GenerateRandomQuadix(5, 5, Graphic.Width - 10, Graphic.Height - 10, aQ)
  
  Call GenerateRandomPoints2DXY(5, 5, Graphic.Width - 10, Graphic.Height - 10, aP)
  For i = 0 To UBound(aP)
    cP(i) = ClosestPointOnQuadixFromPoint2D(aQ, aP(i))
  Next
  
  Graphic.Pen.Width = 3
  Call Graphic.DrawQ(aQ)
  Graphic.Pen.Width = 1
  For i = 0 To UBound(aP)
    Graphic.Pen.Color = clBlack
    Call Graphic.DrawCirclePR(aP(i), 4)
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clRed
    Call Graphic.DrawCirclePR(cP(i), 4)
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clLime
    Call Graphic.DrawLineP(aP(i), cP(i))
    'Call Sleep(AnimT)
  Next
  Graphic.Pen.Color = TmpPC
  Graphic.Pen.Width = TmpPW
End Sub
Public Sub CPOnAABBFromPoint(Graphic As TFastGEOGraphics)
Dim i As Long
Dim TmpPC As TColor, TmpPW As Long
Dim aR As TRectangle
Dim aP() As TPoint2D
Dim cP() As TPoint2D
Const MAX_POINTS As Long = 30
'Const AnimT As Long = 0 '10
Const D As Long = 6
  Call Graphic.Clear(clWhite)
  TmpPC = Graphic.Pen.Color
  TmpPW = Graphic.Pen.Width
  ReDim aP(0 To MAX_POINTS - 1)
  ReDim cP(0 To MAX_POINTS - 1)
  Call GenerateRandomRectangle(5, 5, Graphic.Width - 10, Graphic.Height - 10, aR)
  
  Call GenerateRandomPoints2DXY(5, 5, Graphic.Width - 10, Graphic.Height - 10, aP)
  For i = 0 To UBound(aP)
    cP(i) = ClosestPointOnAABBFromPoint(aR, aP(i))
  Next
  
  Graphic.Pen.Width = 3
  Call Graphic.DrawR(aR)
  Graphic.Pen.Width = 1
  For i = 0 To UBound(aP)
    Graphic.Pen.Color = clBlack
    Call Graphic.DrawCirclePR(aP(i), 4)
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clRed
    Call Graphic.DrawCirclePR(cP(i), 4)
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clLime
    Call Graphic.DrawLineP(aP(i), cP(i))
    'Call Sleep(AnimT)
  Next
  Graphic.Pen.Color = TmpPC
  Graphic.Pen.Width = TmpPW
End Sub
Public Sub CPOnCircleFromPoint(Graphic As TFastGEOGraphics)
Dim i As Long
Dim TmpPC As TColor, TmpPW As Long
Dim aCi As TCircle
Dim aP() As TPoint2D
Dim cP() As TPoint2D
Const MAX_POINTS As Long = 30
'Const AnimT As Long = 0 '10
Const D As Long = 6
  Call Graphic.Clear(clWhite)
  TmpPC = Graphic.Pen.Color
  TmpPW = Graphic.Pen.Width
  ReDim aP(0 To MAX_POINTS - 1)
  ReDim cP(0 To MAX_POINTS - 1)
  Call GenerateRandomCircle(5, 5, Graphic.Width - 10, Graphic.Height - 10, aCi)
  
  Call GenerateRandomPoints2DXY(5, 5, Graphic.Width - 10, Graphic.Height - 10, aP)
  For i = 0 To UBound(aP)
    cP(i) = ClosestPointOnCircleFromPoint(aCi, aP(i))
  Next
  Graphic.Pen.Width = 3
  Call Graphic.DrawC(aCi)
  Graphic.Pen.Width = 1
  For i = 0 To UBound(aP)
    Graphic.Pen.Color = clBlack
    Call Graphic.DrawCirclePR(aP(i), 4)
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clRed
    Call Graphic.DrawCirclePR(cP(i), 4)
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clLime
    Call Graphic.DrawLineP(aP(i), cP(i))
    'Call Sleep(AnimT)
  Next
  Graphic.Pen.Color = TmpPC
  Graphic.Pen.Width = TmpPW
End Sub
Public Sub CPOnCircleFromSegment(Graphic As TFastGEOGraphics)
Dim i As Long
Dim TmpPC As TColor, TmpPW As Long
Dim aCi As TCircle         'der Kreis
Dim aSA As TSegment2DArray 'die Strecken
Dim cP As TPoint2DArray    'der Punkt auf dem Kreis
Dim cPs As TPoint2DArray   'der Punkt auf der Strecke
Const MAX_SEGMENTS As Long = 5
'Const AnimT As Long = 0 '10
Const D As Long = 6
  Call Graphic.Clear(clWhite)
  TmpPC = Graphic.Pen.Color
  TmpPW = Graphic.Pen.Width
  ReDim aSA.Arr(0 To MAX_SEGMENTS - 1)
  ReDim cP.Arr(0 To MAX_SEGMENTS - 1)
  ReDim cPs.Arr(0 To MAX_SEGMENTS - 1)
  
  Call GenerateRandomCircle(5, 5, Graphic.Width - 10, Graphic.Height - 10, aCi)
  
  'Call GenerateRandomsegment(5, 5, Graphic.Width - 10, Graphic.Height - 10, aP)
  
  Randomize
  
  For i = 0 To UBound(aSA.Arr) '- 1
    aSA.Arr(i).p(0).x = RandomL(Graphic.Width - 10) + 5
    aSA.Arr(i).p(0).y = RandomL(Graphic.Height - 10) + 5
    aSA.Arr(i).p(1).x = RandomL(Graphic.Width - 10) + 5
    aSA.Arr(i).p(1).y = RandomL(Graphic.Height - 10) + 5
  Next

  For i = 0 To UBound(aSA.Arr)
    cP.Arr(i) = ClosestPointOnCircleFromSegment(aCi, aSA.Arr(i))
  Next
  
  'so now, but which point is it then?
  For i = 0 To UBound(aSA.Arr)
    'it can be the first or the second point or any other between
    cPs.Arr(i) = ClosestPointOnLineFromPoint2DXY(aSA.Arr(i).p(0).x, aSA.Arr(i).p(0).y, aSA.Arr(i).p(1).x, aSA.Arr(i).p(1).y, cP.Arr(i).x, cP.Arr(i).y)
  Next
  Graphic.Pen.Width = 3
  Call Graphic.DrawC(aCi)
  Graphic.Pen.Width = 1
  For i = 0 To UBound(aSA.Arr)
    Graphic.Pen.Color = clBlack
    Call Graphic.DrawSegment(aSA.Arr(i))
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clRed
    Call Graphic.DrawCirclePR(cP.Arr(i), 4)
    Call Graphic.DrawCirclePR(cPs.Arr(i), 4)
    'Call Sleep(AnimT)
    Graphic.Pen.Color = clLime
    Call Graphic.DrawLineP(cP.Arr(i), cPs.Arr(i))
    'Call Sleep(AnimT)
  Next
  Graphic.Pen.Color = TmpPC
  Graphic.Pen.Width = TmpPW
End Sub
