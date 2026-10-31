Attribute VB_Name = "FGDSTriangles"
Option Explicit
'Double Precision!

Public Sub GenerateRandomTriangles(W As Long, H As Long, Triangle As TTriangle2DArray)
Dim i As Long
  For i = 0 To UBound(Triangle.Arr) '- 1
    If (i Mod 3) = 0 Then
    'jedes dritte soll ein gleichseitiges Dreieck sein
      'With Triangle.Arr(i)
      '  .P(1).x = RandomL(W - 10) + 5#
      '  .P(1).y = RandomL(H - 10) + 5#
      '  .P(2).x = Triangle.Arr(i).P(1).x + RandomL(200)
      '  .P(2).y = Triangle.Arr(i).P(1).y - RandomL(200)
      'End With
      Call GenerateRandomTriangle(5, 5, W - 10, H - 10, Triangle.Arr(i))
      Call CreateEquilateralTriangleS2DP(Triangle.Arr(i).p(0), Triangle.Arr(i).p(1), Triangle.Arr(i).p(2))
    Else
      Do
        'With Triangle.Arr(i)
        '  .P(1).x = RandomL(W - 10) + 5#
        '  .P(1).y = RandomL(H - 10) + 5#
        '  .P(2).x = Triangle.Arr(i).P(1).x + RandomL(300) - 50 '+ 50
        '  .P(2).y = Triangle.Arr(i).P(1).y + RandomL(200) - 50 '+ 50
        '  .P(3).x = Triangle.Arr(i).P(1).x - RandomL(200) + 50 '+ 50
        '  .P(3).y = Triangle.Arr(i).P(1).y - RandomL(300) + 50 '+ 50
        'End With
        Call GenerateRandomTriangle(5, 5, W - 10, H - 10, Triangle.Arr(i))
      Loop Until Not IsDegenerate2DT(Triangle.Arr(i))
    End If
  Next
End Sub


Public Sub TriangleCircumcircle(Graphic As TFastGEOGraphics)
Const MAX_TRIANGLES = 15 '30
Dim i As Long, Triangle As TTriangle2DArray
Dim TmpPC As TColor
  
  Randomize
  
  Graphic.Clear (clWhite)
  TmpPC = Graphic.Pen.Color
  ReDim Triangle.Arr(MAX_TRIANGLES - 1)
  Call GenerateRandomTriangles(Graphic.Width, Graphic.Height, Triangle)
  
  For i = 0 To MAX_TRIANGLES - 1
    Call Graphic.DrawT(Triangle.Arr(i))
    Graphic.Pen.Color = clLime
    Call Graphic.DrawC(CircumcircleT(Triangle.Arr(i)))
    Graphic.Pen.Color = clBlack
  Next
  Graphic.Pen.Color = TmpPC
  Erase Triangle.Arr
End Sub


Public Sub TriangleInscribedCircle(Graphic As TFastGEOGraphics)
Const MAX_TRIANGLES = 10 '30
Dim i As Long, Triangle As TTriangle2DArray
Dim TmpPC As TColor
  
  Randomize
  
  Call Graphic.Clear(clWhite)
  TmpPC = Graphic.Pen.Color
  ReDim Triangle.Arr(MAX_TRIANGLES - 1)
  
  Call GenerateRandomTriangles(Graphic.Width, Graphic.Height, Triangle)
  
  For i = 0 To MAX_TRIANGLES - 1
    Graphic.Pen.Color = clBlack
    Call Graphic.DrawT(Triangle.Arr(i))
    Graphic.Pen.Color = clLime
    Call Graphic.DrawC(InscribedCircleT(Triangle.Arr(i)))
  Next
  Graphic.Pen.Color = TmpPC
  Erase Triangle.Arr
End Sub


Public Sub TriangleTorricelliPoint(Graphic As TFastGEOGraphics)
Dim TmpPC As TColor
Const MAX_TRIANLGES = 15 '30
Dim i As Long, Triangle As TTriangle2DArray, TorPnt As TPoint2D
  
  Randomize
  TmpPC = Graphic.Pen.Color
  Call Graphic.Clear(clWhite)
  
  ReDim Triangle.Arr(MAX_TRIANLGES - 1)
  
  Call GenerateRandomTriangles(Graphic.Width, Graphic.Height, Triangle)
  
  For i = 0 To MAX_TRIANLGES - 1
    Call Graphic.DrawT(Triangle.Arr(i))
    Graphic.Pen.Color = clRed
    TorPnt = TorricelliPoint2DT(Triangle.Arr(i))
    Call Graphic.DrawCirclePR(TorPnt, 4)
    Graphic.Pen.Color = clLime
    Call Graphic.DrawSegmentP(TorPnt, Triangle.Arr(i).p(0))
    Call Graphic.DrawSegmentP(TorPnt, Triangle.Arr(i).p(1))
    Call Graphic.DrawSegmentP(TorPnt, Triangle.Arr(i).p(2))
    Graphic.Pen.Color = clBlack
  Next
  Graphic.Pen.Color = TmpPC
  Erase Triangle.Arr
End Sub

