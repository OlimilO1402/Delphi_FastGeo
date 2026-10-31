Attribute VB_Name = "FGDSIntersections"
Option Explicit
'Double Precision!

Public Sub SegmentIntersect00(Graphic As TFastGEOGraphics)
Dim i As Long, j As Long, ix As Double, iy As Double
ReDim Segment(1 To 8) As TSegment2D
Dim Point As TPoint2D

  Call Graphic.Clear(clWhite)
  
  Segment(1) = EquateSegment2DXY(100, 100, 300, 300)
  Segment(2) = EquateSegment2DXY(300, 100, 100, 300)
  
  Segment(3) = EquateSegment2DXY(400, 300, 550, 100)
  Segment(4) = EquateSegment2DXY(700, 300, 550, 100)
  
  Segment(5) = EquateSegment2DXY(100, 350, 300, 350)
  Segment(6) = EquateSegment2DXY(300, 350, 500, 350)
  
  Segment(7) = EquateSegment2DXY(100, 50, 600, 50)
  Segment(8) = EquateSegment2DXY(350, 50, 350, 200)
  
  For i = 1 To 8
    Call Graphic.DrawSegment(Segment(i))
  Next
  For i = 1 To 8
    For j = i + 1 To 8
      If SimpleIntersectS(Segment(i), Segment(j)) Then
        Call Intersect2DSXY(Segment(i), Segment(j), ix, iy)
        Call Graphic.DrawCircleXYR(ix, iy, 4)
        Point = IntersectionPoint2DS(Segment(i), Segment(j))
        Call Graphic.DrawCirclePR(Point, 10)
      End If
    Next
  Next
End Sub

Public Sub SegmentIntersect01(Graphic As TFastGEOGraphics)
Const MAX_SEGMENTS = 100
Dim i As Long, j As Long, ix As Double, iy As Double
Dim Segment() As TSegment2D
  
  Call Graphic.Clear(clWhite)
  
  ReDim Segment(MAX_SEGMENTS - 1)
  
tryE:
  
  Randomize
  
  For i = 0 To UBound(Segment) '- 1
    Segment(i).p(0).x = RandomL(Graphic.Width - 10) + 5
    Segment(i).p(0).y = RandomL(Graphic.Height - 10) + 5
    Segment(i).p(1).x = RandomL(Graphic.Width - 10) + 5
    Segment(i).p(1).y = RandomL(Graphic.Height - 10) + 5
    Call Graphic.DrawSegment(Segment(i))
  Next
  
  Graphic.Pen.Color = clRed
  
  For i = 0 To UBound(Segment) '- 1
    For j = (i + 1) To UBound(Segment) '- 1
      If Intersect2DSXY(Segment(i), Segment(j), ix, iy) Then
        Call Graphic.DrawCircleXYR(ix, iy, 4)
      End If
    Next j
  Next i
  Graphic.Pen.Color = clBlack
finallyE:
  Erase Segment
End Sub


Public Sub CircleIntersect00(Graphic As TFastGEOGraphics)
Const MAX_CIRCLES = 100
Dim i As Long, j As Long, Circles() As TCircle
Dim IPnt1 As TPoint2D, IPnt2 As TPoint2D

 Call Graphic.Clear(clWhite)
 ReDim Circles(MAX_CIRCLES - 1)

tryE:

  Randomize

  For i = 1 To UBound(Circles) '- 1
    Circles(i).Radius = RandomL(50) + 20
    Circles(i).x = RandomL(Graphic.Width - (CLng(Circles(i).Radius) * 2)) + Circles(i).Radius
    Circles(i).y = RandomL(Graphic.Height - (CLng(Circles(i).Radius) * 2)) + Circles(i).Radius
    Call Graphic.DrawCircle(Circles(i))
  Next

  Graphic.Pen.Color = clRed

  For i = 0 To UBound(Circles) '- 1
    For j = (i + 1) To UBound(Circles) '- 1
      If Intersect2DCiCi(Circles(i), Circles(j)) And _
        (Distance2DXY(Circles(i).x, Circles(i).y, Circles(j).x, Circles(j).y) >= Abs(Circles(i).Radius - Circles(j).Radius)) Then
        Call IntersectionPoint2DCiP(Circles(i), Circles(j), IPnt1, IPnt2)
        Call Graphic.DrawCirclePR(IPnt1, 4)
        Call Graphic.DrawCirclePR(IPnt2, 4)
      End If
    Next j
  Next i

  Graphic.Pen.Color = clBlack

finallyE:
  Erase Circles
End Sub

'This shows a multi Segment multi Circle Intersect
'This is not activated because, there is another
'Segment Circle intersection testroutine: SegmentCircleClip in module FGDSClipping
Public Sub CircleIntersect01(Graphic As TFastGEOGraphics)
Const MAX_SEGMENTS = 50
Const MAX_CIRCLES = 20
Dim i As Long, j As Long, Segment() As TSegment2D, Circles() As TCircle
Dim IPnt1 As TPoint2D, IPnt2 As TPoint2D
Dim ICnt As Long

  Call Graphic.Clear(clWhite)
 
  ReDim Segment(MAX_SEGMENTS - 1)
  ReDim Circles(MAX_CIRCLES - 1)

tryE:

  Randomize

  For i = 1 To UBound(Segment) ' - 1
    Segment(i).p(1).x = RandomL(Graphic.Width)
    Segment(i).p(1).y = RandomL(Graphic.Height)
    Segment(i).p(2).x = RandomL(Graphic.Width)
    Segment(i).p(2).y = RandomL(Graphic.Height)
    Call Graphic.DrawSegment(Segment(i))
  Next
  
  For i = 1 To UBound(Circles) '- 1
    Circles(i).Radius = RandomL(50) + 20
    Circles(i).x = RandomL(Graphic.Width - CLng(Circles(i).Radius / 2)) + (Circles(i).Radius / 2)
    Circles(i).y = RandomL(Graphic.Height - CLng(Circles(i).Radius / 2)) + (Circles(i).Radius / 2)
    Call Graphic.DrawCircle(Circles(i))
  Next
  
  Graphic.Pen.Color = clRed

  For i = 0 To UBound(Segment) '- 1
    For j = 0 To UBound(Circles) ' - 1
      If Intersect2DSCi(Segment(i), Circles(j)) Then 'and _
        '(Distance(segment(i).x,segment(i).y,Circles(j).x,Circles(j).y) >= abs(Circle[i].Radius - Circle[j].Radius)) then
        Call IntersectionPoint2DSCi(Segment(i), Circles(j), ICnt, IPnt1, IPnt2)
        Call Graphic.DrawCirclePR(IPnt1, 4)
        Call Graphic.DrawCirclePR(IPnt2, 4)
      End If
    Next
  Next

  Graphic.Pen.Color = clBlack
  
finallyE:
  Erase Segment
  Erase Circles
End Sub

