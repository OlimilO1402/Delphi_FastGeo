Attribute VB_Name = "FGDSClipping"
Option Explicit
'Double Precision!
Public Sub SegmentTriangleClip(Graphic As TFastGEOGraphics)
Const MAX_SEGMENTS = 50
Dim i As Long, Triangle As TTriangle2D, ClippedSeg As TSegment2D, Segment() As TSegment2D

  Randomize

  Call Graphic.Clear(clWhite)

  Do
    Triangle.p(0).x = RandomL(Graphic.Width - 10) + 5#
    Triangle.p(0).y = RandomL(Graphic.Height - 10) + 5#

    Triangle.p(1).y = RandomL(Graphic.Height - 10) + 5#
    Triangle.p(1).x = RandomL(Graphic.Width - 10) + 5#

    Triangle.p(2).y = RandomL(Graphic.Height - 10) + 5#
    Triangle.p(2).x = RandomL(Graphic.Width - 10) + 5#
  Loop Until Not IsDegenerate2DT(Triangle)

  Graphic.Pen.Width = 3
  Call Graphic.DrawT(Triangle)
  Graphic.Pen.Width = 1

  ReDim Segment(0 To MAX_SEGMENTS - 1)

  For i = 0 To MAX_SEGMENTS - 1
    Do
      Segment(i).p(0).x = RandomL(Graphic.Width - 10) + 5#
      Segment(i).p(0).y = RandomL(Graphic.Height - 10) + 5#
      Segment(i).p(1).x = RandomL(Graphic.Width - 10) + 5#
      Segment(i).p(1).y = RandomL(Graphic.Height - 10) + 5#
    Loop Until Not IsDegenerate2DS(Segment(i))
    Call Graphic.DrawS(Segment(i))
    If Clip2DST(Segment(i), Triangle, ClippedSeg) Then
      Graphic.Pen.Color = clLime
      Call Graphic.DrawS(ClippedSeg)
      Graphic.Pen.Color = clRed
      Call Graphic.DrawCirclePR(ClippedSeg.p(0), 4)
      Call Graphic.DrawCirclePR(ClippedSeg.p(1), 4)
      Graphic.Pen.Color = clBlack
    End If
  Next
  Erase Segment
End Sub

Public Sub SegmentQuadixClip(Graphic As TFastGEOGraphics)
Const MAX_SEGMENTS = 50
Dim i As Long, Quadix As TQuadix2D, ClippedSeg As TSegment2D, Segment() As TSegment2D

  Randomize
  Call Graphic.Clear(clWhite)
  Do
    Quadix.p(0).x = RandomL(Graphic.Width - 10) + 5#
    Quadix.p(0).y = RandomL(Graphic.Height - 10) + 5#
    Quadix.p(1).y = RandomL(Graphic.Height - 10) + 5#
    Quadix.p(1).x = RandomL(Graphic.Width - 10) + 5#
    Quadix.p(2).y = RandomL(Graphic.Height - 10) + 5#
    Quadix.p(2).x = RandomL(Graphic.Width - 10) + 5#
    Quadix.p(3).y = RandomL(Graphic.Height - 10) + 5#
    Quadix.p(3).x = RandomL(Graphic.Width - 10) + 5#
  Loop Until (Not IsDegenerate2DQ(Quadix)) And ConvexQuadix(Quadix)

  Graphic.Pen.Width = 3
  Call Graphic.DrawQ(Quadix)
  Graphic.Pen.Width = 1

  ReDim Segment(0 To MAX_SEGMENTS - 1)

  For i = 0 To MAX_SEGMENTS - 1
    Do
      Segment(i).p(0).x = RandomL(Graphic.Width - 10) + 5#
      Segment(i).p(0).y = RandomL(Graphic.Height - 10) + 5#
      Segment(i).p(1).x = RandomL(Graphic.Width - 10) + 5#
      Segment(i).p(1).y = RandomL(Graphic.Height - 10) + 5#
    Loop Until Not IsDegenerate2DS(Segment(i))

    Call Graphic.DrawS(Segment(i))

    If Clip2DSQ(Segment(i), Quadix, ClippedSeg) Then
      Graphic.Pen.Color = clLime
      Call Graphic.DrawS(ClippedSeg)
      Graphic.Pen.Color = clRed
      Call Graphic.DrawCirclePR(ClippedSeg.p(0), 4)
      Call Graphic.DrawCirclePR(ClippedSeg.p(1), 4)
      Graphic.Pen.Color = clBlack
    End If
  Next
  Erase Segment
End Sub

Public Sub SegmentRectangleClip(Graphic As TFastGEOGraphics)
Const MAX_SEGMENTS = 60
Dim i As Long, Rectangle As TRectangle, ClippedSeg As TSegment2D, Segment() As TSegment2D, Temp As Double

  Randomize
  Call Graphic.Clear(clWhite)
  
  Do
    Rectangle.p(0).x = RandomL(Graphic.Width - 10) + 5#
    Rectangle.p(0).y = RandomL(Graphic.Height - 10) + 5#
    
    Rectangle.p(1).y = RandomL(Graphic.Height - 10) + 5#
    Rectangle.p(1).x = RandomL(Graphic.Width - 10) + 5#
    
    If Rectangle.p(0).x > Rectangle.p(1).x Then
      Temp = Rectangle.p(1).x
      Rectangle.p(1).x = Rectangle.p(0).x
      Rectangle.p(0).x = Temp
    End If

    If Rectangle.p(0).y > Rectangle.p(1).y Then
      Temp = Rectangle.p(1).y
      Rectangle.p(1).y = Rectangle.p(0).y
      Rectangle.p(0).y = Temp
    End If
  Loop Until Not IsDegenerate2DR(Rectangle)

  Graphic.Pen.Width = 3
  Call Graphic.DrawR(Rectangle)
  Graphic.Pen.Width = 1

  ReDim Segment(0 To MAX_SEGMENTS - 1)

  For i = 0 To MAX_SEGMENTS - 1
    Do
      Segment(i).p(0).x = RandomL(Graphic.Width - 10) + 5#
      Segment(i).p(0).y = RandomL(Graphic.Height - 10) + 5#
      Segment(i).p(1).x = RandomL(Graphic.Width - 10) + 5#
      Segment(i).p(1).y = RandomL(Graphic.Height - 10) + 5#
    Loop Until Not IsDegenerate2DS(Segment(i))

    Call Graphic.DrawS(Segment(i))

    If Clip2DSR(Segment(i), Rectangle, ClippedSeg) Then
      Graphic.Pen.Color = clLime
      Call Graphic.DrawSegment(ClippedSeg)
      Graphic.Pen.Color = clRed
      Call Graphic.DrawCirclePR(ClippedSeg.p(0), 4)
      Call Graphic.DrawCirclePR(ClippedSeg.p(1), 4)
      Graphic.Pen.Color = clBlack
    End If
  Next
  Erase Segment
End Sub


Public Sub SegmentCircleClip(Graphic As TFastGEOGraphics)
Const MAX_SEGMENTS = 60
Dim i As Long, aCircle As TCircle, ClippedSeg As TSegment2D, Segment() As TSegment2D

  Randomize
  Call Graphic.Clear(clWhite)

  aCircle.Radius = RandomL(150) + 60
  aCircle.x = RandomL(Graphic.Width - (CLng(aCircle.Radius) * 2)) + aCircle.Radius
  aCircle.y = RandomL(Graphic.Height - (CLng(aCircle.Radius) * 2)) + aCircle.Radius

  Graphic.Pen.Width = 3
  Call Graphic.DrawC(aCircle)
  Graphic.Pen.Width = 1

  ReDim Segment(0 To MAX_SEGMENTS - 1)

  For i = 0 To MAX_SEGMENTS - 1
    Do
      Segment(i).p(0).x = RandomL(Graphic.Width - 10) + 5#
      Segment(i).p(0).y = RandomL(Graphic.Height - 10) + 5#
      Segment(i).p(1).x = RandomL(Graphic.Width - 10) + 5#
      Segment(i).p(1).y = RandomL(Graphic.Height - 10) + 5#
    Loop Until Not IsDegenerate2DS(Segment(i))
    
    Call Graphic.DrawS(Segment(i))

    Call Clip2DSCi(Segment(i), aCircle, ClippedSeg)
    
    If Clip2DSCi(Segment(i), aCircle, ClippedSeg) Then
      Graphic.Pen.Color = clLime
      Call Graphic.DrawSegment(ClippedSeg)
      Graphic.Pen.Color = clRed
      Call Graphic.DrawCirclePR(ClippedSeg.p(0), 4)
      Call Graphic.DrawCirclePR(ClippedSeg.p(1), 4)
      Graphic.Pen.Color = clBlack
    End If
  Next
  Erase Segment
End Sub

