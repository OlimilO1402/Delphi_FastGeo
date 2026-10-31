Attribute VB_Name = "FGDSRandomize"
Option Explicit
'Double Precision!
Private Const MAX_POINTS As Long = 1000

Public Sub RandomPointsAABB(Graphic As TFastGEOGraphics)
Dim i As Long, Rectangle As TRectangle, Point() As TPoint2D
Dim mess As String
  Randomize
  Call Graphic.Clear(clWhite)
'  Rectangle.P(1).x = RandomL(Graphic.Width - 305) + 5#
'  Rectangle.P(1).y = RandomL(Graphic.Height - 305) + 5#
'  Rectangle.P(2).x = Rectangle.P(1).x + RandomL(300) + 100#
'  Rectangle.P(2).y = Rectangle.P(1).y + RandomL(300) + 100#
  Call GenerateRandomRectangle(CSng(5), CSng(5), CSng(Graphic.Width - 10), CSng(Graphic.Height - 10), Rectangle)
  Graphic.Pen.Color = clRed 'Black
  Call Graphic.DrawR(Rectangle)
  ReDim Point(MAX_POINTS - 1)
  Call GenerateRandomPoints2DR(Rectangle, Point)
  Graphic.Pen.Color = clBlack
  
  For i = 0 To UBound(Point) '- 1
    If PointInRectangle(Point(i), Rectangle) Then
      Call Graphic.DrawP(Point(i))
    Else
      mess = vbNullString
      mess = mess & "Point("
      mess = mess & CStr(Point(i).X) & " | " & CStr(Point(i).Y)
      mess = mess & ") has been generated outside of Rectangle("
      mess = mess & CStr(Rectangle.P(0).X) & " | " & CStr(Rectangle.P(0).Y) & " | " & CStr(Rectangle.P(1).X) & " | " & CStr(Rectangle.P(1).Y)
      mess = mess & ")"
      MsgBox (mess)
    End If
  Next
  Erase Point
End Sub


Public Sub RandomPointsTriangle(Graphic As TFastGEOGraphics)
Dim i As Integer, Triangle As TTriangle2D, Point() As TPoint2D
Dim mess As String
  Randomize
  Call Graphic.Clear(clWhite)
  Call GenerateRandomTriangle(5#, 5#, Graphic.Width - 10#, Graphic.Height - 10#, Triangle)
  Graphic.Pen.Color = clRed 'Black
  Call Graphic.DrawT(Triangle)
  ReDim Point(MAX_POINTS - 1)
  Call GenerateRandomPoints2DT(Triangle, Point)
  Graphic.Pen.Color = clBlack
  
  For i = 0 To UBound(Point)
    If PointInTriangle(Point(i), Triangle) Then
      Call Graphic.DrawP(Point(i))
    Else
      mess = vbNullString
      mess = "Point("
      mess = mess & CStr(Point(i).X) & " | " & CStr(Point(i).Y)
      mess = mess & ") has been generated outside of Triangle("
      mess = mess & CStr(Triangle.P(1).X) & " | " & CStr(Triangle.P(1).Y) & " | " & ")"
      'Triangle.P(2).x,Triangle.P(2).y, Triangle.P(3).x,Triangle.P(3).y,
      MsgBox mess
    End If
  Next
  Erase Point
End Sub


Public Sub RandomPointsCircle(Graphic As TFastGEOGraphics)
Dim i As Long, aCircle As TCircle, Point() As TPoint2D
Dim mess As String
  Randomize
  Call Graphic.Clear(clWhite)
  Call GenerateRandomCircle(5#, 5#, Graphic.Width - 10#, Graphic.Height - 10#, aCircle)
  Graphic.Pen.Color = clRed 'Black
  Call Graphic.DrawC(aCircle)
  ReDim Point(MAX_POINTS - 1)
  Call GenerateRandomPoints2DCi(aCircle, Point)
  
  Graphic.Pen.Color = clBlack

  For i = 0 To UBound(Point) '- 1
    If PointInCircle(Point(i), aCircle) Then
      Call Graphic.DrawP(Point(i))
    Else
      mess = vbNullString
      mess = mess & "Point("
      mess = mess & " has been generated outside of Circle("
      mess = mess & CStr(aCircle.X) & CStr(aCircle.Y) & CStr(aCircle.Radius) & CStr(Point(i).X) & CStr(Point(i).Y)
      MsgBox mess
    End If
  Next
  Erase Point
End Sub


Public Sub RandomPointsQuadix(Graphic As TFastGEOGraphics)
Dim i As Long, Quadix As TQuadix2D, Point() As TPoint2D
Dim mess As String
  Randomize
  Call Graphic.Clear(clWhite)
  Call GenerateRandomQuadix(5#, 5#, Graphic.Width - 10, Graphic.Height - 10, Quadix)
  Graphic.Pen.Color = clRed 'Black
  Call Graphic.DrawQ(Quadix)
  ReDim Point(MAX_POINTS - 1)
  Call GenerateRandomPoints2DQ(Quadix, Point)
  Graphic.Pen.Color = clBlack
  
  For i = 0 To UBound(Point) - 1
    If PointInQuadix(Point(i), Quadix) Then
      Call Graphic.DrawP(Point(i))
    Else
      mess = vbNullString
      mess = mess & "Point(" & " has been generated outside of Quadix("
      'mess = mess & Quadix(%9.5f,%9.5f,%9.5f,%9.5f,%9.5f,%9.5f,%9.5f,%9.5f) Point(%9.5f,%9.5f) ',[Quadix.P(1).x,Quadix.P(1).y, Quadix.P(2).x,Quadix.P(2).y, Quadix.P(3).x,Quadix.P(3).y, Quadix.P(4).x,Quadix.P(4).y, Point[i].x ,Point[i].y]), mtError,[mbOk], 0)
      MsgBox mess
    End If
  Next
  Erase Point
End Sub

