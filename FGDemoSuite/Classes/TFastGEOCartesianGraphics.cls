VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TFastGEOCartesianGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Double Precision!
Private FGraphicsContext  As TFastGEOGraphics

Private Sub Class_Initialize()
  '
End Sub

Private Sub Class_Terminate()
  Set FGraphicsContext = Nothing
End Sub

Public Sub Create(GraphicsContext As TFastGEOGraphics)
  Set FGraphicsContext = GraphicsContext
End Sub

Public Sub Clear(Color As TColor)
  FGraphicsContext.Clear Color
End Sub

Public Function ConvertXCoordinate(X As Double) As Double
  ConvertXCoordinate = FGraphicsContext.CenterX + X
End Function
'(* End Of Convert X Coordinate *)


Public Function ConvertYCoordinate(Y As Double) As Double
  ConvertYCoordinate = FGraphicsContext.CenterY - Y
End Function
'(* End Of Convert Y Coordinate *)


Friend Sub DrawCi(aCircle As TCircle)
  Call DrawCircle(aCircle)
End Sub

Friend Sub DrawQ(Quadix As TQuadix2D)
  Call DrawQuadix(Quadix)
End Sub

Friend Sub DrawPg(Polygon As TPolygon2D)
  Call DrawPolygon(Polygon)
End Sub

'Public Sub DrawO(Obj As TGeometricObject)
'  '
'End Sub

'Friend Sub DrawCu(Curve As TPoint2DArray)
'Dim i As Long
'  'For i = 1 To UBound(Curve.Arr())
'    '
'  'Next
'End Sub

Friend Sub DrawL(Line As TLine2D)
  Call DrawLine(Line)
End Sub

Friend Sub DrawP(Point As TPoint2D)
  Call DrawPoint(Point)
End Sub

Friend Sub DrawS(Segment As TSegment2D)
  Call DrawSegment(Segment)
End Sub

Friend Sub DrawR(Rectangle As TRectangle)
  Call DrawRectangle(Rectangle)
End Sub

Friend Sub DrawT(Triangle As TTriangle2D)
  Call DrawTriangle(Triangle)
End Sub


Friend Sub DrawCirclePR(Point As TPoint2D, Radius As Double)
  Call DrawCircleXYR(Point.X, Point.Y, Radius)
End Sub
Public Sub DrawCircleXYR(X As Double, Y As Double, Radius As Double)
  Call FGraphicsContext.DrawCircleXYR(ConvertXCoordinate(X), ConvertYCoordinate(Y), Radius)
End Sub
Friend Sub DrawCircle(aCircle As TCircle)
  Call DrawCircleXYR(aCircle.X, aCircle.Y, aCircle.Radius)
End Sub

'Friend Sub DrawCurve(Curve As TPoint2DArray)
'  '
'End Sub

Public Sub DrawEllipse(X As Double, Y As Double, a As Double, b As Double)
  Call FGraphicsContext.DrawEllipse(ConvertXCoordinate(X), ConvertYCoordinate(Y), a, b)
End Sub

Public Sub DrawLineXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
  Call FGraphicsContext.DrawLineXY(ConvertXCoordinate(X1), ConvertYCoordinate(Y1), ConvertXCoordinate(X2), ConvertYCoordinate(Y2))
End Sub
Friend Sub DrawLineP(Point1 As TPoint2D, Point2 As TPoint2D)
  Call DrawLineXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Sub
Friend Sub DrawLine(Line As TLine2D)
  Call DrawLineP(Line.P(0), Line.P(1))
End Sub

Public Sub DrawPixelXY(X As Double, Y As Double)
  Call FGraphicsContext.DrawPixel(ConvertXCoordinate(X), ConvertYCoordinate(Y))
End Sub
Friend Sub DrawPixel(Point As TPoint2D)
  Call DrawPixelXY(Point.X, Point.Y)
End Sub

Public Sub DrawPointXY(X As Double, Y As Double)
  Call FGraphicsContext.DrawPoint(ConvertXCoordinate(X), ConvertYCoordinate(Y))
End Sub
Friend Sub DrawPoint(Point As TPoint2D)
  Call DrawPointXY(Point.X, Point.Y)
End Sub

Friend Sub DrawPolygon(Polygon As TPolygon2D)
  Call FGraphicsContext.DrawPolygon(ConvertPolygon(Polygon))
End Sub
Friend Sub DrawPolygonA(Point As TPoint2DArray)
  '
End Sub

Public Sub DrawQuadixXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double)
  Call FGraphicsContext.DrawQuadixXY(ConvertXCoordinate(X1), ConvertYCoordinate(Y1), _
                                     ConvertXCoordinate(X2), ConvertYCoordinate(Y2), _
                                     ConvertXCoordinate(X3), ConvertYCoordinate(Y3), _
                                     ConvertXCoordinate(X4), ConvertYCoordinate(Y4))
End Sub
Friend Sub DrawQuadix(Quad As TQuadix2D)
  Call FGraphicsContext.DrawQuadixP(Quad.P(1), Quad.P(2), Quad.P(3), Quad.P(4))
End Sub
Friend Sub DrawQuadixP(Pnt1 As TPoint2D, Pnt2 As TPoint2D, Pnt3 As TPoint2D, Pnt4 As TPoint2D)
  Call FGraphicsContext.DrawQuadixXY(Pnt1.X, Pnt1.Y, _
                                     Pnt2.X, Pnt2.Y, _
                                     Pnt3.X, Pnt3.Y, _
                                     Pnt4.X, Pnt4.Y)
End Sub


Friend Sub DrawRectangleP(Point1 As TPoint2D, Point2 As TPoint2D)
  Call DrawRectangleXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Sub
Friend Sub DrawRectangle(Rectangle As TRectangle)
  Call DrawRectangleP(Rectangle.P(0), Rectangle.P(1))
End Sub
Friend Sub DrawRectangleXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
  Call FGraphicsContext.DrawRectangleXY(ConvertXCoordinate(X1), ConvertYCoordinate(Y1), _
                                        ConvertXCoordinate(X2), ConvertYCoordinate(Y2))
End Sub


Friend Sub DrawSegmentP(Point1 As TPoint2D, Point2 As TPoint2D)
  Call DrawSegmentXY(Point1.X, Point1.Y, Point2.X, Point2.Y)
End Sub
Friend Sub DrawSegmentXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
  Call FGraphicsContext.DrawSegmentXY(ConvertXCoordinate(X1), ConvertYCoordinate(Y1), _
                                      ConvertXCoordinate(X2), ConvertYCoordinate(Y2))
End Sub
Friend Sub DrawSegment(Seg As TSegment2D)
  Call DrawSegmentP(Seg.P(0), Seg.P(1))
End Sub


Friend Sub DrawTriangle(Triangle As TTriangle2D)
  Call DrawTriangleP(Triangle.P(0), Triangle.P(1), Triangle.P(2))
End Sub
Public Sub DrawTriangleXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double)
  Call FGraphicsContext.DrawTriangleXY(ConvertXCoordinate(X1), ConvertYCoordinate(Y1), _
                                       ConvertXCoordinate(X2), ConvertYCoordinate(Y2), _
                                       ConvertXCoordinate(X3), ConvertYCoordinate(Y3))
End Sub
Friend Sub DrawTriangleP(Point1 As TPoint2D, Point2 As TPoint2D, Point3 As TPoint2D)
  Call DrawTriangleXY(Point1.X, Point1.Y, _
                      Point2.X, Point2.Y, _
                      Point3.X, Point3.Y)
End Sub

Private Function ConvertPolygon(Polygon As TPolygon2D) As TPolygon2D
Dim i As Long, n As Long: n = UBound(Polygon.Arr())
  ReDim ConvertPolygon.Arr(n - 1)
  For i = 0 To UBound(Polygon.Arr) - 1
    ConvertPolygon.Arr(i).X = ConvertXCoordinate(Polygon.Arr(i).X)
    ConvertPolygon.Arr(i).Y = ConvertYCoordinate(Polygon.Arr(i).Y)
  Next
End Function
