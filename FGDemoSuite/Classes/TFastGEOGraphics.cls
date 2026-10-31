VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TFastGEOGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Double Precision!
Private Canvas As TCanvas
Private FWidth As Long
Private FHeight As Long

Private Sub Class_Initialize()
  '
End Sub

Private Sub Class_Terminate()
  Call Free
End Sub

Public Sub Create(Can As TCanvas, Optional cWidth As Long, Optional cHeight As Long)
  Set Canvas = Can
  Width = cWidth
  Height = cHeight
End Sub
'(* End Of Create *)
'
Public Property Let Width(cWidth As Long)
  FWidth = cWidth
End Property
Public Property Get Width() As Long
  Width = FWidth
End Property
'(* End Of Get Width *)
'
Public Property Let Height(cHeight As Long)
  FHeight = cHeight
End Property
Public Property Get Height() As Long
  Height = FHeight
End Property
'(* End Of Get Height *)
'
Public Sub Clear(Color As TColor)
Dim TmpBrushColor As TColor, TmpBrushStyle As TBrushStyle
  TmpBrushColor = Canvas.Brush.Color
  TmpBrushStyle = Canvas.Brush.Style
  Canvas.Brush.Style = bsSolid
  Canvas.Brush.Color = Color
  Call Canvas.FillRect(0&, 0&, Width, Height)
  Canvas.Brush.Color = TmpBrushColor
  Canvas.Brush.Style = TmpBrushStyle
End Sub
'(* End Of Clear *)
'
'##################################'  D R A W  '##################################'
Friend Sub DrawP(Pnt As TPoint2D)
  Call DrawPoint(Pnt.X, Pnt.Y)
End Sub
'(* End of DrawP *)
'
Friend Sub DrawL(ln As TLine2D)
  Call DrawLine(ln)
End Sub
'(* End of DrawL *)
'
Friend Sub DrawS(Seg As TSegment2D)
  Call DrawSegment(Seg)
End Sub
'(* End of DrawS *)
'
Friend Sub DrawT(Tri As TTriangle2D)
  Call DrawTriangle(Tri)
End Sub
'(* End of DrawT *)
'
Friend Sub DrawR(Rec As TRectangle)
  Call DrawRectangle(Rec)
End Sub
'(* End of DrawR *)
'
Friend Sub DrawQ(Quad As TQuadix2D)
  Call DrawQuadix(Quad)
End Sub
'(* End of DrawQ *)
'
Friend Sub DrawC(Cir As TCircle)
  Call DrawCircle(Cir)
End Sub
'(* End of DrawC *)
'
Friend Sub DrawPg(Poly As TPolygon2D)
  Call DrawPolygon(Poly)
End Sub
'(* End of DrawPg *)
'
'Public Sub Draw(Obj As TGeometricObject)
'  Select Case Obj.ObjectType
'  Case geoPoint2D:        Call DrawP(Obj.Point2D)
'  Case geoLine2D:         Call DrawL(Obj.Line2D)
'  Case geoSegment2D:      Call DrawS(Obj.Segment2D)
'  Case geoTriangle2D:     Call DrawT(Obj.Triangle2D)
'  Case geoQuadix2D:       Call DrawQ(Obj.Quadix2D)
'  Case geoRectangle:      Call DrawR(Obj.Rectangle)
'  Case geoCircle:         Call DrawC(Obj.Circle)
'  Case geoPolygon2D:      Call DrawPg(Obj.Polygon2D)
' End Select
'End Sub
'(* End of Draw *)
'
Public Sub DrawPixel(X As Double, Y As Double)
  Canvas.Pixels(CLng(X), CLng(Y)) = Canvas.Pen.Color
End Sub
Friend Sub DrawPixelP(Pnt As TPoint2D)
  Call DrawPixel(Pnt.X, Pnt.Y)
End Sub
'(* End Of DrawPixel *)
'
Public Sub DrawPoint(ByVal X As Double, ByVal Y As Double)
  Call Canvas.MoveTo(CLng(X), CLng(Y))
  Call Canvas.LineTo(CLng(X) + 1, CLng(Y) + 1)
  'wie wärs denn eigentlich mit einem Circle?
  'Call Canvas.Ellipse(CLng(x), CLng(y), CLng(x) + 1, CLng(y) + 1)
End Sub
Friend Sub DrawPointP(Pnt As TPoint2D)
  Call DrawPoint(Pnt.X, Pnt.Y)
End Sub
'(* End Of DrawPoint *)
'
Public Sub DrawLineXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
  Call Canvas.MoveTo(CLng(X1), CLng(Y1))
  Call Canvas.LineTo(CLng(X2), CLng(Y2))
End Sub
Friend Sub DrawLineP(Pnt1 As TPoint2D, Pnt2 As TPoint2D)
  Call DrawLineXY(Pnt1.X, Pnt1.Y, Pnt2.X, Pnt2.Y)
End Sub
Friend Sub DrawLine(ln As TLine2D)
  Call DrawLineP(ln.P(0), ln.P(1))
End Sub
'(* End Of DrawLine *)
'
Public Sub DrawSegmentXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
  Call DrawLineXY(X1, Y1, X2, Y2)
End Sub
Friend Sub DrawSegmentP(Pnt1 As TPoint2D, Pnt2 As TPoint2D)
  Call DrawLineP(Pnt1, Pnt2)
End Sub
Friend Sub DrawSegment(Seg As TSegment2D)
  Call DrawLineP(Seg.P(0), Seg.P(1))
End Sub
'(* End Of Draw Segment *)
'
Public Sub DrawTriangleXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double)
  Call DrawLineXY(X1, Y1, X2, Y2)
  Call DrawLineXY(X1, Y1, X3, Y3)
  Call DrawLineXY(X2, Y2, X3, Y3)
End Sub
Friend Sub DrawTriangleP(Pnt1 As TPoint2D, Pnt2 As TPoint2D, Pnt3 As TPoint2D)
  Call DrawLineP(Pnt1, Pnt2)
  Call DrawLineP(Pnt1, Pnt3)
  Call DrawLineP(Pnt2, Pnt3)
End Sub
Friend Sub DrawTriangleL(Ln1 As TLine2D, Ln2 As TLine2D, Ln3 As TLine2D)
  Call DrawLine(Ln1)
  Call DrawLine(Ln2)
  Call DrawLine(Ln3)
End Sub
Friend Sub DrawTriangle(Tri As TTriangle2D)
  Call DrawLineP(Tri.P(0), Tri.P(1))
  Call DrawLineP(Tri.P(1), Tri.P(2))
  Call DrawLineP(Tri.P(2), Tri.P(0))
End Sub
'(* End Of Draw Triangle *)
'
'
Public Sub DrawRectangleXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
  Call DrawLineXY(X1, Y1, X2, Y1)
  Call DrawLineXY(X2, Y1, X2, Y2)
  Call DrawLineXY(X2, Y2, X1, Y2)
  Call DrawLineXY(X1, Y2, X1, Y1)
  'Call Canvas.FrameRect(x1, y1, x2, y2)
End Sub
Friend Sub DrawRectangleP(Pnt1 As TPoint2D, Pnt2 As TPoint2D)
  Call DrawRectangleXY(Pnt1.X, Pnt1.Y, Pnt2.X, Pnt2.Y)
End Sub
Friend Sub DrawRectangle(Rec As TRectangle)
  Call DrawRectangleP(Rec.P(0), Rec.P(1))
End Sub
'(* End Of Draw Rectangle *)
'
'
Public Sub DrawQuadixXY(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, X3 As Double, Y3 As Double, X4 As Double, Y4 As Double)
  Call DrawLineXY(X1, Y1, X2, Y2)
  Call DrawLineXY(X2, Y2, X3, Y3)
  Call DrawLineXY(X3, Y3, X4, Y4)
  Call DrawLineXY(X4, Y4, X1, Y1)
End Sub
Friend Sub DrawQuadixP(Pnt1 As TPoint2D, Pnt2 As TPoint2D, Pnt3 As TPoint2D, Pnt4 As TPoint2D)
  Call DrawLineP(Pnt1, Pnt2)
  Call DrawLineP(Pnt2, Pnt3)
  Call DrawLineP(Pnt3, Pnt4)
  Call DrawLineP(Pnt4, Pnt1)
End Sub
Friend Sub DrawQuadix(Quad As TQuadix2D)
  Call DrawQuadixP(Quad.P(0), Quad.P(1), Quad.P(2), Quad.P(3))
End Sub
'(* End Of Draw Quadix *)

'
Public Sub DrawCircleXYR(X As Double, Y As Double, Radius As Double)
  Call Canvas.Ellipse(CLng(X - Radius), CLng(Y - Radius), CLng(X + Radius), CLng(Y + Radius))
End Sub
Friend Sub DrawCirclePR(Pnt As TPoint2D, Radius As Double)
  Call Canvas.Ellipse(CLng(Pnt.X - Radius), CLng(Pnt.Y - Radius), CLng(Pnt.X + Radius), CLng(Pnt.Y + Radius))
End Sub
Friend Sub DrawCircle(Cir As TCircle)
  Call DrawCircleXYR(Cir.X, Cir.Y, Cir.Radius)
End Sub
'(* End Of Draw Circle *)
'
Public Sub DrawEllipse(X As Double, Y As Double, a As Double, b As Double)
  Call Canvas.Ellipse(CLng(X - a), CLng(Y - b), CLng(X + a), CLng(Y + b))
End Sub
'(* End Of Draw Ellipse *)
'
Friend Sub DrawPolygonA(Points As TPoint2DArray)
Dim i As Long, j As Long, n As Long
  j = UBound(Points.Arr()): n = j
  For i = 0 To n
    Call DrawLineP(Points.Arr(j), Points.Arr(i))
    j = i
  Next
  'Call Canvas.PolylineP(VarPtr(Points.Arr(0)))
End Sub
Friend Sub DrawPolygon(Polygon As TPolygon2D)
Dim i As Long, j As Long, n As Long
  n = UBound(Polygon.Arr): j = n
  For i = 0 To n
    Call DrawLineP(Polygon.Arr(j), Polygon.Arr(i))
    j = i
  Next
End Sub
'(* End Of Draw Polygon *)
'
'
Friend Sub DrawCurve(Curve As TPoint2DArray)
Dim i  As Long, n As Long: n = UBound(Curve.Arr())
  For i = 0 To n - 2
    Call DrawLineP(Curve.Arr(i), Curve.Arr(i + 1))
  Next
End Sub
'(* End Of Draw Curve *)

Public Property Set Pen(cPen As TPen)
  Set Canvas.Pen = cPen
End Property
Public Property Get Pen() As TPen
  Set Pen = Canvas.Pen
End Property
'(* End Of Get Pen *)
'
Public Property Set Brush(cBrush As TBrush)
  Set Canvas.Brush = cBrush
End Property
Public Property Get Brush() As TBrush
  Set Brush = Canvas.Brush
End Property
'(* End Of SetGet Brush *)
'
Public Property Get CenterX() As Double
  CenterX = FWidth * 0.5
End Property
Public Property Get CenterY() As Double
  CenterY = FHeight * 0.5
End Property
'(* End Of Get Center Y *)

Public Sub Free()
  Set Canvas = Nothing
End Sub
